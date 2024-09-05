//go:build windows

package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/fs"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/devlights/goxcel"
)

type (
	Args struct {
		dir     string
		text    string
		json    bool
		onlyHit bool
		verbose bool
		debug   bool
	}

	SimpleResult struct {
		Path  string `json:"path"`
		Sheet string `json:"sheet"`
		Text  string `json:"text"`
	}

	DetailResult struct {
		Path  string `json:"path"`
		Sheet string `json:"sheet"`
		Row   int32  `json:"row"`
		Col   int32  `json:"col"`
		Text  string `json:"text"`
	}
)

var (
	args Args
)

var (
	appLog = log.New(os.Stdout, "", 0)
)

func init() {
	flag.StringVar(&args.dir, "dir", ".", "directory path")
	flag.StringVar(&args.text, "text", "", "search text")
	flag.BoolVar(&args.json, "json", false, "output as JSON")
	flag.BoolVar(&args.onlyHit, "only-hit", true, "show ONLY HIT")
	flag.BoolVar(&args.verbose, "verbose", false, "verbose mode")
	flag.BoolVar(&args.debug, "debug", false, "debug mode")
}

func newSimpleResult(path string, name string, text string) *SimpleResult {
	return &SimpleResult{path, name, text}
}

func newDetailResult(path string, name string, foundRange *goxcel.XlRange) *DetailResult {
	col, _ := foundRange.Column()
	row, _ := foundRange.Row()
	value, _ := foundRange.Value()

	return &DetailResult{path, name, row, col, value.(string)}
}

func abs(p string) string {
	abs, err := filepath.Abs(p)
	if err != nil {
		log.Panic(err)
	}

	return abs
}

func genErr(procName string, err error) error {
	return fmt.Errorf("%s failed: %w", procName, err)
}

func main() {
	log.SetFlags(0)
	flag.Parse()

	if args.text == "" {
		flag.PrintDefaults()
		os.Exit(1)
	}

	if args.dir == "" {
		args.dir = "."
	}

	if err := run(); err != nil {
		log.Fatal(err)
	}
}

func run() error {
	quit := goxcel.MustInitGoxcel()
	defer quit()

	excel, excelRelease := goxcel.MustNewGoxcel()
	defer excelRelease()

	excel.MustSilent(false)

	wbs := excel.MustWorkbooks()

	rootDir := abs(args.dir)
	err := filepath.WalkDir(rootDir, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}

		if d.IsDir() {
			return nil
		}

		if strings.Contains(filepath.Base(path), "~$") {
			return nil
		}

		if !strings.HasSuffix(path, ".xlsx") {
			return nil
		}

		absPath := abs(path)
		wb, wbRelease, err := wbs.Open(absPath)
		if err != nil {
			return genErr("wbs.Open(absPath)", err)
		}
		defer wbRelease()

		if args.debug {
			appLog.Printf("Document Open: %s", absPath)
		}

		sheets, err := wb.WorkSheets()
		if err != nil {
			return genErr("wb.WorkSheets()", err)
		}

		relPath, _ := filepath.Rel(rootDir, absPath)
		_, err = sheets.Walk(func(ws *goxcel.Worksheet, index int) error {
			rng, err := ws.UsedRange()
			if err != nil {
				return genErr("ws.UsedRange()", err)
			}

			startCell, err := rng.Cells(1, 1)
			if err != nil {
				return genErr("rng.Cells(1, 1)", err)
			}

			foundRange, found, err := rng.Find(args.text, startCell)
			if err != nil {
				return genErr("rng.Find(args.text, startCell)", err)
			}

			name, _ := ws.Name()
			if !found {
				if !args.onlyHit {
					err = outSimple(newSimpleResult(relPath, name, "NO HIT"))
					if err != nil {
						return err
					}
				}

				return nil
			}

			if !args.verbose {
				err = outSimple(newSimpleResult(relPath, name, "HIT"))
				if err != nil {
					return err
				}

				return nil
			}

			err = outDetail(newDetailResult(relPath, name, foundRange))
			if err != nil {
				return err
			}

			startCell, _ = foundRange.Cells(1, 1)
			for i := 0; found; i++ {
				after, _ := foundRange.Cells(1, 1)

				foundRange, found, err = rng.FindNext(after)
				if err != nil {
					return genErr("rng.FindNext(after)", err)
				}

				if !found {
					break
				}

				if args.debug {
					printCellPos(startCell, after)
				}

				if i > 0 && sameCell(startCell, after) {
					break
				}

				if args.verbose {
					err = outDetail(newDetailResult(relPath, name, foundRange))
					if err != nil {
						return err
					}
				}
			}

			return nil
		})

		if err != nil {
			return err
		}

		return nil
	})

	if err != nil {
		return err
	}

	return nil
}

func sameCell(c1, c2 *goxcel.Cell) bool {
	col1, _ := c1.Column()
	row1, _ := c1.Row()
	col2, _ := c2.Column()
	row2, _ := c2.Row()

	return col1 == col2 && row1 == row2
}

func printCellPos(startCell, after *goxcel.Cell) {
	col1, _ := startCell.Column()
	row1, _ := startCell.Row()
	col2, _ := after.Column()
	row2, _ := after.Row()

	appLog.Printf("FindNext: startCell=(%d,%d)\tafter=(%d,%d)", row1, col1, row2, col2)
}

func outSimple(result *SimpleResult) error {
	var (
		message = fmt.Sprintf("%s %q: %s", result.Path, result.Sheet, result.Text)
		err     error
	)

	if args.json {
		message, err = toJson(result)
		if err != nil {
			return genErr("toJson(result)", err)
		}
	}

	appLog.Println(message)

	return nil
}

func outDetail(result *DetailResult) error {
	var (
		message = fmt.Sprintf("%s %q (%d,%d): %q", result.Path, result.Sheet, result.Row, result.Col, result.Text)
		err     error
	)

	if args.json {
		message, err = toJson(result)
		if err != nil {
			return genErr("toJson(result)", err)
		}
	}

	appLog.Println(message)

	return nil
}

func toJson(v any) (string, error) {
	b, err := json.Marshal(v)
	if err != nil {
		return "", genErr("json.Marshal(result)", err)
	}

	return string(b), nil
}
