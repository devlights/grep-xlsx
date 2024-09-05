# これは何？

指定ディレクトリの下に存在する MS EXCEL ファイル (*.xlsx) をGREPコマンドのように検索します。

対象のEXCELファイルが存在していて、どこに書いてあるのかが分からない時などに良かったらご利用ください。

Windowsのみで動作します。

## インストール

```sh
go install github.com/devlights/grep-xlsx/cmd/grep-xlsx@latest
```

## 使い方

```sh
$ ./grep-xlsx.exe -h
Usage of ./grep-xlsx.exe:
  -debug
        debug mode
  -dir string
        directory path (default ".")
  -json
        output as JSON
  -only-hit
        show ONLY HIT (default true)
  -text string
        search text
  -verbose
        verbose mode
```

ヒットした文書のパスが知りたい場合は以下のようにします。

```sh
$ grep-xlsx.exe -dir ディレクトリパス -text "検索文字列(ワイルドカード利用可)"
```

ヒットした箇所も見たい場合は ```-verbose``` オプションを付与するとみることが出来ます。

```sh
$ ./grep-xlsx.exe -dir ~/path/to/documents -text "データベース*サイズ" -verbose  
```

結果をjsonで出力したい場合は ```-json``` オプションを付与します。

```sh
$ ./grep-xlsx.exe -dir ~/path/to/documents -text "データベース*サイズ" -verbose -json
```

## ビルド方法

[Task](https://taskfile.dev/#/) を使っています。詳細は [Taskfile.yml](./Taskfile.yml) を参照ください。

```sh
$ task build
```
