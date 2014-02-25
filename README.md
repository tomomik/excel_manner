# ExcelManner

Excleファイルを納品用や提出用に体裁を整えます。  
こうしておかないと怒る人向けです。  
(Win32OLEを利用してますので、windowsでしか利用出来ません。)

体裁を整えること
- 最初のシートを選択します
- 全てのシートはセル"A1"を選択状態にします
- プロパティを空にします(title、表題、作成者、カテゴリ、キーワード、コメント）
- 表示率を設定します
- 表示を標準にします



## Installation

Add this line to your application's Gemfile:

    gem 'excel_manner'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install excel_manner

## Usage

### 設定ファイルの作成

YAML形式で以下のように、体裁を整えるExcelファイルのパスと表示率を書きます。

    read_path: ../data/
    zoom_ratio: 75

read_pathは相対パスでも絶対パスでもいいです。  
設定ファイルは spec/data/config001.yml を参考にしてください。

### 実行方法

以下のようにして実行します。
startメソッドの引数に設定ファイルを指定します。

    require 'excel_manner'
    ExcelManner.start('../data/config001.yml')



## Contributing

1. Fork it ( http://github.com/<my-github-username>/excel_manner/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request

