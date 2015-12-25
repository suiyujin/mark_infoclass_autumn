## 使用方法
##### Ruby, bundlerをインストール:
Ruby 2.1以降を推奨 (開発時はRuby 2.2.3)

##### 依存ライブラリをインストール:
Excelファイルをrubyから扱うためのライブラリをインストール
```shell
bundle install --path vendor/bundle
```

##### 設定ファイルを作成:
サンプルファイルをコピーして編集
```shell
cp config/config.yml.sample config/config.yml
```
作成後、`config/`ディレクトリ内の`config.yml`を編集
- 例年通りであれば、設定する項目は下記の3つのみ
  - reports_dir
  - file_prefix
  - evaluations

##### テンプレートファイルを作成:
レポートディレクトリへコピーして編集
```shell
cp template/evaluation_sample.xlsx template/list_sample.xlsx <your reports directory>
```
作成後、レポートディレクトリ内の上記2ファイルを編集
- 評価結果は、テンプレートファイルを使って作成されます
- テンプレートファイルの形式に合わせて設定ファイルを修正してください

##### 実行開始:
レポートを読み込み、クラス全体の評価ファイルと学生個人への評価ファイルを書き出す
```shell
bundle exec ruby src/main.rb
```
- 実行後、下記のファイルがレポートディレクトリ内に保存されます
  - evaluation_**************(作成年月日時分秒).xlsx 【クラス全体の評価ファイル】
  - lists/*********(学籍番号)####(氏名).xlsx【学生個人への評価ファイル】
