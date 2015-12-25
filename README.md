## 使用方法
##### Ruby, bundlerをインストール:
Ruby 2.1以降を推奨 (開発時はRuby 2.2.3)

##### 依存ライブラリをインストール:
```shell
bundle install --path vendor/bundle
```

##### 設定ファイルを作成:
```shell
cp config/config.yml.sample config/config.yml
```
作成後、`config/`ディレクトリ内の`config.yml`を編集
- 例年通りであれば、設定する項目は下記の3つのみ
  - reports_dir
  - file_prefix
  - evaluations

##### テンプレートファイルを作成:
```shell
cp template/evaluation_sample.xlsx template/list_sample.xlsx <your reports directory>
```
作成後、レポートディレクトリ内の上記2ファイルを編集
- 評価結果は、テンプレートファイルを使って作成されます
- テンプレートファイルの形式に合わせて設定ファイルを修正してください

##### 処理を開始:
```shell
bundle exec ruby src/main.rb
```
