## 使用方法
Ruby, bundlerをインストール

依存ライブラリをインストール:
```shell
bundle install --path vendor/bundle
```

設定ファイルを作成:
```shell
cp config/config.yml.sample config/config.yml
```
その後、`config/`ディレクトリ内の`config.yml`を編集

処理を開始:
```shell
bundle exec ruby src/main.rb
```
