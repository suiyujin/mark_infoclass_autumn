## Usage
Install Ruby and bundler (e.g. gem install bundler)
then, install dependencies libraries:
```shell
bundle install --path vendor/bundle
`````

Make config file:
```shell
cp config/config.yml.sample config/config.yml
```
then, edit `config.yml` in `config/` directory

Start execution:
```shell
bundle exec ruby src/main.rb
`````
