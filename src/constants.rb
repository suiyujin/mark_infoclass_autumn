require 'yaml'

module Constants
  ### configファイルを読み込んでチェックする
  def self.read_and_check_config_file(file_dir)
    config = YAML.load_file("#{file_dir.sub(/src/, 'config')}/config.yml")

    config.each do |key, value|
      raise StandardError, "#{key} is invalid." if value == ''
      unless value.is_a?(Enumerable) || value.instance_of?(TrueClass) || value.instance_of?(FalseClass)
        raise StandardError, "#{key} is invalid." unless value.instance_of?(String)
      end
    end

    config['number'].each do |key, value|
      raise StandardError, "#{key} is invalid." unless value.instance_of?(Fixnum)
    end

    raise StandardError, 'reports_dir is not absolute path.' unless config['reports_dir'] =~ /\A\/\w+/

    config
  end

  ### configから定数を作成する
  def self.const_set_from_config(config)
    config.each do |key, value|
      if value.instance_of?(Hash)
        const_set_from_config(value)
        next
      end
      const_set(key.upcase, value)
      p "set const: #{key.upcase} = #{value}" if $DEBUG
    end
  end

  file_dir = File.expand_path(File.dirname(__FILE__))
  config = read_and_check_config_file(file_dir)
  const_set_from_config(config)
  puts "set all consts."
end
