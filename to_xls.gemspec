# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)
require "to_xls/version"

Gem::Specification.new do |s|
  s.name        = "to_xls"
  s.version     = ToXls::VERSION
  s.authors     = ["Enrique Garcia Cota", "Francisco de Juan", "Sergio Díaz Álvarez"]
  s.email       = %q{egarcia@splendeo.es}
  s.homepage    = "https://github.com/splendeo/to_xls"
  s.summary     = %q{to_xls for Enumerations}
  s.description = %q{Adds a to_xls method to all enumerations, which can be used to generate excel files conveniently. Can rely on ActiveRecord sugar for obtaining attribute names.}

  s.rubyforge_project = "to_xls"

  s.files         = Dir['**/*'].select{|f| File.file?(f)}
  s.executables   = s.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  s.test_files    = s.files.grep(%r{^(test|spec|features)/})

  s.require_paths = ["lib"]

  s.add_runtime_dependency "spreadsheet"
  s.add_development_dependency "rspec", "~> 2.11"
  s.add_development_dependency "bundler"
  s.add_development_dependency "rake"
end
