# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'excel_manner/version'

Gem::Specification.new do |spec|
  spec.name          = "excel_manner"
  spec.version       = ExcelManner::VERSION
  spec.authors       = ["tomomik"]
  spec.email         = ["tomomik210@gmail.com"]
  spec.summary       = %q{Excel_manner is changing to Excel for deliverables.}
  spec.description   = %q{Excel_manner is changing to Excel for deliverables.}
  spec.homepage      = "https://github.com/tomomik/excel_manner"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.5"
  spec.add_development_dependency "rake"
end
