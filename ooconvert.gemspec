$:.push File.expand_path('../lib', __FILE__)
require 'ooconvert/version'

Gem::Specification.new do |s|
  s.name        = 'ooconvert'
  s.version     = OOConvert::VERSION
  s.date        = '2014-11-27'
  s.authors     = ['wvengen']
  s.email       = 'dev-rails@willem.engen.nl'
  s.summary     = "Convert documents using Openoffice"
  s.description = "Convert documents to other formats using Openoffice"
  s.homepage    = 'https://github.com/wvengen/ruby-ooconvert'
  s.license     = 'GPL-3.0+'

  s.extra_rdoc_files = ["README.md", "LICENSE.md"]
  s.files            = Dir["lib/**/*"] + s.extra_rdoc_files

  s.require_paths    = ["lib"]
  s.rdoc_options     = ["--charset=UTF-8"]

  s.add_development_dependency 'rake', '>= 7.5.0'
  s.add_development_dependency 'rspec', '~> 3.1.0'
  s.requirements << 'openoffice, libreoffice or soffice in PATH'
end
