# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)
require "rxls/version"

Gem::Specification.new do |s|
  s.name        = "rxls"
  s.version     = Rxls::VERSION
  s.platform    = Gem::Platform::RUBY
  s.authors     = ["Frank Behrens"]
  s.email       = ["fbehrens@gmail.com"]
  s.homepage    = ""
  s.summary     = %q{access to excel spreadsheeds}
  s.description = %q{like csv}

  s.rubyforge_project = "rxls"

  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
end
