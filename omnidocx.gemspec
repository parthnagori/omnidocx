# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'omnidocx/version'

Gem::Specification.new do |spec|
  spec.name          = "omnidocx"
  spec.version       = Omnidocx::VERSION
  spec.authors       = ["Parth Nagori"]
  spec.email         = ["nagori.parth@gmail.com"]

  spec.summary       = %q{A gem to merge docx files, write images to docx and other utilities.}
  spec.description   = %q{A gem that allows to merge multiple docx (microsoft word) files into one, write images to a docx file and make string replacements in the header, footer or main document content.}
  spec.homepage      = "https://github.com/parthnagori/omnidocx"
  spec.license       = "MIT"

  # Prevent pushing this gem to RubyGems.org. To allow pushes either set the 'allowed_push_host'
  # to allow pushing to a single host or delete this section to allow pushing to any host.
  # if spec.respond_to?(:metadata)
  #   spec.metadata['allowed_push_host'] = "TODO: Set to 'http://mygemserver.com'"
  # else
  #   raise "RubyGems 2.0 or newer is required to protect against " \
  #     "public gem pushes."
  # end

  spec.files         = `git ls-files -z`.split("\x0").reject do |f|
    f.match(%r{^(test|spec|features)/})
  end
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_dependency 'nokogiri', '~> 1.6'
  spec.add_dependency 'mime-types', '~> 3.1'
  spec.add_runtime_dependency 'rubyzip', '~> 1.1', '>= 1.1.6'

  spec.add_development_dependency "bundler", "~> 1.13"
  spec.add_development_dependency "rake", "~> 10.0"
  spec.add_development_dependency "rspec", "~> 3.0"

end
