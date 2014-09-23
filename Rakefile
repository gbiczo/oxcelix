require 'bundler/setup'
require 'rspec'

Bundler.require

task :default => [:test]

task :test do
  rspec "spec/cell_spec.rb"
  ruby "spec/fixnum_spec.rb"
  ruby "spec/matrix_spec.rb"
  ruby "spec/oxcelix_spec.rb"
  ruby "spec/string_spec.rb"
end
