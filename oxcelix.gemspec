# -*- encoding: utf-8 -*-

require 'rake'
Gem::Specification.new do |s|
  s.name	= 'oxcelix'
<<<<<<< HEAD
  s.version	= '0.3.0'
  s.date	= '2013-11-12'
=======
  s.version	= '0.2.4'
  s.date	= '2013-10-10'
>>>>>>> origin/master
  s.summary	= 'A fast Excel 2007/2010 file parser'
  s.description	= 'A fast Excel 2007/2010 (.xlsx) file parser that returns a collection of Matrix objects'
  s.authors	= 'Giovanni Biczo'
  s.homepage	= 'http://github.com/gbiczo/oxcelix'
  s.rubyforge_project = 'oxcelix'

  s.files	= FileList["LICENSE", "README.rdoc", "README.md",
		  "lib/oxcelix.rb", "lib/oxcelix/cellhelper.rb",
		  "lib/oxcelix/cell.rb", "lib/oxcelix/numformats.rb",
		  "lib/oxcelix/nf.rb",
		  "lib/oxcelix/sheet.rb", "lib/oxcelix/workbook.rb",
		  "lib/oxcelix/sax/*",
		  "oxcelix.gemspec", "spec/*", ".yardopts", "CHANGES"].to_a
  s.license	= 'MIT'
  
  s.add_runtime_dependency "ox", [">= 2.0.6"]
  s.add_runtime_dependency "rubyzip", [">= 0.9.9"]
  s.rdoc_options << '--all'
end
