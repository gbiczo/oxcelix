Oxcelix
=======
<a href="http://badge.fury.io/rb/oxcelix"><img src="https://badge.fury.io/rb/oxcelix@2x.png" alt="Gem Version" height="18"></a>
[![Build Status](https://travis-ci.org/gbiczo/oxcelix.svg?branch=0.4.0)](https://travis-ci.org/gbiczo/oxcelix)

Oxcelix - A fast and simple .xlsx file parser

Description
-----------

Oxcelix is an xlsx (Excel 2007/2010) parser. The result of the parsing is a
Workbook which is an array of Sheet objects, which in turn store the data in
Matrix objects. Matrices consist of Cell objects to maintain comments and
formatting/style data

Oxcelix uses the great Ox gem (http://rubygems.org/gems/ox) for fast SAX-parsing.

Synopsis
--------

  To process an xlsx file:

    `require 'oxcelix'`
    `w = Oxcelix::Workbook.new('whatever.xlsx')`

  To omit certain sheets:

    `w = Oxcelix::Workbook.new('whatever.xlsx', :exclude => ['sheet1', 'sheet2'])`

  Include only some of the sheets:

    `w = Oxcelix::Workbook.new('whatever.xlsx', :include => ['sheet1', 'sheet2', 'sheet3'])`

  To have the values of the merged cells copied over the mergegroup:

    `w = Oxcelix::Workbook.new('whatever.xlsx', :copymerge => true)`
  
  Convert a Sheet object into a collection of ruby values or formatted ruby strings:
  
    `require 'oxcelix'`
    `w = Oxcelix::Workbook.new('whatever.xlsx', :copymerge => true)`
    `w.sheets[0].to_ru # returns a Matrix of DateTime, Integer, etc objects`
    `w.sheets[0].to_fmt # returns a Matrix of formatted Strings based on the above.`
  OR:
    `require 'oxcelix'`
    `w = Oxcelix::RuValueWorkbook.new('whatever.xlsx', :copymerge => true)`
    `w = Oxcelix::FormattedWorkbook.new('whatever.xlsx', :copymerge => true)`

  You can parse an Excel file partially to save memory:
    `require 'oxcelix'`
    `w = Oxcelix::Workbook.new('whatever.xlsx', :cellrange => ('A3'..'R42')) # will only parse the cells included in the given range on every included sheet`
    `w = Oxcelix::Workbook.new('whatever.xlsx', :paginate => [5,2]) # will only parse the second five-row group of every included sheet.`

Installation
------------

  `gem install oxcelix`


Advantages over other Excel parsers
-----------------------------------

Excel file processing involves XML document parsing. Usually, this is achieved by DOM-parsing the data with a suitable XML library.

The main drawbacks of this approach are memory usage and speed. The resulting object tree will be roughly as big as the original file, and during the parsing, they will both be stored in the memory, which can cause quite some complications when processing huge files. Also, interpreting every bit of an excel spreadsheet will slow down unnecessarily the process, if we only need the data stored in that file.

The solution for the memory-consumption problem is SAX stream parsing. This ensures that only the relevant XML elements get processed, 
without having to load the whole XML file in memory.

Oxcelix uses the SAX parser offered by Peter Ohler's Ox gem. I found Ox SAX parser quite fast, so to further speed up the parsing.

For a comparison of XML parsers, please consult the Ox homepage[http://www.ohler.com/dev/xml_with_ruby/xml_with_ruby.html].


TODO
----
  * include/exclude mechanism should extend to cell areas inside Sheet objects
  * Further improvement to the formatting algorithms. Theoretically, to_fmt should be able to
    split conditional-formatting strings and to display e.g. thousands separated number strings
