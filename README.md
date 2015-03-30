# UK TIMES EXCEL TOOLS

A mishmash of scripts to help working with UK TIMES excel sheets:

To use, please install:

1. The [Ruby scripting language](www.ruby-lang.org)
2. The [RubyZip Gem](http://rubyzip.sourceforge.net) - (usually by `gem install rubyzip`)

The scripts are:

* [make-xls-xlsx.rb](./make-xls-xlsx.rb) - Turns all xls spreadsheets in a folder and its subfolders into xlsx spreadsheets. Windows only. Requires Excel.
* [check-for-external-links.rb](./check-for-external-links.rb) - Looks at all xlsx spreadsheets in a folder and subfolders and reports any links that are not relative (i.e., they start with C: or U: or some such). Only works on xlsx files, not xls
* [make-external-links-relative.rb](make-external-links-relative.rb) - Looks at all xlsx spreadsheets in a folder and subfloders and changes any absolute links (i.e., they start with a C:) to relative links (i.e., they say that excel file is in the same folder, or a subfolder, or the folder above, as appropriate). Only works on xlsx files, not xls


