# UK TIMES EXCEL TOOLS

A mishmash of scripts to help working with UK TIMES excel sheets:

To use, please install:

1. The [Ruby scripting language](www.ruby-lang.org)
2. The [RubyZip Gem](http://rubyzip.sourceforge.net) - (usually by `gem install rubyzip`)

The scripts are:

* [make-xls-xlsx.rb](./make-xls-xlsx.rb) - Turns all xls spreadsheets in a folder and its subfolders into xlsx spreadsheets. Windows only. Requires Excel.
* [check-for-external-links.rb](./check-for-external-links.rb) - Looks at all xlsx spreadsheets in a folder and subfolders and reports any links that are not relative (i.e., they start with C: or U: or some such). Only works on xlsx files, not xls
* [propose-replacements-for-external-links.rb](propose-replacements-for-external-links.rb) - Looks at all xlsx spreadsheets in a folder and subfloders and proposes changes to any absolute links (i.e., they start with a C:) to relative links (i.e., they say that excel file is in the same folder, or a subfolder, or the folder above, as appropriate). Only works on xlsx files, not xls. It also proposes converting links from .xls to .xlsx if a worksheet with the same name but more file format exists.
* [make-replacements-for-external-links.rb](make-replacements-for-external-links.rb) - Takes the output from propose-replacements-for-external-links.rb and actually makes the changes
* [propose-replacements-for-external-links-for-versioned-files.rb](propose-replacements-for-external-links-for-versioned-files.rb) - Look at all xlsx spreadsheets in a folder and subfloders and proposes changes to any external links that point at spreadsheets of the form residential_v0.2.xlsx to point at things like residential.xlsx if (and only if) the version-less file exists.
