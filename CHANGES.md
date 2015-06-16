# 2015-06-16

* Fixed a bug and added documentation to update-all-external-links.rb. The default number of passes has increased to 12.

# 2015-06-15

* Added update-all-external-links.rb which opens all the worksheets in a folder in turn, updates their external links, then cloese them.
* Updated the default number of passes in update-all-external-links to 10

# 2015-05-07

* Add propose-replacements-for-external-links-for-versioned-files.rb to help alter references when eliminating version numbers on individual worksheets

# 2015-05-06

* make-xls-xlsx.rb can now operate on any folder, defaults to the current working directory

# 2015-04-27

* Now also proposes changes to relative external references if it can spot an xlsx in place of an xls file
