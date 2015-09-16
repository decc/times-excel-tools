# UK TIMES EXCEL TOOLS

A mishmash of scripts to help working with UK TIMES excel sheets:

To use, please install:

1. The [Ruby scripting language](www.ruby-lang.org)
2. The [RubyZip Gem](http://rubyzip.sourceforge.net) - (usually by `gem install rubyzip`)

The scripts are:

* [explore-links] (explore-links.rb) - Uses Tarjan's algorithm to explore the update links between all the .xlsx files in a folder and subfolders. Detects circular loops. Produces a topological sort sequence of all files such that they can be updated just once, in that order, and all updated values will propagate correctly (only feasible where there are no loops). (Currently we do not have a script which actually does that single-pass update - work is ongoing.)
* [check-for-external-links.rb](./check-for-external-links.rb) - Looks at all xlsx spreadsheets in a folder and subfolders and reports any links that are not relative (i.e., they start with C: or U: or some such, or are foreign in that they are http://). This list should be checked for local references to files which are only present on your own machine. Only works on xlsx files, not xls
* [propose-replacements-for-external-links.rb](propose-replacements-for-external-links.rb) - Looks at all xlsx spreadsheets in a folder and subfloders and proposes changes to any absolute links (i.e., they start with a C:) to relative links (i.e., they say that excel file is in the same folder, or a subfolder, or the folder above, as appropriate). Only works on xlsx files, not xls. It also proposes converting links from .xls to .xlsx if a worksheet with the same name but more file format exists.
* [make-replacements-for-external-links.rb](make-replacements-for-external-links.rb) - Takes the output from propose-replacements-for-external-links.rb and actually makes the changes. You can hand-edit the  external-links-to-be-replaced.tsv (produced earlier) so that this script makes those replacements too. This is commonly used to turn ephemeral links, e.g. to file:\\\Q:\.. into dummy links file:\\\A:\.. so that machines which do not have a Q drive do not produce different results from those which do.
* [external-link-depth](external-link-depth.rb) - Attempts to measure the nesting depth of the update graph. It fails to do this accurately, but it idoes produce output listing missing files which are linked from the files examined.
* [propose-replacements-for-external-links-for-versioned-files.rb](propose-replacements-for-external-links-for-versioned-files.rb) - Look at all xlsx spreadsheets in a folder and subfloders and proposes changes to any external links that point at spreadsheets of the form residential_v0.2.xlsx to point at things like residential.xlsx if (and only if) the version-less file exists.
* [update-all-external-links.rb](update-all-external-links.rb) - Opens each workbook in a folder in unspecified order, updates its external links and saves it. It takes a command-line argument for the number of times to open all the files.
* [external-link-graph](external-link-graph.rb)Creates an HTML file with JavaScript using the D3 library to make a force graph showing the dependencies between all the .xlsx files.
* [make-xls-xlsx.rb](./make-xls-xlsx.rb) - Turns all xls spreadsheets in a folder and its subfolders into xlsx spreadsheets. Windows only. Requires Excel.

