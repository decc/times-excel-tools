# UK TIMES EXCEL TOOLS

Several scripts to help working with UK TIMES excel sheets, or any large set of Excel sheets with multiple update dependencies across a complex folder tree where the name of the root folder changes (usually between versions).

To use, please install:

1. The [Ruby scripting language](www.ruby-lang.org)
2. The [RubyZip Gem](http://rubyzip.sourceforge.net) - (usually by `gem install rubyzip`)

The scripts are:

* [propose-replacements-for-external-links.rb](propose-replacements-for-external-links.rb) - Looks at all xlsx spreadsheets in a folder and subfloders and proposes changes to any absolute links (i.e., they start with a C:) to relative links (i.e., they say that excel file is in the same folder, or a subfolder, or the folder above, as appropriate). Only works on xlsx files, not xls. It also proposes converting links from .xls to .xlsx if a worksheet with the same name but more file format exists.
* [make-replacements-for-external-links.rb](make-replacements-for-external-links.rb) - Takes the output from propose-replacements-for-external-links.rb and actually makes the changes. You can hand-edit the  external-links-to-be-replaced.tsv (produced earlier) so that this script makes those replacements too. This is commonly used to turn ephemeral links, e.g. to file:\\\Q:\.. into dummy links file:\\\A:\.. so that machines which do not have a Q drive do not produce different results from those which do.
* [propose-links-update] (propose-links-update.rb) - Uses Tarjan's algorithm to explore the update links between all the .xlsx files in a folder and subfolders. Detects circular loops. Produces a topological sort sequence of all files such that they can be updated just once, in that order, and all updated values will propagate correctly (only properly feasible where there are no loops). The sorted sequence is in topolist.tsv . This also attempts a hack when it finds loops: it inserts n copies of the files in a cyclic loop of size n into the correct point in the topological sort.
* [make-links-update] (make-links-update.rb) - Reads topolist.tsv and updates all the workbooks in sequence, reading from the external links specified in each workbook. If it hits an error, then it records it in "make-links-errors.tsv" and continues. ALways CLOSE Excel before running this script as it uses WinOLE to run Excel itself.
* [propose-replacements-for-external-links-for-versioned-files.rb](propose-replacements-for-external-links-for-versioned-files.rb) - Look at all xlsx spreadsheets in a folder and subfloders and proposes changes to any external links that point at spreadsheets of the form residential_v0.2.xlsx to point at things like residential.xlsx if (and only if) the version-less file exists.
* [update-all-external-links.rb](update-all-external-links.rb) - Opens each workbook in a folder in unspecified order, updates its external links and saves it. It takes a command-line argument for the number of times to open all the files.
* [external-link-graph](external-link-graph.rb)Creates an HTML file with JavaScript using the D3 library to make a force graph showing the dependencies between all the .xlsx files.
* [make-xls-xlsx.rb](./make-xls-xlsx.rb) - Turns all xls spreadsheets in a folder and its subfolders into xlsx spreadsheets. Windows only. Requires Excel.
* [check-for-external-links.rb](./check-for-external-links.rb) - Looks at all xlsx spreadsheets in a folder and subfolders and reports any links that are not relative (i.e., they start with C: or U: or some such, or are foreign in that they are http://). This list should be checked for local references to files which are only present on your own machine. Only works on xlsx files, not xls
* [external-link-depth](external-link-depth.rb) - DEPRECATED Attempts to measure the nesting depth of the update graph. It fails to do this and the algorithm is completely broken, but it does produce output listing missing files which are linked from the files examined.
