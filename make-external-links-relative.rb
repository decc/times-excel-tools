# This program opens all .xlsx in this directory and the directories below.
# It looks at the external references in each xlsx file.
#
# If the external reference matches the folder structure that the xlsx is in, then 
# it tries to make that external reference relative.
#
# E.g., If there was an excel spreadsheet:
# model/sub/model.xlsx
#
# That had an external reference to:
# C:/VEDA/model/othersub/othermodel.xlsx
#
# Then it would change that external reference to:
# ../othersub/othermodel.xlsx
#
# Written by Tom COunsell 2015 03 16
#

require 'zip' # xlsx files are zip files full of xml
Dir.glob("**/*.xlsx").each do |workbook| # *.xlsx means all xlsx files **/ means in all subfolders. This returns the url for each excel file in turn
	name = File.basename(workbook) # i.e., model/sub/model.xlsx is named model.xlsx
	next if name.start_with?('~') # Temporary files created by excel should be ignored
	root_folder = File.dirname(workbook).gsub('/','\\') # model/sub is the folder containing the spreadsheet. Need to turn the / from Windows into \ for unix
	parents = root_folder.split('\\') # Turns model/sub into ["model", "sub"]
	top_folder = parents.shift # Gets "model"
	path_to_top_folder = parents.map { "..\\" }.join # Works out the relative url from the current xlsx to the top folder

	Zip::File.open(workbook) do |spreadsheet| # Each spreadsheet is a zip file full of xml files
		link_files = spreadsheet.glob("xl/externalLinks/_rels/*") # We only care about the xml files that contain the external links
		link_files.each do |link_file|
			original = spreadsheet.read(link_file)
			original.gsub!('%20',' ') # Some of the links have mangled urls where the spaces have been replaced by %20
			original.gsub!(/Target="[^"]*?#{Regexp.quote(root_folder)}\\/,'Target="') # This matches links to other excel files in this folder or sub folders of this folder
			original.gsub!(/Target="[^"]*?#{Regexp.quote(top_folder)}\\/,"Target=\"#{path_to_top_folder}") # This matches links to files in folders above this folder
			original.gsub!(/Target="[^"]*?#{Regexp.quote(top_folder.gsub('\\','/'))}\//,"Target=\"#{path_to_top_folder.gsub('\\','/')}") # This does the same, but with the other direction for the backslash (because Ruby gets confused between Windows and Unix approaches to slashes
			spreadsheet.get_output_stream(link_file.name) { |f| f.puts original } # This writes the modified version back out
		end
	end
end
