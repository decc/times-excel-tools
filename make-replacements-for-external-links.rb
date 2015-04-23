# This program opens all .xlsx in this directory and the directories below.
# It looks at the external references in each xlsx file.
#
# If the external reference matches the a proposed replacement in 
# external-links-to-be-replaced.tsv
# then it makes the replacement.
#
# You can generate external-links-to-be-replaced.tsv using propose-replacements-for-external-links.rb
#
# Written by Tom Counsell 2015 03 16
#

require 'zip' # xlsx files are zip files full of xml
require 'pathname' # Makes manipulating paths easier
require 'uri' # The links are encoded as urls

paths_to_all_xlsx_files = Dir.glob("**/*.xlsx").map { |s| Pathname.new(s) } # *.xlsx means all xlsx files **/ means in all subfolders. This returns the url for each excel file in turn

unless File.exist?("external-links-to-be-replaced.tsv")
  puts "Can't find external-links-to-be-replaced.tsv"
  puts "You can use propose-replacements-for-external-links.rb to generate this file"
  exit
end

proposed_replacements = Hash.new { |hash, key| hash[key] = {} }

IO.readlines("external-links-to-be-replaced.tsv").each do |line|
  line.strip!
  worksheet, original, replacement = *line.split("\t")
  proposed_replacements[worksheet][original] = replacement
end

replacement_count = 0

paths_to_all_xlsx_files.each do |workbook|
	next if workbook.basename.to_s.start_with?('~') # Temporary files created by excel should be ignored
  next unless proposed_replacements.has_key?(workbook.to_s)

  Zip::File.open(workbook) do |spreadsheet| # Each spreadsheet is a zip file full of xml files
    link_files = spreadsheet.glob("xl/externalLinks/_rels/*") # We only care about the xml files that contain the external links
    link_files.each do |link_file|
      original = spreadsheet.read(link_file)
      replacement = original.gsub(/Target="([^"]*)"/) do |match|
        external_reference = URI.unescape($1)
        possible_replacment = proposed_replacements[workbook.to_s][external_reference]
        if(possible_replacment)
          replacement_count += 1
          "Target=\"#{possible_replacment}\""
        else
          match
        end
      end
      spreadsheet.get_output_stream(link_file.name) { |f| f.puts replacement } # This writes the modified version back out
    end
  end
end
puts "#{replacement_count} replacements made"