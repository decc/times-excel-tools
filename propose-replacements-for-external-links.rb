# This program opens all .xlsx in this directory and the directories below.
# It looks at the external references in each xlsx file.
#
# If the external reference matches the folder structure that the xlsx is in, then 
# it propses how to make that external link relative
#
# E.g., If there was an excel spreadsheet:
# model/sub/model.xlsx
#
# That had an external reference to:
# C:/VEDA/model/othersub/othermodel.xlsx
#
# Then it would propose changing that external reference to:
# ../othersub/othermodel.xlsx
#
# The proposals are written in external-links-to-be-replaced.tsv
# which can be checked and edited in Excel.
#
# Once your are happy with the proposed changes, you can make them by running
# ruby make-external-links-relative.rb 
#
# Written by Tom Counsell 2015 03 16
#

require 'zip' # xlsx files are zip files full of xml
require 'pathname' # Makes manipulating paths easier
require 'uri' # The links are encoded as urls

wsheets = []
targets =[]

paths_to_all_xlsx_files = Dir.glob("**/*.xlsx").map { |s| Pathname.new(s) } # *.xlsx means all xlsx files **/ means in all subfolders. This returns the url for each excel file in turn

proposed_replacements = Hash.new { |hash, key| hash[key] = {} }
not_replaced = Hash.new { |hash, key| hash[key] = {} }
replacement_count = 0
not_replaced_count = 0

def strip(reference)
  case reference
  when /^file:(.*)$/i, /^[A-Z]:(.*)$/i, /^\/(.*)$/, /^\\(.*)$/
    strip($1)
  else
    reference
  end
end

def decompose(reference)
  reference.split(/[\/\\]+/)
end

def reference_is_local(reference)
  File.exist?(reference.join("/"))
end

def replacement_for(external_reference)
  reference = decompose(strip(external_reference))
  until reference.empty?
    if reference_is_local(reference)
      return reference.join("/")
    end
    reference.shift
  end
  return replacement_for(external_reference.gsub(".xls", ".xlsx")) if external_reference.end_with?(".xls")
  nil
end

def xlsx_replacement_for_xls(external_reference)
  return nil unless external_reference.end_with?(".xls")
  return external_reference.gsub(".xls", ".xlsx") if File.exist?(external_reference.gsub(".xls", ".xlsx"))
  nil
end

paths_to_all_xlsx_files.each do |workbook|
	next if workbook.basename.to_s.start_with?('~') # Temporary files created by excel should be ignored

  Zip::File.open(workbook) do |spreadsheet| # Each spreadsheet is a zip file full of xml files
    link_files = spreadsheet.glob("xl/externalLinks/_rels/*") # We only care about the xml files that contain the external links
    link_files.each do |link_file|
      original = spreadsheet.read(link_file)

      original.scan(/Target="([^"]*)"/) do |match|
        external_reference = URI.unescape($1)
        case external_reference
        when /^file:/i, /^[A-Z]:/i, /^\//, /^\\/
          possible_replacement = replacement_for(external_reference)
          if possible_replacement
            relative_link = Pathname.new(possible_replacement).relative_path_from(workbook.parent).to_s
            proposed_replacements[workbook][external_reference] = relative_link
            replacement_count += 1
          else
            not_replaced[workbook.to_s][external_reference] = true
            not_replaced_count += 1
          end
        else # We think it is already a relative link
          absolute_link = (workbook.parent+external_reference).to_s
          possible_replacement = xlsx_replacement_for_xls(absolute_link)
          if possible_replacement
            # Need to turn it back into a relative link
            relative_link = Pathname.new(possible_replacement).relative_path_from(workbook.parent).to_s
            proposed_replacements[workbook][external_reference] = relative_link
            replacement_count += 1
          else
            not_replaced[workbook.to_s][external_reference] = true
            not_replaced_count += 1
          end
        end

      end 
    end
  end
end
puts "#{replacement_count} replacements proposed in external-links-to-be-replaced.tsv"

File.open("external-links-to-be-replaced.tsv","w") do |f|
  proposed_replacements.each do |worksheet, replacements|
    replacements.each do |original, replacement|
      f.puts "#{worksheet}\t#{original}\t#{replacement}"
    end
  end
end

puts "#{not_replaced_count} external references have not been replaced. They are listed in external-links-not-replaced.tsv"
File.open("external-links-not-replaced.tsv","w") do |f|
  not_replaced.each do |worksheet, references|
    references.each do |original, _|
      f.puts "#{worksheet}\t#{original}"
    end
  end
end

File.open("files-with-external-links.tsv","w") do |f|
  not_replaced.each do |worksheet, references|
    wsheets<<worksheet.to_s   
	references.each do |r, _|
		targets<<r.to_s
	end
  end
  proposed_replacements.each do |worksheet, replacements|
    wsheets<<worksheet.to_s  
	replacements.each do |r, _|
		targets<<r.to_s
	end
  end
  f.puts wsheets
 end
puts "#{wsheets.count} files with references. They are listed in files-with-external-links.tsv"


File.open("unique-propose.txt","w") do |f|
  targets.sort!.uniq!
  targets.each do |t|
      f.puts "#{t}"
    end
end
