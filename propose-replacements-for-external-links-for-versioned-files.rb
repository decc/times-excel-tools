# This program opens all .xlsx in this directory and the directories below.
# It looks at the external references in each xlsx file.
#
# If the external reference ends with something that looks like a version number: 
# e.g., residential_v0.2.xlsx
#
# Then checks if there is a file in the same folder without the version number:
# e.g., residential.xlsx
# 
# If there is a file, then proposes a change in  external-links-to-be-replaced.tsv
# which can be checked and edited in Excel.
#
# Once your are happy with the proposed changes, you can make them by running
# ruby make-external-links-relative.rb 
#
# Written by Tom Counsell 2015 05 7
#

require 'zip' # xlsx files are zip files full of xml
require 'pathname' # Makes manipulating paths easier
require 'uri' # The links are encoded as urls

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

def unversioned_replacement_for_versioned_link(external_reference)
  return nil unless external_reference =~ /_?v?[0-9.]+\.xlsx?$/i
  possible_replacement = external_reference.gsub(/_?v?[0-9.]+(\.xlsx?)$/,'\1')
  if File.exist?(possible_replacement)
    return possible_replacement
  else
    return nil
  end
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
          # We don't replace absolute links
          not_replaced[workbook.to_s][external_reference] = true
          not_replaced_count += 1
        else # We think it is already a relative link
          absolute_link = (workbook.parent+external_reference).to_s
          possible_replacement = unversioned_replacement_for_versioned_link(absolute_link)
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
