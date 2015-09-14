#[Originally written by Tom Counsell]
#Looks at all xlsx spreadsheets in a folder and subfolders and reports any links that are not relative (i.e., they start with C: or U: or some such). Only works on xlsx files, not xls
#
# Written in the Ruby language version 2.2
# Requires the rubyzip gem. Usually installed by gem install rubyzip
#

# THIS CODE MISSIES SOME files with links compared with propose-replacements-for-external-links.rb !!
# The regex are slightly different...
# so it finds only 94 of the 109 files with links...

# This is because this should only be run AFTER the propose/make--replacements-for-external-links script
# in which case it will ONLY show the file:/// and http:// https:// external links. OK.
# Philip Sargent 2015/09/06

require 'zip'

wsheets = []
targets =[]


Dir.glob("**/*.xls*").each do |workbook|
name = File.basename(workbook)
File.open("checks.tsv","w") do |f|
	next if name.start_with?('~')
		begin
			Zip::File.open(workbook) do |spreadsheet|
				spreadsheet.glob("xl/externalLinks/_rels/*").each do |link|
					target = link.get_input_stream.read.to_s[/Target="([^"]*)"/,1]
					if target.start_with?('/') || target =~ /^\w+:/
						f.puts "#{workbook}\t#{target}"
						wsheets<<workbook.to_s
						targets<<target.to_s.gsub(/%20/,' ')
					end
				end
			end
		rescue Exception => e
			puts e
		end
	end
end
puts "#{wsheets.uniq!.count} files with references. They are listed in check-files-external-links.tsv"
File.open("check-files-external-links.tsv","w") do |f|
  wsheets.sort!
  f.puts wsheets
end

File.open("unique-check.txt","w") do |f|
  targets.sort!.uniq!
  f.puts targets
end