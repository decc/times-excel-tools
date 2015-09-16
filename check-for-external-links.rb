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

# updated 2015-09-16

require 'zip'

wsheets = []
targets =[]
links ={}

i=0
Dir.glob("**/*.xls*").each do |workbook|
fname = File.basename(workbook)
next if fname.start_with?('~')
	begin
		Zip::File.open(workbook) do |spreadsheet|
			links[workbook.to_s] =[]
			spreadsheet.glob("xl/externalLinks/_rels/*").each do |link|
				target = link.get_input_stream.read.to_s[/Target="([^"]*)"/,1]
				if target.start_with?('/') || target =~ /^\w+:/
					i += 1
					#puts "#{i}\t#{workbook.to_s}\t#{target.to_s.gsub(/%20/,' ')}\n"
					print "\r#{i} links scanned"
					links[workbook.to_s].push(target.to_s.gsub(/%20/,' '))
					wsheets<<workbook.to_s
					targets<<target.to_s.gsub(/%20/,' ')
				end
			end
		end
	rescue Exception => e
		puts e
	end
end

puts "\r#{i} Scanned links. They are listed in check-links.tsv"

File.open("check-links.tsv","w") do |ff|
	links.each do |w,tl|
		tl.each do |l|
			ff.puts "#{w}\t#{l}"
		end
	end
end


puts "#{wsheets.uniq!.count} files with references. They are listed in check-files-external-links.txt"
File.open("check-files-external-links.txt","w") do |f|
  wsheets.sort!
  f.puts wsheets
end

File.open("unique-check.txt","w") do |f|
  targets.sort!.uniq!
  f.puts targets
end
puts "#{targets.count} Referenced files. They are listed in unique-check.txt"
