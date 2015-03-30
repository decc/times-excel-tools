#Looks at all xlsx spreadsheets in a folder and subfolders and reports any links that are not relative (i.e., they start with C: or U: or some such). Only works on xlsx files, not xls
#
# Written in the Ruby language version 2.2
# Requires the rubyzip gem. Usually installed by gem install rubyzip
#
require 'zip'
Dir.glob("**/*.xlsx").each do |workbook|
name = File.basename(workbook)
next if name.start_with?('~')
	begin
		Zip::File.open(workbook) do |spreadsheet|
			spreadsheet.glob("xl/externalLinks/_rels/*").each do |link|
				target = link.get_input_stream.read.to_s[/Target="([^"]*)"/,1]
				if target.start_with?('/') || target =~ /^\w+:/
					puts "#{workbook}\t#{target}"
				end
			end
		end
	rescue Exception => e
		puts e
	end
end
