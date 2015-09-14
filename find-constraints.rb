# This program opens all .xlsx in this directory and the directories below.
# It looks at the external references in each xlsx file.
#
# It looks for constraints ACT_BND and CAP_BND
#


# Written by Philip Sargent 2015 09 14
#

require 'zip' # xlsx files are zip files full of xml
require 'pathname' # Makes manipulating paths easier
require 'uri' # The links are encoded as urls

dir = File.expand_path(File.dirname(ARGV[0] || '.'))
$dirx = dir+"/"

puts "Scanning files ..."
puts $dirx.gsub!('/','\\')

founds = {}
ts = []
skips=[]
i=0

# NOTE: this skips .xlsm files because they may have password-protected macros and thus will not be readable.
# Hence the number of files seen does not match those searched for using *.xls* 
Dir.glob("**/*.xls*").each do |workbook|
	name = File.basename(workbook)
	i=i+1
	print "\r#{i}\tworkbooks scanned "
	if name.start_with?('~') # Temporary files should be ignored
		skips<<name
		next
	end
	if name.end_with?('xlsm') # Macro files are ignored - though probably should not be...
		skips<<name
		next
	end
	if name.end_with?('xls') # Macro files are ignored - though probably should not be...
		skips<<name
		next
	end
	begin # start exception block
		Zip::File.open(workbook) do |subfile|
			subfile.glob("xl/externalLinks/_rels/*").each do |link|
				t = link.get_input_stream.read.to_s[/Target="([^"]*)"/,1]
				t = URI.unescape(t)
				unless t =~ /^\w+:/ # keeps http:// and file:// targets as full pathnames but DO NOT add to graph
					
					t = File.basename(t, '.*').downcase
					ts<<t
					
				end						
			end
			subfile.glob("xl/sharedStrings.xml").each do |sheet|
				
				content = sheet.get_input_stream.read
				t = content[/[A-Z]*\_BND[A-Z]*/i]
				unless t.nil?
					found=File.basename(workbook, '.*').downcase.gsub(/%20/,' ')
					founds[found] = workbook.gsub(/%20/,' ')
				end
			end
			subfile.glob("xl/worksheets/*").each do |sheet|
				
				content = sheet.get_input_stream.read
				t = content[/[A-Z]*\_BND[A-Z]*/i]
				unless t.nil?
					found=File.basename(workbook, '.*').downcase.gsub(/%20/,' ')
					founds[found] = workbook.gsub(/%20/,' ')
				end
			end
		end
	rescue Exception => e
		puts "\nEXCEPTION\n#{workbook}"
		puts e
	end
end
puts "\n"

puts "Files containing _BND:"
founds.each do |ff,fp|
	puts "\t#{fp}"
end


# temporary files skipped
puts "\nNOTE that these files have been ignored:"
skips.sort.each do |s|
	puts "\t#{s}"
end
