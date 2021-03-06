# This opens each spreadsheet in a folder in turn and updates its external links.
# It skips any external links that don't appear to be on the local file system.
# It only works in Windows and must have Excel installed.
# It doesn't try and order the updating. Therefore if spreadsheet A links to 
# spreadsheet B which links to spreadsheet C, then to be sure that all the 
# updates have happened, the utility has to open and update all the spreadsheets twice.
# By default it opens them all 12 times, because that is the maximum chain length 
# in the UK TIMES energy model.
#
# Originally written by Tom Counsell
#
# Usage:
# ruby path/to/whereever/this/is/update-all-external-links.rb
#
# You may want to specify the directory full of spreadsheets:
#
# ruby path/to/whereever/this/is/update-all-external-links.rb directory/with/the/spreadsheets
#
# You may want to specify the number of times to update each spreadsheet:
#
# ruby path/to/whereever/this/is/update-all-external-links.rb directory/with/the/spreadsheets number-of-times-to-update
#
# 2015-09-16 PMS

require 'win32ole'
require 'digest'

excel = WIN32OLE.new('Excel.Application')
dir = File.expand_path(File.dirname(ARGV[0] || '.'))
number_of_iterations = ARGV[1] ? ARGV[1].to_i : 8
puts dir
puts "Number of iterations: #{number_of_iterations}"
excel.Visible = 0
excel.ScreenUpdating = 0
excel.DisplayAlerts = 0

durits = []
empties={}
cksums = {}

timer = Time.now.to_i
iter = 0
number_of_iterations.times do
	iter = iter+1
	timit = Time.now.to_i
	Dir.glob("**/*.xls*").each do |workbook|
		cksums[workbook] = Digest::SHA2.file(workbook).hexdigest
		puts workbook
		name = File.join(dir,workbook).gsub('/','\\')
		next if File.basename(name).start_with?('~')
		next if empties[name]=="empty"
		file = excel.Workbooks.Open(name, 0)
		external_links = excel.ActiveWorkbook.LinkSources
		if external_links
			found_links, missing_links = external_links.partition { |f| File.exist?(f) }
			unless missing_links.empty?
				puts "Not updating these links:"
				puts missing_links
			end
			unless found_links.empty?
				puts "Updating #{found_links.length} external links. Iteration: #{iter}"
				excel.ActiveWorkbook.UpdateLink( 'Name' => found_links )
				excel.Calculate
				file.Save
			end
		else
			empties[name]="empty"
		end
		
		file.Close
		puts
	end
	durit = Time.now.to_i - timit
	puts "Did #{iter}  in #{durit} secs."
	durits[iter]=durit
	
	File.open("cksums-#{iter}.tsv","w") do |f|
	cksums.each do |w,ck|
		f.puts "#{w}\t#{ck}"
	end
end

end
duration = Time.now.to_i - timer
puts "Finished #{number_of_iterations} iterations in #{duration} secs."
File.open("durations.txt","w") do |ff|
      ff.puts durits 
end
excel.Visible = 1
excel.DisplayAlerts = 1
excel.ScreenUpdating = 1
