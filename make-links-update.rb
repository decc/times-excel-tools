# This opens each spreadsheet from a supplied list in turn and updates its external links.

# It only works in Windows and must have Excel installed.

# ruby path/to/whereever/this/is/make-links-update.rb


# 2015-09-19 PMS

require 'win32ole'

excel = WIN32OLE.new('Excel.Application')
excel.Visible = 1
excel.ScreenUpdating = 1
excel.DisplayAlerts = 1

dir = File.expand_path(File.dirname(ARGV[0] || '.'))
puts "\n#{dir}"

topolist =[]

topolist = File.readlines("topolist.tsv")
puts "#{topolist.count} files to update"

timer = Time.now.to_i

topolist.each do |workbook|
	name=workbook.strip
	unless File.exist?(name)
		puts "File does not exist (and will be skipped):\t#{name}"
	else
		# puts "File exists:\t#{name}"
	end
end

topolist.each do |workbook|
	puts "\nWorkbook: #{workbook}"
	name=workbook.strip

	name=File.expand_path(workbook.strip)
	
		
	unless File.exist?(name)
		puts "File does not exist:\n\t#{name}"
	else
		file = excel.Workbooks.Open(name, 0)
		external_links = excel.ActiveWorkbook.LinkSources
		if external_links
			found_links, missing_links = external_links.partition { |f| File.exist?(f) }
			unless missing_links.empty?
				puts "Not updating using missing link(s):"
				puts missing_links
			end
			unless found_links.empty?
				puts "Updating using #{found_links.length} external link(s):"
				puts found_links
				excel.ActiveWorkbook.UpdateLink( 'Name' => found_links )
				#excel.ActiveWorkbook.UpdateLink
				excel.Calculate
				file.Save
			end
		end
	end
end
	
	

duration = Time.now.to_i - timer
puts "Finished in #{duration} secs."

excel.Visible = 1
excel.DisplayAlerts = 1
excel.ScreenUpdating = 1
