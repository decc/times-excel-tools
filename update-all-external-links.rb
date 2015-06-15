 # Turns all xls spreadsheets in a folder and its subfolders into xlsx spreadsheets. Windows only. Requires Excel.
require 'win32ole'
excel = WIN32OLE.new('Excel.Application')
dir = File.expand_path(File.dirname(ARGV[0] || '.'))
number_of_iterations = ARGV[1].to_i || 12
puts dir
excel.Visible = 0
excel.ScreenUpdating = 0
number_of_iterations.times do 
	Dir.glob("**/*.xls*").each do |workbook|
		puts workbook
		name = File.join(dir,workbook).gsub('/','\\')
		next if File.basename(name).start_with?('~')
		file = excel.Workbooks.Open(name, 0)
		external_links = excel.ActiveWorkbook.LinkSources 
		if external_links
			found_links, missing_links = external_links.partition { |f| File.exist?(f) }
			unless missing_links.empty?
				puts "Not updating these links:"
				puts missing_links
			end
			unless found_links.empty?
				puts "Updating #{found_links.length} external links"
				excel.ActiveWorkbook.UpdateLink( 'Name' => found_links )
				excel.Calculate
				file.Save
			end
		end
		file.Close
		puts
	end
end
excel.Visible = 1
excel.ScreenUpdating = 1
