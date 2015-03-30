 # Turns all xls spreadsheets in a folder and its subfolders into xlsx spreadsheets. Windows only. Requires Excel.
require 'win32ole'
excel = WIN32OLE.new('Excel.Application')
dir = File.dirname(File.expand_path(__FILE__))
puts dir
excel.Visible = 1
Dir.glob("**/*.xls").each do |workbook|
	puts workbook
	name = File.join(dir,workbook).gsub('/','\\')
	new_name = name.gsub('.xls','.xlsx')
	next if File.exist?(new_name)
	file = excel.Workbooks.Open(name, 0)
	file.saveAs(new_name, 51)
	file.Close
end

