 # Turns all xls spreadsheets in a folder and its subfolders into xlsx spreadsheets. Windows only. Requires Excel.
require 'win32ole'
excel = WIN32OLE.new('Excel.Application')
dir = File.expand_path(File.dirname(ARGV[0] || '.'))
puts dir
excel.Visible = 1
Dir.glob("**/*.xls").each do |workbook|
	puts workbook
	name = File.join(dir,workbook).gsub('/','\\')
	file = excel.Workbooks.Open(name, 0)
  excel.ActiveWorkbook.UpdateLink( 'Name' => excel.ActiveWorkbook.LinkSources )
	file.Save
	file.Close
end
