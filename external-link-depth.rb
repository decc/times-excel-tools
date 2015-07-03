 # Turns all xls spreadsheets in a folder and its subfolders into xlsx spreadsheets. Windows only. Requires Excel.
require 'win32ole'
excel = WIN32OLE.new('Excel.Application')
dir = File.expand_path(File.dirname(ARGV[0] || '.'))
excel.Visible = 0
excel.ScreenUpdating = 0

require 'tsort'

class Hash
	include TSort
	alias tsort_each_node each_key
	def tsort_each_child (node, &block)
		fetch(node).each(&block)
	end
end

links = {}

def depth(name, links, stack = [])
	stack << name
	children = links[name]
	return 0 if children.empty?
	max_depth = 0
	children.each do |child|
		next if stack.include?(child)
		child_depth = depth(child, links, stack)
		if child_depth > max_depth
			max_depth = child_depth
		end
	end
	max_depth + 1
end

	Dir.glob("**/*.xls*").each do |workbook|
		# puts File.absolute_path(workbook)
		name = File.join(dir,workbook).gsub('/','\\')
		next if File.basename(name).start_with?('~')
		file = excel.Workbooks.Open(name, 0)
		external_links = excel.ActiveWorkbook.LinkSources 
		if external_links
			found_links, missing_links = external_links.partition { |f| File.exist?(f) }
			links[File.absolute_path(name).gsub('/','\\')] = found_links
		else
				links[File.absolute_path(name).gsub('/','\\')] = []
		end
		file.Close
	end
	puts links.keys.map { |f| [f, depth(f, links)].join("\t") }.join("\n")
excel.Visible = 1
excel.ScreenUpdating = 1
