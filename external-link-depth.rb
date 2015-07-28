 # Scans for external links in a folder tree of spreadsheet files. Requires Excel.
 # Edited Philip Sargent, 28 July 2015
require 'win32ole'
excel = WIN32OLE.new('Excel.Application')
excel.Visible = 0
excel.ScreenUpdating = 0
excel.DisplayAlerts = 1
msoAutomationSecurityForceDisable = 3 # https://msdn.microsoft.com/EN-US/library/office/ff862064.aspx
excel.AutomationSecurity = msoAutomationSecurityForceDisable


dir = File.expand_path(File.dirname(ARGV[0] || '.'))
puts "Scanning files for depth of reference link nesting..."

require 'tsort'

class Hash
	include TSort
	alias tsort_each_node each_key
	def tsort_each_child (node, &block)
		fetch(node).each(&block)
	end
end

links = {}
missing = {}
empties = []

def depth(name, links, stack = [])
	stack << name
	#puts name
	return 0 unless links.key?(name) # catches linked filenames which are not found
	children = links[name]
	return 0 if children.none? # catches [nil] as well as []
	#puts "-> 	"
	#puts "   	#{children}\n"
	max_depth = 0
	children.each do |child|
		next if stack.include?(child)
		child_depth = depth(child, links, stack)
		if child_depth > max_depth
			max_depth = child_depth
		end
	end
	max_depth = max_depth + 1
	#puts max_depth
	return max_depth
end

	i=0
	
	Dir.glob("**/*.xls*").each do |workbook|
		i=i+1
		#puts "    Finding..."
		name = File.join(dir,workbook).gsub('/','\\')
		#puts "#{i} #{File.absolute_path(workbook)}"
		print "\r#{i}   files scanned  "
		next if File.basename(name).start_with?('~')
		#puts "    Opening...#{File.absolute_path(workbook)}"
		file = excel.Workbooks.Open(name, 0)
		external_links = excel.ActiveWorkbook.LinkSources
		if external_links
			found_links, missing_links = external_links.partition { |f| File.exist?(f) }
			links[File.absolute_path(name).gsub('/','\\')] = found_links
			#puts "- Links found: #{found_links.count}"
			#puts "- Links missing: #{missing_links.count}"
			missing[File.absolute_path(name).gsub('/','\\')] = missing_links if missing_links.count >0
		else
			links[File.absolute_path(name).gsub('/','\\')] = []
			empties<<File.absolute_path(name).gsub('/','\\')
			#puts "  No External Links"
		end
		#puts "   #{empties.count} empties found so far" 
		#puts "    Closing...#{File.absolute_path(workbook)}"
		file.Close(0)
	end
	puts "\n\nMissing links in these files:"

	missing.keys.each do |miss|
		puts "\n",miss
		missing[miss].each do |gone|
			puts "-> MISSING #{gone} "
		end
	end
	
	puts "\nTotal links: #{links.keys.count}"
	if links.keys.count !=links.values.count then puts "(Mismatched no. of keys and values)" end
	
	puts "\nDepth traverse:"
	
	mymax=0
	depths = links.keys.map { |f| [f, depth(f, links)]}
	depths.each{ |_,n| mymax=n if n>mymax }

	puts "Max depth = #{mymax}"
	
	File.open("link-map.tsv","w") do |ff|
      ff.puts links.keys.map { |f| [f, depth(f, links)].join("\t") }.join("\n")
    end

	puts "Written file links-map.tsv "
	excel.Visible = 1
	excel.ScreenUpdating = 1
	excel.DisplayAlerts = 1
	