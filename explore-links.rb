# This program opens all .xlsx in this directory and the directories below.
# It looks at the external references in each xlsx file.
#
# It explores the link structure between workbooks, looking for loops and
# calculating the depth of the link tree.
#

# TO DO - fix the folder formatting so that targets are in the same format as parents - otherwise Tsort can't do its stuff.
# TO DO as we are interested in the link graph, EXCLUDE all the file:/// and http:// links.

# Written by Philip Sargent 2015 09 06
#

require 'zip' # xlsx files are zip files full of xml
require 'pathname' # Makes manipulating paths easier
require 'uri' # The links are encoded as urls

require 'tsort'

class TsortableHash < Hash
	include TSort
	alias tsort_each_node each_key
	def tsort_each_child (node, &block)
		fetch(node).each(&block)
	end
end

dir = File.expand_path(File.dirname(ARGV[0] || '.'))
$dirx = dir+"/"

puts "Scanning files for depth of reference link nesting..."
puts $dirx.gsub!('/','\\')

parents = []
skips=[]

arcs = TsortableHash.new
#arcs = Hash.new
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
	begin # start exception block
		Zip::File.open(workbook) do |subfile|
			parent = File.basename(workbook, '.*').downcase.gsub(/%20/,' ')
			# parent = workbook.to_s.gsub(/%20/,' ')

			targets =[]
			subfile.glob("xl/externalLinks/_rels/*").each do |link|
				t = link.get_input_stream.read.to_s[/Target="([^"]*)"/,1]
				t = URI.unescape(t)
				unless t =~ /^\w+:/ # keeps http:// and file:// targets as full pathnames but DO NOT add to graph
					parents << parent
					t = File.basename(t, '.*').downcase
					targets << t
					arcs[parent] = targets
				end						
			end
		end
	rescue Exception => e
		puts "\nEXCEPTION\n#{workbook}"
		puts e
	end
end
puts "\n"

parents.sort!.uniq!
puts "#{parents.count}\tworkbooks have links to other files."
File.open("explore-parents.tsv","w") do |f|
  parents.sort!
  f.puts parents
end

alltargets = []
arcs.each do |p,t|
	unless t.nil? 
		alltargets.push(*t) # flatten before adding
	end
end

#WHY is this not sorting !??!
newt = alltargets.sort.uniq
puts "#{newt.count}\tlocal files are targets of these links."
newt.sort!
File.open("unique-explore.tsv","w") do |f|
	newt.each do |t|
		f.puts t
	end
end

# The Tsort capability requires that the targets are formatted exactly as the parents.
# This is a problem, as within a directory, the target is a simple filename, e.g. transport_v0.2.xlsx  
# whereas links from workbooks in other folders may say "../transport/transport_v0.2.xlsx" 
# or "uk_times_data/transport_v0.2.xlsx" . 

# In UK TIMES all filenames are unique, so this is not a problem.

lc = 0
newarcs=[]
File.open("links-explore.tsv","w") do |f|
	arcs.each do |p,_|
		arcs[p].each do |t|
			lc += 1
			f.puts "#{p}\t#{t}"
			unless arcs.has_key?(t)
				newarcs << t
			end
		end
	end
end

puts "#{lc}\ttotal number of links between local files."

# Now ensure that every target is also in the hash
newarcs.each do |a|
	arcs[a]=[]
end

puts "\nAll file-to-file links:"
arcs.each do |p,t|
	unless t.empty? 
		puts "\n#{p}"
		puts "#{t}"
	end
end

# Using the Tsort methods will expose cyclic depdendency
begin # start exception block
	puts "\nStrongly connected components:"
	arcs.each_strongly_connected_component do |p,t|
		#puts "parent:#{p}"
		#puts "\ttargets: #{t}"
	end
	puts "\nTopological sort:"
	puts arcs.tsort
rescue Exception => e
	puts "\nEXCEPTION\n"
	puts e
end


# temporary files skipped
puts "\nNOTE that these files have been ignored:"
skips.each do |s|
	puts "\t#{s}"
end
