# This program opens all .xlsx in this directory and the directories below.
# It looks at the external references in each xlsx file.
#
# It explores the link structure between workbooks, looking for loops and
# calculating the depth of the link tree.
#

# As we are interested in the link graph of local files only, EXCLUDE all the file:/// and http:// links.

# Still need to check that base filenames are globally unique !
# Still need to sort out proper paths for export list of sorted filenames
# Still need to deal with links to files which do not actually exist.

# Originally written by Philip Sargent 2015 09 06
# updated 2015-09-17 PMS

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
fullpaths = {}
cksums = {}

arcs = TsortableHash.new
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
	if name.end_with?('xlsm') # Macro files are ignored - though probably should not be in general..! OK in UK TIMES I hope.
		skips<<name
		next
	end
	begin # start exception block
		Zip::File.open(workbook) do |subfile|
			parent = File.basename(workbook, '.*').downcase.gsub(/%20/,' ')
			# parent = workbook.to_s.gsub(/%20/,' ')
			fullpaths[parent]=workbook.gsub(/%20/,' ')

			targets =[]
			subfile.glob("xl/externalLinks/_rels/*").each do |link|
				t = link.get_input_stream.read.to_s[/Target="([^"]*)"/,1]
				t = URI.unescape(t)
				unless t =~ /^\w+:/ # keeps http:// and file:// targets as full pathnames but DO NOT add to graph
					parents << parent
					tb = File.basename(t, '.*').downcase
					#if fullpaths[tb].nil?
					#	fullpaths[tb] = t.gsub(/%20/,' ').gsub!('/','\\')
					#end
					targets << tb
					arcs[parent] = targets
					
				end						
			end
		end
	rescue Exception => e # This happens if it can't open a .xlsm file
		puts "\nEXCEPTION\n#{workbook}"
		puts e
	end
end
puts "\n"

parents.sort!.uniq!
puts "#{parents.count}\tworkbooks have links to other local files."
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

alltargets.each do |t| 
	if fullpaths[t].nil?
		puts "\nMISSING Full path for this workbook: #{t}"
	else
		unless File.exist?(fullpaths[t])
			puts "\nMISSING FILE: #{t} #{fullpaths[t]}"
		end
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

puts "#{lc}\ttotal number of direct links between local files."

# Now ensure that every target is also in the hash
newarcs.each do |a|
	arcs[a]=[]
end

puts "\nAll file-to-file links in dependencies.txt"
File.open("dependencies.txt","w") do |f|
	arcs.each do |p,t|
		unless t.empty? 
			f.puts "\n#{p}"
			f.puts "#{t}"
		end
	end
end

cyclics = {}
# Using the Tsort methods will expose cyclic depdendency
begin # start exception block
	nl=0
	puts "\nStrongly connected components aka cyclic dependencies (with more than one member):"
	# arcs.each_strongly_connected_component {|scc| p scc }
	arcs.each_strongly_connected_component do |scc| # iterates through arrays of nodes
		if scc.length > 1 then 
			nl = nl +1 
			cyclics[nl]=scc
			puts "\nCyclic loop #{nl}:"
			scc.each do |g|
				printf "   %-40s %s\n",  g, fullpaths[g]
			end
		end
		#puts "parent:#{p}"
		#puts "\ttargets: #{t}"
	end
	# puts "\nTopological sort:"
	# puts arcs.tsort
rescue Exception => e
	puts "\nEXCEPTION Begin\n"
	puts e
	puts "\nEXCEPTION End\n"
end

listids ={}
# Now compact each loop into one node, and produce a clean topo sort.
# First create new parent nodes
cyclics.each do |nl,scc|
	id="~SCC-" + nl.to_s
	listids[nl] = id.to_s
	arcs[id] = [] # create new node for the compactified scc
	scc.each do |p|
		# puts id, p, "->",arcs[p]
		
		# copy target lists of all loop members to new parent
		tl = arcs[id] + arcs[p]
		arcs[id] = tl
		arcs.delete(p)
		
		# Next replace all targets with the loop names
		arcs.each do |pp,_|
			if arcs[pp].include?(p)
				#puts "replace #{pp} - with link to #{id}"
				arcs[pp].delete(p)
				arcs[pp] << id
			end
			
		
		end
	end
	#puts "\n#{id} -> #{arcs[id]}"
	


end



# Next remove duplicates in target lists 
arcs.each do |p,t|
	targets = arcs[p]
	arcs[p] =[]
	arcs[p] = targets.uniq
end

cyclics2 = {}
topolist = []
# Using the Tsort methods will expose cyclic depdendency
begin # start exception block
	nl=0
	puts "\nTry again having compacted the loops \n(aka Strongly connected components aka cyclic dependencies with more than one member):"
	# arcs.each_strongly_connected_component {|scc| p scc }
	arcs.each_strongly_connected_component do |scc| # iterates through arrays of nodes
		if scc.length > 1 then 
			nl = nl +1 
			cyclics2[nl]=scc
			puts "\nCyclic loop #{nl}:"
			scc.each do |g|
				printf "   %-40s %s\n",  g, fullpaths[g]
			end
		end
		#puts "parent:#{p}"
		#puts "\ttargets: #{t}"
	end
	#puts "\nTopological sort:"
	topolist = arcs.tsort
	#puts "#{topolist}"
rescue Exception => e
	puts "\nEXCEPTION Begin\n"
	puts e
	puts "\nEXCEPTION End\n"
end


#puts "\nNow replace the SCC items"
#Now replace each SCC with the real files, but duplicate by the size of the SCC

# First prepare the replacements for ~SCC-1 etc
fscc = {}

cyclics.each do |nl, _|
	fscc[listids[nl]]=[]
	#puts "\n#{listids[nl]}"
	cyclics[nl].each do |_| # repeat n times for cyclic loop of size n
		fscc[listids[nl]].push(cyclics[nl])
	end
	#puts fscc[listids[nl]]
end

#puts "#{topolist}\n"
#puts "\n"
#puts listids
#puts "\nNow replace SCCs with multiplied original workbook names"
newtopo = []

substitutes=[]
topolist.each do |t|
	found=false
	listids.each do |nl, _|
		if listids[nl].match(t)
			found=true
			substitutes = fscc[listids[nl]]
		end
	end
	
	if found 
		#puts "Found #{t}"
		newtopo.push(substitutes.flatten)
		#puts "Replacing with #{substitutes}"
		#puts "#{newtopo}\n"
		next
	else
		newtopo << t
		#puts "Adding #{t}"
	end
end
#puts "#{newtopo}\n"
	
newtopo.flatten!

puts "\nTopological sorted filelist in topolist.tsv"
File.open("topolist.tsv","w") do |f|
	newtopo.each do |n|
		f.puts "#{n}\t#{fullpaths[n]}"
	end
end

# temporary files skipped
puts "\nNOTE that these files have been ignored (either .xlsM or ~temporaries):"
skips.each do |s|
	puts "\t#{s}"
end
