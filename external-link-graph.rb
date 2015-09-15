# This writes out link-graph.html which shows the links between all the spreadsheets in a folder
# More for fun than for any practical value

# Edited Philip Sargent 2015-09-13
# updated 2015-0915

output_filename = 'link-graph.html'

require 'win32ole'
require 'json'

links = []
nodes = []
used_nodes=[]

excel = WIN32OLE.new('Excel.Application')
dir = File.expand_path(File.dirname(ARGV[0] || '.'))

excel.Visible = 0
excel.ScreenUpdating = 0
excel.DisplayAlerts = 0
i = 0

Dir.glob(File.join(dir,"**/*.xls*")).each do |workbook|
	path = File.absolute_path(workbook).gsub('/','\\')
	name = File.basename(path, '.*').downcase
	nodes << {name: name, path: path}
	i=i+1
	print "\r#{i}\tworkbooks scanned "
	next if name.start_with?('~')
	file = excel.Workbooks.Open(path, 0)
	external_links = excel.ActiveWorkbook.LinkSources 
	if external_links
		external_links.each do |link|
			nodes << { name: File.basename(link, '.*'), path: File.absolute_path(link).gsub('/','\\') }
			links << { source: path, target: link, value: 1 }
		end
	end
	file.Close
end

excel.Visible = 1
excel.ScreenUpdating = 1
excel.DisplayAlerts = 1

def normalise(path)
	File.absolute_path(path).gsub('/','\\').downcase
end

$node_lookup = {}

puts "\n#{nodes.count}\ttotal nodes"
nodes.uniq!
puts "#{nodes.count}\tunique nodes"
puts "#{links.count}\tlinks"

nodes.each.with_index do |node,i|
	$node_lookup[normalise(node[:path])] = i
end

def lookup(name)
	index = $node_lookup[normalise(name)]
	return index if index
	puts "Can't find node for reference "+name+" normalised as "+normalise(name)
end

reformatted_links = links.map do |link|
	{
		source: lookup(link[:source]),
		target: lookup(link[:target]),
		value: link[:value]
	}
end

links.each do |link|
	used_nodes << lookup(link[:source])
	used_nodes << lookup(link[:target])
end

# puts "#{used_nodes.count}\ttotal nodes in links"
used_nodes.uniq!
puts "#{used_nodes.count}\tunique nodes in links"

# Now look for nodes that are neither sources nor targets, and try remove them from the nodes list?
nodes.each.with_index do |node,i|
	unless used_nodes.include?($node_lookup[normalise(node[:path])])
		# puts "#{node[:path]}\tnot part of any link"
		# nodes.delete(node)	
	end
end

# We can't simply remove them from the nodes list as the numbering seems to be implicitly used by d3
# so we would need to renumber everything.

html =<<END
<!DOCTYPE html>
<meta charset="utf-8">
<style>

.node {
  stroke: #fff;
  stroke-width: 1.5px;
}

.link {
  stroke: #999;
  stroke-opacity: .6;
}

</style>
<body>
<script src="https://cdnjs.cloudflare.com/ajax/libs/d3/3.5.5/d3.min.js"></script>
<script>

var graph = {
	nodes: #{nodes.to_json},
	links: #{reformatted_links.to_json}
};

var width = window.innerWidth;
var height = window.innerHeight;

var color = d3.scale.category20();

var force = d3.layout.force()
    .charge(-120)
    .linkDistance(30)
    .size([width, height]);

var svg = d3.select("body").append("svg")
    .attr("width", width)
    .attr("height", height);

  force
      .nodes(graph.nodes)
      .links(graph.links)
      .start();

  var link = svg.selectAll(".link")
      .data(graph.links)
    .enter().append("line")
      .attr("class", "link")
      .style("stroke-width", function(d) { return Math.sqrt(d.value); });

  var node = svg.selectAll(".node")
      .data(graph.nodes)
    .enter().append("circle")
      .attr("class", "node")
      .attr("r", 5)
      .style("fill", function(d) { return color(d.group); })
      .call(force.drag);

  node.append("title")
      .text(function(d) { return d.name; });

  force.on("tick", function() {
    link.attr("x1", function(d) { return d.source.x; })
	.attr("y1", function(d) { return d.source.y; })
	.attr("x2", function(d) { return d.target.x; })
	.attr("y2", function(d) { return d.target.y; });

    node.attr("cx", function(d) { return d.x; })
	.attr("cy", function(d) { return d.y; });
  });

</script>
END

File.open(output_filename,'w') { |f| f.puts html }


