# Produces the SHA256 checksum digest for the worksheets in an Excel .xlsx file
# To be used for re-factoring large Excel projects where the results in teh final workbooks
# is to be kept unchanged.

# Checked: the SHA256 digest produces the same string as the SHA256 tool built into 7-zip.

# Version 1.0.0
# 2016-01-01 Philip Sargent, philip.sargent@decc.gsi.gov.uk  


# TO DO: implement ARGF mechanism to handle "-" and piped files as per 
# https://robots.thoughtbot.com/rubys-argf
# to make it easier to build this script into complex toolchains.

# TO DO: implement -t, --truncate option to return only first 32 characters of 
# the 64 character SHA256 digest.  For ease of use when used by hand.

# TO DO: replace the 'choices' command flag handling with optparse or Trollop.
# http://docs.ruby-lang.org/en/2.1.0/OptionParser.html
# http://trollop.rubyforge.org/

require 'choice' # https://github.com/defunkt/choice
require 'zip' # xlsx files are zip files full of xml
require 'pathname' # Makes manipulating paths easier
require 'uri' # The links are encoded as urls

require 'digest/sha2'
require 'base64'

VERSION = "1.0.0"

Choice.options do
 
  banner ''
  banner ''
  banner 'This produces the SHA256 checksum digest for the worksheets in an Excel .xlsx file'
  header ''
  header 'Options:'
  
  option :verbose do
    short '-v'
    long '--verbose'
    desc 'Produces digest values for each component file of the .xlsx workbook.'
  end

  option :all do
    short '-a'
    long '--all'
    desc 'Produces digest value for the whole file, including /docProps/core.xml'
  end
 
   option :help do
   long '--help'
    desc 'Show this message only'
  end


  separator ''
  separator "Usage: ruby " + File.basename($0) + " [-v][-a] [--help] <filename.xlsx>"
  separator ''

end

if !Choice.choices[:verbose] 
	# print "Not verbose\n"
end
if !Choice.choices[:all] 
	# print "Not all\n"
end

if ARGV[ARGV.count-1].nil? or !File.file?(ARGV[ARGV.count-1])
	puts "\nFile does not exist: #{ARGV[ARGV.count-1]}"
	puts "\nUsage: ruby " + File.basename($0) + " [-v][-a] [--help] <filename.xlsx>"
	exit
end

cksums ={}
digsheets = 0

name = ARGV[ARGV.count-1]
#puts "\nFile exists: #{name}"

#       :: the digest for the worksheets collectively
# -a    :: the digest for the whole file (without unzipping it)
# -a -v :: the digests for each and every subfile
# -v    :: the digests for ONLY the worksheets subfiles

allfile= Digest::SHA256.file(name).hexdigest
	
if Choice.choices[:all] and !Choice.choices[:verbose]
	puts allfile
	exit
end

begin # start exception block
	Zip::File.open(name) do |zipfile|
		# subfile has the same name as the .xlsx file but it is unzipped
		
		worksheets=""
		zipfile.each do |ff|
			fname = "#{name}/#{ff}"
			
			if !Choice.choices[:all] and ff.to_s.end_with?('docProps/core.xml') # Always changes if file opened
				# skips<<"#{name}/#{ff}"
				next 
			end
			contents = zipfile.read(ff)
			dig=Digest::SHA2.hexdigest(contents)
			if ff.to_s.start_with?('xl/worksheets/') # a subset of the consituents
				worksheets = worksheets + contents
				cksums[fname] = dig
			end
			if Choice.choices[:all]
				cksums[fname] = dig
			end
		end
		digsheets = Digest::SHA2.hexdigest(worksheets)
		
	end
rescue Exception => e # This happens if it can't open a .xlsm file
	puts "\nEXCEPTION\n#{name}"
	puts e
end

if Choice.choices[:verbose] 
	cksums.each do |ck|
		puts "#{ck[1]} #{ck[0]}"
	end
	if Choice.choices[:all] 
		puts allfile + " " + name + " (zipped)"
	end
else	
	puts digsheets 
end
