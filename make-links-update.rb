# This opens each spreadsheet from a supplied list in turn and updates its external links.

# It only works in Windows and must have Excel installed.

# ruby path/to/whereever/this/is/make-links-update.rb

# It is quicker to update all links at once, but if it fails on a single update, we need
# to know which it waas. Which is why we iterate through the links and accept an
# update from each one individually.

# Don't know why the excel application is not shutting down, even though we catch the
# failure with an exception handler and see the message.

# 2015-10-07 PMS

require 'win32ole'

excel = WIN32OLE.new('Excel.Application')
excel.Visible = 0
excel.ScreenUpdating = 0
excel.DisplayAlerts = 1

dir = File.expand_path(File.dirname(ARGV[0] || '.'))
puts "\n#{dir}"
dir = dir.gsub('/','\\')+"\\"

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

# ensure error file is empty
ferror = "make-links-errors.tsv"
f = File.new(ferror, "w")
f.close

topolist.each do |workbook|
	puts "\nWorkbook: #{workbook}"
	name=workbook.strip

	name=File.expand_path(workbook.strip)
	
		
	unless File.exist?(name)
		puts "File does not exist:\n\t#{name}"
	else
			
		begin # start exception block
			wb = excel.Workbooks.Open(name, 0)
			external_links = excel.ActiveWorkbook.LinkSources
			if external_links
				found_links, missing_links = external_links.partition { |f| File.exist?(f) }
				unless missing_links.empty?
					puts "Not updating using missing link(s):"
					puts missing_links
				end
				unless found_links.empty?
					puts "Updating using #{found_links.length} external link(s):"
					found_links.each do |k|
						 #puts k.gsub(dir,'')
							unless File.exist?(k)
								puts "Linked file does not exist:\n\t#{k}"
							end
					end
					#puts " "
					begin # start inner exception block
					
					found_links.each do |k|
						begin # start inner exception block
							# puts k.gsub(dir,'')
							excel.ActiveWorkbook.UpdateLink( 'Name' => k )
							excel.Calculate
						rescue Exception => e # This happens if it fails due to a very obscure OLE error..
							puts "\nEXCEPTION with \n\t#{name.gsub('/','\\').gsub(dir,'')} \nPROBLEM updating using \n\t#{k.gsub(dir,'')}"
							puts e
							
							File.open(ferror, 'a') { |f| f.puts "#{name.gsub('/','\\').gsub(dir,'')}\t#{k.gsub(dir,'')}" }
							
							#wb.Close
							#excel.Quit()
							
							#excel2 = WIN32OLE.connect('Excel.Application')
							#excel2.Quit()
						end
					end

					#excel.ActiveWorkbook.UpdateLink( 'Name' => found_links )
					#excel.ActiveWorkbook.UpdateLink
					excel.Calculate
					wb.Save
					rescue Exception => e # This happens if it fails due to a very obscure OLE error..
						puts "\nEXCEPTION with \n\t#{name.gsub(dir,'')} \nPROBLEM updating (outer loop - all UpdateLink files) "
						puts e
						#wb.Close
						#excel.Quit()
					end
				end
			wb.Close
			end
		rescue Exception => e # This happens if it fails due to obscure OLE error..
			puts "\nEXCEPTION with \n\t#{name.gsub(dir,'')} - \nPROBLEM GENERIC - problem opening it ?"
			puts e
			# wb.Close
			#excel.Quit()
		end
	end
end
	
	

duration = Time.now.to_i - timer
puts "Finished in #{duration} secs."

excel.Visible = 1
excel.DisplayAlerts = 0
excel.ScreenUpdating = 1

unless File.zero?(ferror)
	puts "\nERRORS - update incomplete - see #{ferror}"
else
	File.delete(ferror)
end

# If there are errors, then Excel will have been left running with the files open. 
# This saves and shuts them down without asking the user.
excel.Quit()
excel.DisplayAlerts = 1
