###
# This program reads the JSON specification files and creates the Markdown files (minus the examples). 
# Location: https://github.com/sumurthy/Js-Spec-Gen
###
require 'pathname'
require 'logger'
require 'json'
require 'FileUtils'

module SpecMaker

	# Initialize 
	NEWLINE = "\n"
	JS_SOURCE_FILES = "jsonFiles/source"		
	SEGMENT_END = '*/'
	FIRST_LINE = 0
	SECOND_LINE = 1
	NAMESPACE = 'Office.context.'

	# Conversion to specification 
	def self.convert_to_json (js_lines=[])
		js_lines.each_with_index do |line, i|
			# Skip first line
			next if i == FIRST_LINE
			line_parts = line.split(" ")

			if i == SECOND_LINE
				if line_parts[1] == '@namespace'
					# Object found
				end
			end

			if line_parts[0] == SEGMENT_END
				# End of segment
			end


		end
	end

	# Main loop. 
	processed_files = 0
	Dir.foreach(JS_SOURCE_FILES) do |item|
		next if item == '.' or item == '..'
		fullpath = JS_SOURCE_FILES + '/' + item.downcase

		if File.file?(fullpath)
			convert_to_json File.readlines(fullpath)
			processed_files = processed_files + 1
		end
	end
	
	# Write the README output file. 
	outfile = MARKDOWN_OUTPUT_FOLDER + '$changes.md'
	file=File.new(outfile,'w')
	@changes.each do |line|
		file.write line
	end	
	puts ""
	puts "*** OK. Processed #{processed_files} input files. Check #{File.expand_path(LOG_FOLDER)} folder for results. ***"
end