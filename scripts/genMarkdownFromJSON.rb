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
	OBJ_FLAGS = w[@interface, @namespace, @typedef]
	REQ_FLAG = '@since'
	PERMISSION_FLAG = '@permission'
	READMODE = '@readmode'
	COMPOSEMODE = '@composemode'
	SUMMARY = '@summary'
	DESCRIPTION = '@desc'
	EXAMPLE = '@example'
	MEMBEROF = '@memberof'


	obj_block, read_mode, compose_mode, in_desc, in_example = false, false, false, false, false

	key, comment, val, req_ver, permission, example_caption, non_comment = '', '', '', '', '', '', ''

	eg_arr, desc_arr = [], []
	# Conversion to specification 
	def self.convert_to_json (js_lines=[])
		js_lines.each_with_index do |line, i|
			# Skip first line
			next if i == FIRST_LINE
			next if key == MEMBEROF

			# Overrides
			line.gsub!('<code>','`')
			line.gsub!('</code>','`')
			#
			line_parts = line.strip.split(" ")
			comment = line_parts[0]
			non_comment = line[1..-1]
			key = line_parts[1]
			val = line_parts[2]
			rest = line_parts[2..-1].join(' ')

			## 
			# End of the segment.
			###
			if comment == SEGMENT_END
				eg_arr.push '```' if in_example
				obj_block ? (assign to object = req_ver) : (assign to member = req_ver)

				# End of segment resets
				obj_block, read_mode, compose_mode, in_desc, in_example = false, false, false, false, false
				key, comment, val, req_ver, permission, example_caption = '', '', '', '', '', ''
				eg_arr, desc_arr = [], []

				# End of segment
			end

			if i == SECOND_LINE
				obj_block = true if OBJ_FLAGS.include? key 
			end

			if key == DESCRIPTION
				in_desc = true
				desc_arr.push rest # add the description line
				next
			end

			if key == EXAMPLE 
				in_example = true
				in_desc = false
				if rest.include? '<caption>'
					rest.gsub!('<caption>','')
					rest.gsub!('</caption>','')
					eg_arr.push '```js' 
					eg_arr.push rest
				else
					eg_arr.push rest # Caption of example
					eg_arr.push '```js'
				end
				next
			end

			if in_desc
				if comment 
				desc_arr.push non_comment
				next
			end

			if in_example
				eg_arr.push non_comment
				next
			end

			req_ver = val if key == REQ_FLAG 
			permission = val key == PERMISSION_FLAG
			read_mode = true if key == READMODE
			compose_mode = true if key == COMPOSEMODE
			summary = val if key == SUMMARY

		end
	end

	# Main loop. 
	processed_files = 0
	Dir.foreach(JS_SOURCE_FILES) do |item|
		next if item == '.' or item == '..'
		fullpath = JS_SOURCE_FILES + '/' + item.downcase

		if File.file?(fullpath)
			convert_to_json File.readlines(fullpath)
			# Write the README output file. 
			outfile = JSON_OUTPUT_FOLDER + item.tstrip('.js')
			file=File.new(outfile,'w')
			@changes.each do |line|
				file.write line
			end				
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






