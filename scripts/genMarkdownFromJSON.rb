###
# This program reads the JSON specification files and creates the Markdown files (minus the examples). 
# Location: https://github.com/sumurthy/Js-Spec-Gen
###
require 'pathname'
require 'logger'
require 'json'
require 'FileUtils'
require 'base64'


module SpecMaker

	# Initialize 
	NEWLINE = "\n"
	JS_SOURCE_FILES = "../../data/outlook"		
	JSON_OUTPUT_FOLDER = "jsonFiles/source/"
	SEGMENT_END = '*/'
	FIRST_LINE = 0
	SECOND_LINE = 1
	NAMESPACE = 'Office.context.'
	OBJ_FLAGS = %w[@interface @namespace @typedef]
	IGNORE_KEYS = %w[@memberof @alias]
	ITEM_TYPES = %w[appointment.js message.js]
	REQ_FLAG = '@since'
	PERMISSION_FLAG = '@permission'
	READMODE = '@readmode'
	COMPOSEMODE = '@composemode'
	SUMMARY_FLAG = '@summary'
	DESCRIPTION_FLAG = '@desc'
	METHOD_FLAG = '@method'
	PROPERTY_FLAG = '@member'
	EXAMPLE_FLAG = '@example'
	MEMBEROF_FLAG = '@memberof'
	CALLBACK_FLAG = '@standardcallback'
	PARAM_FLAG = '@param'
	RETURNTYPE = '@return'
	STD_TEXT = "When the method completes, the function passed in the callback parameter is called with a single parameter, asyncResult, which is an AsyncResult object. For more information, see Using asynchronous methods. "
	@json_object = nil

	# Read config and json_struct files 

	CONFIG = "../config/config.json"
	@config = JSON.parse(File.read(CONFIG, :encoding => 'UTF-8'), {:symbolize_names => true})

	JSTRUCT = "../config/json_structure.json"
	@jstruct = JSON.parse(File.read(JSTRUCT, :encoding => 'UTF-8'), {:symbolize_names => true})

	puts "....Starting run for the app #{@config}"
	puts

	#


	def self.deep_copy(o)
		Marshal.load(Marshal.dump(o))
	end

	def self.encode(arr=[])
		return Base64.encode64(arr.join('|'))
	end

	def self.get_type(val="", type="")

		return type
	end

	def self.get_name(name="")

		return name
	end

	def self.clean_desc(desc="")

		return desc
	end


	# Conversion to specification 
	def self.convert_to_json (js_lines=[])
		in_object, read_mode, compose_mode, in_desc, in_example, in_method, in_prop = false, false, false, false, false, false, false
		s_callback_tag = false
		key, comment, val, req_ver, permission, example_caption, non_comment, return_type = '', '', '', '', '', '', '', ''
		summary = ""
		eg_arr = []
		desc_arr = []				
		prop_copy, method_copy, param_copy = nil, nil, nil
		js_lines.each_with_index do |line, i|
			# Skip first line
			if i == FIRST_LINE
				@json_object = deep_copy(@jstruct[:object])
				next
			end

			# Overrides
			line.gsub!('<code>','`')
			line.gsub!('</code>','`')
			#
			line_parts = line.strip.split(" ")
			comment = line_parts[0]
			non_comment = line[1..-1]
			key = line_parts[1]
			val = line_parts[2]
			type = line_parts[3]
			rest = line_parts[2..-1]
			param_desc = line_parts[4..-1]

			next if IGNORE_KEYS.include? key  


			## 
			# End of the segment.
			###
			if comment == SEGMENT_END	
				eg_arr.push "```\n" if in_example == true
				if in_object
					@json_object[:description] =  summary 
					@json_object[:longDesc] = encode (desc_arr + eg_arr)
					@json_object[:reqSet].push req_ver
					(@json_object[:modes].push "Read") if read_mode
					(@json_object[:modes].push "Compose") if compose_mode
					@json_object[:minPermission] = permission
					in_object = false
				end

				if in_method
					method_copy[:description] = summary 
					method_copy[:longDesc] = encode desc_arr
					method_copy[:reqSet].push req_ver
					(method_copy[:modes].push "Read") if read_mode
					(method_copy[:modes].push "Compose") if compose_mode
					method_copy[:codeSnippet] = encode eg_arr
					method_copy[:minPermission] = permission
					@json_object[:methods].push method_copy
					in_method = false
				end

				if in_prop
					prop_copy[:description] = summary 
					prop_copy[:longDesc] = encode desc_arr
					prop_copy[:reqSet].push req_ver
					(prop_copy[:modes].push "Read") if read_mode
					(prop_copy[:modes].push "Compose") if compose_mode
					prop_copy[:minPermission] = permission
					@json_object[:properties].push prop_copy
					in_prop = false
				end

				# End of segment resets
				in_object, read_mode, compose_mode, in_desc, in_example, in_method, in_prop = false, false, false, false, false, false, false
				s_callback_tag = false				
				key, comment, val, req_ver, permission, example_caption, non_comment, return_type = '', '', '', '', '', '', '', ''
				summary = ""
				eg_arr, desc_arr = [], []
				# End of segment
			end

			# Check if this is an object/type
			if i == SECOND_LINE
				if OBJ_FLAGS.include? key 
					in_object = true 
					# Get the name without the namespace
					@json_object[:name] = val.split('.')[-1]
				end
			end

			# Method 
			if key == METHOD_FLAG
				method_copy = deep_copy(@jstruct[:method])
				method_copy[:name] = val
				in_method = true
			end

			# Parameter of method
			if key == PARAM_FLAG
				param_copy = deep_copy(@jstruct[:parameter])
				param_copy[:name] = get_name val
				param_copy[:description] = 
				param_copy[:dataType] = get_type val, type 
				method_copy[:parameters].push param_copy
			end

			# Property 
			if key == PROPERTY_FLAG
				prop_copy = deep_copy(@jstruct[:property])
				prop_copy[:name] = val
				prop_copy[:description] = clean_desc param_desc
				if s_callback_tag
					prop_copy[:description] = STD_TEXT + prop_copy[:description]
				end
				# handle data type
				in_prop = true 
			end


			if key == DESCRIPTION_FLAG
				in_desc = true
				desc_arr.push(clean_desc rest) # add the description line
				next
			end

			if key == CALLBACK_FLAG
				s_callback_tag = true
				next
			end

			if key == EXAMPLE_FLAG 
				in_example = true
				eg_arr = []
				in_desc = false
				if rest.include? '<caption>'
					rest.gsub!('<caption>','')
					rest.gsub!('</caption>','')
					eg_arr.push "\n```js"
					eg_arr.push rest
				else
					eg_arr.push rest # Caption of example
					eg_arr.push "\n```js"
				end
				next
			end

			if in_desc
				desc_arr.push clean_desc non_comment
				next
			end

			if key == RETURNTYPE
				return_type = get_type val, type
			end

			if in_example
				eg_arr.push non_comment
				next
			end

			req_ver = val if key == REQ_FLAG 
			permission = val if key == PERMISSION_FLAG
			read_mode = true if key == READMODE
			compose_mode = true if key == COMPOSEMODE
			(summary = clean_desc val) if key == SUMMARY_FLAG

		end
	end

	# Main loop. 
	processed_files = 0
	lines = []
	
	FileUtils.rm Dir.glob(JSON_OUTPUT_FOLDER + '/*')

	Dir.foreach(JS_SOURCE_FILES) do |item|
		next if item == '.' or item == '..' or item == '.DS_Store'
		# Skip types
		next if ITEM_TYPES.include? item

		puts "** Processing #{item}"
		fullpath = JS_SOURCE_FILES + '/' + item.downcase

		if File.file?(fullpath)

			lines = File.readlines(fullpath)

			# Append sub-types of "item" at the end.
			if item == 'item.js'
				ITEM_TYPES.each do |subtype|
					fullpath = JS_SOURCE_FILES + '/' + subtype
					lines = lines + File.readlines(fullpath)
				end
			end
			# Converty to JSON
			convert_to_json lines

			# Write JSON Output Files

			File.open("#{JSON_OUTPUT_FOLDER}#{(@json_object[:name]).downcase}.json", "w") do |f|
				f.write(JSON.pretty_generate @json_object)
			end


			processed_files = processed_files + 1
		end
	end

	puts ""
	puts "*** OK. Processed #{processed_files} input files. ***"
end

#####
# todos
# 1. Handle @link; [Body.getAsync]{@linkcode Body#getAsync}
# 2. There is no indicator to arrays -- we should add that? 
# 3. Objects that have known structure should be of their own type; example event.js > @member source {Object} 
#
#
#####




