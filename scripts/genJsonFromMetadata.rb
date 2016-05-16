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
	IGNORE_KEYS = %w[@alias]
	COLLECTION_FLAGS = %w[Array array []]
	ITEM_TYPES = %w[appointment.js message.js]
	REQ_FLAG = '@since'
	PERMISSION_FLAG = '@permission'
	READMODE_FLAG = '@readmode'
	COMPOSEMODE_FLAG = '@composemode'
	SUMMARY_FLAG = '@summary'
	DESCRIPTION_FLAG = '@desc'
	METHOD_FLAG = '@method'
	PROPERTY_FLAG = '@member'
	ENUM_FLAG = '@enum'
	EXAMPLE_FLAG = '@example'
	MEMBEROF_FLAG = '@memberof'
	CALLBACK_FLAG = '@standardcallback'
	PARAM_FLAG = '@param'
	PARAM2_LEVEL_FLAG = '@param2'
	PARAM3_LEVEL_FLAG = '@param3'
	RETURNTYPE_FLAG = '@return'
	STD_TEXT = "When the method completes, the function passed in the callback parameter is called with a single parameter, asyncResult, which is an AsyncResult object. For more information, see Using asynchronous methods. "
	@json_object = nil
	SIMPLETYPES = %w[int string object object[][] object[] Date double float bool number void]

	# Read config and json_struct files 

	CONFIG = "../config/config.json"
	@config = JSON.parse(File.read(CONFIG, :encoding => 'UTF-8'), {:symbolize_names => true})

	JSTRUCT = "../config/json_structure.json"
	@jstruct = JSON.parse(File.read(JSTRUCT, :encoding => 'UTF-8'), {:symbolize_names => true})

	puts "....Starting run for the app #{@config[:app]}"
	puts

	#


	def self.deep_copy(o)
		Marshal.load(Marshal.dump(o))
	end

	def self.encode(arr=[])
		return Base64.encode64(arr.join('|'))
	end
    
    def self.is_optional(text="")
    	
		if text.include?('{')
		    text = text.scan(/{(.*?)}/)[0].join
            (text[0] == '?') ? (return true) : (return false)
		else
			return false
		end    
    end 

	def self.get_type(text="")
		if text.include?('{')
		    text = text.scan(/{(.*?)}/)[0].join
            text[0] = '' if text[0] == '?'
            return text
		else
			return nil
		end
	end

	def self.get_name(name="")

		if name[0] == '['
			name[0] = ''
			name[-1] = ''
		end
		return name
	end

	def self.clean_desc(desc="")
		
		return desc
	end


	# Conversion to specification 
	def self.convert_to_json (js_lines=[])
		in_object, read_mode, compose_mode, in_desc, in_example, in_method, in_prop = false, false, false, false, false, false, false
		s_callback_tag, is_blank, in_xmode, return_nullable, in_enum  = false, false, false, false, false
		key, comment, val, req_ver, permission, example_caption, after_comment_text, return_type = '', '', '', '', '', '', '', ''
		summary = ""
		member_of = nil
		eg_arr = []
		desc_arr = []				
		xmode_arr = []				
		prop_copy, method_copy, param_copy, enums_copy, enum_copy = nil, nil, nil, nil, nil
		enum_string = ""


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
			n = line_parts.length
			comment = line_parts[0]

			after_comment_text = line_parts[1..-1].join(' ') if n > 0
			key = line_parts[1]
			val = line_parts[2]
			datatype = line_parts[3..-1].join(' ') if n > 2
			rest_text = line_parts[2..-1].join(' ') if n > 1
			param_desc = line_parts[4..-1].join(' ') if n > 3
			if (i+1) < js_lines.length
				next_line = js_lines[i+1]
			else
				next_line = 'EOF'
			end

			next if n == 0
			next if IGNORE_KEYS.include?(key)
			is_blank = (comment == '*') ? true : false

			## 
			# End of the segment.
			###

			if comment == SEGMENT_END	
				eg_arr.push "```\n" if in_example
				# Object 
				if in_object
					@json_object[:description] =  summary
					@json_object[:longDesc] = encode (desc_arr + xmode_arr + eg_arr)

					puts desc_arr
					@json_object[:reqSet].push req_ver
					(@json_object[:modes].push "Read") if read_mode
					(@json_object[:modes].push "Compose") if compose_mode
					@json_object[:namespace] = member_of
					@json_object[:minPermission] = permission
					in_object = false
				end

				# Method
				if in_method
					method_copy[:description] = summary 
					method_copy[:longDesc] = encode desc_arr
					method_copy[:reqSet].push req_ver
					(method_copy[:modes].push "Read") if read_mode
					(method_copy[:modes].push "Compose") if compose_mode
					method_copy[:codeSnippet] = encode eg_arr
					method_copy[:minPermission] = permission
					method_copy[:returnType] = return_type
					
					if COLLECTION_FLAGS.any? { |word| method_copy[:returnType].include?(word) }
						method_copy[:isCollection] = true 
					end
					method_copy[:returnNullable] = return_nullable
					@json_object[:methods].push method_copy
					in_method = false
				end

				# Property
				if in_prop
					prop_copy[:description] = summary 
					if s_callback_tag
						prop_copy[:description] = STD_TEXT + prop_copy[:description]
					end
					prop_copy[:longDesc] = encode desc_arr
					prop_copy[:reqSet].push req_ver
					(prop_copy[:modes].push "Read") if read_mode
					(prop_copy[:modes].push "Compose") if compose_mode
					prop_copy[:minPermission] = permission				
					@json_object[:properties].push prop_copy
					in_prop = false
				end

				# Enum 
				if in_enum
					enums_copy[:description] = summary 
					enums_copy[:reqSet].push req_ver
					(enums_copy[:modes].push "Read") if read_mode
					(enums_copy[:modes].push "Compose") if compose_mode
					###
					# Since Ennums are defined outside of comments, we need a way to extract the enum + its comments
					# Read ahead until all enum lines read
					##
					if next_line.include?('var') && next_line.include?('=') && next_line.include?('{')
						k = i + 1
						loop do 				
						  # Until the Enum definition ends, concat and eat new lines. 
						  # Replace /** with <start> and */ with <end> to help extract the comment.
						  enum_string = enum_string + js_lines[k].chomp.gsub('/**','<start>').gsub('*/','<end>')
						  break if ((k + 1) == js_lines.length || js_lines[k].strip == '};')
						  k = k + 1
						end								
						###
						# First get {<value>} (value between {})
						# Then split based on <start> and <end> values; discard first [0] value which one has blanks
						# Then get descriptions in one array (evens)
						# Then get key: values in another array (odds)
						# Yeah, this is crazy! 
						##
						e_arr = (get_type enum_string).split(/<start>(.*?)<end>/)[1..-1].map {|i| i.strip.gsub('"','').gsub(',','')}
						e_desc_arr = e_arr.select.with_index { |_, i| i.even? }
						e_value_arr = e_arr.select.with_index { |_, i| i.odd? }

						e_value_arr.each_with_index do |e_kv, j| 
							enum_copy = deep_copy(@jstruct[:enumInstance])
							enum_copy[:name] = e_kv.split(':')[0]
							enum_copy[:value] = e_kv.split(':')[1].strip
							enum_copy[:description] = e_desc_arr[j]
							
							enums_copy[:items].push enum_copy
						end
					end

					@json_object[:enums].push enums_copy
					in_enum = false
				end

				# End of segment resets
				in_object, read_mode, compose_mode, in_desc, in_example, in_method, in_prop = false, false, false, false, false, false
                s_callback_tag, is_blank, in_xmode, return_nullable, in_enum  = false, false, false, false, false
                key, comment, val, req_ver, permission, example_caption, after_comment_text, return_type = '', '', '', '', '', '', '', ''
				summary = ""
				member_of = nil
				eg_arr, desc_arr, xmode_arr = [], [], []
				prop_copy, method_copy, param_copy, enums_copy, enum_copy = nil, nil, nil, nil, nil
				enum_string = ""
				
				# End of segment
				next
			end

			# Check if this is an object/type
			if i == SECOND_LINE
				if OBJ_FLAGS.include? key 
					in_object = true 
					# Get the name without the namespace
					@json_object[:name] = val.split('.')[-1]
					if val.split('.').length > 0
						member_of = val
					end
				end
				next
			end

			# Method 
			if key == METHOD_FLAG
				method_copy = deep_copy(@jstruct[:method])
				method_copy[:name] = val
				in_method = true
				next
			end

			# Parameter of method
			if key == PARAM_FLAG
				param_copy = deep_copy(@jstruct[:parameter])
				param_copy[:name] = get_name val
				param_copy[:description] = clean_desc param_desc
				param_copy[:dataType] = get_type rest_text
				if (is_optional rest_text)
                	param_copy[:isRequired] = false 
                end
				method_copy[:parameters].push param_copy                					
				next
			end


			# 2nd LEVEL Parameter of method
			if key == PARAM2_LEVEL_FLAG
				param_copy = deep_copy(@jstruct[:parameter])
				param_copy[:name] = get_name val
				param_copy[:description] = clean_desc param_desc
				param_copy[:dataType] = get_type rest_text
				if (is_optional rest_text)
                	param_copy[:isRequired] = false 
                end
                # Group generic {object} details into subParams so we can make it part of the parameter's description 
                # Check if the previous param is same as current param name. If so, attach it to subParams	
				method_copy[:parameters][-1][:subParms].push param_copy                	
				next
			end

			# 3RD LEVEL Parameter of method
			if key == PARAM3_LEVEL_FLAG
				param_copy = deep_copy(@jstruct[:parameter])
				param_copy[:name] = get_name val
				param_copy[:description] = clean_desc param_desc
				param_copy[:dataType] = get_type rest_text
				if (is_optional rest_text)
                	param_copy[:isRequired] = false 
                end
                # Group generic {object} details into subParams so we can make it part of the parameter's description 
                # Check if the previous param is same as current param name. If so, attach it to subParams	
				method_copy[:parameters][-1][:subParms][-1][:subParms].push param_copy                	
				next
			end

			# Property 
			if key == PROPERTY_FLAG
				prop_copy = deep_copy(@jstruct[:property])
				prop_copy[:name] = val
				prop_copy[:dataType] = (get_type rest_text).split('|').map {|s| s.strip}
				if COLLECTION_FLAGS.any? { |word| prop_copy[:dataType].to_s.include?(word) }
					prop_copy[:isCollection] = true 
				end					
             	prop_copy[:isNullable] = is_optional rest_text
				in_prop = true 
				next
			end

			if key == ENUM_FLAG
				enums_copy = deep_copy(@jstruct[:enums])
				enums_copy[:name] = @json_object[:namespace] + '.' + val
				enums_copy[:dataType] = 'string'
				in_enum = true
				next
			end


			if key == DESCRIPTION_FLAG
				in_desc = true
				desc_arr.push(clean_desc rest_text) # add the description line
				next
			end

			if key == CALLBACK_FLAG
				s_callback_tag = true
				next
			end

			if key == MEMBEROF_FLAG
				member_of = val.strip
			end

			if key == EXAMPLE_FLAG 
				in_example = true
				eg_arr = []
				eg_arr.push "##### Example \n"
				in_desc = false
				if rest_text.include? '<caption>'
					rest_text.gsub!('<caption>','')
					rest_text.gsub!('</caption>','')
					eg_arr.push "\n```js"
					eg_arr.push rest_text
				else
					eg_arr.push rest_text # Caption of example
					eg_arr.push "\n```js"
				end
				next
			end

			if key == READMODE_FLAG 
				read_mode, in_xmode = true, true
				# Add to xmode array if there is any comment for compose, read modes.
				if n > 2
					xmode_arr.push ("##### Read mode \n" + clean_desc(rest_text)) 
				end
				next
			end

			if key == COMPOSEMODE_FLAG 
				compose_mode, in_xmode = true, true
				# Add to xmode array if there is any comment for compose, read modes.
				if n > 2
					xmode_arr.push ("##### Compose mode \n" + clean_desc(rest_text))
				end
				next
			end
			if key.to_s.length > 0 && key.start_with?('@')
				in_xmode = false 
			end
			# if in_desc or in_xmode
			# 	desc_arr.push clean_desc after_comment_text
			# end

			if in_xmode
				xmode_arr.push clean_desc after_comment_text
			end

			if key == RETURNTYPE_FLAG
				return_type = get_type rest_text
				if (rest_text.to_s.length > 2) &&  (rest_text[0] == '?' || rest_text[1] == '?')
					return_nullable = true
				end
				next
			end

			if in_example
				eg_arr.push after_comment_text
				next
			end

			req_ver = val if key == REQ_FLAG 
			permission = val if key == PERMISSION_FLAG
			if key == SUMMARY_FLAG
				(summary = clean_desc rest_text) 			
			end
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

		next if item != 'item.js'

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
# 4. Code snippets in the @readmode and @composemode don't have ``` block and there is no way to know that
 # * @readmode The `subject` property returns a string. Use the [`normalizedSubject`]{@link Office.context.mailbox.item#normalizedSubject} property to get the subject minus any leading prefixes such as `RE:` and `FW:`.
 # * 
 # *     var subject = Office.context.mailbox.item.subject;
 # * 
 # * @composemode The `subject` property returns a `Subject` object that provides methods to get and set the subject. 
 # * 
 # *     Office.context.mailbox.item.subject.getAsync(callback);
 # * 
 # *     function callback(asyncResult) {
 # *       var subject = asyncResult.value;
 # *     }
 # *
# 5. JJ change Array.<( to Array<(
# 6. Inconsistent casing all over the place
# 7. displayReplyForm -- has optional parameters.. This is really hard to read.	See my email. 
# 8. 
#
#
#
#
#
#
#
#
#
#
#
#####
