###
# This program reads the .CS metadata file and generates the JSON representaiton as separate files for each resources.
# Location: https://github.com/sumurthy/Js-Spec-Gen
####

module SpecMaker

require 'logger'
require 'json'
require 'FileUtils'

# Log file
	LOG_FOLDER = '../../../logs'
	Dir.mkdir(LOG_FOLDER) unless File.exists?(LOG_FOLDER)

	if File.exists?("#{LOG_FOLDER}/#{$PROGRAM_NAME.chomp('.rb')}.txt")
		File.delete("#{LOG_FOLDER}/#{$PROGRAM_NAME.chomp('.rb')}.txt")
	end
	@logger = Logger.new("#{LOG_FOLDER}/#{$PROGRAM_NAME.chomp('.rb')}.txt")
	@logger.level = Logger::DEBUG
# End log file


@processed_files = 0
@json_files_created = 0
METADATA_FILE_SOURCE = '../../data/VisioApi.cs'
ENUMS = 'jsonFiles/settings/enums.json'
LOADMETHOD = 'jsonFiles/settings/loadMethod.json'
JSONOUTPUT_FOLDER = 'jsonFiles/source/'
OBJECTKEYS = 'jsonFiles/settings/objectkeys.json'

@csarray = []
@csarray_pure = []
@csarray_out = []
@current_object = ''
@resource = {}
@jsonHash = {}
# Master json containers
@json_hash = {}
@json_array = []
@json_object = {}
@json_object[:name] = ''
@json_object[:description] = ''
@json_object[:isCollection] = false
@json_object[:collectionOf] = nil
@json_object[:restPath] = []
@json_object[:info] = {}
@json_object[:info][:version] = '1.1'
@json_object[:info][:reqSet] = '1.1'
@json_object[:info][:addinTypes] = ["Visio"]
@json_object[:info][:nameSpace] = "Visio"
@json_object[:info][:addinHosts] = ["Task pane"]
@json_object[:info][:title] = 'Office JavaScript Add-in API'
@json_object[:info][:description] = 'Office JavaScript Add-in API'
@json_object[:properties] = []
@json_object[:methods] = []

# Sub json containers
# Method = Struct.new(:name, :returnType, :description, :parameters, :syntax, :vbaInfo, :signature)
# Property = Struct.new(:name, :dataType, :description, :isReadOnly, :enumNameJs, :isCollection, :vbaInfo, :possibleValues, :isRelationship)
Method = Struct.new(:name, :returnType, :description, :syntax, :signature, :restfulName, :notes, :httpSuccessResponse, :parameters, :reqSet)
Property = Struct.new(:name, :dataType, :description, :isReadOnly, :enumNameJs, :isCollection, :isRelationship, :reqSet, :isKey, :notes)
ParamStr = Struct.new(:name, :dataType, :isCollection, :description, :isRequired, :enumNameJs, :notes)

SIMPLETYPES = %w[int string object object[][] object[] double bool float number void]

def csarray_write (line=nil)
	@csarray_out.push line
end

###
# Load up all the known existing enums. Remove leading Visio.
##
@enumHash = {}
tempEnumHash = JSON.parse File.read(ENUMS)
@enumHash = Hash[tempEnumHash.map {|k, v| [k.gsub('Visio.',''), v] }]

###
# Load the "load()" method to be added to all items that have at least one property.
##
@loadMethodHash = {}
@loadMethodHash = JSON.parse(File.read(LOADMETHOD), {:symbolize_names => true})

###
# Load the keys of the collections. Righ now, there is no way to identify keys from the metadata file.
##

@objectKeyHash = {}
@objectKeyHash = JSON.parse(File.read(OBJECTKEYS))


###
# Read the file & create a transit file by removing existing comments from the .CS File.
##
@csarray_pure = File.readlines(METADATA_FILE_SOURCE)
handle_getItem = ''

FileUtils.rm Dir.glob(JSONOUTPUT_FOLDER + '/*')

@csarray_pure.each do |line|
	if line.include?('this[')
		handle_getItem = line
		#handle_getItem = handle_getItem[0,handle_getItem.index('{')]
		#handle_getItem = handle_getItem.gsub('this[','getItem(').gsub(']',');')
		handle_getItem = handle_getItem.gsub('this[','getItem(')
		handle_getItem = handle_getItem.rpartition(']').first + ')' + ';'
		line = handle_getItem + "\n"
	end
	@csarray.push line
end

def self.uncapitalize (str="")
	if str.length > 0
		str[0, 1].downcase + str[1..-1]
	else
		str
	end
end


###
# Forward Pass: Write to the output array
##

@json_objects = nil
in_region = false
member_summary = ''
member_ahead = false
readOnly = false
property_array = []
method_array = []
parm_array = []
parm_array_metadata = []
enumName = ''
parm_hash_array = []
restfulName = nil
req_set = '1.1'
object_req_set = ''


def self.parser input

  input.split(',').each_with_object({}) do |s,output|
          k,v = s.split('=')
          output[v.chomp(')').gsub('"','').strip]=k
  end
end

@csarray.each_with_index do |line, i|

	if line.strip.start_with?('[ApiSet(')
		#req_set = line.split('=')[1].gsub(']','').gsub(')','').strip
		req_set = (parser line[/\(.*?\)/]).keys.join(', ')
	end

	## For new object, load its resource and fill the description
	if line.include?('public interface') || line.include?('public struct')
		# Get the third Visio
		@json_object[:name] = line.split.first(3).join(' ').split.last(1).join(' ').gsub(':','')
		@json_object[:info][:reqSet] = req_set
		object_req_set = req_set
		puts "*-> Processing #{@json_object[:name]}"
		if @json_object[:name].include?('Collection')
			@json_object[:isCollection] = true
			# Define what it is a collection of. Extract object name between < >
			# public interface ChartSeriesCollection : IEnumerable<ChartSeries>
			@json_object[:collectionOf] = line.split('<')[1].chomp("\n").chomp(">")
		else
			@json_object[:isCollection] = false
			@json_object[:collectionOf] = nil
		end
		in_region = true
		@member_summary = ''
	end

	if line.strip.start_with?('[ClientCallableComMember', '[ClientCallableOperation')
		member_ahead = true
		restfulName = nil
		# Extract the Restfull name, which usually strips off get prefix from method names.
		if line.include?('RESTfulName')
			lineSplitArray = line.split(',')
			restNameIndex = lineSplitArray.index {|w| w.include?('RESTfulName')}
			restfulName = lineSplitArray[restNameIndex].split('=')[1].gsub('"','').gsub(')','').gsub(']','').strip

		end
	end

	# This signals end of an object. Time to write stuff to file.
	if in_region && line.start_with?("\t}")
		# If this is a collection, add the 'items' property as that is not listed in the .CS file for some reason!
		if 	@json_object[:isCollection] == true
			prop_name = 'items'
			isRel = false

			objectMadeOf = uncapitalize (@json_object[:name][0, @json_object[:name].index('Collection')])
			makeDesc = "A collection of " + objectMadeOf + " objects."
			readOnly = true
			enumName = nil
			isItCollection = true
			itemReturnType = @json_object[:name][0,@json_object[:name].index('Collection')] + '[]'

			#Because chartpoints is not named correctly, it needs an override.
			if itemReturnType == 'ChartPoints[]'
				itemReturnType = 'ChartPoint[]'
			end

			property = Property.new(prop_name, itemReturnType, makeDesc, readOnly, enumName, isItCollection, isRel, object_req_set, nil )
			property_array.push property.to_h
			property = nil
		end


		# Write the buffer
		if property_array.length == 0
			@json_object[:properties] = nil
		else
			@json_object[:properties] = property_array
		end

		# Add the .load method to the method array.
		if property_array.length == 0
			if method_array.length == 0
				@json_object[:methods] = nil
			else
				@json_object[:methods] = method_array
			end
		else
			method_array.push @loadMethodHash
			@json_object[:methods] = method_array
		end

		# Seed the restPath if its the parent object (workbook)
		if @json_object[:name] == 'Workbook'
			@json_object[:restPath] = ['/workbook']
		else
			@json_object[:restPath] = nil
		end


		File.open("#{JSONOUTPUT_FOLDER}#{(@json_object[:name]).downcase}.json", "w") do |f|
			f.write(JSON.pretty_generate @json_object)
		end
		@json_files_created = @json_files_created + 1
		# Reset the variables.
		in_region = false
		# Bug fix. Caused issue with Visio API.
		member_ahead = false
		# End bug fix
		parm_hash_array = []
		parm_array = []
		property_array = []
		method_array = []
	end

	# Load up either the object or member summary.
	if line.include?('/// <summary>') &&
		if !in_region
			@json_object[:description] = @csarray[i+1].delete!('///').strip
		else
			member_summary = @csarray[i+1].delete!('///').strip

			if member_summary.index('See Visio.') != nil
				enumName = member_summary[member_summary.index('See Visio.')..-1].split[1]
				member_summary = member_summary[0,member_summary.index('See Visio.')-1]
				enumName = enumName.chomp('.')
			elsif member_summary.index('Refer to Visio.') != nil
				enumName = member_summary[member_summary.index('Refer to Visio.')..-1].split[2]
				member_summary = member_summary[0,member_summary.index('Refer to Visio.')-1]
				enumName = enumName.chomp('.')
			else
				enumName = nil
			end

			# if member_summary.include?('Read-only')
			# 	readOnly = true
			# 	member_summary = member_summary.gsub('. Read-only.', '.')
			# 	member_summary = member_summary.gsub(' Read-only.', '.')
			# else
			# 	readOnly = false
			# end
		end
	end

	# if method has params, then load them and also mark the enumName.
	if line.include?('/// <param name=')
		param_summary = line.split('>')[1].gsub('</param', '')

		if param_summary.index('See Visio.') != nil
			enumName = param_summary[param_summary.index('See Visio.')..-1].split[1]
			param_summary = param_summary[0,param_summary.index('See Visio.')-1]
			enumName = enumName.chomp('.')
		elsif param_summary.index('Refer to Visio.') != nil
			enumName = param_summary[param_summary.index('Refer to Visio.')..-1].split[2]
			param_summary = param_summary[0,param_summary.index('Refer to Visio.')-1]
			enumName = enumName.chomp('.')
		else
			enumName = nil
		end

		param_name = line.split('"')[1]
		parameter = ParamStr.new(param_name, nil, false, param_summary, nil, enumName, nil)
		parm_array.push parameter

	end

	# Presence of { would indicate that it is a property or a relation
	if member_ahead && !line.include?('_') && line.include?('{')
		prop_name = line.split[1]
		prop_name = uncapitalize prop_name
		member_ahead = false

		if line.include?('Collection')
			isItCollection = true
		else
			isItCollection = false
		end
		if line.include?('set;')
		 	readOnly = false
		 	member_summary = member_summary.gsub('. Read-only.', '.')
		 	member_summary = member_summary.gsub(' Read-only.', '.')
		else
		 	readOnly = true
		 	member_summary = member_summary.chomp(' Read-only.')
		 	member_summary = member_summary.chomp('Read-only.')
		end

		proDataType = line.split[0].gsub('?','')

		if (proDataType != 'object[]') && (proDataType != 'object[][]')
			proDataType = proDataType.chomp('[][]')
			proDataType = proDataType.chomp('[]')
		end

		# If the return type is primitive or one of the enums, then it's a property. else a relation.
		if (SIMPLETYPES.include? proDataType) || (@enumHash.has_key? proDataType)
			isRel = false
		else
			isRel = true
		end
		if enumName != nil
			proDataType = 'string'
		end

		if @enumHash.has_key? proDataType
			enumName = 'Visio.' + proDataType
			proDataType = 'string'
		end

		property = Property.new(prop_name, proDataType, member_summary, readOnly, enumName, isItCollection, isRel, req_set, nil)

		property_array.push property.to_h
		property = nil
		req_set = '1.1'
	end

	# If member is a method and has param, capture its optional param and data type.
	line = line.chomp

	if member_ahead && line.include?(');')  && !line.include?('();') && !line.include?('_')

		# Capture the first part of the parameter definition inside method definition to see if it has readonly flag and also note down its data type.

		parm_array_metadata = line[line.index('(')+1, line.index(');')].chomp(');').split(',')
		opt_array = parm_array_metadata.map {|n| n.split[0]}
		opt_array.each_with_index do |metadata, j|
			if metadata.include?('Optional')
				parm_array[j][:isRequired] = false
			else
				parm_array[j][:isRequired] = true
			end
		end

		# Now we can thrash the metadata array to figure out the actual data type
		parm_array_metadata.map! {|n| n.gsub('[Optional]','')}
		parm_array_metadata.map! {|n| n.gsub('?','')}
		#parm_array_metadata.map! {|n| n.gsub(']',' ')}
		parm_array_metadata.map! {|n| n.gsub(';',' ')}

		parm_array_metadata.each_with_index do |dataType, j|
			suffix = ''
			if dataType.include?('Array')
				if dataType.include?('Array<Array')
					suffix = '[][]'
				else
					suffix = '[]'
				end
			end
			if dataType.include?('TypeScriptType')
				typeScriptDataStart = dataType.index('TypeScriptType') + 15
				typeScriptData = dataType[typeScriptDataStart..-1]
				typeScriptDataEnd = typeScriptData.index(')') - 1
				typeScriptData = typeScriptData[0..typeScriptDataEnd]
				if typeScriptData.include?('>>')
					typeScriptData = typeScriptData[typeScriptData.index('>>|')+3..-1]
				else
					if typeScriptData.include?('>')
						typeScriptData = typeScriptData[typeScriptData.index('>')+2..-1]
					end
				end
				typeScriptDataArray = typeScriptData.gsub('"','').gsub(')','').gsub('Visio.','').split('|').join(' or ')
				if suffix != ''
					parm_array[j][:dataType] = "(" + typeScriptDataArray +")" + suffix
				else
					parm_array[j][:dataType] = typeScriptDataArray + suffix
				end
			else
				parm_array[j][:dataType] = dataType.split[0] + suffix
			end

			# Add Collection string at the end instead of "[]"
			parm_array[j][:isCollection] = false
			if parm_array[j][:dataType].include?('[]')
				parm_array[j][:isCollection] = true
			end
			if parm_array[j][:dataType] == 'int'
				parm_array[j][:dataType] = 'number'
			end

			# If the enum still slips through to the data type, then overwrite and set the enum correctly.
			if @enumHash.has_key? parm_array[j][:dataType]
				parm_array[j][:enumNameJs] = 'Visio.' + parm_array[j][:dataType]
				parm_array[j][:dataType]  = 'string'
			end

			# Enum data type should be documented as strings.
			#if enumName != nil
			if parm_array[j][:enumNameJs] != nil
				parm_array[j][:dataType] = 'string'
			end
			parm_array[j][:dataType] = parm_array[j][:dataType].gsub('?', '')

		end

		parm_array.each do |parmStruct|
			parm_hash_array.push parmStruct.to_h
		end
	end

	# If it's a method, dump its informaiton
	if member_ahead && !line.include?('_') && line.include?(');')
		#Method = Struct.new(:name, :returnType, :description, :parameters)
		#@json_object[:methods] = []
		member_ahead = false
		temp = line.split[1]
		mthd_name = temp[0,temp.index('(')]
		mthd_name = uncapitalize mthd_name
		if parm_hash_array.length == 0
			parm_hash_array = nil
			signature = mthd_name + '()'
			syntax = "#{uncapitalize @json_object[:name]}Object.#{mthd_name}();"
		else
			signature = mthd_name + '('
			syntax = "#{uncapitalize @json_object[:name]}Object.#{mthd_name}("
			parm_hash_array.each_with_index do |parmhash, k|
				signature = signature + parmhash[:name] + ': ' + parmhash[:dataType].to_s
				syntax = syntax + parmhash[:name]
				if k < (parm_hash_array.length - 1)
					signature = signature + ', '
					syntax = syntax + ', '
				end
			end
			signature = signature + ')'
			signature = signature.gsub('?','')
			syntax = syntax + ');'
			syntax = syntax.gsub('?','')
		end
		line.split[0] = line.split[0].gsub('?','')

		# Finally, hanlde the restful names
		if restfulName.to_s.length == 0
			restfulName = mthd_name
			if restfulName.start_with?('get')
				restfulName = restfulName[3..-1]
			end
		end

		restfulName.slice(0,1).capitalize + restfulName.slice(1..-1)

		# Create method hash and push the values.
		method = Method.new(mthd_name, line.split[0], member_summary, syntax, signature, restfulName, nil, nil, parm_hash_array, req_set)
		method_array.push method.to_h

		# Reset the variables.
		method = nil
		req_set = '1.1'
		parm_array = []
		parm_hash_array = []
	end
end

# Reccursively add restPath

def self.add_restpath (item=nil, restPath=[], pathToWriteBack)

	@processed_files = @processed_files + 1
	jsonHash = JSON.parse(item, {:symbolize_names => true})

	@logger.debug("-----> Recursive called from, #{pathToWriteBack}, for #{jsonHash[:name]}")

	# Max of 5 paths are enough for display.
	if jsonHash[:restPath] && jsonHash[:restPath].length > 5
		return
	end

	# Assign path. If one already exists, merge and remove dups.
	jsonHash[:restPath] = jsonHash[:restPath] ? (restPath | jsonHash[:restPath]) : restPath
	#@logger.debug(".... After restpath: #{jsonHash[:restPath]}")
	#resource = uncapitalize(jsonHash[:name])

	propreties = jsonHash[:properties]
	methods = jsonHash[:methods]
	printName = jsonHash[:name]
	# Process if the resource has properties.
	if propreties

		propreties.each do |prop|
			# Process only if its a relation.
			# Avoid infinite loop by avoiding circular reference with worksheet > range > worksheet
			if prop[:isRelationship] && !((prop[:name] == 'worksheet') && (jsonHash[:name] == 'Range')) &&  \
					!((prop[:name] == 'worksheet') && (jsonHash[:name] == 'Table'))  && \
					!((prop[:name] == 'worksheet') && (jsonHash[:name] == 'Chart'))
				relFilePath = JSONOUTPUT_FOLDER + prop[:dataType].downcase + '.json'
				if File.file?(relFilePath)
					@logger.debug(".... Relation: Going recursive with #{prop[:name]} for @>#{printName}|")

					pathToSendArray = jsonHash[:restPath].map {|d| d + '/' + prop[:name].downcase }
					add_restpath File.read(relFilePath), pathToSendArray, relFilePath

				end
				# If it's a collection, add the RESTful path to it's item. Ex: /table from /tables
				if prop[:isCollection]
					collectionItem = prop[:dataType].chomp('Collection').downcase
					# Special case chartpoint because it is not named correctly.
					if collectionItem == 'chartpoints'
						collectionItem = 'chartpoint'
					end
					collectionItemFilePath = JSONOUTPUT_FOLDER + collectionItem + '.json'
					# Append the keys of collection. e.g. worksheets({id|name})
					lastSegment = (@objectKeyHash.has_key?(prop[:name].downcase)) ? ('({' + @objectKeyHash[prop[:name].downcase].join('|') + '})') : '/{undefined}'
					collectionItemRestPath = jsonHash[:restPath].map { |d| d + '/' + prop[:name].downcase + lastSegment}
					if File.file?(collectionItemFilePath)
						@logger.debug(".... Collection Item: Going recursive with #{collectionItem}")
						add_restpath File.read(collectionItemFilePath), collectionItemRestPath, collectionItemFilePath

					end
				end
			end
		end
	end
	# Now process methods to get things like range.
		# Process if the resource has properties.
	if methods
		methods.each do |method|
			# Process only if its a relation.
			methodFilePath = JSONOUTPUT_FOLDER + method[:restfulName].to_s.downcase + '.json'
			if File.file?(methodFilePath)
				@logger.debug(".... Method: Going recursive with #{method[:restfulName]}")
				parmForMethod = method[:parameters] ? '({' + method[:parameters][0][:name] + '})' : ''
				pathToSendArray = jsonHash[:restPath].map {|d| d + '/' + method[:restfulName].downcase + parmForMethod}
				add_restpath File.read(methodFilePath), pathToSendArray, methodFilePath
			end
		end
	end

	# Write the file back with the REST path.
	File.open(pathToWriteBack, "w") do |f|
		f.write(JSON.pretty_generate jsonHash)
	end
end

# Add REST Path to the resources.

	fullpath = JSONOUTPUT_FOLDER + 'workbook.json'
	if File.file?(fullpath)
		# add_restpath File.read(fullpath), ["/workbook"], fullpath
	end
puts
puts "*** Done. Created #{@json_files_created} JSON intermediary files. Next, run the markdown spec creation step. ***"
puts
#end module
end
