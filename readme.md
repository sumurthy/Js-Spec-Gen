# Office JavsScript Spec Maker

This tool is designed to generate Office JavaScript API specification docs based on internal definition/design. 

Running this tool involves 3 things: 

1. Setup the input data as described below. 
2. Run intermediary JSON file creation step (`genJsonFromMetadata.rb` script). 
3. Run the markdown creation step (`genMarkdownFromJSON.rb` script), which produces the final API spec files. You can copy these files into final destination such as another repository. 

## Pre-requisites

Ruby interpreter. Version 2.1+ 

## Input data 

### Metadata definition file

This is an internal metadata defintion file that describes the API structure. This file contains all of the objects, proeprties, relationships, methods, parameters, etc. along with their description. Additional information such as version supported in, etc. are also provided. 

##### File Location

This file is not part of the repository since it contains some internal data. 
The location of this file is defined in `genJsonFromMetadata.rb` script file. The path is relative to this script file. If you wish to change the location, do update the line below prior to running the script. 

`METADATA_FILE_SOURCE = '../../data/WdJscomApi.cs'`

**Please do not make this part of the repository as it doesn't belong here**

### Code snippet (example) files

For each of the resources in your object model, create a new file in the specificed folder and add example code that goes with each of the method. You can also define getter and setter examples. 

##### formatting rules for adding example code

* Name the file as "object name".md. Example `workbook.md`.
* For each method, begin with three `#` symbols and include method signature. Example `### getNotebookById(id: string)` 
* Under this line, define the example you wish to include. 
* For getter and setter: begin with three `#` symbols and follow with the text `getter` and `setter` or `getter or setter` depending on what the object supports.
* Between code blocks or within a code snippets - DO NOT INCLUDE # symbol for formatting purpose. This will throw off the script.

When the final markdown spec file is created for the resource, these code snippets get included underneath method definition. Getter and setter code snippets are added at the end of the spec file. 

**Note: Code snippet is an optional information; though highly recommended. If you don't include a code snippet file, the script shows a warning when you run it.**

##### File Location

The location of this folder is at `Js-Spec-Gen/api-examples-to-merge`. Add one file per resource. 

## Output markdown spec

The final output files are generated in markdown format. 

##### File Location

The location of final spec files are at `Js-Spec-Gen/markdown`. It includes one file per object. 

## Run steps

1. Setup the input data 
2. change directory to `Js-Spec-Gen/scripts` folder
2. Run Json intermediary file creation step: `ruby genJsonFromMetadata.rb`
3. Once the above step completes, run the markdown creation step. `ruby genMarkdownFromJSON.rb`
4. Find your markdown files in the output folder. 

## Next version of this tool.

This tool is being enhanced to combine all steps into one sript and provider better support for changing output format through settings. For the time being, please ignore the `v2` folder under scripts and also ignore the `config` folder. 
