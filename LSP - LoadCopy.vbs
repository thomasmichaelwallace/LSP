$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Load Copy
'Copyright (C) 2010-2013 Thomas Michael Wallace <http://www.thomasmichaelwallace.co.uk>

' This file is part of the LSP.

'    The LSP is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    The LSP is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with The LSP.  If not, see <http://www.gnu.org/licenses/>.

'Write load attribute definitions to file for pasting using the LSP Load Paste command.

'collections
dim loadings			'total loadings in model
dim copy_ids			'ids of loads to copy
dim names			'name of load variables

'objects
dim copy_id			'id of load being copied
dim loading			'load being copied
dim name			'name of current variable being copied
dim value			'value of variable to be copied
dim pos				'formated position coordinate array

'counters
dim rows			'rows of additional information
dim row				'current row

'storage
dim loading_type		'loading type being copied
dim loading_name		'name of loading being copied
dim values			'string of values to be copied
dim copy_string			'concercated copying string

'filesytem objects
dim filesystem			'file system access
dim text_file			'text file to save copying data

'options
dim copy_ids_string		'ids to copy as a delimited string
dim copy_ids_default		'default ids
dim file_path			'file path to save copied data

'script level defaults and options
copy_ids_default = "All"
file_path = GetSystemString("SCRIPTS") & "LSP\LSP - LoadCopy.txt"
copy_id_string = inputbox("Load ids to copy (1;2...)/All", "Load Copy", copy_ids_default)

'provide "all" functionality
if copy_id_string = "All" then
	loadings = db.getAttributes("Loading")
	redim copy_ids(ubound(loadings))
	for index = 0 to ubound(loadings)
		copy_ids(index) = loadings(index).getID()
	next
else
	copy_ids = split(copy_id_string, ";")
end if

'prepare file
set filesystem = CreateObject("Scripting.FileSystemObject")
if not filesystem.FileExists(file_path) then
	set text_file = filesystem.CreateTextFile(file_path)
	text_file.close
	set text_file = nothing
End If 
set text_file = filesystem.OpenTextFile (file_path, 2, true)

'get loading to copy
for each copy_id in copy_ids
	set loading = database.getAttribute("Loading", clng(copy_id))

	'reset memory
	copy_string = ""
	values =  ""
	
	'get loading information
	loading_type = loading.getSubType()	
	loading_name = loading.getName()
	names = loading.getValueNames()
	
	select case loading_type
	
		case "Discrete Point Load"
		
			'write basic values
			values = loading.getValue("dirType")			
			for index = 0 to 2
				value = loading.getValue("pDir")(index)
				values = values & ";" & value
			next			
			values = values & ";" & loading.getValue("nGridX")
			values = values & ";" & loading.getValue("nGridY")
			
			'dump header
			rows = loading.countRows("P") - 1
			values = values & ";" & rows			
			copy_string = loading_type & ";" & loading_name & ";" & values
			text_file.WriteLine(copy_string)
			
			for row = 0 to rows				
				values = ""
				
				'cope with bad return types
				if IsArray(loading.getValue("pos", row)(0)) then
					pos = loading.getValue("pos", row)(0)
				else
					pos = loading.getValue("pos", row)
				end if
				
				'fetch individual load entries
				for each value in pos
					values = values & ";" & value
				next
				value = loading.getValue("P", row)
				values = values & ";" & value
				
				copy_string = "|-" & row & values
				text_file.WriteLine(copy_string)				
			next

		case "Discrete Patch Load"
		
			'write basic values
			values = loading.getValue("patchType")			
			values = values & ";" & loading.getValue("sweptAngle")
			values = values & ";" & loading.getValue("xDivisions")
			values = values & ";" & loading.getValue("yDivisions")			
			values = values & ";" & loading.getValue("dirType")			
			for index = 0 to 2
				value = loading.getValue("pDir")(index)
				values = values & ";" & value
			next			
			values = values & ";" & loading.getValue("nGridX")
			values = values & ";" & loading.getValue("nGridY")
			
			'dump header
			rows = loading.countRows("P") - 1
			values = values & ";" & rows			
			copy_string = loading_type & ";" & loading_name & ";" & values
			text_file.WriteLine(copy_string)
			
			for row = 0 to rows				
				values = ""
				
				'cope with bad return types
				if IsArray(loading.getValue("pos", row)(0)) then
					pos = loading.getValue("pos", row)(0)
				else
					pos = loading.getValue("pos", row)
				end if
				
				'fetch individual load entries
				for each value in pos
					values = values & ";" & value
				next
				value = loading.getValue("P", row)
				values = values & ";" & value
				
				copy_string = "|-" & row & values
				text_file.WriteLine(copy_string)				
			next

		case "Discrete Compound Load"
		
			'dump header
			rows = loading.countLoading() - 1
			copy_string = loading_type & ";" & loading_name & ";" & rows
			text_file.WriteLine(copy_string)				
			
			for row = 0 to rows
				values = ""
				
				'fetch default arguments
				values = loading.getLoading(row).getName()				
				for each value in loading.getOffsetCoordinates(row)
					values = values & ";" & value
				next			
				
				'fetch optional arguments
				if loading.hasTransformation(row) then					
					values = values & ";" & loading.getTransformation(row).getName()
				else
					values = values & ";" & "__Nothing__"
				end if
			
				copy_string = "|-" & row & ";" & values
				text_file.WriteLine(copy_string)							
			next
			
		case else 
	
			'copy all defining variables
			for each name in names
				value = loading.getValue(name)
				values = values & ";" & value
			next		
	
			'save copied information
			copy_string = loading_type & ";" & loading_name & values		
			text_file.WriteLine(copy_string)		
	
	end select
	
	'report
	call textwin.writeLine("Copied loading " & copy_id & ": " & loading_name)	
	
next

'clean up
text_file.close