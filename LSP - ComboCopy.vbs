$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Combo Copy
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

'Write load combination defintiions to file for use with LSP Combo Paste command.

'collections
dim loadsets		'total loadings in model
dim copy_ids		'ids of loads to copy
dim names		'name of load variables

'objects
dim copy_id		'id of load being copied
dim loadset		'load being copied
dim value		'value of variable to be copied

'storage
dim loadset_type	'loading type being copied
dim loadset_name	'name of loading being copied
dim values		'string of values to be copied
dim copy_string		'concercated copying string

'filesytem objects
dim filesystem		'file system access
dim text_file		'text file to save copying data

'options
dim copy_ids_string	'ids to copy as a delimited string
dim copy_ids_default	'default ids
dim file_path		'file path to save copied data

'script level defaults and options
copy_ids_default = "All"
file_path = GetSystemString("SCRIPTS") & "LSP-User\LSP - ComboCopy.txt"	
copy_id_string = inputbox("Loadset ids to copy (1;2...)/All", "Combo Copy", copy_ids_default)

'provide "all" functionality
if copy_id_string = "All" then
	loadsets = db.getLoadsets("Basic Combination", "All")
	redim copy_ids(ubound(loadsets))
	for index = 0 to ubound(loadsets)
		copy_ids(index) = loadsets(index).getID()
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
	set loadset = database.getLoadset("Basic Combination", clng(copy_id))

	'reset memory
	copy_string = ""
	values =  ""
	
	'get loading information
	loadset_type = loadset.getTypeCode()	
	loadset_name = loadset.getName()
	
	select case loadset_type
	
		case 2 'basic combination

			'dump header
			copy_string = loadset_type & ";" & loadset_name & ";" & copy_id
			text_file.WriteLine(copy_string)
			
			'factors
			values = ""
			for each value in loadset.getFactors()
				values = values & ";" & value
			next
			copy_string = "|-#F" & values
			text_file.WriteLine(copy_string)

			'loadcases
			values = ""
			for each value in loadset.getLoadcaseIDs()
				values = values & ";" & database.getLoadset(value).getName()
			next
			copy_string = "|-#L" & values
			text_file.WriteLine(copy_string)						
			
			'report
			call textwin.writeLine("Copied basic combination " & copy_id & ": " & loadset_name)	
			
		case else 	
			call textwin.writeLine("Skiped loadset " & copy_id & ": " & loadset_name)
	
	end select
	
	
next

'clean up
text_file.close