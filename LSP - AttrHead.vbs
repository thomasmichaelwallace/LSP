$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Create Attribute Import Header
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

'Prepares a new LSP attribute import file with header names/types/descriptions.

dim attr_type		'attribute type
dim attr_id		'attribute unique id
dim template		'attribute to build header

dim proper_type		'proper lusas type
dim proper_subtype	'proper lusas sub-type

dim proper_line		'formed line for proper type header
dim name_line		'formed line of value names
dim type_line		'formed line of value types
dim desc_line		'formed line of value descriptions
dim value_line		'formed line of example values

dim names		'value names
dim i			'counter for current name

dim name		'current value name
dim vbtype		'current value type
dim desc		'current value description
dim value		'current value example value

dim file_path		'header file path
dim filesystem		'file system connection
dim text_file		'header file

'fetch template by type/id
attr_type = inputbox("Attribute Type", "Attr Header", "Support")
attr_id = inputbox("Base Attribute ID", "Attr Header", "1")
set template = db.getAttributes(attr_type, attr_id)(0)

'form proper identifier
proper_type = template.getAttributeType()
proper_subtype = template.getSubType()
proper_line = proper_type & "," & proper_subtype

'prepare headers
name_line = "Name"
type_line = "String"
desc_line = "Attr. Name"
value_line = "[TEMPLATE]"

'switch error handling off to cope with complex return values
on error resume next

'itterate through value names
names = template.getValueNames()
for i = lbound(names) to ubound(names)		
	name = names(i)
	
	'fetch additional information
	vbtype = template.getValueType(name)
	desc = template.getValueDescription(name)
	value = "[COMPLEX]"
	value = template.getValue(name)
	
	'populate headers
	name_line = name_line & "," & name
	type_line = type_line & "," & vbtype
	desc_line = desc_line & "," & desc
	value_line = value_line & "," & value

next

'switch error handling back on
on error goto 0

'open file
file_path = "attr_import.csv"
set filesystem = CreateObject("Scripting.FileSystemObject")
if not filesystem.FileExists(file_path) then
	set text_file = filesystem.CreateTextFile(file_path)
	text_file.close
	set text_file = nothing
end if	
set text_file = filesystem.OpenTextFile(file_path, 2)

'write file
text_file.writeLine(proper_line)
text_file.writeLine(name_line)
text_file.writeLine(type_line)
text_file.writeLine(desc_line)
text_file.writeLine(value_line)

'close file
set text_file = nothing
call textwin.writeLine("Created header file: " & file_path)