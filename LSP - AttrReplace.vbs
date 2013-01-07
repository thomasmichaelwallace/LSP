$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Bulk Attribute Find and Replace
'	Copyright (C) 2010-2012 Thomas Michael Wallace <http://www.thomasmichaelwallace.co.uk>

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

'Undertakes a bulk replacement of attribute asignments throughout the model.

'objects and collections
dim attr_finds			'array of attributes to find
dim attr_replaces		'array of attributes to replace found with
dim attr_find				'attribute to find
dim attr_replace		'attribute as replacement

'counters
dim indicies				'total attributes to find
dim index						'current attribute index

'user input
dim mode_string			'attribute type to perform find/replace within
dim find_string			'string array of attribute ids to find
dim replace_string	'string array of attribute ids as replacements

'get user options
mode_string = cstr(Inputbox("Attribute type (Load, Geometric...)", "Attribute Replace", "Material"))
find_string = cstr(Inputbox("Find attribute ids (1;2;...)", "Attribute Replace", "6;7;8;9;10;11"))
replace_string = cstr(Inputbox("Replacement attribute ids (1;2;...)", "Attribute Replace", "26;27;28;29;30;31"))

'convert arrays
attr_finds = split(find_string, ";")
attr_replaces = split(replace_string, ";")
indicies = ubound(attr_finds)

for index = 0 to indicies

	'garbage collect, and establish model
	call selection.remove("all")
	set attr_find = nothing
	call assignment.setAllDefaults()
	call assignment.setLoadsetOff()
	
	'find all visible assignments
	set attr_find = database.getAttribute(mode_string, clng(attr_finds(index)))
	call selection.add(attr_find)	
	
	'replace all visible assignments
	set attr_replace = database.getAttribute(mode_string, clng(attr_replaces(index)))
	call attr_replace.assignTo(selection, assignment)

next