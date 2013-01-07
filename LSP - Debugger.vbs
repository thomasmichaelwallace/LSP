$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Debug Scripts
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

'Debug and development functions and modules for LSP

'manual switch board
LPI_Ls_Dir
LPI_Attr_Dir


Sub LPI_Attr_Dir()
	'__dir__ for LUSAS attributes		
	
	'variables
	Dim dir					'returned string
	Dim index				'index of names
	Dim names				'colleciton of variable names
	Dim variable		'variable type
	Dim value				'value
	Dim value_temp	'tempoaray memory
	
	'user options
	Dim parent			'attribute type
	Dim child				'attribute name
	Dim family(1) 	'defaults
	
	'list of sensible options
		'Discrete Compound Load
		'Discrete Patch Load
		'Discrete Point Load		
		'Geometry

	'set defaults for fast running
	family(0) = "Geometric"
	family(1) = "Top Chord"
	
	'get object
	parent = inputbox("Attribute Family","(__dir__)", family(0))	
	child = inputbox("Attribute Name","(__dir__)", family(1))	
	set attr = database.getAttribute(parent, child)
	
	'header row
	dir = "(__dir__)"
	call textwin.writeLine(dir)
	dir = attr.getAttributeType() & "." & attr.getSubType()
	call textwin.writeLine(dir)
	
	'scroll through namespaces
	names = attr.getValueNames()	
	for index = 0 to ubound(names)
		name = names(index)
		
		'set sensible defaults
		variable = "Unknown"
		value = "Unknown"
		
		'attempt to get information, LUSAS allows this to fail, so protect us against bad coding
		on error resume next			
			variable = attr.getValueType(name)
			value_temp = attr.getValue(name)

			'attempt to rip from complex arrays
			if IsArray(value_temp) then
				if IsArray(value_temp(0)) then
					value = value_temp(0)(0)
				else
					value = value_temp(0)
				end if
			else
			  value = value_temp
			end if
			
			'attempt to correct variable using vartype
			if variable = "Unknown" then variable = "VBType " & cstr(VarType(value_temp))
			
		'add row to data
		dir = "|-" & name & " (" & variable & ") - " & attr.getValueDescription(name) & " = " & value
		call textwin.writeLine(dir)
		
	next
	
	'respond with directory
	'msgbox dir, vbinformation, "(__dir__)"
	
End Sub

Sub LPI_Ls_Dir()
	'__dir__ for LUSAS load sets
	
	'variables
	Dim dir					'returned string
	Dim index				'index of names
	Dim names				'colleciton of variable names
	Dim variable		'variable type
	Dim value				'value
	Dim value_temp	'tempoaray memory
	
	'user options
	Dim load_name		'load name
	
	'get object
	load_name = inputbox("Loadset","(__dir__)", "ULS Combo 1 (Breaking) (Centre)(Expansion)(Up)")	
	set loadset = database.getLoadset(load_name)
	
	'header row
	dir = "(__dir__)"
	call textwin.writeLine(dir)
	dir = "Type Code: " & loadset.getTypeCode()
	call textwin.writeLine(dir)
	
	'scroll through namespaces
	names = loadset.getValueNames()
	for index = 0 to ubound(names)
		name = names(index)
		
		'set sensible defaults
		variable = "Unknown"
		value = "Unknown"
		
		'attempt to get information, LUSAS allows this to fail, so protect us against bad coding
		on error resume next			
			variable = loadset.getValueType(name)
			value_temp = loadset.getValue(name)

			'attempt to rip from complex arrays
			if IsArray(value_temp) then
				if IsArray(value_temp(0)) then
					value = value_temp(0)(0)
				else
					value = value_temp(0)
				end if
			else
			  value = value_temp
			end if
			
			'attempt to correct variable using vartype
			if variable = "Unknown" then variable = "VBType " & cstr(VarType(value_temp))
			
		'add row to data
		dir = "|-" & name & " (" & variable & ") - " & loadset.getValueDescription(name) & _
			" = " & value & " [" & loadset.countRows(name) & "]"
		call textwin.writeLine(dir)
		
	next
	
	'respond with directory
	'msgbox dir, vbinformation, "(__dir__)"
		
End Sub