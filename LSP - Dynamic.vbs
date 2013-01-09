$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Apply Dynamic Factor
'Copyright (C) 2010-2012 Thomas Michael Wallace <http://www.thomasmichaelwallace.co.uk>

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

'Applies the dynamic factor to live loading at a node and appends results to file.

'arrays
dim factored_values	'running total of factored additions
dim base_values		'initial values
dim testcases		'loadcase ids subject to factoring
dim livecases		'loadcase ids potentially subject to factoring
dim basecase_ids	'loadcases included in base loadcase set

'objects
dim node		'node to factor
dim element		'element to factor
dim component		'component currently being factored
dim basecase		'base loadcase to factor from

'itterators
dim row			'current row
dim id			'current loadcase id
dim delta_component	'factored addition

'file objects
dim text_file		'report file object
dim filesystem		'filesystem access object
dim file_path		'path of report file

'reports
dim report_string	'string to append to report file
dim delta_string	'summary of changes
dim factored_string	'summary of factored values
dim base_string		'summary of original values
dim header_string	'summary header

'user options
dim element_id		'selected element id
dim node_id		'selected node id
dim delta_phi		'addative multiplcation factor
dim basecase_id		'base loadcase id
dim livecases_string	'string defining live loadcases to factor
dim mode_repeat		'repeater mode

'script options
dim entity		'entity type to factor
dim components		'entity components to factor
dim report_places	'number of decimal places to show in message box

mode_repeat = vbYes

do while mode_repeat = vbYes

	'get user options
	element_id = cint(Inputbox("Element id", "Dynamic Factor", "11621"))
	node_id = cint(Inputbox("Node id", "Dynamic Factor", "11743"))
	basecase_id = cint(Inputbox("Base loadcase id", "Dynamic Factor", "45"))
	livecases_string = cstr(Inputbox("Live loadcases ids (1;2;...)", "Dynamic Factor", "3;18;24;25;26;27"))
	delta_phi = cdbl(Inputbox("Addative factor", "Dynamic Factor", "1"))

	'set default options
	entity = "Force/Moment - Thick 3D Beam"
	components = Array("Fx", "Fy", "Fz", "Mx", "My", "Mz")
	file_path = db.getDBFilenameNoExtension & " - Dynamic Factored.txt"
	report_places = 3

	'configure modeller
	setManualRefresh(true)

	'load loadcases
	livecases = split(livecases_string, ";")
	set basecase = database.getLoadset(basecase_id)
	basecase_ids = basecase.getLoadcaseIDs()

	'identify inclusive live loadcases
	redim testcases(0)
	for id = 0 to ubound(basecase_ids)
		for row = 0 to ubound(livecases)

		if cint(livecases(row)) = basecase_ids(id) then
				'save matching loadcases
				redim preserve testcases(ubound(testcases) + 1)
				testcases(ubound(testcases)) = basecase_ids(id)
				exit for	
			end if
		
		next
	next

	'setup reference
	set element = db.getObject("Element", element_id)
	set node = db.getObject("Node", node_id)

	'get basecase
	call view.setActiveLoadset(basecase_id)
	redim factored_values(ubound(components))
	for row = 0 to ubound(components)
		factored_values(row) = element.getNodeResults(node, entity, components(row))
	next
	base_values = factored_values

	'fetch test cases
	for id = 1 to ubound(testcases)
		call view.setActiveLoadset(testcases(id))
		
		'check all listed components
		for row = 0 to ubound(components)
			component = element.getNodeResults(node, entity, components(row))
			delta_component = delta_phi * component
			
			'save addative change
			factored_values(row) = factored_values(row) + delta_component
		
		next
	next

	'report results
	report_string = element_id & vbtab & node_id
	header_string = ""
	base_string = ""
	factored_string = ""
	delta_string = ""	
	for row = 0 to ubound(factored_values)
		header_string = header_string & components(row) & ";"
		base_string = base_string & formatnumber(base_values(row), report_places) & ";"
		factored_string = factored_string & formatnumber(factored_values(row), report_places) & ";"
		delta_string = delta_string & _
			formatnumber(((factored_values(row) - base_values(row))/base_values(row))*100, report_places) & "%;"
		report_string = report_string & vbtab & factored_values(row)
	next

	'return modeller
	call view.setActiveLoadset(basecase_id)
	setManualRefresh(false)

	'show results
	msgbox "The following results have been added to " & file_path & vblf & _
		"Components: " & header_string & vblf & _
		"Base Value:" & base_string & vblf & _
		"Factored Value: " & factored_string & vblf & _
		"Delta: " & delta_string,	vbInformation, "Dynamic Factor"
		
	'create text file if required
	set filesystem = CreateObject("Scripting.FileSystemObject")
	if not filesystem.FileExists(file_path) then
		set text_file = filesystem.CreateTextFile(file_path)
		text_file.close
		set text_file = nothing
	End If 

	'append result to text file
	set text_file = filesystem.OpenTextFile (file_path, 8, true)
	text_file.WriteLine(report_string)
	text_file.close
	
	mode_repeat = msgbox("Factor more results?", vbYesNo, "Dynamic Factor")
	
loop