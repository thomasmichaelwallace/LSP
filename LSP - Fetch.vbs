$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Fetch Results
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

'Fetch and report selected results for a node/element.

'arrays
dim base_values			'initial values
dim basecase_ids		'loadcases included in base loadcase set

'objects
dim node			'node to factor
dim element			'element to factor
dim basecase			'base loadcase to factor from

'itterators
dim row				'current row

'reports
dim base_string			'summary of original values
dim header_string		'summary header

'user options
dim element_id			'selected element id
dim node_id			'selected node id
dim basecase_id			'base loadcase id
dim mode_repeat

'script options
dim entity			'entity type to factor
dim components			'entity components to factor
dim report_places		'number of decimal places to show in message box

mode_repeat = vbYes

do while mode_repeat = vbYes

	'get user options
	element_id = cint(Inputbox("Element id", "Fetch Results", "11621"))
	node_id = cint(Inputbox("Node id", "Fetch Results", "11743"))
	basecase_id = cint(Inputbox("Base loadcase id", "Fetch Results", "45"))

	'set default options
	entity = "Force/Moment - Thick 3D Beam"
	components = Array("Fx", "Fy", "Fz", "Mx", "My", "Mz")
	report_places = 3

	'configure modeller
	setManualRefresh(true)

	'load loadcases
	set basecase = database.getLoadset(basecase_id)
	basecase_ids = basecase.getLoadcaseIDs()

	'setup reference
	set element = db.getObject("Element", element_id)
	set node = db.getObject("Node", node_id)

	'get basecase
	call view.setActiveLoadset(basecase_id)
	redim base_values(ubound(components))
	for row = 0 to ubound(components)
		base_values(row) = element.getNodeResults(node, entity, components(row))
	next

	'report results
	for row = 0 to ubound(base_values)		
		header_string = components(row) & ": "
		base_string = formatnumber(base_values(row), report_places)
		report_string = report_string & header_string & base_string & vblf
	next

	'return modeller
	call view.setActiveLoadset(basecase_id)
	setManualRefresh(false)

	'show results
	mode_repeat = msgbox("Fetched results for element " & element_id & ", node " & node_id & vblf & _
		report_string & "- Fetch more results?",	vbYesNo, "Fetch Results")
	
loop