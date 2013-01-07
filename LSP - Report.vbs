$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Bulk Reporter
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

'Automate bulk result reporting for envolopes over named groups

'collections
dim loadcase_results()	'selected loadcase results filenames
dim loadcase_names()		'selected loadcase names
dim loadcase_ids		'selected loadcase ids
dim groups					'model groups

'objects
dim group						'current group
dim loadcase				'current loadcase

'data
dim group_name			'current group name
dim report_path			'report file name and path

'counters
dim count						'number of reports generated
dim id							'current loadcase id
dim index						'current loadcase index

'user options
dim loadcase_string	'loadcase selection string array
dim entitiy					'entity to report on
dim mode_suffix			'mode to report in
dim component_strnig'string array of components to report

'script options
dim prefix					'reportable group prefix indentifier
dim components			'componenets to report
dim mode						'mode string

'user options
loadcase_string = inputbox("Loadcases (1;2;...)", "Report Generator", "1364;1365;1368;1369")
mode_suffix = inputbox("Result level (Element, Node)", "Report Generator", "Node")
entity = inputbox("Result type (Reaction, ...)", "Report Generator", "Reaction")
component_string = inputbox("Components", "Report Generator", "Fx;Fy;Fz;Mx;My;Mz")
prefix = inputbox("Report for groups prefixed by", "Report Generator", "_Rep_")

'script options
components = split(component_string, ";")
mode = "Results Component " & mode_suffix

'provide "all" functionality
if loadcase_string = "Basic Combination" then
	loadsets = db.getLoadsets("Basic Combination", "All")
	redim loadcase_ids(ubound(loadsets))
	for index = 0 to ubound(loadsets)
		loadcase_ids(index) = loadsets(index).getID()
	next
else
	loadcase_ids = split(loadcase_string, ";")
end if

'get loadcase references
index = -1
for each id in loadcase_ids	
	
	'resize arrays
	index = index + 1
	redim preserve loadcase_names(index)
	redim preserve loadcase_results(index)

	'fetch identification data
	set loadcase = database.getLoadset(cint(id))
	loadcase_names(index) = loadcase.getName()
	loadcase_results(index) = loadcase.getResultsFileName()

	'cope with lusas' poor handling of envolopes
	if loadcase.getTypeCode = 3 then
		if loadcase.isMax then
			loadcase_names(index) = loadcase_names(index) & " (Max)"
		else
			loadcase_names(index) = loadcase_names(index) & " (Min)"
		end if
	end if		
	
next

'identify valid groups
count = 0
groups = db.getObjects("Group")

for each group in groups

	'test for identifying prefix
	group_name = group.getName()
	if prefix = left(group_name, len(prefix)) then		
		count = count + 1
		
		'setup default print options
		call prwOptions.setAllDefaults()
		call prwOptions.setSigFig(6)
		call prwOptions.showCoordinates(false)
		
		'configure to current results set
		call prwOptions.setID(group.getID)
		call prwOptions.setLoadcases(loadcase_names, loadcase_results)
		call prwOptions.setPrimaryComponents(components)
		call prwOptions.setExtent("Group", group_name)
		
		'get results
		call printWizard(entity, mode, prwOptions)	
		call getPrintResultsWindowByID(group.getID).setCurrent()
		
		'save results file
		report_path = mid(group_name, len(prefix) + 1) & ".xls"
		call getPrintResultsWindowByID(group.getID).saveAllAs(report_path, "Microsoft Excel")
		
		'close window
		call getPrintResultsWindowByID(group.getID).close()
		
	end if
next

'report results
msgbox count & " result files generated.", vbinformation, "Report Generator"