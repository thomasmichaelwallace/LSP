$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Match Properties
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

'Match attribute assignments of first selected objects to rest.

Option Explicit

dim master		'object to match
dim slave		'object to change
dim assignments		'properties to match
dim assignment		'current matching property
dim assignattr		'current assigning attribute
dim slave_id		'slave properties matching

'get master properties
set master = selection.getObjects("all")(0)

'cycle through slaves
for slave_id = 1 to ubound(selection.getObjects("all"))
	set slave = selection.getObjects("all")(slave_id)
	call textwin.writeLine("Matching properties of " & master.getName() & " to " & slave.getName())	

	'itterate though all assignments
	assignments = master.getAssignments
	for each assignment in assignments
		set assignattr = assignment.getAttribute

		'apply attributes
		call assignattr.assignTo(slave, assignment)
	next

next
	
'update mesh to reflect changes
call db.meshSelected()