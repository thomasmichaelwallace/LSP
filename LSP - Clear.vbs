$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Clear Properties
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

'Clear all attribute assignments from selected objects.

Option Explicit

dim elements			'objects selected
dim element				'object to change
dim assignments		'properties to match
dim assignment		'current matching property
dim assignattr		'current assigning attribute

'get master properties
elements = selection.getObjects("all")

'cycle through elements
for each element in elements
	call textwin.writeLine("Clearing properties of " & element.getName())

	'itterate though all assignments
	assignments = element.getAssignments
	for each assignment in assignments
		set assignattr = assignment.getAttribute

		'apply attributes
		call assignattr.deassignFrom(element, assignment)
	next

next
	
'update mesh to reflect changes
call db.meshSelected()