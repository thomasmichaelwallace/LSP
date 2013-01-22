$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Show Parent
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

'Make parent geometries of selected visible.

dim geometries			'selected geometries
dim geometry			'working geometry
dim parents			'parents
dim parent			'working parent

'fetch selected geometries
geometries = selection.getObjects("All")

'itterate through selected geometries
for each geometry in geometries
	
	'get parents
	parents = line.getHOFs()	

	'make parents visible
	for each parent in parents
		call visible.add(parent)

	next
next