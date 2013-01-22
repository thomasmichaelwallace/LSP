$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Show Neighbour
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

'Make neighbour geometries of selected visible.

dim geometries			'selected geometries
dim geometry			'working geometry
dim lines			'defining lines
dim line			'current line
dim adj_geometries		'adjacent geometries
dim adj_geometry		'working adjacent geometry

'fetch selected geometries
geometries = selection.getObjects("All")

'itterate through selected geometries
for each geometry in geometries
	
	'for each line get geometries
	lines = geometry.getLOFs()			
	for each line in lines	
		adj_geometries = line.getHOFs()
	
		'make geometries visible
		for each adj_geometry in adj_geometries	
			call visible.add(adj_geometry)

		next
	next
next