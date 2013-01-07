$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Show Neighbour
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

'Make neighbour surfaces of selected visible.

dim surfaces				'selected surfaces
dim surface					'working surface
dim lines						'defining lines
dim line						'current line
dim adj_surfaces		'adjacent surfaces
dim adj_surface			'working adjacent surface

'fetch selected surfaces
surfaces = selection.getObjects("Surface")

'itterate through selected surfaces
for each surface in surfaces
	
	'for each line get surfaces
	lines = surface.getLOFs()			
	for each line in lines	
		adj_surfaces = line.getHOFs()
	
		'make surfaces visible
		for each adj_surface in adj_surfaces	
			call visible.add(adj_surface)

		next
	next
next