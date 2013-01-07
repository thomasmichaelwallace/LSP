$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Pin Beams
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

'Converts beam ends to pins of selected group where intersected with columns.

'collections
dim lines						'lines forming beams
dim points					'points forming current beam
dim branches				'lines attached to beam forming points

'objects
dim line						'current beam line
dim point						'current beam line end
dim branch					'current beam/column attached to beam line end

'fixity data
dim fixity(1)				'applied fixity matrix (0 is start, 1 is end)
dim fixity_point		'index for apply fixity

'column attributes
dim line_mesh				'beam mesh assignments
dim line_fixity			'existing beam fixity assignment

'mesh applications
dim pinned					'mesh applied when beam is released at both ends
dim pinned_start		'mesh applied when beam is released at start
dim pinned_end			'mesh applied when beam is released at end
dim fixed						'mesh applied when beam is no end releases

'script options
dim pinned_id				'id for pinned mesh
dim pinned_start_id	'id for pinned at start mesh
dim pinned_end_id		'id for pinned at end mesh
dim fixed_id				'id for fixed mesh
dim beam_group			'beam group name
dim column_group		'column group name

'script options
fixed_id = inputbox("Fixed beam connection attribute ID", "Pin Beam", "1")
pinned_id = inputbox("Pined at start beam connection attribute ID", "Pin Beam", "2")
pinned_start_id = inputbox("Pined at end beam connection attribute ID", "Pin Beam", "3")
pinned_end_id = inputbox("Pined both sides beam connection attribute ID", "Pin Beam", "4")
beam_group = inputbox("Beam group", "Pin Beam", "Beams")
column_group = inputbox("Column group", "Pin Beam", "Columns")

'setup attributes
set pinned = db.getAttribute("Mesh", clng(pinned_id))
set pinned_start = db.getAttribute("Mesh", clng(pinned_start_id))
set pinned_end = db.getAttribute("Mesh", clng(pinned_end_id))
set fixed = db.getAttribute("Mesh", clng(fixed_id))

'fetch beams
call selection.remove("All")
call selection.add("Group", beam_group)
lines = selection.getObjects("Lines")
call selection.remove("All")

'cycle through beams
for each line in lines

	'reset matrix
	fixity(0) = True
	fixity(1) = True

	'maintain existing fixity options
	line_mesh = line.getAssignments("Mesh")
	set line_fixity = line_mesh(0).getAttribute()					
	if line_fixity is pinned then
		exit for
	elseif line_fixity is pinned_start then
		fixity(0) = False
	elseif line_fixity is pinned_end then
		fixity(1) = False
	end if

	'fetch points
	points = line.getLOFs()		
	for fixity_point = 0 to ubound(points)
		set point = points(fixity_point)
	
		'fetch branches
		branches = point.getHOFs()			
		for each branch in branches
		
			'release end if attached to column
			if branch.isMemberOfGroup(column_group) then														
				fixity(fixity_point) = 0
				exit for
			end if

		next
	next

	'assign correct fixity to column
	if fixity(0) and fixity(1) then
		fixed.assignTo(line)
	elseif fixity(0) then
		pinned_end.assignTo(line)
	elseif fixity(1) then
		pinned_start.assignTo(line)
	else
		pinned.assignTo(line)
	end if
				
next

'remesh
call database.closeAllResults()
call database.updateMesh()