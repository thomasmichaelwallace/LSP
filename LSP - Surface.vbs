$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Surface Fix
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

'Attempt to fix perimeter for gaps in surface definitions on 'Fix' layer.

'counts
dim surface_count		'count of surfaces checked
dim interface_count		'count of interfaces 
dim fix_count			'count of surfaces redefiend
dim crash_count			'count to prevent crashing

'options
dim fix_group			'name of group to be fixed
dim fixed_group			'where to dump fixed surfaces
dim crash_max			'runs until mesh is allowed
dim crash_protect		'attempt to hone apparent access violations
dim debug_mode			'further debugging for vbscripting
dim edge_protect		'focused debugging for 'edge' members

'configuration
fix_group = inputbox("Group name to fix surfaces on", "Surface Fix", "Fix")
fixed_group = inputbox("Group name to place fixed surfaces on", "Surface Fix", "Fixed")

'testing options
crash_max = 5
crash_protect = false
debug_mode = false
edge_protect = false

'set counters
surface_count = 0
interface_count = 0
fix_count = 0	
crash_count = 0

'debug mode overrides
if edge_protect then crash_max = 1
if edge_protect then debug_mode = true

'setup environment
call database.setMeshLock(true)
setManualRefresh(true)
if debug_mode then call setManualRefresh(false)

'programme loop
do while SurfFix = true
	crash_count = crash_count + 1
	if crash_protect then msgbox "Run count: " & crash_count & "/" & crash_max
	if crash_count = crash_max and (crash_protect or edge_mode) then exit do
loop

'provide script statistics
stats = "Surface fix complete" & vbNewLine & _
	" Group Included - " & fix_group & vbNewLine & _
	" Surfaces Checked - " & surface_count & vbNewLine & _
	" Interfaces Tested - " & interface_count & vbNewLine & _
	" Voids Fixed - " & fix_count
if crash_protect or debug_mode or edge_mode then
	stats = stats & vbNewLine & _
	" Crash Protection - " & crash_protect & vbNewLine & _
	" Debug Mode - " & debug_mode & vbNewLine & _
	" Edge Protection " & edge_mode
end if

'restore session
call selection.remove("all")
call database.updateMesh()
call database.setMeshLock(false)
call setManualRefresh(false)

function SurfFix()
	'fix selected surfaces where voids have formed due to modelling errors
	
	'collections
	dim surfaces			'collection of selected surfaces
	dim define_lines		'collection of defining lines
	dim define_surfaces		'collection of surfaces linked to the line	
	dim define_points		'collection of points defining surface line
	dim point_lines			'collection of lines defined by point
	dim check_points		'collection of points linked to a line
	dim check_lines			'collection of lines for check-back	
	dim back_points			'collections of points for check-back
	
	'members
	dim surface			'surface being assesed
	dim define_line			'definition line being assesed
	dim master_point		'point to investigate
	dim slave_point			'point to check-back to
	dim point_line			'line defined by master point being assesed
	dim check_point			'point checked as mid-point	
	dim check_line			'final line implied by match
	dim back_point			'point to compare for check-back
	dim fix_surface			'surface to be fixed
		
	'setup environment
	SurfFix = false
	
	'populate surface collection
	call selection.remove("All")
	call selection.add("Group", fix_group)
	surfaces = selection.getObjects("Surfaces")
	call selection.remove("All")
	
	'check selected surfaces
	for each surface in surfaces
		surface_count = surface_count + 1
		
		'[debug] increase visibility of surfaces
		if debug_mode then
			call selection.remove("All")
			call selection.add(surface)
			msgbox "Debug [Surface]: " & surface_count
		end if
		
		'populate and iterate definition lines collection
		define_lines = surface.getLOFs()		
		for each define_line in define_lines

			'only continue if line defines one surface
			define_surfaces = define_line.getHOFs()
			if ubound(define_surfaces) = 0 then
					
				'set checking/checked points
				define_points = define_line.getLOFs()
				set master_point = define_points(0)
				set slave_point = define_points(1)
			
				'populate and iterate through master point lines
				point_lines = master_point.getHOFs()
				for each point_line in point_lines
					interface_count = interface_count + 1

					'[debug] increase visibility of lines
					if debug_mode then
						call selection.remove("All")
						call selection.add(point_line)
						msgbox "Debug [Interface]: " & interface_count
					end if					
					
					'locate check-point potentially defining check-back
					check_points = point_line.getLOFs()
					for each check_point in check_points
						if not check_point.getName() = master_point.getName() then

							'populate and iterate through valid check-point lines
							check_lines = check_point.getHOFs()
							
							if ubound(check_lines) = 2 then
								for each check_line in check_lines

									'locate true back-point for check-back
									back_points = check_line.getLOFs()
									if back_points(0).getName() = check_point.getName() then
										set back_point = back_points(1)
									else
										set back_point = back_points(0)
									end if									
									
									'test to see if line creates void
									if back_point.getName() = slave_point.getName() then
										fix_count = fix_count + 1
										
										'trim surface
										call CleverTrim(surface, define_line, point_line, check_line)
										
										'restart to fix bad handling of variable collection sizes
										SurfFix = true
										exit function
										
									end if
								next
							end if
							
						end if
					next
					
				next
				
			end if

		next
		
	next	
	
end function

sub CleverTrim(surface, remove_line, off_line, back_line)
	'manually trim a surface

	Dim new_surface_lines		'outline of trimmed surface
	Dim new_combined_line		'combined top lines
	Dim assignments			'collection of assignments to match		
	Dim assignment			'specific assignment matching
	Dim new_surface			'trimmed surface
	Dim attr			'attribute for assignement
	
	'develop outline for replacement surface
	set new_surface_lines = newObjectSet()
	call new_surface_lines.add(surface.getLOFs())
	call new_surface_lines.remove(remove_line)
	call new_surface_lines.add(off_line)
	call new_surface_lines.add(back_line)
		
	'select old surface
	call selection.remove("All")
	call selection.add(surface)
	
	'[debug] highlight old surfaces
	if debug_mode then msgbox "Debug [Fixed]: " & fix_count
	
	'backup assignments
	assignments = surface.getAssignments()
	
	'remove surface and trim line; [debug] highlight steps
	call selection.delete("Surface")	
	if edge_protect then msgbox "Debug [Edge]: Surface Removed"	
	call selection.remove("All")
	call selection.add(remove_line)
	if edge_protect then msgbox "Debug [Edge]: Line Selected"	
	call selection.delete("Line")
	if edge_protect then msgbox "Debug [Edge]: Line Removed"		
	
	'create new surface
	call geometryData.setAllDefaults()
	call geometryData.setCreateMethod("Planar")
	call geometryData.setLowerOrderGeometryType("Lines")
	set new_surface = new_surface_lines.createSurface(geometryData)

	'reapply assignments
	for assignment = 0 to ubound(assignments)
		set attr = assignments(assignment).getAttribute()
		call attr.assignTo(new_surface, assignments(assignment))
	next
	
	'combine lines for easy meshing
	call geometryData.setCreateMethod("Combined")
	call geometryData.useSelectionOrder(true)
	set new_combined_line = newObjectSet()
	call new_combined_line.add(off_line)
	call new_combined_line.add(back_line)
	call new_combined_line.createCombinedLine(geometryData)
	
	're-group
	call selection.add(new_surface)
	call database.getGroupByName(fix_group).remove(selection)
	call database.getGroupByName(fixed_group).add(selection)	

	'[debug] clarify view
	if debug_mode then
		call visible.keep("Group", group_fix, "Recurse")
		call visible.add("Group", group_fix, "Recurse")
	end if
	
end sub