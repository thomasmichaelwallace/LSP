$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Geometry Paste
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

'Paste geometry as defined by the LSP Geometry Copy command. Note that assignment definitions, are
'not supported.

dim DEFINITIONS			'geometric definitions
dim POINT_MAP			'map between old ids and new points
dim LINE_MAP			'map between old ids and new lines
dim SURFACE_MAP			'map between old ids and new surfaces

'switch board for assignments, [not implemented]
if PasteGeometry() = "@A" then
	'if msgbox( _ 
		'"The current definition file includes assignments, do you wish to paste these as well?", _
		'vbYesNo + vbQuestion, "Paste Geometry") = vbYes then

		'PasteAssignments
		textwin.writeLine("Assignment definitions ignored.")
		
	'end if
end if

function PasteGeometry()
	'paste geometry from definition file

	dim filesystem		'file system access
	dim file_path		'path to definition file
	dim text_file		'definition file
	
	dim surface_types	'map between lpi type codes and surface names

	dim def_string		'current definition string

	dim type_code		'geometry type code
	dim subtype_code	'geometry subtype code	
	dim id			'old id
	dim def_mode		'line definition mode
	
	dim point		'current point
	dim line		'current line
	dim surface		'current surface
	
	dim boundaries		'surface boundaries
	dim boundary		'current surface boundary
	
	dim from_index		'position from index
	dim index		'current position index
	dim to_index		'position to index
	
	dim i			'iterator
	dim count		'total pasted
	
	dim total_lines		'total number of lines to process

	'script level defaults and options
	file_path = GetSystemString("SCRIPTS") & "LSP\LSP - GeoCopy.def"
	
	'setup map between surface type codes and names
		'LPI surface type map: [plannar, cylindrical, spherical, conical_
		'nurb, coons, web, transsweep, rotsweep, isoparametric, offset, ruled]
		'where lpi map does not function as expected, or surface is untested, unset is used to allow
		'lusas to choose sensibly.
	surface_types = array( _ 	
		"plannar", "cylindrical", "spherical", "conical", _ 
		"unset", "coons", "unset", "unset", "unset", "unset", "unset", "unset") 

	'prepare file
	set filesystem = CreateObject("Scripting.FileSystemObject")
	if not filesystem.FileExists(file_path) then
		set text_file = filesystem.CreateTextFile(file_path)
		text_file.close
		set text_file = nothing
	end if 
	
	'get line count
	set text_file = filesystem.OpenTextFile(file_path, 8)
	total_lines = text_file.line
	text_file.close
	set text_file = nothing
	
	set text_file = filesystem.OpenTextFile(file_path, 1)

	'init values
	PasteGeometry = "@G"
	redim POINT_MAP(0)
	redim LINE_MAP(0)	
	count = 0
		
	'surpress messages and report
	call suppressMessages(1)
	call setManualRefresh(True)
	call textwin.writeLine("Pasting geometric data from " & file_path & "...")		

	'init progress bar
	call initStatusBarProgressCtrl("Pasting geometry...", cint(total_lines))
	
	'input copying file by line
	do until text_file.AtEndOfStream
		
		'common preperation
		call geometryData.setAllDefaults()
		call statusBarProgressCtrlStep()

		'convert def string format
		def_string = text_file.ReadLine
		DEFINITIONS = split(def_string, ";")
		def_mode = DEFINITIONS(0)
		
		'set line mode
		if def_mode = "@G" then
		
			'extract common aspects
			type_code = DEFINITIONS(1)
			subtype_code = DEFINITIONS(2)
			id = DEFINITIONS(3)
			
			select case type_code
			
				case 1
					'point - main type; 0; id; x; y; z
					call geometryData.addCoords(DEFINITIONS(4), DEFINITIONS(5), DEFINITIONS(6))
					call geometryData.setLowerOrderGeometryType("coordinates")
					set point = database.createPoint(geometryData).getObject("Point")
									
					set_map point, id, POINT_MAP	
					count = count + 1
			
				case 2
					'lines
					select case subtype_code
					
						case 1
							'straight line - 2; 1; id; start_id; end_id		
							get_map 4, 5, POINT_MAP, "Points"
							
							call geometryData.setCreateMethod("straight")
							call geometryData.useSelectionOrder(true)
							call geometryData.setLowerOrderGeometryType("points")

						case 2
							'arc line - 2; 2; id; start_id; end_id; bulge_coords x; y; z; plane_coords x; y; z; angle
							call geometryData.addCoords(map_coords(DEFINITIONS(4)))
							call geometryData.addCoords(DEFINITIONS(6), DEFINITIONS(7), DEFINITIONS(8))
													
							if DEFINITIONS(12) = 360 then
								'create circle from interpolated end point
								call geometryData.addCoords(DEFINITIONS(9), DEFINITIONS(10), DEFINITIONS(11))							
								call geometryData.makeCircle()
							elseif DEFINITIONS(12) <= 180 then
								'distinguish between minor and major arcs
								call geometryData.addCoords(map_coords(DEFINITIONS(5)))
								call geometryData.keepMinor()
							else
								call geometryData.addCoords(map_coords(DEFINITIONS(5)))
								call geometryData.keepMajor()
							end if 
											
							call geometryData.setCreateMethod("arc")					
							call geometryData.setStartMiddleEnd()					
							call geometryData.createArcCentrePoint(False)
							call geometryData.setLowerOrderGeometryType("coordinates")

						case 3	
							'splice line - 2; 3; id; start_id; end_id; facet_count; def coords				
							for i = 0 to DEFINITIONS(6) - 1
								
								'create spline from facet points; provindg accuracy to those visualsied
								call geometryData.addCoords( _
									DEFINITIONS((i * 3) + 1 + 6), _
									DEFINITIONS((i * 3) + 2 + 6), _
									DEFINITIONS((i * 3) + 3 + 6))
							next
							
							call geometryData.setCreateMethod("spline")
							call geometryData.setLowerOrderGeometryType("coordinates")
												
					end select

					'blanket line creation code
					set line = selection.createLine(geometryData).getObject("Line")
					set_map line, id, LINE_MAP
					count = count + 1
					
					'remap start/end points if replaced by merge
					set POINT_MAP(DEFINITIONS(4)) = line.getLOFs()(0)
					set POINT_MAP(DEFINITIONS(5)) = line.getLOFs()(1)			

				case 3
					'combined line - 3; 4; id; lofs_count; lof_id...					
					get_map 5, DEFINITIONS(4) + 4, LINE_MAP, "Line"			
					
					call geometryData.setCreateMethod("combined")
					call geometryData.useSelectionOrder(true)
					call geometryData.setLowerOrderGeometryType("lines")				
					call selection.createCombinedLine(geometryData)							
					
					set line = selection.createLine(geometryData).getObject("Combined Line")
					set_map line, id, LINE_MAP
					count = count + 1

				case 4
					'surface - 4; surface_type; id; boundaries; b0_lofs_count; b0_lof_id...; b0_check_order; b1...			
					
					call geometryData.setCreateMethod(surface_types(cint(DEFINITIONS(2))))
					call geometryData.useSelectionOrder(true)
					call geometryData.setLowerOrderGeometryType("Combined Line")	

					'create outer boundary
					get_map 6, DEFINITIONS(5) + 5, LINE_MAP, "Line"
					set surface = selection.createSurface(geometryData).getObject("Surface")			

					'setup hole definition locations
					boundaries = DEFINITIONS(4)			
					index = DEFINITIONS(5) + 7
					
					for boundary = 1 to boundaries
						
						'setup lusas to remove holes from main surface
						call geometryData.setAllDefaults()
						call geometryData.trimOuterBoundaryOff()				
						call geometryData.trimDeleteOuterBoundaryOff()
						call geometryData.trimDeleteTrimmingLinesOff()					
						call selection.remove("All")
						call selection.add(surface)
						
						'reposition parser
						from_index = index + 1
						index = from_index + DEFINITIONS(index)
						to_index = index - 1
						
						'define hole
						for i = from_index to to_index		
							call selection.add(LINE_MAP(cint(DEFINITIONS(i))))
						next				
						call selection.trim(geometryData)
						
						'remap merged lines
						for i = from_index to to_index
							set LINE_MAP(cint(DEFINITIONS(i))) = surface.getBoundaryLOFs(boundary)(i - from_index)
						next
					
					next
					
					'manually orientate surface if not implied; (LUSAS Bug)
					if DEFINITIONS(DEFINITIONS(5) + 6) = "Check" then
						manual_surface surface, 6, DEFINITIONS(5) + 5
					end if
					
				case else
					textwin.writeLine "Unhandled definition: " & def_string	& "; ignored"
					
			end select
	
		elseif def_mode = "@S" then
			PasteGeometry = "@S"
		end if
	
	loop

	'restore messages and report
	call suppressMessages(0)
	call unInitStatusBarProgressCtrl()
	call setManualRefresh(False)
	call textwin.writeLine(cstr(count) & " geometries pasted successfully.")		
	
end function

sub set_map(geometry, map_id, byref map)
	'maintain mapping between copied and pasted ids	
	
	'map old id to new object
	if cint(map_id) > ubound(map) then
		redim preserve map(cint(map_id))
	end if	
	set map(cint(map_id)) = geometry

end sub

sub get_map(from_index, to_index, map, name)
	'fetch matched type from definttions array

	dim i	'counter	
	
	call selection.remove("All")

	'select objects by old id
	for i = from_index to to_index		
		call selection.add(map(cint(DEFINITIONS(i))))
	next	

end sub

function map_coords(coord_id)
	'return coordinates for mapped point
	
	map_coords = array(POINT_MAP(cint(coord_id)).getX, _
		POINT_MAP(cint(coord_id)).getY, _
		POINT_MAP(cint(coord_id)).getZ)

end function

sub manual_surface(surface, from_index, to_index)
	'manually set surface orientation
	
	dim cycle_max		'maximum cycle orientations
	dim cycle_order		'correct lof definition order
	
	dim lofs		'surface lofs
	dim lof			'current lof being tested
	
	dim correct		'correct orientation flag
	
	dim i			'counter
	dim mode		'mode switcher for reversal
	
	call geometryData.setAllDefaults()
	
	'fetch expeected lof order for correct orientation
	cycle_max = to_index - from_index
	redim cycle_order(cycle_max)	
	for i = 0 to cycle_max
		cycle_order(i) = LINE_MAP(cint(DEFINITIONS(from_index + i))).getID()
	next
	
	'test reversal
	for mode = 0 to 1
		surface.reverse()
	
		'test all cycle locations
		for i = 0 to cycle_max
			surface.cycle()
			lofs = surface.getLOFs()
		
			'test if order is now as defiend
			correct = true			
			for lof = 0 to cycle_max
				if not (lofs(lof).getID() = cycle_order(lof)) then
					correct = false
					exit for
				end if
			next
			
			'continue until correct, or cycles compelte
			if correct = true then exit for
		next		
		if correct = true then exit for
	next				
end sub