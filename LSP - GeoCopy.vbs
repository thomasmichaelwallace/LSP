$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Geometry Copy
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

'Copy selected LUSAS geometry definition to file; for later use by LSP Geometry Paste.
'Copy support also provided for assignments, but not fully implemented.

dim TEXT_FILE				'text file object to write data to	
dim copy_mode				'result of msgbox

copy_mode = msgbox("Do you want to copy assignments in addition to the geometry?", _
	vbYesNoCancel + vbQuestion, "Copy Geometry")

'switch board
if copy_mode = vbYes then
	CopyGeometry(True)
elseif copy_mode = vbNo then
	CopyGeometry(False)	
else
	textwin.writeLine("Geomertry copy aborted.")
end if


sub CopyGeometry(assign)
	'copy selected geometry to a definition file

	dim filesystem		'file system access
	dim file_path			'path to copy file

	dim points				'selected points
	dim lines					'selected lines
	dim combined_lines'selected combined lines
	dim surfaces			'selected surfaces	

	dim point					'current point
	dim line					'current line
	dim surface				'current surface

	dim def_array			'array of definitions

	dim lofs					'lower order features
	dim lof_count			'total lofs to write
	dim lof_start			'start position of lof definition
	
	dim bulge_coords	'coordinates of arc/circle bulge
	dim end_coords		'coordinates of arc/circle end

	dim facet_coords	'coordinates of spline facets
	dim facet_count		'total of spline facets
	dim facet_start		'start position of spline facet definitions

	dim boundaries		'total surface boundaries
	dim boundary			'current surface boundary
	dim trust_order		'flag for surface orientation reliablilty
	
	dim i							'counter
	dim count					'count of objects copied
	dim total_lines		'total number of lines to create	

	'script level defaults and options
	file_path = GetSystemString("SCRIPTS") & "LSP\LSP - GeoCopy.def"
	count = 0

	'surpress messages and report
	call suppressMessages(1)
	call textwin.writeLine("Copying geometric data to " & file_path & "...")
	
	'capture all geometric features upper and lower bounds
	call selection.addLOF()
	points = selection.getObjects("Points")
	lines = selection.getObjects("Lines")
	combined_lines = selection.getObjects("Combined Lines")
	surfaces = selection.getObjects("Surfaces")
	
	'init progress bar
	total_lines = ubound(points) + ubound(lines) + ubound(combined_lines) + ubound(surfaces)
	if assign then total_lines = total_lines * 2
	call initStatusBarProgressCtrl("Copying geometry...", cint(total_lines))
	
	'warn about unhandled object types
	if selection.count("Volumes") > 0 then
		call textwin.writeLine("Volume objects are currently unhandled; ignored.")
	end if
		
	'prepare file
	set filesystem = CreateObject("Scripting.FileSystemObject")
	if not filesystem.FileExists(file_path) then
		set TEXT_FILE = filesystem.CreateTextFile(file_path)
		TEXT_FILE.close
		set TEXT_FILE = nothing
	End If 
	set TEXT_FILE = filesystem.OpenTextFile (file_path, 2, true)

	'get points
	for each point in points
		'point - 1; 0; id; x; y; z
		def_array = array(point.getTypeCode(), 0, point.getID(), _
			point.getX(), point.getY(), point.getZ())

		write_def(def_array)
		count = count + 1
	next

	'get lines
	for each line in lines
		call statusBarProgressCtrlStep()
		select case line.getLineTypeCode()
		
			case 1
				'straight line - 2; 1; id; start_id; end_id
				lofs = line.getLOFs()			
				def_array = array(line.getTypeCode(), line.getLineTypeCode(), line.getID(), _
					lofs(0).getID(), lofs(1).getID())

			case 2
				'arc line - 2; 2; id; start_id; end_id; bulge_coords x; y; z; plane_coords x; y; z; angle							
				lofs = line.getLOFs()
				
				'define end coordinate at interpolated (75% length) position to cope with arcs and cricles
				bulge_coords = line.getArcBulge()
				end_coords = line.getInterpolatedPosition(0.75)

				def_array = array(line.getTypeCode(), line.getLineTypeCode(), line.getID(), _
					lofs(0).getID(), lofs(1).getID(), _
					bulge_coords(0), bulge_coords(1), bulge_coords(2), _
					end_coords(0), end_coords(1), end_coords(2), _
					line.getArcAngleDegrees())

			case 3
				'splice line - 2; 3; id; start_id; end_id; facet_count; def_coords x; y; z...
				lofs = line.getLOFs()			
				
				'define coordinates at facets to reduce spline curve as visualised by user
				facet_coords = line.getFacetCoordinates()
				facet_count = (ubound(facet_coords) + 1) / 3
				
				def_array = array(line.getTypeCode(), line.getLineTypeCode(), line.getID(), _
					lofs(0).getID(), lofs(1).getID(), facet_count)
							
				facet_start = ubound(def_array) + 1
				redim preserve def_array(facet_start + (facet_count * 3))
				
				for i = 0 to ubound(facet_coords)
					def_array(facet_start + i) = facet_coords(i)
				next
				
		end select

		write_def(def_array)		
		count = count + 1
	next

	'get combined lines
	for each line in combined_lines
		call statusBarProgressCtrlStep()	
		select case line.getLineTypeCode()
		
			case 4
				'combined line - 3; 4; id; lofs_count; lof_id...
				lofs = line.getLOFs()
				lof_count = ubound(lofs) + 1
				
				def_array = array(line.getTypeCode(), line.getLineTypeCode(), line.getID(), lof_count)
				
				'define combined lines by their lower-order lines
				lof_start = ubound(def_array)
				redim preserve def_array(lof_start + lof_count)
				
				for i = 1 to lof_count
					def_array(lof_start + i) = lofs(i - 1).getID()
				next

				write_def(def_array)
				count = count + 1
		end select
	next

	'surfaces
	for each surface in surfaces
		call statusBarProgressCtrlStep()
		
		'surface - 5; surface_type; id; boundaries; b0_lofs_count; b0_lof_id...; b0_check_order; b1...
		boundaries = surface.countBoundaries() - 1		
		def_array = array(surface.getTypeCode(), surface.getSurfaceTypeCode(), surface.getID(), boundaries)
		
		'define surface as a series of boundaries (#0 is outer; rest are holes.)
		for boundary = 0 to boundaries				
			lofs = surface.getBoundaryLOFs(boundary)
			
			'manage definition array positions
			lof_count = ubound(lofs) + 1
			lof_start = ubound(def_array)	+ 1
			redim preserve def_array(lof_start + lof_count +1)

			def_array(lof_start) = lof_count
			
			'orientation is not implicit when combined lines are used; (LUSAS bug)
			trust_order = True
			for i = 1 to lof_count
				def_array(lof_start + i) = lofs(i - 1).getID()
				if lofs(i - 1).getTypeCode() = 3 then trust_order = False
			next

			'convert trust flag as definition is string only
			if trust_order = False then
				def_array(ubound(def_array)) = "Check"
			else
				def_array(ubound(def_array)) = "Trust"
			end if
		next
		
		write_def(def_array)	
		count = count + 1
	next	

	'copy assignments if required
	if assign then
		textwin.writeLine("Adding assignments to definition file...")
		for each point in points
			call statusBarProgressCtrlStep()
			write_assign(point)
		next
		for each line in lines
			call statusBarProgressCtrlStep()
			write_assign(line)
		next
		for each surface in surfaces
			call statusBarProgressCtrlStep()
			write_assign(surface)
		next
	end if
	
	'report
	call suppressMessages(0)
	call textwin.writeLine(cstr(count) & " geometries copied successfully.")
	call unInitStatusBarProgressCtrl()		
	
end sub
	
sub write_def(definitions)
	'expand a definition array into a full string, and write to file
	
	dim def_string		'developed definition string
	dim definition		'current definition
	
	'set definition line mode to geometry
	def_string = "@G"
	
	'expand and delete superflous seperators
	for each definition in definitions
		def_string = def_string & ";" & definition
	next	
	
	TEXT_FILE.writeline(def_string)	
end sub

sub write_assign(geometry)
	'write assignments to file
	
	dim assign_string	'developed assignment string
	dim assignments		'attributes assigned to geometry
	dim assignment		'current assignment being defined
	
	'provide header and end definition
	assignments = geometry.getAssignments()	
	assign_string = "@S;" + cstr(geometry.getID())	
	
	if ubound(assignments) > -1 then
	
		for each assignment in assignments
			assign_string = assign_string + ";" + _
				assignment.getAttributeType() + ";" + _
				assignment.getAttribute().getName()
		next
	
		TEXT_FILE.writeline(assign_string)
	
	end if	
end sub