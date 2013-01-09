$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Load Length
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

'Apply a load to a series of connected lines until a certain length is reached

'run as a sub as a work around for m$'s poor exception handling
call SetupLength()

sub SetupLength()
	'establishes the scenario to load

	'database objects
	dim line		'spawnning line element
	dim points		'points of line element
	dim loading		'loading attribute pointer
	
	'geometry
	dim east_point		'east point index
	dim west_point		'west point indext
	dim actual_length	'actual length counter

	'user options
	dim load_name		'name of load to apply
	dim direction		'direction to follow
	dim target_length	'length to aim for

	'check selection for line
	if selection.count("Line") <> 1 then
		msgbox "Select line first", vbcritical, "Length Loader"
		exit sub
	end if

	'get user options
	load_name = inputbox("Load name", "Length Loader", "Traction")
	target_length = cint(inputbox("Target length", "Length Loader", "30.303"))
	direction = inputbox("Follow direction ([W]est/[E]ast/[B]oth)", "Length Loader", "E")	
	
	'fetch database objects
	set line = selection.getObject("Line")
	points = line.getLOFs()
	
	'find east (also first point if x1 == x2)
	if points(0).getX() >= points(1).getX() then
		east_point = 0
		west_point = 1
	else
		east_point = 1
		west_point = 0
	end if		
	
	'setup assignment
	set loading = database.getAttribute("Loading", load_name)
	
	'organise application	
	select case direction
		case "E"
			'head east
			actual_length = LoadLength(line, points(east_point), target_length, loading)
		case "W"
			'head west
			actual_length = LoadLength(line, points(west_point), target_length, loading)			
		case "B"
			'go half the length in both directions
			actual_length = LoadLength(line, points(0), 0.5 * target_length, loading)
			actual_length = actual_length + LoadLength(line, points(1), 0.5 * target_length, loading)	
	end select
			
	msgbox "Length Loader Complete." & vbNewLine & _
		"  Target length: " & cstr(target_length) & vbNewLine & _
		"  Actual length: " & left(cstr(actual_length), len(cstr(target_length))+3) & vbNewLine & _
		"  Error: " & left(cstr(((actual_length - target_length)/target_length)*100),5) & "%", _
		vbinformation, "Length Loader"
			
end sub

	
function LoadLength(first_line, follow_point, target_length, loading)
	'follows a line along a direction and loads all following lines

	dim lines		'collection of lines associated at point
	dim line		'current line
	dim points		'collection of points associated with line
	dim point		'current point
	dim total_length	'length count
	dim i			'array counter
	
	'initial conditions
	total_length = 0.0	
	set line = first_line
	set point = follow_point	
	
	'continue until target is met or exceeded
	while total_length < target_length
			
		'mark line for loading and add length
		selection.add(line)		
		total_length = total_length + line.getLineLength()

		'select next line
		lines = point.getHOFs()	
		for i = 0 to ubound(lines)
			'only add visible points to prevent confusion and allow non-lines to be changed
			if lines(i).getId() <> line.getId() and lines(i).isVisible() then
				set line = lines(i)
				exit for
			end if
		next
		
		'select next point (details as line)
		points = line.getLOFs()
		for i = 0 to ubound(points)
			if points(i).getId() <> point.getId() and points(i).isVisible() then
				set point = points(i)
				exit for
			end if
		next
		
	wend		

	'asign loading
	call assignment.setAllDefaults()
	call loading.assignTo(selection, assignment)
	
	'return actual length achieved
	LoadLength = total_length

end function	