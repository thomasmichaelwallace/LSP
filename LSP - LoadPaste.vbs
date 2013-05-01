$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Load Paste
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

'Paste load attribute definitions from definition file defined by the LSP Load Copy command.

'collections
dim values		'array of loaded variables

'objects
dim loading		'load being copied
dim value		'value of variable to be copied

'counters
dim row			'current internal loading row being pasted
dim rows		'total number of internal loading rows

'storage
dim paste_string	'complete data string from file
dim loading_type	'loading type being copied
dim loading_name	'name of loading being copied

'filesytem objects
dim filesystem		'file system access
dim text_file		'text file to save copying data

'options
dim file_path		'file path to save copied data

'script level defaults and options
file_path = GetSystemString("CONFIGDIR") & "LSP\LSP - LoadCopy.txt"

'prepare file
set filesystem = CreateObject("Scripting.FileSystemObject")
if not filesystem.FileExists(file_path) then
	set text_file = filesystem.CreateTextFile(file_path)
	text_file.close
	set text_file = nothing
end if 
set text_file = filesystem.OpenTextFile (file_path, 1)

'input copying file
do until text_file.AtEndOfStream

	'get copying line data
	paste_string = text_file.ReadLine
	values = split(paste_string, ";")	
	loading_type = values(0)
	loading_name = values(1)
	
	'check for existing
	if db.existsAttribute("Loading", loading_name) then
		if not isNumeric(loading_name) then _
			call textwin.writeLine("Skipping existing: " & loading_name)
	
	else		
		if not isNumeric(loading_name) then _
			call textwin.writeLine("Pasting loading: " & loading_name)
	
		select case loading_type
		
			case "Body Force Load"
				set loading = db.createLoadingBody(loading_name)
				'(accX, accY, [accZ], [angVelX], [angVelY], [angVelZ], [angAccX], [angAccY], [angAccZ])
				call loading.setBody(values(2), values(3), values(4), _
					values(5), values(6), values(7), _
					values(8), values(9), values(10))			
			
			case "Global Distributed Load"
				set loading = db.createLoadingGlobalDistributed(loading_name)
				'(type, wx, wy, [wz], [mx], [my], [mz], [loof1], [loof2])		
				call loading.setGlobalDistributed(values(2), _
					values(3), values(4), values(5), _
					values(6), values(7), values(8), _
					values(9), values(9))
				if values(9) <> 0 then _
					call textwin.writeLine("Assumed loof 1 and loof 2 are equal to hinge rotation.")	
			
			case "Distributed Load"
				set loading = db.createLoadingLocalDistributed(loading_name)
				'(wx, wy, [wz])
				call loading.setLocalDistributed(values(2), values(3), values(4))
			
			case "Temperature Load"
				set loading = db.createLoadingTemperature(loading_name)
				'(type, temp, [dT/dX], [dT/dY], [dT/dZ], [T0], [dT0/dX], [dT0/dY], [dT0/dZ])
				call loading.setTemperature(values(2), _
					values(3), values(4), values(5), values(6), _
					values(7), values(8), values(9), values(10))		
			
			case "Discrete Point Load"
				set loading = db.createLoadingDiscretePoint(loading_name)
				'(dirType, Dir, [nGridX], [nGridY])
				call loading.setDiscrete(values(2), array(values(3), values(4), values(5)), _
					values(6), values(7))
							
				if values(2) = "XYZ" then call textwin.writeLine("Dirtype XYZ loading ignored.")
				
				'fetch individual points
				rows = values(8)			
				for row = 0 to rows			
					paste_string = text_file.ReadLine
					values = split(paste_string, ";")
					
					'(coordX, coordY, coordZ, load, [load2], [load3])
					call loading.addRow(values(1), values(2), values(3), values(4))
				next

			case "Discrete Patch Load"
				set loading = db.createLoadingDiscretePatch(loading_name)
				'(type, dirType, [Dir])
				call loading.setDiscretePatch(values(2), _
					values(6), array(values(7), values(8), values(9)))

				if values(6) = "XYZ" then call textwin.writeLine("Dirtype XYZ loading ignored.")
				if values(10) <> 0 or values(11) <> 0 then _
					call textwin.writeLine("Grid X/Y for patch loading ignored.")
					
				'set optional arguments
				if values(4) <> 0 or values(5) <> 0 then call loading.setDivisions(values(4), values(5))
				if values(3) <> 0 then call loading.setSweptAngleDegrees(values(3))
										
				'set individual definitions
				rows = values(12)
				for row = 0 to rows
					paste_string = text_file.ReadLine
					values = split(paste_string, ";")			

					'(coordX, coordY, coordZ, load, [load2], [load3])
					call loading.addRow(values(1), values(2), values(3), values(4))
				next

			case "Discrete Compound Load"
				set loading = db.createLoadingDiscreteCompound(loading_name)
				
				rows = values(2)
				for row = 0 to rows
					paste_string = text_file.ReadLine
					values = split(paste_string, ";")	
					
					if db.existsAttribute("Loading", values(1)) and _
						(db.existsAttribute("Transformation", values(5)) or values(5) = "__Nothing__") then
					
						'(pLoadingAttr, [pOffsetCoord], [pTransAttr])					
						if values(5) = "__Nothing__" then						
							call loading.addLoading(db.getAttribute("Loading", values(1)), _
								array(values(2), values(3), values(4)))
						else
							call loading.addLoading(db.getAttribute("Loading", values(1)), _
								array(values(2), values(3), values(4)), _
								db.getAttribute("Transformation", values(5)))
						end if
					
					'copy with loads/transformations missing from compound definition
					else
						if db.existsAttribute("Loading", values(1)) then
							call textwin.writeLine(values(1) & " loading does not exist, aborted paste.")
						else
							call textwin.writeLine(values(5) & " transformation does not exist, aborted paste.")
						end if						
						db.deleteAttribute(loading)
						exit for					
					end if
				next

			case else
				if not isNumeric(loading_name) then _
					call textwin.writeLine("Loading type unhandled.")				
				
		end select
	end if
loop

'clean up
text_file.close