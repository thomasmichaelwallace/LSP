$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Autocant
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

'Automagically distribute a vertical load accoridng to cant and centrigual proportion

'objects and properties
dim compound			'compound loading object
dim discrete			'discrete loading object
dim pos				'current position data
dim P 				'current loading data
dim y				'y position to point

'iterators
dim index			'index counter for compound library
dim row				'row of position data

'workings
dim shift			'shift in centre of gravity
dim cant_cos			'vertical component of in-plane moment reaction

'factors
dim high_factor			'factor to be applied to high loads
dim low_factor			'factor to be appleid to low loads
dim static_high			'high factor due to static loading
dim static_low			'low factor due to static loading
dim dynamic			'factor due to dynamic loading
dim load_factor			'factor applied to current load

'user options
dim name			'name of compound load to apply to
dim cant			'cant of rail
dim centrigual			'proprtion of vertical load that is centrfugal
dim north			'high rail is north
dim e				'eccentricity

'functional constants
dim split			'split already applied by load generator
dim equator			'point where high/low are divided
dim gauge			'standard gauge width
dim offset			'height at which centrigual loading is applied

'set user options
name = Inputbox("Compound Name","Quick Canter","LM71 + Cf Vert Train (Centre)")
cant = cdbl(Inputbox("Rail Cant (mm)","Quick Canter","150"))
centrigual = cdbl(Inputbox("Centrigual Coefficient (of vertical)","Quick Canter","0.32"))
e = cdbl(Inputbox("Eccentricity (mm) towards high","Quick Canter","155"))
if cstr(Inputbox("High Rail (North/South)","Quick Canter","North")) = "North" then
	north = true
else
	north = false
end if

'set default options
split = 0.5			'default by load generator
equator = 0.0			'default by load generator
gauge = 1435.0			'standard gauge
offset = 1800.0			'as eurocode

'develop plane dimensions
cant_cos = sqr(1.0 - (cant / gauge)^2.0)
	'sin = opposite/adjacent; sin^2 + cos^2 = 1; cos = sqr(1 - sin^2) = sqr(1 - (opposite/adjacent)^2)
shift = (((cant / gauge) * offset) / cant_cos) - e
	
'develop static factors
static_low = (0.5 * gauge + shift) / gauge
static_high = (0.5 * gauge - shift) / gauge

'develop centrifugal factors
dynamic = (centrigual * offset) / gauge * cant_cos

'calculate factors
high_factor = (static_high + dynamic) / split
low_factor = (static_low - dynamic) / split

'reporting
msgbox "Applied Distribution: " & vbNewLine & _
	"   High - " & high_factor & " / " & (high_factor * split) & vbNewLine & _
	"   Low - " & low_factor & " / " & (low_factor * split) & vbNewLine & _
	"   Check - " & (high_factor + low_factor) & " / " & ((high_factor + low_factor)*split), _
	vbInformation, "Quick Canter"

'get main compound object
set compound = database.getAttribute("Discrete Compound Load", name)

'scroll through loading
for index = 0 to compound.countLoading() - 1
	set discrete = compound.getLoading(index)

	'get existing coordinate points
	for row = 0 to discrete.countRows("P") - 1
		pos = discrete.getValue("pos", row)
		
		call textwin.writeline(discrete.getDiscreteLoadType())

		'test array dimensions to cope with LUSAS' random return values
		if IsArray(pos(0)) then
			call textwin.writeLine("double")
			y = pos(0)(1)
		else
			call textwin.writeLine("single")
			y = pos(1)
		end if
	
		'set conversion mode
		if not north then y = -y
	
		'determine conversion type
		if y > equator then
			load_factor = high_factor
		elseif y < equator then
			load_factor = low_factor
		else
			load_factor = (high_factor + low_factor) / 2.0
		end if
		
		'factor loading
		P = discrete.getValue("P", row)
		P = P * load_factor
		call discrete.setValue("P", P, row)
				
	next		
next