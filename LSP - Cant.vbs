$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Cant
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

'Convert a simple train load into one which is factored by North/South rail.

'objects and properties
dim compound			'compound loading object
dim discrete			'discrete loading object
dim pos				'current position data
dim P 				'current loading data
dim y				'y position to point

'iterators
dim index			'index counter for compound library
dim row				'row of position data

'factors
dim north_factor		'factor to be applied to north loads
dim south_factor		'factor to be appleid to south loads
dim load_factor			'factor applied to current load

'user options
dim name			'name of compound load to apply to
dim factor			'global factor to apply to all loads
dim north_ratio			'ratio of north track loading
dim south_ratio			'ratio of south track loading
dim axis			'axis to rotate to

'functional constants
dim split			'split already applied by load generator
dim equator			'point where north/south are divided

'set user options
name = Inputbox("Compound Name","Canter","Cmp")
factor = cdbl(Inputbox("Global Factor","Canter","1.0"))
north_ratio = cdbl(Inputbox("North Ratio","Canter","1.25"))
south_ratio = cdbl(Inputbox("South Ratio","Canter","1.0"))
axis = ucase(Inputbox("Set Projection","Canter","O"))

'set default options
split = 0.5
equator = 0.0

'calculate factors
north_factor = (north_ratio / (north_ratio + south_ratio)) * (factor / split)
south_factor = (south_ratio / (north_ratio + south_ratio)) * (factor / split)

'get main compound object
set compound = database.getAttribute("Discrete Compound Load", name)

'scroll through loading
for index = 0 to compound.countLoading() - 1
	set discrete = compound.getLoading(index)

	'rotate if required
	if axis = "X" or axis = "Y" or axis = "Z" then call discrete.setValue("dirType", axis)
	
	'get existing coordinate points
	for row = 0 to discrete.countRows("P") - 1
		pos = discrete.getValue("pos", row)
		
		'test array dimensions to cope with LUSAS' random return values
		if IsArray(pos(0)) then
			y = pos(0)(1)
		else
			y = pos(1)
		end if
	
		'determine conversion type
		if y > equator then
			load_factor = north_factor
		elseif y < equator then
			load_factor = south_factor
		else
			load_factor = factor
		end if
		
		'factor loading
		P = discrete.getValue("P", row)
		P = P * load_factor
		call discrete.setValue("P", P, row)	
		
	next		
next