$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Loadcase Generator
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

'Automagically define loadcases from loading arrays.

'arrays
dim suffixes()			'suffixes to array
dim switches()			'blocks that switch
dim commons()			'blocks that don't switch
dim permutations()		'total suffix permutations
dim blocks()			'blocks for switching input

'loadcase objects
dim def_case 			'loadcase object
dim def_ids			'load id
dim def_eigenvalues		'eigenvalue id
dim def_factors			'factor
dim def_harmonics		'harmonics id
dim def_results 		'result id
dim def_load			'load object

'counters
dim suffix_count		'count of defined suffix arrays
dim switch_count		'count of switching blocks
dim switch_index		'current switching block
dim row				'current row
dim branches			'number of prefixes
dim branch			'current prefix
dim is_common			'flag for common blocks
dim perm_count 			'permutations count

'objects
'dim switch_array()		'array defining switching load
'dim common_array()		'array defining common load
dim permutation()		'permutation object

'user options
dim def_name			'defining case to duplicate
dim def_string			'string defining suffix array
dim def_suffix			'array from string definition

'get basic user options
def_name = Inputbox("Definition case","Loadcase Array","ULS Combo 1")
def_string = Inputbox("Loading arrays (define;suffixes;...", "Loadcase Array", "Centre;Left;Right")
base_name = def_name

'get additional user options
suffix_count = -1
do	
	'add new suffix branch
	suffix_count = suffix_count + 1
	def_suffix = split(def_string, ";")
	redim preserve suffixes(suffix_count)
	suffixes(suffix_count) = def_suffix
	
	'continue until no more definitions
	def_string = Inputbox("Additional loading arrays", "Loadcase Array", "Expansion;Contraction")		
	
loop until def_string = ""

'load deinition data
set def_case = database.getLoadset(def_name)
def_ids = def_case.getLoadcaseIDs()
def_eigenvalues = def_case.getEigenvalueIDs()
def_factors = def_case.getFactors()
def_harmonics = def_case.getHarmonicIDs()
def_results = def_case.getResultsFileIDs()

'parse into common/switching blocks
branches = suffix_count
switch_count = -1
common_count = -1
switch_max = -1
for row = 0 to ubound(def_ids)
	
	'get load infromation
	set def_load = database.getLoadset(def_ids(row))
	def_name = def_load.getName()

	'check for switching key
	is_common = True
	for branch = 0 to branches
		if instr(def_name, "(" & suffixes(branch)(0) & ")") then
			
			'create switching object
			switch_array = array(def_name, def_eigenvalues(row), def_factors(row), _
				def_harmonics(row), def_results(row))
			
			'add switching row to branch
			switch_count = switch_count + 1
			if switch_count > switch_max then switch_max = switch_max + 1
			redim preserve switches(branches, switch_max)
			switches(branch, switch_count) = switch_array
			is_common = False
			
			exit for
			
		end if
	next
	
	'record common blocks
	if is_common then
		
		'create common object
		common_array = array(def_ids(row), def_eigenvalues(row), def_factors(row), def_harmonics(row), _
			def_results(row))	
	
		'add to common block
		common_count = common_count + 1
		redim preserve commons(common_count)						
		commons(common_count) = common_array

	end if
next

'size permultation array
perm_count= 1
for branch = 0 to branches
	perm_count = perm_count * (ubound(suffixes(branch)) + 1)
next
perm_count = perm_count - 1
redim permutations(perm_count)
redim permutation(branches)

'get permultations
Permutate suffixes, 0, -1, permutation, permutations

'write permutations
for perm_index = 0 to perm_count
	case_name = base_name & " "
	block_count = -1
	
	'create name
	for branch = 0 to branches
		case_name = case_name & "(" & permutations(perm_index)(branch) & ")"
		
		'consolidate switcher into block
		for switch_index = 0 to switch_max
			if not isEmpty(switches(branch, switch_index)) then
				block_count = block_count + 1				
				redim preserve blocks(block_count)
				
				'locate new load
				load_name = switches(branch, switch_index)(0)
				load_name = replace(load_name, "(" & suffixes(branch)(0) & ")", _
					"(" & (permutations(perm_index)(branch)) & ")")
				set new_load = database.getLoadset(load_name)
				load_id = new_load.getId()
				
				'create common array
				common_array = array(load_id, switches(branch, switch_index)(1), _
					switches(branch, switch_index)(2), switches(branch, switch_index)(3), _
					switches(branch, switch_index)(4))
				blocks(block_count) = common_array
				
				'debug message
				'msgbox "Object: " & cstr(switches(branch, switch_index)(0)) & vblf & _
				'	"Mode: " & cstr(permutations(perm_index)(branch)) & vblf & _
				'	"Replace: " & load_name
			end if
		next
	next
	'msgbox "LC : " & case_name
	
	'write case	
	set load_combo = database.createCombinationBasic(case_name)

	'add common aspects
	for index = 0 to common_count
		call load_combo.addEntry(commons(index)(2), commons(index)(0), commons(index)(4), _
			commons(index)(1), commons(index)(3))
	next
	
	'add switching blocks
	for index = 0 to block_count
		call load_combo.addEntry(blocks(index)(2), blocks(index)(0), blocks(index)(4), _
			blocks(index)(1), blocks(index)(3))	
	next
	
next


sub Permutate(set_space, set_number, offset, memory, permutations)
	'populate a permutation matrix

	dim i		'force counter blank
	
	'scroll through current set level
	for i = 0 to ubound(set_space(set_number))
		memory(set_number) = set_space(set_number)(i)
		
		'add to permutations library at final level
		if set_number = ubound(set_space) then					
			offset = offset + 1
			permutations(offset) = memory		

		'resurse until lowest level
		else
			Permutate set_space, set_number + 1, offset, memory, permutations

		end if
	next
	
end sub