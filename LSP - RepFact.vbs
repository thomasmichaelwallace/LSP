$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Factor Find and Replace
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

'Find and replace factors in loadsets.

'collections
dim loadsets			'complete list of loadsets
dim ids				'loadcase ids
dim results			'result files ids
dim eignevalues			'eigenvalue ids
dim harmonics			'harmonic ids
dim factors			'existing factors
dim replacements		'replacement factors

'objects
dim loadset			'current loadset
dim index			'current row

'flags
dim count			'number of matches
dim match			'matching flag

'options
dim find_factor			'original factor to check for
dim replace_factor		'factor to replace with

'get user options
find_factor = cdbl(inputbox("Find factor", "Factor Replace", "2.35"))
replace_factor = cdbl(inputbox("Replace factor", "Factor Replace", "1.75"))
count = 0

'fetch all combinations
loadsets = db.getLoadsets("Combination", "All")
for each loadset in loadsets

	'only work on valid combinations
	if loadset.getTypeCode = 2 then

		'fetch associated factors
		match = false
		factors = loadset.getFactors
		redim replacements(ubound(factors))
		for index = 0 to ubound(factors)
		
			'replace any matching factors
			if factors(index) = find_factor then
				replacements(index) = replace_factor
				count = count + 1
				match = true

			'preserve existing matrix
			else
				replacements(index) = factors(index)
				
			end if
		next
		
		'apply any changes
		if match then

			'duplicate data
			ids = loadset.getLoadcaseIDs
			results = loadset.getResultsFileIDs
			eigenvalues = loadset.getEigenvalueIDs
			harmonics = loadset.getHarmonicIDs
				
			'delete and re-add to cope with limited editing scope of the LPI
			loadset.removeEntries
			
			'deal with lusas' inability to size arrays
			for index = 0 to ubound(replacements)
				call loadset.addEntry(replacements(index), ids(index), results(index), _
					eigenvalues(index), harmonics(index))
			next
			
		end if
	end if
next

'report back
msgbox count & " replacements made.", vbinformation, "Factor Replace"