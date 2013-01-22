$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Sort Groups
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

'Sort groups alphabetically
Option Explicit

'group objects
dim groups		'model groups
dim group		'current working group
dim slot		'slot for working group
dim buffer		'buffer to copy

'maps
dim group_map()		'current/slot id map
dim isempty_map()	'history of created groups

'counters and boundaries
dim group_id		'current group id
dim map_id		'current map id
dim map_base		'map starting point
dim map_count		'total number of groups to sort
dim map_bound		'allocation of groups to sort

'bubble sort objects
dim bubble_move		'flag for end of bubble sort
dim compare_name	'name to compare...
dim to_name		'...against name

'layer 1 is the master layer; cannot be sorted.
map_base = 2

'fetch database groups
groups = db.getObjects("Group")

'create group map
map_count = ubound(groups) - lbound(groups)
map_bound = map_count + map_base
redim group_map(map_bound)
redim isempty_map(map_bound)
map_id = map_base

'populate group maps
for each group in groups
	group_map(map_id) = group.getID()
	map_id = map_id + 1
next

'provide worse-case sort-time
call initStatusBarProgressCtrl("Determining sort order...", cint(map_count))

'flag bubble move
bubble_move = true
do while bubble_move = true
	bubble_move = false

	'sweep through group_map
	for group_id = map_base to map_bound - 1
		
		'fetch names
		compare_name = db.getObject("Group", clng(group_map(group_id))).getName()
		to_name = db.getObject("Group", clng(group_map(group_id + 1))).getName()
		
		'prevent case level seraching
		compare_name = ucase(compare_name)
		to_name = ucase(to_name)
		
		'compare sort and flag bubble if out-of-order
		if compare_name > to_name then
			bubble_move = true
					
			'bubble swap into order
			map_id = group_map(group_id + 1)
			group_map(group_id + 1) = group_map(group_id)
			group_map(group_id) = map_id
			
		end if
	next	

	'run/escape progress
	call statusBarProgressCtrlStep()	
loop
call unInitStatusBarProgressCtrl()		

'create defrag'd groups
for group_id = map_base to map_bound	
	if db.exists("Group", clng(group_id)) then
		isempty_map(group_id) = false	
	else
		set group = db.createEmptyGroup("_LSP_SPACE_" & cstr(group_id))
		call textwin.writeLine("Group Create: " & cstr(group_id))
		isempty_map(group_id) = true
	end if
next

'create buffer group
set buffer = db.createEmptyGroup("_LSP_BUFFER")

'debug-fake sort:
for group_id = map_base to map_bound
	call textwin.writeLine("Group " & cstr(group_id) & ": " & _
		db.getObject("Group", clng(group_map(group_id))).getName())
next

'prepare system for faster sort time
call initStatusBarProgressCtrl("Sorting groups...", cint(map_count))
call setManualRefresh(true)
call suppressMessages(1)						
call selection.remove("All")
call visible.add("All")

for map_id = map_base to map_bound
	
	set group = db.getObject("Group", clng(group_map(map_id)))
	set slot = db.getObject("Group", clng(map_id))
	
	if group.getID() = slot.getID() then
		'skip groups already in order		
		call textwin.writeLine("[" & cstr(map_id) & "] Ignr : (" & _
			cstr(group_map(map_id)) & ") " & slot.getName())
	
	else
	
		if isempty_map(map_id) = true then
			'populate defrag groups

			call textwin.writeLine("[" & cstr(map_id) & "] Move : (" & _
				cstr(group_map(map_id)) & ") " & group.getName())
			
			'duplicate group into slot
			call slot.remove("All")
			call selection.remove("All")
			call selection.add("Group", group.getID(), "Recurse")			
			call slot.add(selection)
			call slot.setName(group.getName())
			
			'wipe group
			call group.ungroup()
		
		else
			'use buffer to swap occupied slot
	
			call textwin.writeLine("[" & cstr(map_id) & "] Swap : (" & _
				cstr(group_map(map_id)) & ") " & group.getName() & "/" & slot.getName())
	
			'duplicate group into buffer
			call buffer.remove("All")
			call selection.remove("All")			
			call selection.add("Group", slot.getID(), "Recurse")			
			call buffer.add(selection)
			call buffer.setName(slot.getName())

			'move group to slot
			call slot.remove("All")
			call selection.remove("All")
			call selection.add("Group", group.getID(), "Recurse")			
			call slot.add(selection)
			call slot.setName(group.getName())
			
			'write buffer to group
			call group.remove("All")
			call selection.remove("All")
			call selection.add("Group", buffer.getID(), "Recurse")			
			call group.add(selection)
			call group.setName(buffer.getName())
			call buffer.setName("_LSP_BUFFER")
			
			'locate pivot point
			for group_id = map_id to map_bound
				if map_id = group_map(group_id) then exit for
			next
			
			'update map
			group_map(group_id) = group_map(map_id)
			
		end if
			
	end if

	call statusBarProgressCtrlStep()	
next

'restore system
call setManualRefresh(false)
call suppressMessages(0)
call unInitStatusBarProgressCtrl()		

'kill buffer
call buffer.ungroup()