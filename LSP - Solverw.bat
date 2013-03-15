@ECHO OFF

ECHO The LUSAS Scriping Pack (LSP): Solver Daemon (Waiter)
ECHO Copyright (C) 2010-2013 Thomas Michael Wallace (http://www.thomasmichaelwallace.co.uk)
ECHO.
ECHO  This file is part of the LSP.
ECHO.
ECHO     The LSP is free software: you can redistribute it and/or modify
ECHO     it under the terms of the GNU General Public License as published by
ECHO     the Free Software Foundation, either version 3 of the License, or
ECHO     (at your option) any later version.
ECHO.
ECHO     The LSP is distributed in the hope that it will be useful,
ECHO     but WITHOUT ANY WARRANTY; without even the implied warranty of
ECHO     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
ECHO     GNU General Public License for more details.
ECHO.
ECHO     You should have received a copy of the GNU General Public License
ECHO     along with The LSP.  If not, see (http://www.gnu.org/licenses/)

REM Solver Daemon waiter object for running machine

REM Move to daemon space
SET daemon_path=%1
SET daemon_drive=%daemon_path:~1,2%
%daemon_drive%
CD %daemon_path%

REM Ensure solver daemon is waiting
IF EXIST %3 (
		GOTO SKIPLOCK
	) ELSE (
		DEL %2
	)

:WAITLOCK
	ECHO.
	ECHO Engaging solver daemon...
	ECHO.
	PING -n 10 localhost > NUL

	IF EXIST %2 (
		ECHO Solver Daemon is waiting..
	) ELSE (		
		ECHO Solver Daemon does not appear to be running.
		ECHO.
		ECHO Please run LSP - Solverd.bat on the machine you want to use as a solver daemon.
		ECHO It should be located within the project directory.
		ECHO.
		PAUSE	
		GOTO WAITLOCK
	)

:SKIPLOCK
	
ECHO.
ECHO Waiting for daemon to become available...
ECHO.
	
:QUEUEMODEL
	IF EXIST %3 (
		ECHO Solver daemon is still busy...
		PING -n 10 localhost > NUL
		GOTO QUEUEMODEL
	) ELSE (
		ECHO Daemon is available.
	)
	
	
REM Create solver daemon instructions
ECHO %4 > "LSP - Solverd.dta"
ECHO %5 > "LSP - Solverd.res"
ECHO %6 > "LSP - Solverd.lgr"
ECHO %7 > "LSP - Solverd.dri"
ECHO %8 > "LSP - Solverd.pat"
ECHO %9 > "LSP - Solverd.thd"

ECHO "LOCKED" > "LSP - Solverw.tmp"

ECHO.
ECHO Waiting for solver daemon...
ECHO.

:WAITSOLVE

	IF EXIST "LSP - Solverw.tmp" (
		ECHO Solver Daemon is still solving...
		PING -n 10 localhost > NUL
		GOTO WAITSOLVE
	) ELSE (		
		ECHO Solver Daemon has finished solving.		
	)

REM Tidy up
DEL "LSP - Solverd.dta"
DEL "LSP - Solverd.res"
DEL "LSP - Solverd.dri"
DEL "LSP - Solverd.lgr"
DEL "LSP - Solverd.pat"
DEL "LSP - Solverd.thd"