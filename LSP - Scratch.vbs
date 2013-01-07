$ENGINE=VBScript
' LUSAS Modeller session file
' Created by LUSAS 14.7-4 - Modeller Version 14.7.1665.13586
' Created at 14:12 on Friday, January 04 2013
' (C) Finite Element Analysis Ltd 2013
'
call setCreationVersion("14.7-4, 14.7.1665.13586")
'

'*** Print Results Wizard

redim loadcases(0)
redim lcResFiles(0)
loadcases(0) = "Wind"
lcResFiles(0) = "P:\eng\300.Western Structures\302. SAVEx West Superstructure\04. Calcs\Lusas Model - MASTER\Lusas - MASTER\TMW\TMW.mys"
call prwOptions.setAllDefaults()
call prwOptions.setID(1)
call prwOptions.setLoadcases(loadcases, lcResFiles)
call prwOptions.setSigFig(6)
call prwOptions.showCoordinates(false)
call prwOptions.setExtent("Visible model", "")
call printWizard("Reaction", "Results Component Node", prwOptions)
erase loadcases
erase lcResFiles

'*** Print Results Wizard

call getPrintResultsWindowByID(1).setCurrent()

'*** Print Results Wizard

call getPrintResultsWindowByID(1).close()

