; Muchtall Window Arranger
; Version 20210211 - Bug fix to expand the default editor variable so that it actually fires notepad.exe
; Version 20200820 - Bug fix to better handle the situation where the "New monitor arrangement" is not dismissed, and then the screen layout changes. Now detects shanges between the loaded profile and the current layout.
; Version 20200817 - Now automatically creates and loads separate profiles based upon monitor arrangement
;                  - Removed static reference to PSPad as a text editor, and instead try to determine default text editor from the registry
; Version 20190529 - Added support for multiple saved monitor configurations
;                  - Added support for multiple monitor count triggers (probably could be reworked to use a list of values)
; Version 20170516 - Add option to close window after resize/reposition (for close-to-tray applications)
; Version 20130423 (initial release)

#Include %A_ScriptDir%

GetMonitorLayout() {
	MonitorLayout = 
	SysGet, MonitorCount, 80
	i := MonitorCount
	while i > 0 {
	    SysGet, CurrentMonitor, Monitor, %i%
	    ;MsgBox, Left: %CurrentMonitorLeft% -- Top: %CurrentMonitorTop% -- Right: %CurrentMonitorRight% -- Bottom %CurrentMonitorBottom%.
	    MonitorLayout := MonitorLayout CurrentMonitorLeft "_" CurrentMonitorTop "x" CurrentMonitorRight "_" CurrentMonitorBottom
	    i := i - 1
	    if (i > 0) {
	    	MonitorLayout := MonitorLayout "+"
	    }
	}
	;MsgBox, %MonitorLayout%
	Return %MonitorLayout%
}
MonitorLayout := GetMonitorLayout()
SavedPositionsFullFilePath := A_ScriptDir "`\SavedWindowPositions." MonitorLayout ".ahk"

if !FileExist(SavedPositionsFullFilePath) {
	FileAppend, `n, %SavedPositionsFullFilePath%
	MsgBox,, Muchtall Window Arranger, New monitor arrangement detected.`nGenerated a new profile at %SavedPositionsFullFilePath%
}
Run, %SavedPositionsFullFilePath%
LoadedMonitorLayout := MonitorLayout

Menu, Tray, Icon, %SystemRoot%\system32\SHELL32.dll, 99
Menu, Tray, Tip, Muchtall Window Arranger
Menu, tray, add, Save Active Window Position, SaveWindowPos
Menu, tray, Default, Save Active Window Position
Menu, tray, add, Open Saved Window Positions List, OpenWindowPosList
Menu, tray, add, Re-arrange Windows (Reload), ReArrangeWindows

SetTitleMatchMode, 1
Loop {
	MonitorLayout := GetMonitorLayout()
	If ( MonitorLayout != LoadedMonitorLayout ) {
		MsgBox,, Muchtall Window Arranger, Monitor arrangement change detected.`nRe-arranging windows in 5 seconds., 5
		Reload
	}
	Sleep, 1000
}

SaveWindowPos:
	MsgBox, Activate the window that you wish to save the postion of and click OK
	Sleep, 100
	WinGetActiveTitle, MyWinTitle
	WinGetText, MyWinText, %MyWinTitle%
	WinGetPos, MyWinX, MyWinY, MyWinW, MyWinH, A

	; START Design and Layout of GUI
	; Set up control dimensions and positions
	Column1 =	6
	Row1 =		6
	Column2Margin =	16
	RowMargin =	0
	TextRowOffset =	4
	TextHeight =	20
	TextWidth =	120
	EditHeight =	20
	EditWidth =	400
	ButtonWidth =	80
	ButtonHeight =	20
	ButtonMargin =	16

	Column2Pos := (Column2Margin + TextWidth)
	TextRow1 := ( Row1 + TextRowOffset )

	Gui, Add, Text, x%Column1% y%TextRow1% w%TextWidth% h%TextHeight% , Window Title
	Gui, Add, Edit, x%Column2Pos% y%Row1% w%EditWidth% h%EditHeight% vMyWinTitle, %MyWinTitle%
	NextRow := ( EditHeight + Row1 + RowMargin )
	NextRowText := ( NextRow + TextRowOffset )

	MyWinTextEditHeight := ( EditHeight * 3 )
	MyWinTextHeight := ( TextHeight * 3 )
	Gui, Add, Text, x%Column1% y%NextRowText% w%TextWidth% h%MyWinTextHeight% , Window Text`n (USE ONLY ONE`nOR NONE`nof these lines!!!)
	Gui, Add, Edit, x%Column2Pos% y%NextRow% w%EditWidth% h%MyWinTextEditHeight% vMyWinText, %MyWinText%
	NextRow := ( NextRow + MyWinTextEditHeight + RowMargin )
	NextRowText := ( NextRow + TextRowOffset )

	Gui, Add, Text, x%Column1% y%NextRowText% w%TextWidth% h%TextHeight% , Window Title Exclusions
	Gui, Add, Edit, x%Column2Pos% y%NextRow% w%EditWidth% h%EditHeight% vMyWinTitleExcl
	NextRow := ( NextRow + EditHeight + RowMargin )
	NextRowText := ( NextRow + TextRowOffset )

	Gui, Add, Text, x%Column1% y%NextRowText% w%TextWidth% h%TextHeight% , Window Text Exclusions
	Gui, Add, Edit, x%Column2Pos% y%NextRow% w%EditWidth% h%EditHeight% vMyWinTextExcl
	NextRow := ( NextRow + EditHeight + RowMargin )
	NextRowText := ( NextRow + TextRowOffset )

	Gui, Add, Text, x%Column1% y%NextRowText% w%TextWidth% h%TextHeight% , Title Matching Mode
	Gui, Add, DropDownList, x%Column2Pos% y%NextRow% w%EditWidth% r%EditHeight% Choose1 vMyTitleMatchMode, Starts with|Contains|Is Exactly|RegEx
	NextRow := ( NextRow + EditHeight + RowMargin )
	NextRowText := ( NextRow + TextRowOffset )

	Gui, Add, Text, x%Column1% y%NextRowText% w%TextWidth% h%TextHeight% , Bring to front?
	Gui, Add, DropDownList, x%Column2Pos% y%NextRow% w%EditWidth% r%EditHeight% Choose1 vMyWinActivate, No|Yes
	NextRow := ( NextRow + EditHeight + RowMargin )
	NextRowText := ( NextRow + TextRowOffset )

	Gui, Add, Text, x%Column1% y%NextRowText% w%TextWidth% h%TextHeight% , Close window?
	Gui, Add, DropDownList, x%Column2Pos% y%NextRow% w%EditWidth% r%EditHeight% Choose1 vMyWinClose, No|Yes
	NextRow := ( NextRow + EditHeight + RowMargin )
	NextRowText := ( NextRow + TextRowOffset )

	Gui, Add, Button, x%Column1% y%NextRowText% w%ButtonWidth% h%ButtonHeight% Default, &Save
	NextButtonPos := ( ButtonWidth + ButtonMargin )
	Gui, Add, Button, x%NextButtonPos% y%NextRowText% w%ButtonWidth% h%ButtonHeight% , &Cancel
	; END Design and Layout of GUI

	Gui, Show,, New GUI Window
	Return

ButtonSave:
	Gui, Submit
	Gui, Destroy
	if ( MyTitleMatchMode == "Starts with" ) {
		MyTitleMatchModeArg = 1
	}
	if ( MyTitleMatchMode == "Contains" ) {
		MyTitleMatchModeArg = 2
	}
	if ( MyTitleMatchMode == "Is Exactly" ) {
		MyTitleMatchModeArg = 3
	}
	if ( MyTitleMatchMode == "RegEx" ) {
		MyTitleMatchModeArg = RegEx
	}

	FileAppend, SetTitleMatchMode`, %MyTitleMatchModeArg%`n, %SavedPositionsFullFilePath%
	FileAppend, WinRestore`, %MyWinTitle%`, %MyWinText%`, %MyWinTitleExcl%`, %MyWinTextExcl% `n, %SavedPositionsFullFilePath%
	FileAppend, WinMove`, %MyWinTitle%`, %MyWinText%`, %MyWinX%`, %MyWinY%`, %MyWinW%`, %MyWinH%`, %MyWinTitleExcl%`, %MyWinTextExcl% `n, %SavedPositionsFullFilePath%
	if ( MyWinActivate == "Yes" ) {
		FileAppend, WinActivate`, %MyWinTitle%`, %MyWinText%`, %MyWinTitleExcl%`, %MyWinTextExcl% `n, %SavedPositionsFullFilePath%
	}
	if ( MyWinClose == "Yes" ) {
		FileAppend, WinClose`, %MyWinTitle%`, , %MyWinText%`, %MyWinTitleExcl%`, %MyWinTextExcl% `n, %SavedPositionsFullFilePath%
	}
	FileAppend, `n, %SavedPositionsFullFilePath%
	MsgBox, Saved to %SavedPositionsFullFilePath%
return

ButtonCancel:
	Gui, Submit
	Gui, Destroy
Return

OpenWindowPosList:
	; Get the default editor from the registry
	RegRead, TextEditorCommandLine, HKEY_CLASSES_ROOT\txtfile\shell\open\command
	MyRegEx := ` `%1
	MyEditor := RegExReplace(TextEditorCommandLine, MyRegEx)
	Transform, MyEditor, Deref, %MyEditor%
	Run, %MyEditor% "%SavedPositionsFullFilePath%"
return


ReArrangeWindows:
	Reload
return
