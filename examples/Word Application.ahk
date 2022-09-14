#Include ..\ObjectFromWindow.ahk

#HotIf WinActive("ahk_exe winword.exe")
F12::
{
	try {
		wd := ObjectFromWindow(OBJID_NATIVEOM, "- Word ahk_class OpusApp", "_WwG1")
		wd := wd.Application
		mainSel           := wd.selection
		mainRange         := wd.selection.range

		if mainRange.start = mainRange.end
		{
			wd.ActiveDocument.select()
			mainRange := wd.selection.range
		}

		hwnd := ToolTip(mainSel.Text)
		SetTimer () => WinClose(hwnd), -1500
	}
	catch Error as e
		MsgBox e.What ":`n" e.Message
}
#HotIf