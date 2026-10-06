Set WshShell = CreateObject("WScript.Shell")
' Chay file bat ma khong hien thi cua so den (0 = Hide window)
WshShell.Run chr(34) & "E:\docx_standardizer\Chay_May_Chu_Nha_Truong.bat" & Chr(34), 0
Set WshShell = Nothing
