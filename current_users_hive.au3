#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Google-Noto-Emoji-Food-Drink-32364-hot-pepper.ico
#AutoIt3Wrapper_Outfile_x64=PowerpointFix.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=PowerpointFix
#AutoIt3Wrapper_Res_Fileversion=1.1.0.0
#AutoIt3Wrapper_Res_ProductName=PowerpointFix
#AutoIt3Wrapper_Res_ProductVersion=1.1.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Carm0
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#NoTrayIcon
Local $q = 0, $a
#include <File.au3>
#include <Debug.au3>
Local $sFILE = 'c:\' & @ComputerName & '.txt'
Local $qbefore, $qafter
If UBound(ProcessList(@ScriptName)) > 2 Then Exit
$usn = @UserName
$a = _FileListToArrayRec('c:\users', "*", $FLTAR_FOLDERS, $FLTAR_NORECUR, $FLTAR_SORT, $FLTAR_NOPATH)
;_DebugArrayDisplay($a)
For $i = 1 To UBound($a) - 1
	If $a[$i] = 'Public' Or $a[$i] = 'All Users' Or $a[$i] = 'Default User' Or $a[$i] = 'altirisad' Then ; profiles to ignore
		;
	Else
		hive()
	EndIf
Next

If @ScriptDir = 'c:\windows\temp' Then
	OnAutoItExit()
EndIf
Exit

hive()

Func hive() ; these changes are made to the default user hive
	$sUSER = $a[$i]
	$pHive = "HKEY_USERS\Temp\" ;Path\Name of Loaded Hive
	$cmd = " /c reg.exe load HKU\Temp C:\users\" & $sUSER & "\NTUSER.dat"
	Run(@ComSpec & $cmd, "", "", @SW_HIDE)
	Sleep(1000)
	$RegTest = RegRead("HKEY_USERS\Temp\Environment", "Temp")
	$q = $q + 1
	If $RegTest <> "" Then
		$qbefore = RegRead($pHive & "Software\Microsoft\Office\16.0\PowerPoint\Options", "DisableSetTopology") ; check value before change
		;MsgBox(0,"before", $qbefore)
		If $qbefore <> 1 Then ; if the value is not 1 then change. THIS IS JUST A LOGIC imput and can be used without
			#Region Microsoft Office Powerpoint "disable extend these displays"
			RegWrite($pHive & "Software\Microsoft\Office\16.0\PowerPoint\Options", "DisableSetTopology", "REG_DWORD", "1")
			;RegDelete($pHive & "Software\Microsoft\Office\16.0\PowerPoint\Options", "DisableSetTopology")
			#EndRegion Microsoft Office Powerpoint "disable extend these displays"
			Sleep(50)
		EndIf
		$qafter = RegRead($pHive & "Software\Microsoft\Office\16.0\PowerPoint\Options", "DisableSetTopology") ; check value after
		Do
			Sleep(800)
			Run(@ComSpec & " /c reg.exe unload HKU\Temp", "", "", @SW_HIDE)
			RegRead("HKEY_USERS\Temp\Environment", "Temp")
		Until @error <> ""
		; writes to a local file on the root of C with the host name the changes to any profile and if not
		If $qbefore = $qafter Then
			FileWrite($sFILE, @ComputerName & ' ' & $sUSER & ' no change in value. Current value =' & $qafter & @CRLF)
			;MsgBox(0, "Black Magic Installer", "Casewatch Settings Installed for Default User Profile")
		ElseIf $qafter <> 1 Then
			FileWrite($sFILE, @ComputerName & ' ' & $sUSER & ' hive unable to modifiy after loading. Current value =' & $qafter & @CRLF)
			;MsgBox(0, "Black Magic Installer", "Could not load the Default User Hive " & $sUSER, 60)
			;Exit ('667')
		ElseIf $qafter = 1 Then
			FileWrite($sFILE, @ComputerName & ' ' & $sUSER & ' hive SUCCESSFUL. Current value =' & $qafter & @CRLF)
		EndIf
	Else
		FileWrite($sFILE, @ComputerName & ' ' & $sUSER & ' hive unable load hive; in use.' & @CRLF); if the HIVE is in use somehow then this will be written
	EndIf
EndFunc   ;==>hive


Func OnAutoItExit()
	Local $iDelay = 2
	Run(@ComSpec & ' /c ping -n ' & $iDelay & ' localhost > nul & del /f /q "' & @ScriptFullPath & '"', '', @SW_HIDE)
EndFunc   ;==>OnAutoItExit

; https://www.autoitscript.com/forum/topic/167388-yet-another-_regwriteallusers/?tab=comments#comment-1224842

