' MayrPreprocess.vbs script
' License: GPL
' Version: 1.0
' Date: 24.01.2014
' Author: Andreas Hoesl
' Comment: This script checks if the printed document is a pdf-file and replaces the spoolfile with a blank page.
'          If it is not a pdf-file, Bookmarks are added for each page.
' ChangeLog:
' 0.1: inital version
Option Explicit
On Error Resume Next

Const AppTitle = "PDFMailer - Preprocessor"
Const ForReading = 1, ForAppending = 8
Const EVENTCREATE = "\System32\eventcreate.exe"

Dim objArgs, objFSO, ObjFile, f, pages, i, objEnv, WshShell, fext
Dim TempFileName, SpoolFileDir, SessionID, WinStation, UserName, ClientComputer, SpoolFileName, PrinterName, JobId, JobCounter, DocTitle, SpoolFile, MtdFileSplit, MtdFile, PageBlanked
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set objEnv = WshShell.Environment("Process")

Set objArgs = WScript.Arguments

If objArgs.Count <> 1 Then
	WriteEventLog("Falsche Anzahl an Parametern, benötigt werden: <TempFilename>")
	WScript.Quit
End If

TempFileName = objArgs(0)
Set objFile = objFSO.GetFile(TempFileName)
If Err.Number <> 0 then
	WriteEventLog("Fehler beim Zugriff auf " & TempFileName &" !")
	WScript.Quit(1)
End If
SpoolFileDir = objFSO.GetParentFolderName(objFile)

SessionId = ReadIni( objArgs(0), "1", "SessionID" )
WinStation = ReadIni( objArgs(0), "1", "WinStation" )
UserName = ReadIni( objArgs(0), "1", "UserName" )
ClientComputer = ReadIni( objArgs(0), "1", "ClientComputer" )
SpoolFileName = ReadIni( objArgs(0), "1", "SpoolFileName" )
PrinterName = ReadIni( objArgs(0), "1", "PrinterName" )
JobId = ReadIni( objArgs(0), "1", "JobId" )
JobCounter = ReadIni( objArgs(0), "1", "JobCounter" )
DocTitle = ReadIni( objArgs(0), "1", "DocumentTitle" )

SpoolFile = SpoolFileDir & "\" & SpoolFileName
MtdFileSplit = Split(SpoolFileName, ".")
MtdFile = SpoolFileDir & "\" & MtdFileSplit(0) & ".mtd"

fext = Split(DocTitle, ".")
if fext(UBound(fext)) = "pdf" then
  PageBlanked = "yes"
else
  PageBlanked = "no"
End If

WriteLogFile("start")

If fext(UBound(fext)) = "pdf" Then
	objFSO.CopyFile "C:\\Program Files (x86)\\PDFMailer\\blank.ps", SpoolFile, true
	If Err.Number <> 0 then
		WriteEventLog("blank.ps konnte nicht kopiert werden!")
                WScript.Quit(1)
        End If
Else
        pages = GetCountOfPagesFromPostscriptfile(SpoolFile)
	Set f = objFSO.OpenTextFile(MtdFile, ForAppending, True)
	If Err.Number <> 0 then
		WriteEventLog(SpoolFile & " konnte nicht geöffnet werden!")
                WScript.Quit(1)
        End If
        f.writeline ""
	f.writeline "[/Title (" & DocTitle & ") /Page " & 1 & " /View [/XYZ null null 1] /Count " & pages & " /OUT pdfmark"
	If Err.Number <> 0 then
		WriteEventLog("Fehler beim schreiben der Bookmark Übersicht in " & SpoolFile & " !")
                WScript.Quit(1)
        End If
	For i=1 to pages
		f.writeline "[/Page " & i & " /View [/XYZ null null 1] /Title (Seite " & i & ") /OUT pdfmark"
		If Err.Number <> 0 then
			WriteEventLog("Fehler beim schreiben der Bookmark für Seite " & i & " in " & SpoolFile & " !")
			WScript.Quit(1)
		End If
	Next
	f.WriteLine "%%EOF"
	If Err.Number <> 0 then
		WriteEventLog("Fehler beim schreiben von EOF nach Bookmark in " & SpoolFile & "!")
                WScript.Quit(1)
        End If
	f.Close
End If

WriteLogFile("done")

'******************************************************************************

Sub WriteEventLog(strMessage)
  'Write custom message and information from VBScript Err object to Eventlog.
  Dim strError
  
  strError = strMessage & VbCrLf & VbCrLf &_
	"Laufzeit Informationen" & VbCrLf &_
	"TempFileName  : " & TempFileName & VbCrLf &_
	"SessionId     : " & SessionID & VbCrLf &_
	"WinStation    : " & WinStation & VbCrLf &_
	"UserName      : " & UserName & VbCrLf &_
	"ClientComputer: " & ClientComputer & VbCrLf &_
	"SpoolFileName : " & SpoolFileName & VbCrLf &_
	"SpoolFileDir  : " & SpoolFileDir & VbCrLf &_
	"PrinterName   : " & PrinterName & VbCrLf &_
	"JobId         : " & JobId & VbCrLf &_
	"JobCounter    : " & JobCounter & VbCrLf &_
	"DocTitle      : " & DocTitle & VbCrLf &_
	"MtdFile       : " & DocTitle & VbCrLf & VbCrLf &_
	"Windows Error Info:" & VbCrLf &_
	"Number (dec) : " & Err.Number & VbCrLf &_
	"Number (hex) : 0x" & Hex(Err.Number) & VbCrLf &_
	"Description  : " & Err.Description & VbCrLf &_
	"Source       : " & Err.Source
  Err.Clear
  
  WshShell.Run objEnv("SYSTEMROOT") & EVENTCREATE & " /L Application  /T ERROR /SO " & Chr(34) & "PDF-Drucker (Fehler)" & Chr(34) &_
    		" /ID 111 /D " & Chr(34) & "PDF-Drucker-Skript (" & WScript.ScriptFullName & ")" & vbCrLf & vbCrLf &_
    		strError &_
    		Chr(34),0,True

End Sub

Sub WriteLogFile(strMsgType)
  'Write Status message to logfile
  Dim objFSOLog, LogFile, f, MsgText
  
  strMsgType = lcase(strMsgType)
  LogFile = "d:\temp\PDFMailer.log"
  
  Set objFSOLog = CreateObject("Scripting.FileSystemObject")
  Set f = objFSOLog.OpenTextFile(LogFile, ForAppending, True)
  If Err.Number <> 0 then
    WriteEventLog("LogFile " & LogFile & " konnte nicht geöffnet werden!")
    Exit Sub
  End If
  
  Select Case strMsgType
  	Case "start"
  		MsgText = Date & " - " & Time & AppTitle & " Start --------------" & vbCrLf &_
  		              "                        TempFileName: " & Chr(34) & TempFileName & Chr(34) &_
  		                                     " PageBlanked: " & Chr(34) & PageBlanked & Chr(34) &_
  		                                     " UserName: " & Chr(34) & UserName & Chr(34) &_
  		                                     " ClientComputer: " & Chr(34) & ClientComputer & Chr(34) &_
  		                                     " PrinterName: " & Chr(34) & PrinterName & Chr(34) &_
  		                                     " DocTitle: " & Chr(34) & DocTitle & Chr(34)
  	Case "done"
  	        MsgText = Date & " - " & Time & AppTitle & " Done ---------------" & vbCrLf
  end select
  
  f.writeline MsgText
  If Err.Number <> 0 then
    WriteEventLog("Fehler beim schreiben in das LogFile " & LogFile & " !")
    f.Close
    objFSOLog = nothing
    Exit Sub
  End If
  
  f.Close
  set objFSOLog = nothing
  
End Sub

Private Function GetCountOfPagesFromPostscriptfile(PostscriptFile)
 Dim objFSO, f, fstr, pp
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 Set f = objFSO.OpenTextFile(PostscriptFile, ForReading, True)
 fstr = f.ReadAll
 f.Close
 pp = InstrRev(fstr, "%%Pages:", -1, 1)
 If pp <= 0 Then
  GetCountOfPagesFromPostscriptfile = 1
  Exit Function
 End If
 pp = Instr(pp, fstr," ", 1)
 If pp <= 0 Then
  GetCountOfPagesFromPostscriptfile = 1
  Exit Function
 End If
 fstr = Trim(Mid(fstr,pp))
 fstr = Replace(fstr, chr(10), " ", 1, -1, 1)
 fstr = Replace(fstr, chr(13), " ", 1, -1, 1)
 pp = Instr(1, fstr," ", 1)
 If pp <= 0 Then
  GetCountOfPagesFromPostscriptfile = 1
  Exit Function
 End If
 fstr=mid(fstr,1,pp-1)
 If Not IsNumeric(fstr) Then
  fstr = 1
 End If
 GetCountOfPagesFromPostscriptfile = fstr
End Function

Function ReadIni( myFilePath, mySection, myKey )
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False, -2 ) '-2 is needed because the .inf from PDF-Creator is UTF-16 Encoded
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )
            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WriteEventLog(strFilePath & " doesn't exists. Exiting...")
        Wscript.Quit 1
    End If
End Function