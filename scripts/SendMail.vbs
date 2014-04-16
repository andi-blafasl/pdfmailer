' SendMail.vbs script
' License: GPL
' Version: 1.0
' Date: 23.01.2014
' Author: Andreas Hoesl
' Comment: This scripts sends the Output of PDF-Creator as an attachment to "objArgs(1)"@domain.tld
'          For more information about the CDO-Object and example code see: http://www.paulsadowski.com/wsh/cdo.htm

'##### DO NOT CHANGE ANYTHING BELOW HERE! GO DOWN TO CONFIG SETTINGS BELOW! #####
Option Explicit
On Error Resume Next
Dim Exit_Code
Exit_Code=0
Const AppTitle = "PDFMailer - SMTP Send Mail"
Const EVENTCREATE = "\System32\eventcreate.exe"

Dim objEnv, WshShell
Set WshShell = CreateObject("WScript.Shell")
Set objEnv = WshShell.Environment("Process")


'Get Commandline Arguments
Dim objArgs, attachment, recipient
Set objArgs = WScript.Arguments

If objArgs.Count <> 2 Then
  WriteEventLog("Falsche Anzahl an Parametern, benötigt werden: <OutputFile> <Author>")
  WScript.Quit(1)
End If

attachment = objArgs(0) 'Dokument to attach
recipient = objArgs(1) 'Username to send mail to

'Generate FileName parts
Dim objFSO, objFile, attachment_name, attachment_type, attachment_dir
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(attachment)
If Err.Number <> 0 then
  'File not found or someting
  WriteEventLog("Dokument konnt nicht geöffnet werden!" & vbCrLf &_
                "Dokument: " & attachment )
  WScript.Quit(1)
End If
attachment_name = objFSO.GetFileName(objFile)
attachment_type = objFSO.GetExtensionName(objFile)
attachment_dir = objFSO.GetParentFolderName(objFile)

'define CDO-Object constants
Const cdoSendUsingPickup = 1 'only usable if SMTP-Services are installed on same host
Const cdoSendUsingPort = 2 'Must use this to use Delivery Notification
'SMTP Auth config for "smtpauthenticate" Variable
Const cdoAnonymous = 0
Const cdoBasic = 1 'clear text
Const cdoNTLM = 2 'NTLM
'Delivery Status Notifications
Const cdoDSNDefault = 0 'None
Const cdoDSNNever = 1 'None
Const cdoDSNFailure = 2 'Failure
Const cdoDSNSuccess = 4 'Success
Const cdoDSNDelay = 8 'Delay
Const cdoDSNSuccessFailOrDelay = 14 'Success, failure or delay

'####################################
'##### CONFIGURE SETTINGS HERE! #####
'####################################
Dim smtp_host, smtp_port, smtp_auth, smtp_user, smtp_passwd, smtp_ssl, smtp_timeout
'SMTP Configuration:
smtp_host="exchange.chrmayr.lan"
smtp_port=25
smtp_auth=0 'set to 1 if username and password are required for smtp_auth
smtp_user=""
smtp_passwd=""
smtp_ssl=FALSE 'Use SSL for the connection (False or True)
smtp_timeout=60 'Connection Timeout in seconds

Dim mail_from_display, mail_from_address, mail_to_domain, mail_subj, mail_html, mail_body, mail_dsn
'MAIL Parameter:
mail_from_display=Ucase(attachment_type) & " Mailer"
mail_from_address="pdfmailer@mayr.de"
mail_to_domain="@mayr.de" 'domain to append to usernames, with @ !
mail_subj=Ucase(attachment_type) & " Dokument: " & attachment_name
mail_html=1 'set to 1 if mail_body uses HTML-Tags
mail_body="<p>Anbei ihr " & Ucase(attachment_type) & "-Dokument.<p>" & vbCRLF & _
          vbCRLF & _
          "<p>Bei Fragen oder Problemen wenden Sie sich bitte an ihren Vor-Ort-Service-Mitarbeiter oder per Mail an <a href='mailto:555@mayr.de'>EDV-Hotline, 555</a><p>" & vbCRLF
mail_dsn=cdoDSNDefault 'Delivery Status Notification, see CDO-Object Constants above

'##### DO NOT CHANGE ANYTHING BELOW HERE! #####

Dim objMsg, strBody 
set objMsg = CreateObject("CDO.Message")

With objMsg.Configuration.Fields
  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtp_host
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtp_port
  If smtp_auth = 1 then
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = smtp_user
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = smtp_passwd
  End If
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = smtp_ssl
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = smtp_timeout
  .Update
End With

With objMsg
  .From = mail_from_display & " <" & mail_from_address & ">"
  .To = recipient & mail_to_domain
  .Subject = mail_subj
  If mail_html = 1 then 
    .HTMLBody = mail_body
  Else
    .TextBody = mail_body
  End If
  .Addattachment attachment
  If mail_dsn > 0 then
    .Fields("urn:schemas:mailheader:disposition-notification-to") = mail_from_address
    .DSNOptions = mail_dsn
  End If
  .Fields.update
End With

objMsg.Send
If Err.Number = 0 then
  set objMsg = nothing
  'If Message sent successfull, delete temporary File
  'MsgBox "Deleting Folder: " & attachment_dir & "\", vbExclamation, AppTitle
  objFSO.DeleteFolder attachment_dir
  Exit_Code=0
Else
  'If Mail send error, do nothing
  WriteEventLog("Mailversand fehlgeschlagen!" & vbCrLf &_
               "Empfänger: " & recipient & mail_to_domain & vbCrLf &_
               "Dokument: " & attachment )
  Exit_Code=1
End If

WScript.Quit(Exit_Code)

'******************************************************************************

Sub WriteEventLog(strMessage)
  'Write custom message and information from VBScript Err object to Eventlog.
  Dim strError
  
  strError = strMessage & VbCrLf & VbCrLf &_
	"Laufzeit Informationen:" & VbCrLf &_
	"Attachment      : " & attachment & VbCrLf &_
	"Recipient       : " & recipient & VbCrLf &_
	"Attachment_name : " & attachment_name & VbCrLf &_
	"Attachment_type : " & attachment_type & VbCrLf &_
	"Attachment_dir  : " & attachment_dir & VbCrLf & VbCrLf &_
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

