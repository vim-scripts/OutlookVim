' OutlookVim.bas - Edit emails using Vim from Outlook 
' ---------------------------------------------------------------
' Version:       7.0
' Authors:       David Fishburn <dfishburn dot vim at gmail dot com>
' Last Modified: 2012 Sep 25
' Homepage:      http://vim.sourceforge.net/script.php?script_id=???
'
' This VBScript should be installed as a macro inside of Microsoft Outlook.
' It will create two files and ask Vim to edit the file (body of the email)
' and Vim with use the second file to update the body of the email in 
' Outlook with your changes.
'
' This code came from here originally:
'    http://barnson.org/node/295
' Some links might be useful
'    http://office.microsoft.com/en-us/help/HA010429591033.aspx
' Macro Security settings and creating digit certificates
'    http://www.pcreview.co.uk/forums/thread-854025.php
'
' It may be possible to pull the HTMLBody of an email, rather than 
' plaintext as is currently used (in 5.0).  Instead of using 
' item.body, we can reference item.htmlbody.  We can determine this 
' ahead of time by checking (item.BodyFormat = olFormatHTML).
' I found out that for Outlook 2002, there is a constant called, OlBodyFormat.
'     olFormatUnspecified = 0
'     olFormatPlain = 1
'     olFormatHTML = 2
'     olFormatRichText = 3

'
' See http://msdn.microsoft.com/en-us/library/aa171418(v=office.11).aspx
'   - Documentation on HTMLBody Property (if above link no longer works)
'
' Having tried the htmlbody, the html Outlook produced for a 4 character 
' email body was 300 lines long and extremely difficult to read.


Option Explicit

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Sub Edit()
    On Error Resume Next

    'Dim windir As String

    'windir = Environ("WinDir")

    'Shell (windir & "\system32\cscript.exe " & windir & "\system32\launchvim.vbs")
    
    Const TemporaryFolder = 2
    ' Const VIMLocation = "C:\Vim\vim72\gvim.exe"
    

    Dim ol, insp, item, fso, tempfile, tfolder, tname, tfile, cfile, entryID, appRef, x, index
    Dim Vim, vimKeys, vimResponse, vimServerName
    Dim overwrite As Boolean, unicode As Boolean

    ' MsgBox ("Just starting LaunchVim")
    
    Set ol = Application

    Set insp = ol.ActiveInspector
    If insp Is Nothing Then
        ' MsgBox ("No active inspector")
        Exit Sub
    End If
    
    Set item = insp.CurrentItem
    If item Is Nothing Then
        ' MsgBox ("No current item")
        Exit Sub
    End If
    
    ' MsgBox ("type:" & TypeName(item))
    ' MsgBox ("entryID type:" & TypeName(item.entryID))
    If item.entryID = "" Then
        ' If there is no EntryID, Vim will not be able to update
        ' the email during the save.
        ' Saving the item in Outlook will generate an EntryID
        ' and allow Vim to edit the contents.
       item.Save
       If Err.Number <> 0 Then
           ' Clear Err object fields.
           ' Err.Clear
           MsgBox ("OutlookVim: Cannot edit with Vim, could not save item:" & vbCrLf & Err.Description)
           Exit Sub
       End If
    End If

    ' Create an instance of Vim (if one does not already exist)
    Set Vim = CreateObject("Vim.Application")
    If Vim Is Nothing Then
        MsgBox ("OutlookVim: Could not create a Vim OLE object, please ensure Vim has been installed." & vbCrLf & "Inside Vim, run this command, it should echo 1," & vbCrLf & " :echo has('ole')")
        Exit Sub
    End If

    vimResponse = Vim.Eval("(exists('g:loaded_outlook')?1:0)")
    If vimResponse = 0 Then
        MsgBox ("OutlookVim: Please install the plugin OutlookVim from www.vim.org to continue" & vbCrLf)
        Exit Sub
    End If
        
    vimResponse = Vim.Eval("(exists('g:outlook_servername')?1:0)")
    If vimResponse = 0 Then
        MsgBox ("OutlookVim: Please ensure g:outlook_servername is set in version 3.0 or above of OutlookVim" & vbCrLf)
        Exit Sub
    End If
        
    vimServerName = Vim.Eval("g:outlook_servername")
    If vimServerName <> "" Then
        vimResponse = Vim.Eval("match(serverlist(), '\<" & vimServerName & "\>')")
        If vimResponse = -1 Then
            MsgBox ("OutlookVim: There are no Vim instances running named: " & vbCrLf & vimServerName & vbCrLf & vbCrLf _
                    & "Found these Vim instances: " & vbCrLf & Vim.Eval("serverlist()") & vbCrLf _
                    & "Please start a new Vim instance using " & vbCrLf & "    'gvim --servername " & vimServerName & "'" _
                    )
            Exit Sub
        End If
    End If
        
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        MsgBox ("OutlookVim: Could not create a file system object." & vbCrLf & _
                "Please ensure the Windows Script Host utility has been installed and registered correctly." & vbCrLf & _ 
                "You may want to follow the Upgrade links to correct the problem." & vbCrLf & _
                "See http://msdn.microsoft.com/en-us/library/9bbdkx3k.aspx" _
                )
        Exit Sub
    End If

    Set tfolder = fso.GetSpecialFolder(TemporaryFolder)
    tname = fso.GetTempName
    tname = Left(tname, (Len(tname) - 3)) & "outlook"
    ' MsgBox ("Temp folder:" & tfolder.ShortPath)
    
    ' Write out the body of the message to a temp file
    overwrite = False
    unicode = False
    ' MsgBox ("InternetCodePage:" & item.InternetCodepage)
    
    ' Check if there are any unicode characters in the body
    If InStr(1, item.body, Chr(0), vbBinaryCompare) > 0 Then
        ' MsgBox ("Setting unicode")
        unicode = True
    End If
    
    Set tfile = tfolder.CreateTextFile(tname, overwrite, unicode)
    ' MsgBox ("Created body file:" & tfile.Name)
    
    tfile.Write (Replace(item.body, Chr(13) & Chr(10), Chr(10)))
    If Err.Number <> 0 Then
        ' Clear Err object fields.
        ' Err.Clear
        MsgBox ("OutlookVim: Could not create email file [" & tfolder.ShortPath & "\\" & tname & "] " & vbCrLf & Err.Description)
        tfile.Close
        fso.DeleteFile (tfolder.ShortPath & "\\" & tname)
        ' Quit will close Outlook
        ' Quit
        Exit Sub
    End If
    'Try
    '    tfile.Write (Replace(body, Chr(13) & Chr(10), Chr(10)))
    ' Catch ex As Exception
    'Catch
    '    tfile.Write (body)
    '    MsgBox ("Could not convert CRLFs:" & vbCrLf & ex.Message)
    'End Try

    tfile.Close
    ' MsgBox ("tfile:" & tname)

    ' Write out the control file so the outlookvim javascript file
    ' can tell Outlook which inspector to refresh
    Set cfile = tfolder.CreateTextFile(tname & ".ctl")
    ' MsgBox ("EntryID:" & Replace(item.entryID, Chr(13) & Chr(10), Chr(10)))
    cfile.Write (Replace(item.entryID, Chr(13) & Chr(10), Chr(10)))
    If Err.Number <> 0 Then
        ' Clear Err object fields.
        ' Err.Clear
        MsgBox ("OutlookVim: Could not create control file [" & tfolder.ShortPath & "\\" & tname & "] " & vbCrLf & Err.Description)
        cfile.Close
        fso.DeleteFile (tfolder.ShortPath & "\\" & tname)
        fso.DeleteFile (tfolder.ShortPath & "\\" & tname & ".ctl")
        ' Quit will close Outlook
        ' Quit
        Exit Sub
    End If
    cfile.Close
    ' MsgBox ("id:" & item.EntryID)
    ' MsgBox tfolder.ShortPath & "\\" & tname
        
    
    vimKeys = ":call Outlook_EditFile( '" & tfolder.ShortPath & "\\" & tname & "', '"
    If unicode Then
        vimKeys = vimKeys & "utf-16"
    End If
    vimKeys = vimKeys & "' )<Enter>"
    
    ' MsgBox (vimKeys)
    ' Use Vim's OLE feature to edit the email
    ' This allows us to re-use the same Vim for multiple emails
    Vim.SendKeys vimKeys
    
    ' Force the Vim to the foreground
    Vim.SetForeground

Finished:
End Sub

Public Sub ExecCmd(cmdline$)
    
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ReturnValue As Integer

    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    ' Start the shelled application:
    ReturnValue = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    ' Wait for the shelled application to finish:
    Do
        ReturnValue = WaitForSingleObject(proc.hProcess, 0)
        DoEvents
    Loop Until ReturnValue <> 258
    
    ReturnValue = CloseHandle(proc.hProcess)
    
End Sub
