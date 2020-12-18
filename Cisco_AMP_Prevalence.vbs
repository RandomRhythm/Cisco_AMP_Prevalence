'v 1.3 updated for Cisco URLs. added longer wait time. Added date time
'v 1.2 error handling for IE close
'v 1.1 get computer name(s). Key off hash not file name. Longer delay between page loads. Longer delay for Unknown error and add suggestion to fix.
'Requires a Windows system with IE 11

Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1
Dim DictData
Set DictData = CreateObject("Scripting.Dictionary")
CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strOutPutFile = CurrentDirectory & "\FA_report.txt"

GetFA_Prevalence

for each item in DictData 'FA files
  
    strOutput = strOutput & DictData.Item(Item) & "|" & Item & vbcrlf
next
LogData strOutPutFile, strOutput,False

msgbox "done"

Sub GetFA_Prevalence()

Dim WshShell, objOutputFile, PageCache
Dim docuObj
Dim boolCWRexit
Dim intFApageNumCount
boolCWRexit = False
CurrentDirectory = GetFilePath(wscript.ScriptFullName)
Set WshShell = WScript.CreateObject("WScript.Shell")
Set oIE = CreateObject("InternetExplorer.Application")
oIE.Visible = True
Wscript.Sleep 100
Set objShell = CreateObject("WScript.Shell")
objShell.AppActivate "Internet Explorer"
Wscript.Sleep 100
oIE.Navigate "about:blank"
on error resume next
iHeight = oie.document.parentWindow.screen.height
if err.number <> 0 then boolCWRexit = True
on error goto 0 
if boolCWRexit = False then
  iWidth = oie.document.parentWindow.screen.width
  oIE.Width = iWidth
  oIE.Height = iHeight - 28

  oIE.Navigate "https://console.amp.cisco.com"

  Do While oIE.Busy Or (oIE.READYSTATE <> 4)
      Wscript.Sleep 100
  Loop
  if instr(PageCache, "class=" & chr(34) & "imn-input-login" & chr(34)) then _
   wscript.echo "please log into Cisco AMP"
  
  loope_count = 1
  do while loop_exit = false
    on error resume next
    oIE.Navigate "https://console.amp.cisco.com/executables?page=" & loope_count
    if err.number <> 0 then
      loop_exit = True
      exit sub
    end if
    on error goto 0 
    Do While oIE.Busy Or (oIE.READYSTATE <> 4)
        Wscript.Sleep 100
    Loop
    Wscript.Sleep 11000
    on error resume next
    PageCache = oIE.document.body.innerHTML
    'logdata currentdirectory & "\FA_PageCache.txt",PageCache, false
    if err.number = 0 then
    on error goto 0
    
	  'msgbox instr(PageCache, "class=" & chr(34) & "page-count input-group-addon" & chr(34) & " data-pages=" & chr(34))
      if instr(PageCache, "class=" & chr(34) & "page-count input-group-addon" & chr(34) & " data-pages=" & chr(34)) then
        intFApageNumCount = getdata(PageCache, chr(34), "class=" & chr(34) & "page-count input-group-addon" & chr(34) & " data-pages=" & chr(34))
        if isnumeric(intFApageNumCount) then
          if int(intFApageNumCount) > 1 then
            ParseFA_Prevalence(PageCache)
            if loope_count = int(intFApageNumCount) then
              loop_exit = True
              exit do
            else
              loope_count = loope_count +1 
            end if
          else
            ParseFA_Prevalence(PageCache)
            loop_exit = True
            exit do
          end if
        else
          msgbox "error getting data from Cisco AMP"
          loop_exit = True
        end if
      elseif instr(PageCache, "class=" & chr(34) & "imn-input-login" & chr(34)) then
        msgbox "Session was logged out. Please log back in"
      else
        intAnswer = _
            Msgbox("Unknown error. Do you want to try again? Maybe try refreshing IE (F5) during the delay. If page does not progress try visiting the URL in the address bar again (select and enter)", _
        vbYesNo, "AMP Prevalence")
        If intAnswer = vbYes Then
          wscript.sleep 8000
        else
        loop_exit = True
        end if
        wscript.sleep 3000
      end if
    on error goto 0
    else
		msgbox "error reading HTML from IE!"
    end if
  loop  
end if  
  Set HTMLForm = nothing
oIE.quit
set oIE = Nothing

end Sub

Function GetFilePath (ByVal FilePathName)
found = False

Z = 1

Do While found = False and Z < Len((FilePathName))

 Z = Z + 1

         If InStr(Right((FilePathName), Z), "\") <> 0 And found = False Then
          mytempdata = Left(FilePathName, Len(FilePathName) - Z)
          
             GetFilePath = mytempdata

             found = True

        End If      

Loop

end Function


Function GetData(contents, ByVal EndOfStringChar, ByVal MatchString)
MatchStringLength = Len(MatchString)
x= 0

do while x < len(contents) - (MatchStringLength +1)

  x = x + 1
  if Mid(contents, x, MatchStringLength) = MatchString then
    'Gets server name for section
    for y = 1 to len(contents) -x
      if instr(Mid(contents, x + MatchStringLength, y),EndOfStringChar) = 0 then
          TempData = Mid(contents, x + MatchStringLength, y)
        else
          exit do  
      end if
    next
  end if
loop
GetData = TempData
end Function



Function ParseFA_Prevalence(strFAHTML)
Dim strTmpKey
Dim strTmpValman
Dim StrTmpComputerCount
Dim ArrayFA_HTML
'logdata currentdirectory & "\FA_LineArg.txt", strFAHTML, false
'msgbox instr(strFAHTML, "Fingerprint ")
ArrayFA_HTML = split(strFAHTML, "Fingerprint ")
'logdata currentdirectory & "\FA_LineCache.txt", ubound(ArrayFA_HTML) & "-------------------------------------------------------------------------", false
for each strFA_HTMLine in ArrayFA_HTML
		
        if instr(strFA_HTMLine, "<td class=" & chr(34) & "sha-context-menu"& chr(34)) then
		  logdata currentdirectory & "\FA_LineCache.txt",strFA_HTMLine, false
          strTmpKey = GetData(strFA_HTMLine, chr(34), "data-sha=" & chr(34))
          strTmpValman = GetData(strFA_HTMLine, chr(34), "data-file-name=" & CHR(34))
          strDate = GetData(strFA_HTMLine, "<", chr(34) & "date" & CHR(34) & ">")
          strTime = GetData(strFA_HTMLine, "<", "time" & CHR(34) & ">")
          strTmpKey  = strTmpKey & "|" & GetData(strFA_HTMLine, chr(34), "only executed on <i class=" & Chr(34) & "fa fa-windows" & Chr(34) & "></i> <strong title=" & Chr(34))
          'msgbox "strTmpValman=" & strTmpValman
          'msgbox "strTmpKey=" & strTmpKey
        end if
        if instr(strFA_HTMLine, "data-count=") then
            StrTmpComputerCount = getdata(strFA_HTMLine, chr(34), "data-count=" & CHr(34))
            if DictData.exists(strTmpKey & "|" & strTmpValman) = true then
              'why is it showing the same file again?!?!??!?!
              msgbox strTmpKey & "|" & strTmpValman
            else
              DictData.add strTmpKey & "|" & strTmpValman, StrTmpComputerCount & "|" & strDate & "|" & strTime
			  logdata currentdirectory & "\FA_outD.txt", strTmpKey & "|" & strTmpValman & "|" & StrTmpComputerCount & "|" & strDate & "|" & strTime, false
            end if
            'msgbox "StrTmpComputerCount=" & StrTmpComputerCount
            strTmpValman = ""
            strTmpKey = ""
            StrTmpComputerCount = ""
        else
			logdata currentdirectory & "\FA_mLineCache.txt",strFA_HTMLine, false
        end if
next

end function



function LogData(TextFileName, TextToWrite,EchoOn)
Dim strTmpFilName1
Dim strTmpFilName2
strTmpFilName1 = right(TextFileName, len(TextFileName) - instrrev(TextFileName,"\"))
strTmpFilName2 = replace(strTmpFilName1,"/",".")
'TextFileName = replace(TextFileName,"\",".")
strTmpFilName2 = replace(strTmpFilName2,":",".")
strTmpFilName2 = replace(strTmpFilName2,"*",".")
strTmpFilName2 = replace(strTmpFilName2,"?",".")
strTmpFilName2 = replace(strTmpFilName2,chr(34),".")
strTmpFilName2 = replace(strTmpFilName2,"<",".")
strTmpFilName2 = replace(strTmpFilName2,">",".")
strTmpFilName2 = replace(strTmpFilName2,"|",".")
TextFileName = replace(TextFileName,strTmpFilName1,strTmpFilName2)

Set fsoLogData = CreateObject("Scripting.FileSystemObject")
if EchoOn = True then wscript.echo TextToWrite
  If fsoLogData.fileexists(TextFileName) = False Then
      'Creates a replacement text file 
      fsoLogData.CreateTextFile TextFileName, True
  End If
on error resume next
Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
if err.number <> 0 then
  msgbox "Error writting to " & TextFileName & " perhaps the file is locked?"
  err.number = 0
  Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
  if err.number <> 0 then exit function
end if

on error goto 0
WriteTextFile.WriteLine TextToWrite
WriteTextFile.Close
Set fsoLogData = Nothing
End Function


Function GetFilePath (ByVal FilePathName)
found = False

Z = 1

Do While found = False and Z < Len((FilePathName))

 Z = Z + 1

         If InStr(Right((FilePathName), Z), "\") <> 0 And found = False Then
          mytempdata = Left(FilePathName, Len(FilePathName) - Z)
          
             GetFilePath = mytempdata

             found = True

        End If      

Loop

end Function


function fnShellBrowseForFolderVB()
    dim objShell
    dim ssfWINDOWS
    dim objFolder
    
    ssfWINDOWS = 36
    set objShell = CreateObject("shell.application")
        set objFolder = objShell.BrowseForFolder(0, "Example", 0, ssfDRIVES)
            if (not objFolder is nothing) then
               set oFolderItem = objFolder.items.item
               fnShellBrowseForFolderVB = oFolderItem.Path 
            end if
        set objFolder = nothing
    set objShell = nothing
end function
