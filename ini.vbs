Function GetINIString(Section, KeyName, Default, FileName)
  Dim INIContents, PosSection, PosEndSection, sContents, Value, Found
  
  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    'Separate section contents
    sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)

    If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
      Found = True
      'Separate value of a key.
      Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
    Else
      Found = False
      Value = "0"
  End If
  Else
    Found = False
    Value = "0"
  End If
  If isempty(Found) Then Value = 0
  GetINIString = Value
End Function

'Separates one field between sStart And sEnd
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
  Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(sFrom, PosB, PosE - PosB)
  End If
End Function

'File functions
Function GetFile(ByVal FileName)
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
  'Go To windows folder If full path Not specified.
  If InStr(FileName, ":\") = 0 And Left (FileName,2)<>"\\" Then 
    FileName = FS.GetSpecialFolder(0) & "\" & FileName
  End If
  On Error Resume Next

  GetFile = FS.OpenTextFile(FileName).ReadAll
End Function

Set oArgs=WScript.Arguments ' tableau d'arguments

Dim section: section = oArgs(0)
Dim key: key = oArgs(1)
Dim configFile: configFile = oArgs(2)

If oArgs.Count > 1 Then
  WScript.Echo GetINIString(section, key, "", configFile)
Else
  WScript.Echo "0"
End If