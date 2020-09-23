Attribute VB_Name = "Module1"
Public Function Crypt(File As String, Key As String) As Boolean
'On Error GoTo cerror
Dim iFile As Integer
Dim Buffer As String
Dim eBuffer As String
Buffer = Space(0)
iFile = FreeFile

Open File For Binary As #iFile
    Buffer = Space(Len(Key))
    For x = 1 To LOF(iFile) / Len(Key)
        Get #iFile, , Buffer
            For y = 1 To Len(Key)
                eBuffer = eBuffer & Chr(Asc(Mid(Buffer, y, 1)) Xor Asc(Mid(Key, y, 1)))
                DoEvents
            Next
            DoEvents
            Put #iFile, Loc(iFile) - Len(Key) + 1, eBuffer
            eBuffer = ""
            fMainForm.sbStatusBar.Panels(1).Text = "Encrypting " & Str(CInt(Loc(iFile) / LOF(iFile) * 100)) & "%"
    Next
    
    eBuffer = ""
  If Loc(iFile) < LOF(iFile) Then
    Buffer = Space(LOF(iFile) - Loc(iFile))
    Get #iFile, , Buffer
    For y = 1 To Len(Buffer)
                eBuffer = eBuffer & Chr(Asc(Mid(Buffer, y, 1)) Xor Asc(Mid(Key, y, 1)))
    Next
    Put #iFile, Loc(iFile) - Len(Buffer) + 1, eBuffer
End If

Close
Crypt = True
fMainForm.sbStatusBar.Panels(1).Text = "Status"
End Function

Public Function ViewCrypt(File As String, Key As String) As Boolean
'On Error GoTo cerror
Dim iFile As Integer, iFile2 As Integer
Dim Buffer As String
Dim eBuffer As String
Buffer = Space(0)
iFile = FreeFile


Open File For Binary As #iFile
iFile2 = FreeFile
Open "c:\windows\temp.cfg" For Binary As #iFile2
    Buffer = Space(Len(Key))
    For x = 1 To LOF(iFile) / Len(Key)
        Get #iFile, , Buffer
            For y = 1 To Len(Key)
                eBuffer = eBuffer & Chr(Asc(Mid(Buffer, y, 1)) Xor Asc(Mid(Key, y, 1)))
                DoEvents
            Next
            DoEvents
            Put #iFile2, , eBuffer
            eBuffer = ""
            fMainForm.sbStatusBar.Panels(1).Text = "Decrypting " & Str(CInt(Loc(iFile) / LOF(iFile) * 100)) & "%"
    Next
    
    eBuffer = ""
  If Loc(iFile) < LOF(iFile) Then
    Buffer = Space(LOF(iFile) - Loc(iFile))
    Get #iFile, , Buffer
    For y = 1 To Len(Buffer)
                eBuffer = eBuffer & Chr(Asc(Mid(Buffer, y, 1)) Xor Asc(Mid(Key, y, 1)))
    Next
    Put #iFile2, , eBuffer
End If

Close

fMainForm.sbStatusBar.Panels(1).Text = "Status"

End Function
