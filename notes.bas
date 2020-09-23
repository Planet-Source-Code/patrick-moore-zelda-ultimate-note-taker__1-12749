Attribute VB_Name = "notes32"
Option Explicit
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const EM_GETLINECOUNT = &HBA

Function GetLineCount(txtBox As textbox)
'Use Windows API to send a message to
'the specified textbox.

'That API call returns the number of lines
'in the textbox
GetLineCount = SendMessage(txtBox.hWnd, EM_GETLINECOUNT, 0&, 0&)
End Function

Function FileOpen(Filenamer As String)
Dim TextString As String
On Error GoTo errhandle

'Open the file
Open Filenamer For Input As #1

'Get text the length of the file,
'from the file
TextString = Input(LOF(1), #1)

'Close the file
Close #1


FileOpen = TextString
Exit Function
errhandle:
FileOpen = ""
Close #1
End Function


Function GetFileSize(pSource As String) As Long
Dim vFileNumber As Integer

'Get the free file number
vFileNumber = FreeFile

'Open the file
Open pSource For Binary Access Read As vFileNumber
'Get the length of the file
GetFileSize = LOF(vFileNumber)

'Close the file
Close vFileNumber
End Function

Function GetWordCount(txt As String) As Integer
Dim WordCount As Integer, SpcChar As Integer

'Make sure the last character is a space
If Right(txt, 1) <> " " Then txt = txt & " "

Do
    'Find the space character
    SpcChar = InStr(txt, " ")
    
    'If the next char is a space, ignore it
    If Mid(txt, SpcChar + 1, 1) = " " Then
        SpcChar = SpcChar + 1
    End If
    
    'Add one to the total words in the string
    'so far
    WordCount = WordCount + 1
    
    'Trim the string after the space
    txt = Mid(txt, SpcChar + 1, Len(txt))
    
    'Continue finding words
Loop Until InStr(txt, " ") = 0

GetWordCount = WordCount
End Function
