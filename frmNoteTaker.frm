VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNoteTaker 
   AutoRedraw      =   -1  'True
   Caption         =   "Note Taker - 0 notes"
   ClientHeight    =   4335
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNoteTaker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   663
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   4680
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
   End
   Begin VB.ListBox lstNotes 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtNotes 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   3855
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox txtDocument 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   3855
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmNoteTaker.frx":1042
      Top             =   240
      Width           =   5055
   End
   Begin VB.Menu mnuDocument 
      Caption         =   "Document"
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuNotes 
      Caption         =   "Notes"
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuClearNotes 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlpha 
         Caption         =   "Alphabetize"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmNoteTaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumNotes As Integer

Private Sub Form_Load()
'Default the number of notes to 0
NumNotes = 0
End Sub

Private Sub Form_Resize()
'See if the window is minimized, if so don't resize
'any controls
If Me.WindowState = vbMinimized Then Exit Sub

'Re-position the Notes textbox
txtNotes.Left = Me.ScaleWidth - 279

'Resize the Document/Notes textboxes
txtDocument.Width = Me.ScaleWidth - 326
txtNotes.Height = Me.ScaleHeight - 32
txtDocument.Height = Me.ScaleHeight - 32

'Clear everything drawn on the forms
Me.Cls

'Re-draw the 3d borders
Control_3DBorder txtDocument, Me, vbBlue, 2
Control_3DBorder txtNotes, Me, vbRed, 2
End Sub

Private Sub mnuAlpha_Click()
Dim Notes As String, X As Integer, CrLf As Integer

'Set the Notes string to the notes taken
Notes = txtNotes.text

'Clear the listbox
lstNotes.Clear

'Add each note to the listbox (which has
'the 'Sorted' property set to True so that
'it sorts each list item appropriately
Do
    CrLf = InStr(Notes, vbCrLf)
    lstNotes.AddItem Left(Notes, CrLf - 1)
    Notes = Mid(Notes, CrLf + 2, Len(Notes))
Loop Until InStr(Notes, vbCrLf) = 0

'Clear the textbox
txtNotes.text = ""

'Add each entry back to the notes textbox,
'now that they are alphabetized
For X = 0 To lstNotes.ListCount - 1
    txtNotes.text = txtNotes.text & lstNotes.List(X) & vbCrLf
Next X

'Clear the notes listbox
lstNotes.Clear
End Sub

Private Sub mnuClear_Click()
'Clear the document
txtDocument.text = ""
End Sub

Private Sub mnuClearNotes_Click()
'Default the number of notes back to 0
NumNotes = 0

'Make the caption reflect that
Me.Caption = " Note Taker - " & NumNotes & " notes"

'Clear the notes
txtNotes.text = ""
End Sub

Private Sub mnuOpen_Click()
'Set the title of the common dialog
cmd1.DialogTitle = "Open Document"

'Show it
cmd1.ShowOpen

If cmd1.filename <> "" Then
    'Open the file if the filename isn't blank
    txtDocument.text = FileOpen(cmd1.filename)
End If
End Sub

Private Sub mnuProperties_Click()
Dim FreeNum As Integer

'Get the free file number
FreeNum = FreeFile

'Open the temporary filename for saving
Open "C:\windows\notes.tmp" For Output As #FreeNum

'Save the document to it
Print #FreeNum, txtDocument.text

'Close the file
Close #FreeNum

'Msgbox the document properties
'First line...characters
'Second.......words
'Third........lines
'Fourth.......size of file in bytes
MsgBox "Document Properties:" & vbCrLf & _
"Chars:" & vbTab & Len(txtDocument) _
& vbCrLf & "Words:" & vbTab & GetWordCount(txtDocument) _
& vbCrLf & "Lines:" & vbTab & GetLineCount(txtDocument) _
& vbCrLf & "Filesize:" & vbTab & GetFileSize("C:\windows\notes.tmp") & " bytes" _
, vbExclamation + vbOKOnly
End Sub

Private Sub mnuSave_Click()
Dim FreeNum As Integer

'See if the user has taken any notes
If NumNotes = 0 Then
    'If not, don't let them save
    MsgBox "You haven't taken any notes yet.  Please take at least one note before saving.", vbExclamation + vbOKOnly
    Exit Sub
End If

'Set the common dialog's title
cmd1.DialogTitle = "Save Notes"

'Show it
cmd1.ShowSave

'See if the filename was left blank
If cmd1.filename <> "" Then
    'if not, save the file
    
    'Get the free file number
    FreeNum = FreeFile
    
    'Open the file for saving
    Open cmd1.filename For Output As #FreeNum
    
    'Save to the file
    Print #FreeNum, txtDocument.text
    
    'Close the file
    Close #FreeNum
End If
End Sub

Private Sub txtDocument_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Note As String

'If there's at least one character selected,
'copy that note to the Notes textbox
If txtDocument.SelLength > 0 Then
    'Get the selected text
    Note = txtDocument.SelText
    
    'Remove an enter if it's the last character
    If Right(Note, 2) = vbCrLf Then Note = Left(Note, Len(Note) - 2)
    If Len(Note) = 1 Then Exit Sub
    
    'Copy the note to the Notes textbox
    txtNotes.text = txtNotes.text & "-" & Note & vbCrLf
    
    'Add one to the number of notes taken
    NumNotes = NumNotes + 1
    Me.Caption = " Note Taker - " & NumNotes & " notes"
End If
End Sub
