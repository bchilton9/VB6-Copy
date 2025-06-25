VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PC File Copier"
   ClientHeight    =   3660
   ClientLeft      =   6150
   ClientTop       =   3090
   ClientWidth     =   4770
   FillColor       =   &H00C00000&
   ForeColor       =   &H00C00000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4770
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.DriveListBox drvMP3 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.DirListBox dirMP3 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   450
      Width           =   2175
   End
   Begin VB.FileListBox fileMP3 
      Height          =   2235
      Left            =   2280
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   3135
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FS As New FileSystemObject

Dim MP3FileName


Private Sub Command1_Click()
    Dim FSO, FS
    Set FSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
FS = FSO.CreateFolder(fileMP3.Path & "\" & Text1.Text)
FS = FSO.CopyFile(fileMP3.Path & "\" & fileMP3.Filename, fileMP3.Path & "\" & Text1.Text & "\")
End Sub

Private Sub dirMP3_Change()
  fileMP3 = dirMP3
End Sub

Private Sub drvMP3_Change()
  dirMP3 = drvMP3
  fileMP3 = dirMP3
End Sub



Private Sub fileMP3_DblClick()

    Dim MP3File, MP3Size, frmID3, frmButtons, Title, Artist, Album, Year, Comment, Genre

  If Len(fileMP3.Path) > 3 Then
    MP3FileName = fileMP3.Path & "\"
  Else
    MP3FileName = fileMP3.Path
  End If
  MP3FileName = MP3FileName & fileMP3.Filename

  Dim Buf As String * 128
  Dim tmpStr As String
  Dim i As Byte
  
  MP3File = MP3FileName
  'Get the size of mp3 file(in bytes)
  MP3Size = FileLen(MP3File)
  
  'labLength = labLength & mp3Length & " seconds"
  
  'Open the file for binary access in order to get the ID3 Tag
  Open MP3File For Binary As #1
    'Get last 128 bytes of the file. The size of file is reduced by 127 bytes, because
    'the last byte in file is in fact the size of file
    Get #1, MP3Size - 127, Buf
    'Check if the file has a tag
    If Format(Left(Buf, 3), "<") <> "tag" Then
      'frmID3.Visible = False
      'frmButtons.Visible = False
    Else
      'If it has a tag the separate the info obtained in the buffer string
      Title = Trim(Mid(Buf, 4, 30))
      Artist = Trim(Mid(Buf, 34, 30))
      Album = Trim(Mid(Buf, 64, 30))
      Year = Trim(Mid(Buf, 94, 4))
      Comment = Trim(Mid(Buf, 98, 30))
      'For i = 0 To 148
        'If Genre.ItemData(i) = Trim(Asc(Mid$(Buf, 128, 1))) Then Exit For
      'Next i
      'If i < 149 Then
        'Genre.ListIndex = i
      'End If
    End If
  Close #1

If Artist = "" Then
Artist = fileMP3.Filename
End If

Text1.Text = Artist

End Sub

Private Sub Exit_Click()
    End                                                                             'Quits the program
End Sub

Private Sub Form_Load()
    Main.Show                                                                       'Shows the main form
End Sub
