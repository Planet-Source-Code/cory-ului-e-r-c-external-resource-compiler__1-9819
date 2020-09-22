VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "E.R.C - External Resource Compiler"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmERC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Add File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Timer tmrDone 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   5760
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Extract"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid msfFileList 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4895
      _Version        =   393216
      BackColorSel    =   255
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4895
      _Version        =   393216
      BackColorSel    =   255
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'** E.R.C - External Resource Compiler  **'
'** Code was writen by Cory Watt(mouak@crosswinds.net)
'** Use as you wish, just never sell, unless compiled in
'** a excuting application/program!

'** oh and of course if you want to Add this to a book like those
'** Visualbasic source code books, or source library books.
'** Plzzzzz put me in it then, and my name, +e-mail address..
'** And if you want this for a book, tell me and I'll send it
'** Commented And more tidy! (you have to tell me the book name too,
'** so I can get it when/if it comes out here)

'(As you know I hardly put Comments or REM statments in the work I do, mainly
'because I like to read it visualy in my head, rather rely on comments/rem's to
'remind me. ) Sorry for those who like em...


'Someone had e-mailed me months and months ago, asking on how you do this
'And at the time I had no idea what they were talking about (sorry).
'But one day, I figureed what they wanted (I think) and here it is.
'(You know who you are)...

Option Explicit

Private Type pt_Info
  i_Offset As Long
  i_FileSize As Long
  i_FileType As Long
End Type

Private Type pt_Extension
  i_Extension As String
  i_ID As Long
End Type

Dim t_MEM_Extension() As pt_Extension

Dim AppPath As String, p_Filename As String
Dim t_Header As Long, t_File() As Byte, t_Info As pt_Info, t_FileCount As Long

Private Sub Status(Optional ByVal sCaption As String)
  If IsMissing(sCaption) Then
    lblStatus.Caption = ""
  Else
    lblStatus.Caption = sCaption
  End If
  DoEvents
End Sub
Private Sub Done()
  lblStatus.Caption = "Done"
  DoEvents
  tmrDone.Enabled = True
End Sub
Private Function GetID(giExtension As String) As Long
Dim i_iI As Long
  For i_iI = 0 To UBound(t_MEM_Extension)
    If t_MEM_Extension(i_iI).i_Extension = giExtension Then
      GetID = t_MEM_Extension(i_iI).i_ID
      Exit Function
    End If
  Next
End Function
Private Function GetExtension(ByVal geID As Long) As String
Dim i_iI As Long
  For i_iI = 0 To UBound(t_MEM_Extension)
    If t_MEM_Extension(i_iI).i_ID = geID Then
      GetExtension = t_MEM_Extension(i_iI).i_Extension
      Exit Function
    End If
  Next
  GetExtension = 0
End Function
Private Sub Command1_Click()
Dim i_i1 As Integer, i_Filename As String, i_SaveAs As String, i_String As String

  i_SaveAs = SaveDialog(Me, "All File Types (*.*)|*.*")
  If Len(i_SaveAs) > 4 Then
    Status "Preparing to Compile..." 'Update Status
    If Len(Dir(i_SaveAs)) > 0 Then DeleteFile i_SaveAs 'Delete File if Exists
    
    t_FileCount = msfFileList.Rows - 1
    t_Info.i_Offset = 8 + (t_FileCount * 12)
    Open i_SaveAs For Binary Access Write As #1
      Put #1, , t_Header 'Set Unique Header
      Put #1, , t_FileCount
      
      '** Reference Files **
      For i_i1 = 0 To t_FileCount - 1
        msfFileList.Row = i_i1 + 1: msfFileList.Col = 2
        i_Filename = msfFileList.Text 'Get Filename
        
        Status "Preparing... """ & i_Filename & """" 'Update Status
        
        t_Info.i_FileSize = FileLen(i_Filename)
        t_Info.i_FileType = GetID(Right(UCase(i_Filename), 3))
        Put #1, , t_Info
        t_Info.i_Offset = t_Info.i_Offset + t_Info.i_FileSize
      Next
     
      '** Compilies Files **
      For i_i1 = 0 To t_FileCount - 1
        msfFileList.Row = i_i1 + 1: msfFileList.Col = 2
        i_Filename = msfFileList.Text 'Get Filename
        
        Status "Compiling... """ & i_Filename & """" 'Update Status
        
        If Len(Dir(i_Filename)) > 0 Then
          t_Info.i_FileSize = FileLen(i_Filename)
          If t_Info.i_FileSize > 0 Then
            ReDim t_File(t_Info.i_FileSize - 1) As Byte
          End If
          
          Open i_Filename For Binary Access Read As #2
            Get #2, , t_File()
          Close #2
        End If
        Put #1, , t_File
      Next
    Close #1
    ReDim t_File(0) 'Clear File
    Done 'Update Status
  End If
End Sub
Private Sub Command2_Click()
Dim I As Long, i_Header As Long
  p_Filename = OpenDialog(Me, "All File Types (*.*)|*.*")
  If Len(p_Filename) > 0 Then
    MSFlexGrid1.Redraw = False
    If Len(Dir(p_Filename)) > 0 Then
      Status "Reading External Resource File..."
      Open p_Filename For Binary Access Read As #1
        Get #1, , i_Header
        If i_Header = t_Header Then
          Get #1, , t_FileCount
          DoEvents
          MSFlexGrid1.Rows = t_FileCount + 1
          For I = 0 To t_FileCount - 1
            Get #1, , t_Info
            With MSFlexGrid1
              .Row = I + 1
              .Col = 1: .Text = I + 1
              .Col = 2: .Text = Format(t_Info.i_Offset, "#,###,###,###")
              .Col = 3: .Text = Format(t_Info.i_FileSize, "#,###,###,###")
              .Col = 4: .Text = t_Info.i_FileType
              .Col = 5: .Text = GetExtension(t_Info.i_FileType)
            End With
          Next I
        Else
          MsgBox "File is not a Vaild External Resource File!"
        End If
      Close #1
      DoEvents
      Done
    End If
    MSFlexGrid1.Redraw = True
  End If
End Sub

Private Sub Command3_Click()
Dim I As Long, ii_Filename As String, ii_Extension As String
  If Len(p_Filename) > 0 Then
    If Len(Dir(p_Filename)) > 0 Then
      I = CLng(MSFlexGrid1.Row) - 1
      Open p_Filename For Binary Access Read As #1
        Get #1, (I * 12) + 9, t_Info
        
        ReDim t_File(t_Info.i_FileSize - 1)
        
        Get #1, t_Info.i_Offset + 1, t_File
        MSFlexGrid1.Col = 5
        ii_Extension = MSFlexGrid1.Text
    
        If ii_Extension = "???" Then
          ii_Extension = "er_"
        End If
        ii_Filename = AppPath & "op" & Format((I + 1), "000000") & "." & ii_Extension
          Open ii_Filename For Binary Access Write As #2
            Put #2, , t_File
          Close #2
        ReDim t_File(0)
      Close #1
    End If
  End If
End Sub
Private Sub LoadFileTypes()
Dim i_TextLine, i_I As Long, i_C As Long
  If Dir(AppPath & "filetype.txt") <> "" Then
    Open AppPath & "filetype.txt" For Input As #1
      Do While Not EOF(1)
        Line Input #1, i_TextLine
        ReDim Preserve t_MEM_Extension(i_C)

        i_I = InStr(1, i_TextLine, ":")
        t_MEM_Extension(i_C).i_Extension = UCase(Left(i_TextLine, i_I - 1))
        t_MEM_Extension(i_C).i_ID = CLng(Right(i_TextLine, Len(i_TextLine) - i_I))
        i_C = i_C + 1
        
      Loop
    Close #1
  End If
End Sub

Private Sub Command4_Click()
Dim i_Filename As String, i_i1 As Integer, i_EXT As Integer
  
  i_Filename = OpenDialog(Me, "All File Types (*.*)|*.*", "Select File")

  If Len(i_Filename) > 4 Then
      i_i1 = msfFileList.Rows
      msfFileList.Rows = msfFileList.Rows + 1
      With msfFileList
        .Row = i_i1
        .Col = 1: .Text = i_i1
        .Col = 2: .Text = i_Filename
        .Col = 3: .Text = Format(FileLen(i_Filename), "#,###,###,###")
        i_EXT = GetID(Right(UCase(i_Filename), 3))
        .Col = 4: .Text = i_EXT
        .Col = 5: .Text = GetExtension(i_EXT)
      End With
      
  End If
  
End Sub



Private Sub Form_Load()
Dim i_tx As Integer
  If Right(App.Path, 1) = "\" Then
    AppPath = App.Path
  Else
    AppPath = App.Path & "\"
  End If
  t_Header = 1666336181 'Set header to text µERc

  LoadFileTypes
  
  With msfFileList
    .Cols = 6: .Row = 0: .Rows = 1
    .Col = 1: .Text = "Index"
    .Col = 2: .Text = "Filename"
    .Col = 3: .Text = "Filesize"
    .Col = 4: .Text = "FileType"
    .Col = 5: .Text = "Extension"
  End With
  
  With MSFlexGrid1
    .Cols = 6: .Row = 0: .Rows = 1
    .Col = 1: .Text = "Index"
    .Col = 2: .Text = "OffSet"
    .Col = 3: .Text = "FileSize"
    .Col = 4: .Text = "FileType"
    .Col = 5: .Text = "Extension"
  End With

  i_tx = Screen.TwipsPerPixelX
  MSFlexGrid1.ColWidth(0) = 16 * i_tx
  msfFileList.ColWidth(0) = 16 * i_tx
End Sub

Private Sub Form_Resize()
  lblStatus.Move 2, ScaleHeight - 16, ScaleWidth - 4, 14
End Sub
Private Sub tmrDone_Timer()
  lblStatus.Caption = ""
  tmrDone.Enabled = False
End Sub

