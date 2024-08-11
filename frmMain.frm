VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Convertir a BMP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir a PNG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   120
      Pattern         =   "*.bmp*"
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim dib As Long: Dim bOK As Long
    Dim exts As String: Dim imgFormat As FREE_IMAGE_FORMAT
    Dim sFileName As String
    Dim fFileName As String
        Dim fType As String
    On Local Error GoTo ErrorConversion
    
    PB.Value = 0
    PB.Max = File1.ListCount
    '/// Retrive the formats
    '/// This is only a DEMO, to implement hoter formats see the project
    For i = 0 To File1.ListCount - 1
    fFileName = App.Path & "\BMP\" & File1.List(i)
    exts = GetFilePath(fFileName, Only_Extension)
    
    sFileName = Left$(File1.List(i), Len(File1.List(i)) - 4)
    
    imgFormat = FIF_BMP

    
    '/// Load a image
    dib = FreeImage_Load(imgFormat, fFileName)
    

    
 fType = ".png"

    
    If fType = "." & exts Then
            MsgBox "Sorry, change the convertion file Image!!", vbExclamation, App.Title
        Exit Sub
    End If
    '/// parameters File type to be converted, file to be converted, new image name, image save options
    bOK = FreeImage_Save(FIF_PNG, dib, App.Path + "\PNG\" + sFileName & fType, 0)
  
    '/// Unload the dib
    FreeImage_Unload (dib)
    PB.Value = i
    Next i
Exit Sub
ErrorConversion:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub Command2_Click()
    Dim dib As Long: Dim bOK As Long
    Dim exts As String: Dim imgFormat As FREE_IMAGE_FORMAT
    Dim sFileName As String
    Dim fFileName As String
        Dim fType As String
    On Local Error GoTo ErrorConversion
    
    PB.Value = 0
    PB.Max = File1.ListCount
    '/// Retrive the formats
    '/// This is only a DEMO, to implement hoter formats see the project
    For i = 0 To File1.ListCount - 1
    fFileName = App.Path & "\PNG\" & File1.List(i)
    exts = GetFilePath(fFileName, Only_Extension)
    
    sFileName = Left$(File1.List(i), Len(File1.List(i)) - 4)
    
    imgFormat = FIF_PNG

    
    '/// Load a image
    dib = FreeImage_Load(imgFormat, fFileName)
    

    
 fType = ".bmp"

    
    If fType = "." & exts Then
            MsgBox "Sorry, change the convertion file Image!!", vbExclamation, App.Title
        Exit Sub
    End If
    '/// parameters File type to be converted, file to be converted, new image name, image save options
    bOK = FreeImage_Save(FIF_BMP, dib, App.Path + "\BMP\" + sFileName & fType, 0)
  
    '/// Unload the dib
    FreeImage_Unload (dib)
    PB.Value = i
    Next i
Exit Sub
ErrorConversion:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub File1_Click()

End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\BMP"
End Sub
