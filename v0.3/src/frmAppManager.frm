VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAppManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "App Manager"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstApp 
      Height          =   1230
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5895
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDir 
      Caption         =   "Bulk"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Single"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   6240
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblPath 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   "Install a single app or bulk install from a directory"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Curently very basic.. an actual app manager will be added in a later version."
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   3360
      Width           =   6015
   End
End
Attribute VB_Name = "frmAppManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InstallType As String

Private Sub cmdDir_Click()
    
    '-- Initialize Common Dialog control
    With cdMain
        .Flags = cdlOFNPathMustExist
        .Flags = .Flags Or cdlOFNHideReadOnly
        .Flags = .Flags Or cdlOFNNoChangeDir
        .Flags = .Flags Or cdlOFNExplorer
        .Flags = .Flags Or cdlOFNNoValidate
        .FileName = "*.*"
    End With
    
    Dim x As Integer
    '-- Cheap way to use the common dialog box as a directory-picker
    x = 3

    cdMain.CancelError = True        'do not terminate on error

    On Error Resume Next         'I will hande errors

    cdMain.Action = 1              'Present "open" dialog

    '-- If FileTitle is null, user did not override the default (*.*)
    If cdMain.FileTitle <> "" Then x = Len(cdMain.FileTitle)

    If Err = 0 Then
        ChDrive cdMain.FileName
        lblPath.Caption = Left(cdMain.FileName, Len(cdMain.FileName) - x)
        cmdInstall.Enabled = True
        InstallType = "Dir"
    Else
      '-- User pressed "Cancel"
    End If
    
End Sub

Private Sub cmdFile_Click()

        
    '-- Initialize Common Dialog control
    With cdMain
        .Flags = cdlOFNPathMustExist
        .Flags = .Flags Or cdlOFNHideReadOnly
        .Flags = .Flags Or cdlOFNNoChangeDir
        .Flags = .Flags Or cdlOFNExplorer
        .Flags = .Flags Or cdlOFNNoValidate
        .FileName = "*.apk"
    End With

    Dim x As Integer
    '-- Cheap way to use the common dialog box as a directory-picker
    x = 3

    cdMain.CancelError = True        'do not terminate on error

    On Error Resume Next         'I will hande errors

    cdMain.Action = 1              'Present "open" dialog

    '-- If FileTitle is null, user did not override the default (*.*)
    If cdMain.FileTitle <> "" Then x = Len(cdMain.FileTitle)

    If Err = 0 Then
        ChDrive cdMain.FileName
        lblPath.Caption = cdMain.FileName
        cmdInstall.Enabled = True
        InstallType = "File"
    Else
      '-- User pressed "Cancel"
    End If

End Sub

Private Sub cmdInstall_Click()

    'Install APK's
    MsgBox "Installing, please wait.."
        
    If InstallType = "File" Then
        lstApp.AddItem "Installing " & lblPath.Caption
        frmMain.ADB "install " & Chr(34) & lblPath.Caption & Chr(34)
        If Right(frmMain.ReturnData, 2) = "s)" Then
            lstApp.AddItem "Done"
            lstApp.ListIndex = lstApp.ListCount - 1
        End If
    
    ElseIf InstallType = "Dir" Then
    
        Dim File As String
        Dim Count As Integer
        Count = 0
        If Right$(lblPath.Caption, 1) <> "\" Then lblPath.Caption = lblPath.Caption & "\"
        Extention = "*.apk"
        File = Dir$(lblPath.Caption & Extention)
        Do While Len(File)
            lstApp.AddItem "Installing " & lblPath.Caption & File
            frmMain.ADB "install " & Chr(34) & lblPath.Caption & "\" & File & Chr(34)
            If Right(frmMain.ReturnData, 2) = "s)" Then
                lstApp.AddItem "Done"
                lstApp.ListIndex = lstApp.ListCount - 1
                Count = Count + 1
            End If
            File = Dir$
        Loop
        
        MsgBox Count & " applications installed"
    
    End If

End Sub

Private Sub Form_Load()

    InstallType = ""
    
    cmdInstall.Enabled = False
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmMain.Show
End Sub
