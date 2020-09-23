VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blocked IO"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbBufferSize 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1680
      List            =   "frmMain.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox chkOverwrite 
      Alignment       =   1  'Right Justify
      Caption         =   "Overwrite"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCopyFile 
      Caption         =   "Copy File"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtDestinationFile 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "C:\lady2.mp3"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtSourceFile 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "C:\Lady.mp3"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Buffer Size"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      Caption         =   "0 %"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Destination File"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Source File"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents objBlockedIO As BlockedIO
Attribute objBlockedIO.VB_VarHelpID = -1


Private Sub cmdCancel_Click()

    objBlockedIO.Cancel
    cmdCancel.Enabled = False
    
End Sub

Private Sub cmdCopyFile_Click()

    If cmbBufferSize = "1" Then
        If MsgBox("WARNING: Byte for byte transfer takes a VERY long time for files exceeding 100KB. Are you sure you want to copy using this buffer size?", vbQuestion + vbYesNo, "Buffer Warning") = vbNo Then
            Exit Sub
        End If
    End If

    cmdCancel.Enabled = True
    cmdCopyFile.Enabled = False
    lblStatus = "Copying..."
    If chkOverwrite.Value = vbChecked Then
        objBlockedIO.CopyFile txtSourceFile, txtDestinationFile, True, CLng(cmbBufferSize)
    Else
        objBlockedIO.CopyFile txtSourceFile, txtDestinationFile, , CLng(cmbBufferSize)
    End If

End Sub

Private Sub Form_Load()

    cmbBufferSize.Text = 2048 'The default buffer size

    Set objBlockedIO = New BlockedIO

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objBlockedIO = Nothing

End Sub

Private Sub objBlockedIO_CopyCancelled()

    MsgBox "Copy cancelled by user!", vbCritical, "Copy Cancelled"
    ProgressBar1.Value = 0
    lblPercent.Caption = "0 %"
    lblStatus = "Idle"
    cmdCopyFile.Enabled = True

End Sub

Private Sub objBlockedIO_CopyComplete()

    MsgBox "Copy Complete", vbInformation, "Copy Complete"
    ProgressBar1.Value = 0
    lblPercent.Caption = "0 %"
    lblStatus = "Idle"
    cmdCancel.Enabled = False
    cmdCopyFile.Enabled = True

End Sub

Private Sub objBlockedIO_CopyError(strDescription As String)

    MsgBox strDescription, vbExclamation, "Copy Error"
    ProgressBar1.Value = 0
    lblPercent.Caption = "0 %"
    lblStatus = "Idle"
    cmdCancel.Enabled = False
    cmdCopyFile.Enabled = True
    
End Sub

Private Sub objBlockedIO_CopyProgress(lngPercentDone As Long)

'On Error Resume Next

    ProgressBar1.Value = lngPercentDone
    lblPercent = lngPercentDone & " %"
    
End Sub
