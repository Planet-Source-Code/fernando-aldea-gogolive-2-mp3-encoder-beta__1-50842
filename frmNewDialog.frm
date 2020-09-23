VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New "
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAccept 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1800
      Width           =   915
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2820
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame frmeFile 
      Caption         =   "Archivo"
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5535
      Begin VB.TextBox txtFile 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4815
      End
      Begin VB.CommandButton CmdDialog 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmeFile 
      Caption         =   "Archivo"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton CmdDialog 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmNewDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DialogType As Integer


Private Sub CmdAccept_Click()
    If DialogType = DIALOGTYPE_SAVE Then
        frmGogoLive.cmdNewOnfly Me.txtFile(0).Text
    Else
        frmGogoLive.cmdNewEncodeFile Me.txtFile(0).Text, Me.txtFile(1).Text
    End If
    
    Unload Me
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDialog_Click(Index As Integer)
    If DialogType = DIALOGTYPE_OPENSAVE Then
        If Index = 0 Then
            Me.CommonDialog1.filename = "*.wav"
            Me.CommonDialog1.Filter = "*.wav"
            Me.CommonDialog1.ShowOpen
        End If
        If Index = 1 Then
            Me.CommonDialog1.Filter = "*.mp3"
            Me.CommonDialog1.filename = "newGOGO.mp3"
            Me.CommonDialog1.ShowSave
        End If
    Else
        Me.CommonDialog1.ShowSave
    End If
    
    Me.txtFile(Index).Text = Me.CommonDialog1.filename
End Sub


Private Sub Form_Load()
    If DialogType = DIALOGTYPE_SAVE Then
        Me.Caption = "New Record On-Fly"
        Me.frmeFile(0).Caption = "Out mp3 file:"
        Me.frmeFile(1).Visible = False
    Else
        Me.Caption = "New Encode File"
        Me.frmeFile(0).Caption = "Source wav file:"
        Me.frmeFile(1).Caption = "Out mp3 file:"
        Me.frmeFile(1).Visible = True
    End If
    frmGogoLive.Enabled = False
End Sub

Private Sub Form_Terminate()
    frmGogoLive.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmGogoLive.Enabled = True
End Sub
