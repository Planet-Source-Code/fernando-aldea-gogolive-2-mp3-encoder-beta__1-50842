VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3870
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Wave In:"
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   3615
      Begin VB.ComboBox cmbDevices 
         Height          =   315
         Left            =   720
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ListBox lstSupport 
         Height          =   2010
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label lblSupFormat 
         Caption         =   "Supported Format (in this moment):"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Device:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Live ShoutCast:"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   3615
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   7
         Text            =   "0"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Title in the Winamp"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Listen Port:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2100
      TabIndex        =   3
      Top             =   5100
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   5100
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Text            =   "0"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Buffer Length (bytes) :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''    Module written by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    Release January, 2003                 ''
''                                          ''
''    sorry for not translate completly     ''
''    & sorry for my English!               ''
''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmbDevices_Click()
    GetSupportOfDevice
End Sub

Private Sub Command1_Click()
    'setear los valores
    
    If BUFFER_LENGTH <> Val(Me.Text1(0).Text) Then
        BUFFER_LENGTH = Val(Me.Text1(0).Text)
    End If
    
    If Not Processing Then
        mPipe.Destruir
        If Not mPipe.Crear(BUFFER_LENGTH) Then MsgBox "Error to create Pipe"
    End If
        
    PORT_LISTEN = Me.Text1(1).Text
    TITLE_NAME = Me.Text1(2).Text
    DEVICE_ID = Me.cmbDevices.ListIndex - 1
    
    If (Mode And MODE_INPUT_WAVE) = MODE_INPUT_WAVE Then
        frmGogoLive.lblInput.Caption = GetNameDevice(DEVICE_ID)
    End If
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Text1(0).Text = BUFFER_LENGTH
    Me.Text1(1).Text = PORT_LISTEN
    Me.Text1(2).Text = TITLE_NAME
    Me.Show
    Call GetDevWaveIn
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmGogoLive.Enabled = True
End Sub

Sub GetDevWaveIn()
    Dim a As Integer
    Dim WV As WAVEINCAPS
        
    
    a = waveInGetNumDevs()
    
    Me.cmbDevices.Clear
    For a = -1 To a - 1
        waveInGetDevCaps a, WV, Len(WV)
        Me.cmbDevices.AddItem WV.szPname & " " & WV.wChannels
    Next a
    
    If Me.cmbDevices.ListCount > 0 Then Me.cmbDevices.ListIndex = DEVICE_ID + 1
    
End Sub

Private Sub GetSupportOfDevice()
    Dim Freq As Long

    Me.lstSupport.Clear
    For a = 0 To 7
        Freq = frmGogoLive.Frequency2Long(a)
        
        If modWaveIn.Initialize(Me.cmbDevices.ListIndex - 1, Freq, True, 16, CM_QUERY) Then
            Me.lstSupport.AddItem Freq & " Hz, 16-Bit, Stereo"
        End If
        If modWaveIn.Initialize(Me.cmbDevices.ListIndex - 1, Freq, False, 16, CM_QUERY) Then
            Me.lstSupport.AddItem Freq & " Hz, 16-Bit, Mono"
        End If
        If modWaveIn.Initialize(Me.cmbDevices.ListIndex - 1, Freq, True, 8, CM_QUERY) Then
            Me.lstSupport.AddItem Freq & " Hz, 8-Bit, Stereo"
        End If
        If modWaveIn.Initialize(Me.cmbDevices.ListIndex - 1, Freq, False, 8, CM_QUERY) Then
            Me.lstSupport.AddItem Freq & " Hz, 8-Bit, Mono"
        End If
        DoEvents
    Next a
        
End Sub

