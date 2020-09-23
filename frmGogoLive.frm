VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGogoLive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GOGOLive by Fernando Aldea"
   ClientHeight    =   2760
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3735
   ForeColor       =   &H8000000D&
   Icon            =   "frmGogoLive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":0442
            Key             =   "Rec"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":069E
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":08FA
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":0B56
            Key             =   "CD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":0E12
            Key             =   "Wave"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":132A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":1586
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":17E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":1A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":1C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGogoLive.frx":1EF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   2505
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2990
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Display 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      Height          =   1590
      Left            =   0
      ScaleHeight     =   1530
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   375
      Width           =   3735
      Begin VB.PictureBox Spectrum 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         DragMode        =   1  'Automatic
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   900
         Negotiate       =   -1  'True
         ScaleHeight     =   285
         ScaleWidth      =   1605
         TabIndex        =   3
         Top             =   660
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   225
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   3645
         WordWrap        =   -1  'True
      End
      Begin VB.Image LCD3 
         Height          =   285
         Left            =   2520
         Picture         =   "frmGogoLive.frx":2186
         ToolTipText     =   "Low speed processing"
         Top             =   360
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblDer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<><><>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   15
         Tag             =   "0"
         ToolTipText     =   "Click here to change Frequency"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lblDer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<><><>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   14
         Tag             =   "0"
         ToolTipText     =   "Click here to change Bitrate"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblDer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<><><>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   13
         Tag             =   "0"
         ToolTipText     =   "Click here to change Channels mode"
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblIzq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PSY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   12
         ToolTipText     =   "USe Psycho-acustic mode ON/OFF. (Best Quality)"
         Top             =   300
         Width           =   735
      End
      Begin VB.Label lblIzq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LPF16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   11
         ToolTipText     =   "16KHz low-pass filter ON/OFF"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblIzq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MMX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   10
         ToolTipText     =   "Use MMX capacity ON/OFF (if is available)"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label LCD1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "READY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   900
         TabIndex        =   9
         ToolTipText     =   "Current state"
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label LCD2 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   900
         TabIndex        =   8
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label lblGOGO 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GOGO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   7
         ToolTipText     =   "Gogo.dll version"
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblDer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<><><>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   6
         Tag             =   "0"
         ToolTipText     =   "Select Enphasys mode"
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblIzq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3DNow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "Use 3DNow capacity ON/OFF (if is available)"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblGOGO 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GOGOLive"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "Gogo.dll version"
         Top             =   1260
         Width           =   1755
      End
   End
   Begin MSComctlLib.Toolbar tbControlMP3 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      ButtonWidth     =   529
      ButtonHeight    =   503
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New mp3 file"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rec"
            Object.ToolTipText     =   "Start record to mp3"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop current record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pause"
            Object.ToolTipText     =   "Pause current record"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Monitor"
            Object.ToolTipText     =   "Monitor input level"
            ImageIndex      =   6
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1320
      Picture         =   "frmGogoLive.frx":2474
      ScaleHeight     =   135
      ScaleWidth      =   1470
      TabIndex        =   2
      Top             =   2760
      Width           =   1530
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   120
      Top             =   2640
   End
   Begin VB.Label Label1 
      Caption         =   "Output:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   21
      Top             =   2280
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lInput 
      Caption         =   "Input:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   2040
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInput 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   2040
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOutput 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   2280
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnArchivo 
      Caption         =   "File"
      Begin VB.Menu mnNuevo 
         Caption         =   "New"
         Begin VB.Menu mnuNewOnfly 
            Caption         =   "Record On-Fly..."
         End
         Begin VB.Menu mnuNewEncodeFile 
            Caption         =   "Encode File..."
         End
         Begin VB.Menu mnuNewShoutcast 
            Caption         =   "Live ShoutCast"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnVer 
      Caption         =   "&View"
      Begin VB.Menu mnuOpciones 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmGogoLive"
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

Const AMPLITUD_MAX = 65535


Dim mCountFrame As Long

Private Sub About_Click()
    frmAbout.Show
End Sub

Sub cmdQuickNew()
    Dim sResp As String
    
    'si el modo de salida es un archivo, entonces crear un nombre de archivo simple
    If (Mode And MODE_OUTPUT_FILE) = MODE_OUTPUT_FILE Then
        sResp = InputBox("New out mp3 file: ", "Quick new MP3 file", "c:\newGOGO_" & CLng(Timer) & ".mp3")
        If sResp <> "" Then
            mp3OutFile = sResp
            Me.lblOutput.Caption = mp3OutFile
        End If
    End If
End Sub


Sub CmdMonitor()
    Dim tStereo As Boolean
    
    If Me.tbControlMP3.Buttons("Monitor").value = 1 Then
        
        If Not Monitoring Then
            'determinar modo del wave (stereo o mono)
            tStereo = IIf(mp3Mode <> MC_MODE_MONO, True, False)
            'Initialize Wave Service
            If Not modWaveIn.Initialize(DEVICE_ID, mp3Frequency, tStereo, 16, CM_WINDOWS) Then
                MsgBox "Error in WAVE", vbCritical
                Exit Sub
            End If
            
            'El valor como argumento de StartInput especifica el largo del buffer que debe
            'llenar el driver de sonido. Entre mas pequeñeo mas rapido es la actualizacion
            'de los datos. Cambia este valor a conveniencia
            If modWaveIn.StartInput(1000) = True Then
                frmGogoLive.LCD1.Caption = "MONITOR"
                frmGogoLive.LCD2.Visible = False
                frmGogoLive.Spectrum.Visible = True
                Me.tbControlMP3.Buttons("Rec").Enabled = False
                Me.tbControlMP3.Buttons("Stop").Enabled = False
                Me.tbControlMP3.Buttons("Pause").Enabled = False
                Monitoring = True
            Else
                Me.tbControlMP3.Buttons("Monitor").value = 0
                frmGogoLive.LCD1.Caption = "Error"
                Monitoring = False
                
            End If
        End If
    Else
        modWaveIn.StopInput
        Me.tbControlMP3.Buttons("Rec").Enabled = True
        Me.tbControlMP3.Buttons("Stop").Enabled = True
        Me.tbControlMP3.Buttons("Pause").Enabled = True
        frmGogoLive.LCD2.Visible = True
        frmGogoLive.Spectrum.Visible = False
        Me.LCD1.Caption = "READY"
        Monitoring = False
    End If
    
End Sub


Private Sub CmdPause()
    
    If Me.tbControlMP3.Buttons("Pause").value = 1 And modWaveIn.isPausing Then Exit Sub
    If Me.tbControlMP3.Buttons("Pause").value = 0 And Not modWaveIn.isPausing Then Exit Sub
    
    PauseRecord
    
    If modWaveIn.isPausing Then
        Me.LCD1.Caption = "PAUSE"
    Else
        If Processing Then
            Me.LCD1.Caption = "REC"
        Else
            Me.LCD1.Caption = "READY"
        End If
    End If
    
End Sub


Private Sub CmdRec()
    Dim mControl As Control
    
    If Not Processing Then
        
        'If Not mPipe.estaListo Then MsgBox "error Pipe": Exit Sub
        
        'Habilitar/desabilitar los controles de comando
        Me.tbControlMP3.Buttons("Rec").Enabled = False
        Me.tbControlMP3.Buttons("New").Enabled = False
        Me.tbControlMP3.Buttons("Stop").Enabled = True
        Me.tbControlMP3.Buttons("Monitor").Enabled = False
        
        'Deshabilitamos los controles de configuracion
        For Each mControl In Me.Controls
            If mControl.Name = "lblDer" Then
                mControl.Enabled = False
            End If
            If mControl.Name = "lblIzq" Then
                mControl.Enabled = False
                mControl.Visible = False
            End If
        Next
        
        Me.Timer1.Enabled = True
        
        'mostramos la funcion que esta haciendo
        Me.LCD1.ForeColor = &HFF&
        If Mode = MODE_INPUT_FILE + MODE_OUTPUT_FILE Then
            Me.LCD1.Caption = "ENCODE"
        ElseIf Mode = MODE_INPUT_WAVE + MODE_OUTPUT_FILE Then
            Me.LCD1.Caption = "REC"
            nBytesPCMacum = 0 'Se inica en 0 bytes de datos PCM acumulados
        ElseIf Mode = MODE_INPUT_WAVE + MODE_OUTPUT_SHOUTCAST Then
            Me.LCD1.Caption = "LIVE!"
        End If
        
        'Comienza el proceso y no retorna hasta que haya finalizado
        StartProcess
        
        'Si termino sin que se le indicara
        'Processing=false (finalizar el proceso)
        'entonces hubo un error
        If Processing Then
            Me.LCD1.Caption = "ERROR"
            Me.LCD1.ForeColor = &HFFFFFF
        Else
            Me.LCD1.Caption = "READY"
            Me.LCD1.ForeColor = &HFFFFFF
        End If
        
        'Finalizar el proceso
        EndProcess
        
        Timer1_Timer
        Me.Timer1.Enabled = False
        
        'Habilitar/desabilitar los controles de comando
        Me.tbControlMP3.Buttons("Rec").Enabled = True
        Me.tbControlMP3.Buttons("New").Enabled = True
        'Me.tbControlMP3.Buttons("Stop").Enabled = False
        Me.tbControlMP3.Buttons("Monitor").Enabled = True
        
        
        'Volver a habilitar los controles de configuracion
        For Each mControl In Me.Controls
            If mControl.Name = "lblDer" Then
                mControl.Enabled = True
            End If
            If mControl.Name = "lblIzq" Then
                mControl.Enabled = True
                mControl.Visible = True
            End If
        Next
        
    End If
    
End Sub


Private Sub CmdStop()
    If Processing Then
        Me.LCD1.Caption = "STOP..."
        EndProcess
        Me.tbControlMP3.Buttons("Pause").value = 0
        If (Mode And MODE_OUTPUT_SHOUTCAST) = MODE_OUTPUT_SHOUTCAST Then
            frmShoutCast.sckMain.Close
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim mStr As String
    
    'mostramos en el formularios la configuracion inicial
    Me.lblGOGO(0).Caption = objGogo.Version
    Me.lblGOGO(1).Caption = "GOGOLive " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.lblDer(0).Caption = mp3Frequency & " Hz"
    Me.lblDer(1).Caption = mp3Kbps & " Kbps"
    
    Call EncodeMode2String(mp3Mode, mStr)
    Me.lblDer(2).Caption = mStr
    
    Call Enphasys2String(mp3Enphasis, mStr)
    Me.lblDer(3).Caption = mStr
    
    
    If mp3PSY Then Me.lblIzq(0).Caption = "PSY" Else Me.lblIzq(0).Caption = ""
    If mp3LPF16 Then Me.lblIzq(1).Caption = "LPF16" Else Me.lblIzq(1).Caption = ""
    If mp3MMX Then Me.lblIzq(2).Caption = "MMX" Else Me.lblIzq(2).Caption = ""
    If mp33DNow Then Me.lblIzq(3).Caption = "3DNow" Else Me.lblIzq(3).Caption = ""
    
    'Crear el buffer de intercambio
    If Not mPipe.Crear(BUFFER_LENGTH) Then
        MsgBox "Cannot to create pipe", vbCritical
        Exit Sub
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Processing Then EndProcess
    mPipe.Destruir
    Unload frmNewDialog
    Unload frmShoutCast
    Unload frmOptions
    Unload frmAbout
    
    MsgBox "Please, Vote for me in Planet-Source-Code   ;)", vbInformation
End Sub

Private Sub lblDer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mStr As String
    If Button = 1 Then
        lblDer(Index).Tag = lblDer(Index).Tag + 1
    Else
        lblDer(Index).Tag = lblDer(Index).Tag - 1
    End If
    
    If Index = 0 Then
        mp3Frequency = Frequency2Long(lblDer(Index).Tag)
        Me.lblDer(Index).Caption = mp3Frequency & " Hz"
    End If
    
    If Index = 1 Then
        mp3Kbps = Bitrate2Long(lblDer(Index).Tag)
        Me.lblDer(Index).Caption = mp3Kbps & " Kbps"
    End If
    
    If Index = 2 Then
        mp3Mode = EncodeMode2String(Me.lblDer(Index).Tag, mStr)
        Me.lblDer(Index).Caption = mStr
    End If
    
    If Index = 3 Then
        mp3Enphasis = Enphasys2String(lblDer(Index).Tag, mStr)
        Me.lblDer(Index).Caption = mStr
    End If
    
End Sub


Private Sub lblIzq_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        mp3PSY = Not mp3PSY
        If mp3PSY Then Me.lblIzq(Index).Caption = "PSY" Else Me.lblIzq(Index).Caption = ""
    End If
    If Index = 1 Then
        mp3LPF16 = Not mp3LPF16
        If mp3LPF16 Then Me.lblIzq(Index).Caption = "LPF16" Else Me.lblIzq(Index).Caption = ""
    End If
    If Index = 2 Then
        mp3MMX = Not mp3MMX
        If mp3MMX Then Me.lblIzq(Index).Caption = "MMX" Else Me.lblIzq(Index).Caption = ""
    End If
    If Index = 3 Then
        mp33DNow = Not mp33DNow
        If mp33DNow Then Me.lblIzq(Index).Caption = "3DNow" Else Me.lblIzq(Index).Caption = ""
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNewEncodeFile_Click()
    frmNewDialog.DialogType = DIALOGTYPE_OPENSAVE
    frmNewDialog.Show
End Sub

Private Sub mnuNewOnfly_Click()
    frmNewDialog.DialogType = DIALOGTYPE_SAVE
    frmNewDialog.Show
End Sub

Private Sub mnuNewShoutcast_Click()
    cmdNewShoutCast
    
    'esto lo puedes eliminar si quieres
    MsgBox "Before the pressing Rec Button, check the configuration." & vbNewLine & _
           "The bitrate (in Kbps) must be <= that your Band Width for a good transmission  ;)" & vbNewLine & _
           "The listener use Winamp Player for listen to you. Pressing Ctrl+L and indicate your ip and port (" & Me.lblOutput.Caption & ")"
End Sub

Private Sub mnuOpciones_Click()
    frmOptions.Show
End Sub

Public Sub callBackWave(ByVal pBuffer As Long, ByVal largoBuffer As Long)
    On Error GoTo err
    
    If Monitoring And Not Processing Then
        Monitor pBuffer, largoBuffer
        Exit Sub
    End If
    
    If nBytesPCMacum = -1 Then
        TimerIni = Timer
        nBytesPCMacum = 0
    Else
        nBytesPCMacum = nBytesPCMacum + largoBuffer
    End If
        
    'calcular la diferencia de tiempo entre un contador que lleva el programa y
    'la cantidad de bytes entregados por el dispositivo Wave
    If Abs((Timer - TimerIni) - (nBytesPCMacum / modWaveIn.nAvgBytesPerSec)) >= 0.2 Then
        Me.LCD3.Visible = True
    Else
        Me.LCD3.Visible = False
    End If
        
    Call mPipe.toWrite(pBuffer, largoBuffer)
    
    Exit Sub
err:
    'manejo de errores
End Sub

Public Function callBackMp3(ByVal pBuffer, largoBuffer As Long) As Long
    On Error GoTo err
    
    Dim LeidosPipe As Long
    Static bytesTotal As Long
    Static BytesActual As Long
    
    bytesTotal = mPipe.nBytesTotal
    BytesActual = mPipe.nBytesActual
    
    'condicion de salida
    If Not Processing And BytesActual <= 0 Then callBackMp3 = ME_EMPTYSTREAM: Exit Function
    
    'esperar si el pipe no esta lleno lo suficiente
    If Processing And (BytesActual < largoBuffer) Then callBackMp3 = ME_MOREDATA: Exit Function
    
    'toRead pipe
    LeidosPipe = mPipe.toRead(pBuffer, largoBuffer)
    
    If LeidosPipe < largoBuffer And Not Processing Then
        callBackMp3 = ME_EMPTYSTREAM
    Else
        callBackMp3 = ME_NOERR
    End If
    
    
    Exit Function
err:
    callBackMp3 = ME_INTERNALERROR
    'Me.sbInfo.Panels(1).Text = "Error GOGO"
    
End Function



Private Sub tbControlMP3_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "New"
        cmdQuickNew
    Case "Rec"
        CmdRec
    Case "Stop"
        CmdStop
    Case "Pause"
        CmdPause
    Case "Monitor"
        CmdMonitor
    End Select
End Sub

Private Sub Timer1_Timer()
    
    Dim tiempoAux As Long
    
    
    If (Mode And MODE_OUTPUT_FILE) = MODE_OUTPUT_FILE Then
        'Mostrar numero de frames mp3 generados hasta el momento
        Me.sbInfo.Panels(1).Text = "Mp3: " & objGogo.CurrentFrame & " frames"
    End If
    
    'Si obtenemos los datos de la tarjeta de onda, mostrar la informacion apropiada
    If (Mode And MODE_INPUT_WAVE) = MODE_INPUT_WAVE Then
        
        If mPipe.nBytesTotal = 0 Then 'pipe infinito
            Me.sbInfo.Panels(2).Text = Format$(mPipe.nBytesActual / 1024, "##0.0") & " KB"
        Else
            'mostrar el estado del pipe (buffer)
            Me.sbInfo.Panels(2).Text = "Buffer: " & Format$((mPipe.nBytesTotal - mPipe.nBytesActual) * 100 / mPipe.nBytesTotal, "##0.0") & "% free"
        End If
        'mantener un registro de la cantidad de segundos de grabacion
        'obtenidos de la tarjeta de sonido para posteriormente compararla
        tiempoAux = CLng(nBytesPCMacum \ modWaveIn.nAvgBytesPerSec)
        Me.LCD2.Caption = Format$((tiempoAux \ 60), "#00") & ":" & Format(tiempoAux Mod 60, "0#")
        
        'Si obtenemos los datos de un archivo, mostrar la informacion apropiada
    ElseIf (Mode And MODE_INPUT_FILE) = MODE_INPUT_FILE Then
        'mostrar porcentaje avanzado
        Me.LCD2.Caption = Format$(objGogo.CurrentFrame * 100 / mCountFrame, "##0") & "%"
        
    End If
    
End Sub

Public Function cmdNewOnfly(ByVal mp3OutFile As String)
    'setear modo
    Mode = MODE_INPUT_WAVE + MODE_OUTPUT_FILE
    
    
    'mostrar info en form
    Me.lblMode.Caption = "Record On-Fly"
    Me.lblInput.Caption = GetNameDevice(DEVICE_ID)
    Me.lblOutput.Caption = mp3OutFile
    Me.sbInfo.Panels(1).Text = ""
    
    'desactivar el modulo ShoutCast
    Unload frmShoutCast
    
End Function
Public Function cmdNewEncodeFile(ByVal WavFile As String, ByVal Mp3File As String)
    
    'setear modo
    Mode = MODE_INPUT_FILE + MODE_OUTPUT_FILE
    wavInFile = WavFile
    mp3OutFile = Mp3File
    
    'mostrar archivos en el form
    Me.lblMode.Caption = "Encode File"
    Me.lblInput.Caption = wavInFile
    Me.lblOutput.Caption = mp3OutFile
    Me.sbInfo.Panels(1).Text = ""
    
End Function


Public Function cmdNewShoutCast()
    
    'setear modo
    Mode = MODE_INPUT_WAVE + MODE_OUTPUT_SHOUTCAST
    wavInFile = ""
    mp3OutFile = ""
    
    'mostrar archivos en el form
    Me.lblMode.Caption = "Live ShoutCast - mp3 stream on the net"
    Me.lblInput.Caption = GetNameDevice(DEVICE_ID)
    'Me.lblOutput.Caption = "Output: "
    
    frmShoutCast.StartServer
    
End Function


Sub Monitor(ByVal pWave As Long, Largo As Long)
    On Error GoTo err
    Static waveData(1 To 4) As Byte
    Static L As Long
    Static R As Long
    Static aL As Long
    Static aR As Long
    
    CopyMemory waveData(1), ByVal pWave, 4 'Largo
    
    'If wformat.nChannels = 2 Then
    '    If wformat.wBitsPerSample = 16 Then
    L = Amplitud(waveData(1), waveData(2))
    R = Amplitud(waveData(3), waveData(4))
    '    Else
    '        L = Amplitud(waveData(1))
    '        R = Amplitud(waveData(2))
    '    End If
    'Else
    '    If wformat.wBitsPerSample = 16 Then
    '        L = Amplitud(waveData(1), waveData(2))
    '    Else
    '        L = Amplitud(waveData(1))
    '    End If
    '
    'End If
    
    L = Abs(L)
    R = Abs(R)
    
    If aL - L > 500 Then L = aL - 500
    If aR - R > 500 Then R = aR - 500
    
    'Me.Shape1.Width = Abs(L) * 1515 / 33000
    'Me.Shape2.Width = Abs(R) * 1515 / 33000
    
    
    ' BitBlt Me.Spectrum.hDC, 0, 0, Me.Width, Me.Height, Me.Picture2.hDC, (Abs(R) * Me.Picture2.Width / 33000), (Abs(R) * Me.Picture2.Height / 33000), SRCCOPY
    'BitBlt Me.Picture3, 0, 0, Me.Picture3.Width, Me.Picture3.Height, Me.Picture2.hDC, 0, 0, SRCCOPY
    Me.Spectrum.Cls
    BitBlt Me.Spectrum.hDC, 0, 0, L * Me.Picture2.ScaleX(Me.Picture2.Width) / 65000, Me.Picture2.ScaleY(Me.Picture2.Height), Me.Picture2.hDC, 0, 0, SRCCOPY
    BitBlt Me.Spectrum.hDC, 0, 10, R * Me.Picture2.ScaleX(Me.Picture2.Width) / 65000, Me.Picture2.ScaleY(Me.Picture2.Height), Me.Picture2.hDC, 0, 0, SRCCOPY
    
    aL = L
    aR = R
    
    'Me.Picture2.Refresh
    Exit Sub
err:
    Me.Caption = err.Description
End Sub

Function Amplitud(Byte1 As Byte, Optional Byte2 As Byte) As Long
    'Dim Res As Long
    
    Amplitud = Byte1 + (Byte2 * (2 ^ 8))
    
    'If Byte2 = 0 Then Exit Sub
    
    If Byte2 >= 128 Then
        Amplitud = Amplitud - 65536
    End If
    
End Function



Public Function Frequency2Long(ByVal indice As Long) As Long
    indice = Abs(indice)
    Frequency2Long = Choose((indice Mod 8) + 1, 44100, 32000, 24000, 22050, 16000, 12000, 11025, 8000)
End Function

Public Function Bitrate2Long(ByVal indice As Integer) As Long
    indice = Abs(indice)
    Bitrate2Long = Choose((indice Mod 14) + 1, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
End Function

Public Function EncodeMode2String(indice As Long, strResp As String) As Long
    indice = Abs(indice)
    strResp = CStr(Choose((indice Mod 5) + 1, "MONO", "STEREO", "JOINT", "msSTEREO", "DUAL"))
    EncodeMode2String = indice Mod 5
End Function

Public Function Enphasys2String(indice As Long, strResp As String) As Long
    indice = Abs(indice)
    strResp = CStr(Choose((indice Mod 3) + 1, "No-EMPH", "50/15ms", "CCITT"))
    Enphasys2String = indice Mod 3
End Function

'mostrar la configuracion actual de gogo.dll
Sub getCurrentConfig()
    Dim mStr As String
    
    If objGogo.State = CLOSED Then Exit Sub
    
    mCountFrame = objGogo.CountFrame
    frmGogoLive.lblDer(0).Caption = objGogo.OutputFrequency & " Hz": frmGogoLive.lblDer(0).Visible = True
    frmGogoLive.lblDer(1).Caption = objGogo.BitRate & " Kbps": frmGogoLive.lblDer(1).Visible = True
    
    Call EncodeMode2String(objGogo.EncodeMode, mStr)
    frmGogoLive.lblDer(2).Caption = mStr: frmGogoLive.lblDer(2).Visible = True
    
    Call Enphasys2String(objGogo.Emphasys, mStr)
    frmGogoLive.lblDer(3).Caption = mStr: frmGogoLive.lblDer(3).Visible = True
    
    
    If objGogo.UsePsy Then frmGogoLive.lblIzq(0).Enabled = False: frmGogoLive.lblIzq(0).Visible = True
    If objGogo.UseLPF16 Then frmGogoLive.lblIzq(1).Enabled = False: frmGogoLive.lblIzq(1).Visible = True
    If objGogo.UseMMX Then frmGogoLive.lblIzq(2).Enabled = False: frmGogoLive.lblIzq(2).Visible = True
    
    
End Sub

