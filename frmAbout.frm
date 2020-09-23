VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de MiApli"
   ClientHeight    =   3180
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6120
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2194.893
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   5746.997
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   60
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   4440
      TabIndex        =   0
      Top             =   2700
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "by Fernando Aldea  e-mail:fernando_aldea@terra.cl"
      Height          =   225
      Left            =   720
      TabIndex        =   6
      Top             =   420
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   56.343
      X2              =   5521.624
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Label lblDescription 
      Caption         =   "This is the Fastest MP3 Encoder in Real Time. "
      ForeColor       =   &H00000000&
      Height          =   1650
      Left            =   660
      TabIndex        =   2
      Top             =   960
      Width           =   5325
   End
   Begin VB.Label lblTitle 
      Caption         =   "GOGOLive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   660
      TabIndex        =   4
      Top             =   120
      Width           =   5265
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5408.938
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Label lblVersion 
      Height          =   225
      Left            =   660
      TabIndex        =   5
      Top             =   660
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "No me hago responsable de los eventuales daños del uso de este software y bla, bla.."
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   180
      TabIndex        =   3
      Top             =   2700
      Width           =   3990
   End
End
Attribute VB_Name = "frmAbout"
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

Option Explicit

' Opciones de seguridad de claves del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos principales de claves del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en Null
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long



Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Acerca de " & App.Title
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title & " - Fast MP3 Encoder"
    
    Me.lblDescription = Me.lblDescription & vbNewLine & _
                        "- Record to Mp3 from you Sound Card Directly (analog mode)." & vbNewLine & _
                        "- Send mp3 stream on the net to other user (Listener need Winamp)." & vbNewLine & _
                        "- Encode File wav to mp3." & vbNewLine & vbNewLine & _
                        "   Any Problem or Comment to: fernando_aldea@terra.cl" & vbNewLine & _
                        "   web:  http://orbita.starmedia.com/gogolive/"
End Sub


