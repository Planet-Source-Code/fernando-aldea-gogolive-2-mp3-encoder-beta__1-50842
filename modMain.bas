Attribute VB_Name = "modMain"
''''''''''''''''''''''''''''''''''''''''''''''
''    Module written by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    http://orbita.starmedia.com/gogolive/ ''
''    Release Jan, 2004                     ''
''                                          ''
''   sorry for not translate this completly ''
''    & sorry about my English!             ''
''''''''''''''''''''''''''''''''''''''''''''''

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


'kernel
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Public Const MODE_INPUT_WAVE = 1
Public Const MODE_INPUT_FILE = 2
Public Const MODE_OUTPUT_WAVE = 4
Public Const MODE_OUTPUT_FILE = 8
Public Const MODE_OUTPUT_SHOUTCAST = 16

Public Const DIALOGTYPE_OPENSAVE = 1
Public Const DIALOGTYPE_SAVE = 2

Public mp3Frequency As Long
Public mp3Kbps As Long
Public mp3Mode As Long
Public mp3Enphasis As Long
Public mp3OutFile As String
Public mp33DNow As Boolean
Public mp3PSY As Boolean
Public mp3MMX As Boolean
Public mp3LPF16 As Boolean

Public wavInFile As String


Public TimerIni As Single
Public nBytesPCMacum As Long
Public Processing As Boolean
Public Monitoring As Boolean
Public BUFFER_LENGTH As Long
Public PORT_LISTEN As Long
Public TITLE_NAME As String
Public OUTFILE_DEFAULT As String
Public DEVICE_ID As Long
Public Mode As Integer

Global objGogo As New clsGogo
Global mPipe As New cPipe

Sub Main()
    mp3Frequency = 44100
    mp3Kbps = 128
    mp3Mode = 2 'joint stereo
    mp3Enphasis = 0
    mp3PSY = True
    mp33DNow = True
    mp3LPF16 = False
    mp3MMX = True
    OUTFILE_DEFAULT = "c:\newGOGO.mp3"
    mp3OutFile = OUTFILE_DEFAULT
    BUFFER_LENGTH = 5000000 '10000000
    PORT_LISTEN = 8000
    TITLE_NAME = "GoGoLive stream"
    DEVICE_ID = -1
    
    frmGogoLive.Show
    
    frmGogoLive.cmdNewOnfly OUTFILE_DEFAULT
    
End Sub


Sub StartProcess()
    Dim tStereo As Boolean
    
    'Configuramos el formato de salida del mp3
    objGogo.BitRate = mp3Kbps
    objGogo.EncodeMode = mp3Mode
    objGogo.InputFrequency = mp3Frequency
    objGogo.OutputFrequency = mp3Frequency
    objGogo.Emphasys = mp3Enphasis
    
    'switches
    objGogo.UsePsy = mp3PSY
    objGogo.UseMMX = mp3MMX
    objGogo.UseLPF16 = mp3LPF16
    objGogo.Use3DNow = mp33DNow
    
    'si el modo de salida es un archivo
    If (Mode And MODE_OUTPUT_FILE) = MODE_OUTPUT_FILE Then
        'se indica la ruta al archivo
        objGogo.OutputFile = mp3OutFile
        
    'sino, usa la funcion de usuario.
    Else
        ' una cadena vacia en la propiedad OutputFile indica que
        ' se usara la funcion de usuario
        objGogo.OutputFile = ""
    End If
    
    
    'si el modo de entrada de datos es un archivo
    If (Mode And MODE_INPUT_FILE) = MODE_INPUT_FILE Then
        'se indica la ruta al archivo
        objGogo.InputFile = wavInFile
        
    'sino, usa la funcion de usuario.
    Else
        ' una cadena vacia en la propiedad InputFile indica que
        ' se usara la funcion de usuario
        objGogo.InputFile = ""
        
        'determinar modo del wave (stereo o mono)
        tStereo = IIf(mp3Mode <> MC_MODE_MONO, True, False)
        
        'Initialize Wave Service
        If Not modWaveIn.Initialize(DEVICE_ID, mp3Frequency, tStereo, 16, CM_WINDOWS) Then
            MsgBox "Error in WAVE", vbCritical
            Exit Sub
        End If
        
        TimerIni = Timer  'security for timer record
        nBytesPCMacum = -1 '...
        
    End If
    
    'Preparar el codificador
    If Not objGogo.OpenGogo() Then
        MsgBox "Error in GOGO", vbCritical
        Exit Sub
    End If
    
    'mostrar la configuracion actual
    frmGogoLive.getCurrentConfig
    
    '
    If ((Mode And MODE_INPUT_WAVE) = MODE_INPUT_WAVE) Then
        'Start receiving audio input
        If StartInput() Then
            Processing = True
        Else
            MsgBox "Error in StartInput", vbCritical
            Exit Sub
        End If
    End If
    
    'codificar
    objGogo.StarEncode
    
    'Release Gogo Library
    objGogo.CloseGogo
    
End Sub

Sub EndProcess()
    StopInput
    'objGogo.StopEncode
    Processing = False
End Sub


Public Function PauseRecord()
    PauseInput
End Function


'util
Public Function GetAddressofFunction(ByVal pFunction As Long) As Long
    GetAddressofFunction = pFunction
End Function

