VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''
''    Class written by Fernando Aldea G.    ''
''    e-mail: fernando_aldea@terra.cl       ''
''    http://orbita.starmedia.com/gogolive/ ''
''    Release Jan, 2004                     ''
''                                          ''
''    sorry about my English!               ''
''    your comments are welcome             ''
''''''''''''''''''''''''''''''''''''''''''''''

'diferentes estados de la clase
Private Const ENCODING = 1
Private Const OPENED = 2
Private Const CLOSED = 3

'variables locales que contienen valores de propiedad
Private mvarEncodeMode As Long
Private mvarBitRate As Long
Private mvarInputFrequency As Long
Private mvarOutputFrequency As Long
Private mvarStartOffset As Long
Private mvarUsePsy As Boolean
Private mvarUseLPF16 As Boolean
Private mvarUseMMX As Boolean
Private mvarUse3DNow As Boolean
Private mvarUseKNI As Boolean
Private mvarUseE3DNow As Boolean
Private mvarEmphasys As Long
Private mvarVersion As String
Private mvarMPEGVersion As String
Private mvarState As Long
Private mvarCurrentFrame As Long
Private mvarCountFrame As Long
Private mvarInputFile As String
Private mvarOutputFile As String

'variables locales
Dim Resp As Long
Dim value As Long
Dim ReqStop  As Boolean

':::::::::::::: CountFrame Property (solo lectura)::::::::::::::::::::
Public Property Get CurrentFrame() As Long
    CurrentFrame = mvarCurrentFrame
End Property


':::::::::::::: CountFrame State:::::::::::::::::::::::::::
Public Property Let State(ByVal vData As Long)
    mvarState = vData
End Property

Public Property Get State() As Long
    State = mvarState
End Property


':::::::::::::: CountFrame Property:::::::::::::::::::::::::::
Public Property Let CountFrame(ByVal vData As Long)
    
    'no permitir cambiar la configuracion una vez que empezo la codificacion
    If State = ENCODING Then Exit Property
    
    mvarCountFrame = vData
    
End Property

Public Property Get CountFrame() As Long
    
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_COUNT_FRAME, mvarCountFrame)
    If Resp <> 0 Then
        'err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    CountFrame = mvarCountFrame
    
End Property


Public Property Let MPEGVersion(ByVal vData As String)
    
    mvarMPEGVersion = vData
End Property

Public Property Get MPEGVersion() As String
    
    'obtener la configuracion actual de gogo.dll
    'Resp = MPGE_getConfigureVB(MG_MPEG_VERSION, value)
    'If Resp <> 0 Then
    '    err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
    '    Exit Property
    'End If
    
    MPEGVersion = mvarMPEGVersion
End Property

':::::::::::::: Version Property (solo lectura) :::::::::::::::::::
Public Property Get Version() As String
    Version = mvarVersion
End Property


':::::::::::::: CPU Property:::::::::::::::::::::::::::
Public Property Let CPU(ByVal vData As Long)
    
    'no permitir cambiar la configuracion una vez que empezo la codificacion
    If State = ENCODING Then Exit Property
    
    'establecer la configuracion de gogo.dll
    Resp = MPGE_setConfigureVB3(MC_CPU, vData, 0&)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarCPU = vData
    
End Property

Public Property Get CPU() As Long
    CPU = mvarCPU
End Property


':::::::::::::: Emphasys Property:::::::::::::::::::::::::::
Public Property Let Emphasys(ByVal vData As Long)
        
    'no permitir cambiar la configuracion una vez que empezo la codificacion
    If State = ENCODING Then Exit Property
    
    mvarEmphasys = vData
    
End Property

Public Property Get Emphasys() As Long

    Emphasys = mvarEmphasys
    
End Property


':::::::::::::: UseE3DNow Property:::::::::::::::::::::::::::
Public Property Let UseE3DNow(ByVal vData As Boolean)
    
    'no permitir cambiar la configuracion una vez que empezo a codificar
    If State = ENCODING Then Exit Property
    
    'establecer la configuracion de gogo.dll
    Resp = MPGE_setConfigureVB3(MG_USEE3DNOW, IIf(vData, 1, 0), 0&)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUseE3DNow = vData
End Property

Public Property Get UseE3DNow() As Boolean
        
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_USEE3DNOW, value)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUseE3DNow = CBool(value)
    UseE3DNow = mvarUseE3DNow
End Property


':::::::::::::: UseKNI Property:::::::::::::::::::::::::::
Public Property Let UseKNI(ByVal vData As Boolean)
    
    'no permitir cambiar la configuracion una vez que empezo a codificar
    If State = ENCODING Then Exit Property
    
    'establecer la configuracion de gogo.dll
    Resp = MPGE_setConfigureVB3(MG_USEKNI, IIf(vData, 1, 0), 0&)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUseKNI = vData
End Property

Public Property Get UseKNI() As Boolean
    
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_USEKNI, value)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUseKNI = CBool(value)
    UseKNI = mvarUseKNI
End Property


':::::::::::::: Use3DNow Property:::::::::::::::::::::::::::
Public Property Let Use3DNow(ByVal vData As Boolean)
    
    'no permitir cambiar la configuracion una vez que empezo a codificar
    If State = ENCODING Then Exit Property
    
    'establecer la configuracion de gogo.dll
    Resp = MPGE_setConfigureVB3(MG_USE3DNOW, IIf(vData, 1, 0), 0&)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUse3DNow = vData
End Property

Public Property Get Use3DNow() As Boolean
    
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_USE3DNOW, value)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUse3DNow = CBool(value)
    Use3DNow = mvarUse3DNow
End Property


':::::::::::::: UseMMX Property:::::::::::::::::::::::::::
Public Property Let UseMMX(ByVal vData As Boolean)
    
    'no permitir cambiar la configuracion una vez que empezo a codificar
    If State = ENCODING Then Exit Property
    
    'establecer la configuracion de gogo.dll
    Resp = MPGE_setConfigureVB3(MG_USEMMX, IIf(vData, 1, 0), 0&)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUseMMX = vData
    
End Property

Public Property Get UseMMX() As Boolean
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_USEMMX, value)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUseMMX = CBool(value)
    UseMMX = mvarUseMMX
End Property



Public Property Let UseLPF16(ByVal vData As Boolean)
    
    'no permitir cambiar la configuracion una vez que empezo a codificar
    If State = ENCODING Then Exit Property
    
    'establecer la configuracion de gogo.dll
    Resp = MPGE_setConfigureVB3(MC_USELPF16, IIf(vData, 1, 0), 0&)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUseLPF16 = vData
End Property


Public Property Get UseLPF16() As Boolean
    
    'Syntax: Debug.Print X.UseLPF16
    UseLPF16 = mvarUseLPF16
End Property


':::::::::::::: UsePsy Property:::::::::::::::::::::::::::
Public Property Let UsePsy(ByVal vData As Boolean)
    
    'no permitir cambiar la configuracion una vez que empezo a codificar
    If State = ENCODING Then Exit Property
    
    'establecer la configuracion de gogo.dll
    Resp = MPGE_setConfigureVB3(MC_USEPSY, IIf(vData, 1, 0), 0&)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUsePsy = vData
    
End Property

Public Property Get UsePsy() As Boolean
    
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_USEPSY, value)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    mvarUsePsy = CBool(value)
    UsePsy = mvarUsePsy
End Property



Public Property Let StartOffset(ByVal vData As Long)
    mvarStartOffset = vData
End Property

Public Property Get StartOffset() As Long
    StartOffset = mvarStartOffset
End Property


':::::::::::::: OutputFrequency Property:::::::::::::::::::::::::::
Public Property Let OutputFrequency(ByVal vData As Long)
    mvarOutputFrequency = vData
End Property

Public Property Get OutputFrequency() As Long
    
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_OUTFREQ, mvarOutputFrequency)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    OutputFrequency = mvarOutputFrequency
    
End Property


':::::::::::::: InputFrequency Property:::::::::::::::::::::::::::
Public Property Let InputFrequency(ByVal vData As Long)
    
    'no permitir cambiar la configuracion una vez que empezo la codificacion
    If State = ENCODING Then Exit Property
    
    mvarInputFrequency = vData
    
End Property

Public Property Get InputFrequency() As Long
    
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_INPFREQ, mvarInputFrequency)
    If Resp <> 0 Then
        'err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    InputFrequency = mvarInputFrequency
End Property


':::::::::::::: BitRate Property:::::::::::::::::::::::::::
Public Property Let BitRate(ByVal vData As Long)
    'no permitir cambiar la configuracion una vez que empezo la codificacion
    If State = ENCODING Then Exit Property
    
    mvarBitRate = vData
    
End Property

Public Property Get BitRate() As Long
    
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_BITRATE, mvarBitRate)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    BitRate = mvarBitRate
    
End Property


':::::::::::::: EncodeMode Property:::::::::::::::::::::::::::
Public Property Let EncodeMode(ByVal vData As Long)
    
    'no permitir cambiar la configuracion una vez que empezo la codificacion
    If State = ENCODING Then Exit Property
    
    mvarEncodeMode = vData
    
End Property

Public Property Get EncodeMode() As Long
    
    'obtener la configuracion actual de gogo.dll
    Resp = MPGE_getConfigureVB(MG_ENCODEMODE, mvarEncodeMode)
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Property
    End If
    
    EncodeMode = mvarEncodeMode
    
End Property


':::::::::::::: InputFile Property:::::::::::::::::::::::::::
Public Property Let InputFile(ByVal vData As String)
    
    'no permitir cambiar la configuracion una vez que empezo la codificacion
    If State = ENCODING Then Exit Property
    
    mvarInputFile = vData
    
End Property

Public Property Get InputFile() As String
    InputFile = mvarInputFile
End Property


':::::::::::::: OutputFile Property:::::::::::::::::::::::::::
Public Property Let OutputFile(ByVal vData As String)
    
    'no permitir cambiar la configuracion una vez que empezo la codificacion
    If State = ENCODING Then Exit Property
    
    mvarOutputFile = vData
    
End Property

Public Property Get OutputFile() As String
    OutputFile = mvarOutputFile
End Property


Public Function OpenGogo() As Boolean
    Dim Userfunc As MCP_INPDEV_USERFUNC
    Dim Resp As Long
    
'::::::: Configuracion de  Entrada ::::::::::
    'Hay un archivo de entrada
    If mvarInputFile <> "" Then
        'verificar la existencia del archivo de entrada
        If Dir(InputFile) = "" Then err.Raise 1, "Gogo", "Input file not exist": Exit Function
        
        'Archivo de entrada
        Resp = MPGE_setConfigureVB(MC_INPUTFILE, MC_INPDEV_FILE, InputFile)
        GoSub verify
        
    'Hay una funcion de entrada
    Else
        '
        Userfunc.nFreq = mvarInputFrequency
        Userfunc.nChn = IIf(mvarEncodeMode <> MC_MODE_MONO, 2, 1) 'mvarinputmode
        Userfunc.nBit = 16 'debug
        Userfunc.nSize = MC_INPDEV_MEMORY_NOSIZE
        Userfunc.pUserFunc = GetAddressofFunction(AddressOf InputCallbackFunction)
        
        '
        Resp = MPGE_setConfigureVB2(MC_INPUTFILE, MC_INPDEV_USERFUNC, Userfunc)
        GoSub verify
        
    End If
    
    
    
'::::::: Configuracion de  Salida ::::::::::
    'Hay un archivo de salida
    If OutputFile <> "" Then
        'verificar la existencia de un archivo de salida
        'If Dir(OutputFile) <> "" Then err.Raise 1, "Gogo", "Output file already exist": Exit Function
        'seccion arc out
        Resp = MPGE_setConfigureVB(MC_OUTPUTFILE, MC_OUTDEV_FILE, OutputFile)
        GoSub verify
        
    'Hay un funcion de salida
    Else
        'configurar metodo de salida de gogo.dll
        Resp = MPGE_setConfigureVB3(MC_OUTPUTFILE, MC_OUTDEV_USERFUNC, ByVal GetAddressofFunction(AddressOf OutputCallbackFunction))
        GoSub verify
        
    End If
    
    'Frecuencia
    Resp = MPGE_setConfigureVB3(MC_OUTFREQ, mvarOutputFrequency, 0&)
    GoSub verify
    'Modo: stereo, mono, joint....
    Resp = MPGE_setConfigureVB3(MC_ENCODEMODE, mvarEncodeMode, 0&)
    GoSub verify
    'Tasa de bits
    Resp = MPGE_setConfigureVB3(MC_BITRATE, mvarBitRate, 0&)
    GoSub verify
    'Enfasis
    Resp = MPGE_setConfigureVB3(MC_EMPHASIS, mvarEmphasys, 0&)
    GoSub verify
    'reviza la configuracion actual de gogo.dll
    Resp = MPGE_detectConfigureVB()
    GoSub verify
    
    OpenGogo = True
    mvarState = OPENED
    
    
    Exit Function
    
verify:
    If Resp <> 0 Then
        MsgBox GetError(CInt(Resp))
        Exit Function
    End If
    Return
End Function


Public Function CloseGogo() As Boolean
        
    Call MPGE_closeCoderVB
    Call MPGE_endCoderVB
    
    mvarState = CLOSED
    
End Function

Public Sub StopEncode()
    ReqStop = True
End Sub

Public Sub StarEncode()
    
    'actualizamos el total de frames a codificar (si es el modo de entrada es un archivo)
    Resp = MPGE_getConfigureVB(MG_COUNT_FRAME, mvarCountFrame)
    If Resp <> 0 Then GoTo ssalir
    
    mvarCurrentFrame = 0
    ReqStop = False
    mvarState = ENCODING
    
    'comenzar a codificar
    Do
        'procesar 1 frame
        Resp = MPGE_processFrameVB()
        If Resp = ME_NOERR Then mvarCurrentFrame = mvarCurrentFrame + 1
        
        'si se requiere detener la codificacion...
        If ReqStop Then
            Resp = ME_EMPTYSTREAM
            ReqStop = False
            Exit Do
        End If
        'permitir que otros eventos ocurran (pintar el form, mover, etc)
        DoEvents
        
    Loop Until (Resp <> ME_NOERR And Resp <> ME_MOREDATA)
    
    mvarState = OPENED
    
    If Resp <> ME_EMPTYSTREAM Then
        MsgBox ("ERROR: errcode = " & Resp & Chr$(13) & GetError(Resp))
    End If
    
    
ssalir:
    'end of encoding
    Call MPGE_closeCoderVB
    'free gogod.ll
    Call MPGE_endCoderVB
    
End Sub

Private Sub Class_Initialize()
    Dim mNum As Long
    Dim mStr As String
    
    'obtener la version de gogo.dll
    mStr = Space$(100)
    Resp = MPGE_getVersionVB(mNum, ByVal mStr)
    GoSub VerifyError
    mvarVersion = Left(mStr, Len(RTrim(mStr)) - 1)
    
    'iniciar .....
    Resp = MPGE_initializeWorkVB()
    GoSub VerifyError
    
    Call MPGE_endCoderVB
    
    mvarState = CLOSED
    
    Exit Sub
    
VerifyError:
    If Resp <> 0 Then
        err.Raise Resp, "gogo.dll", GetError(CInt(Resp))
        Exit Sub
    End If
    Return
    
End Sub
