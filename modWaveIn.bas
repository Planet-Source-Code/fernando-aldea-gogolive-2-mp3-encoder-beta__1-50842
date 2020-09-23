Attribute VB_Name = "modWaveIn"
''''''''''''''''''''''''''''''''''''''''''''''
''    Module adapted by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    web: orbita.starmedia.com/gogolive/   ''
''    Release Jan, 2004                     ''
''                                          ''
''    sorry for not translate this completly''
''    & sorry about my English!             ''
''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Const MAXPNAMELEN = 32  '  longitud máx. del nombre del producto (incluido NULL)
Public Const MAXERRORLENGTH = 128  '  longitud máx. del texto de error (incluido NULL final)

Public Const WAVERR_BASE = 32
'  valores de retorno de los errores de audio de forma de onda
Public Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)     '  formato de onda no compatible
Public Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)  '  todavía reproduce algo
Public Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)    '  encabezado no preparado
Public Const WAVERR_SYNC = (WAVERR_BASE + 3)          '  dispositivo síncrono
Public Const WAVERR_LASTERROR = (WAVERR_BASE + 3)     '  último error del intervalo



Public Const CALLBACK_FUNCTION = &H30000
Public Const CALLBACK_WINDOW = &H10000
Private Const MM_WIM_DATA = &H3C0
Private Const MM_WIM_CLOSE      As Long = &H3BF
Private Const MM_WIM_OPEN       As Long = &H3BE

Private Const WHDR_DONE = &H1         '  done bit
'Public Const WIM_DATA = MM_WIM_DATA
Public Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin
Public Const NUM_BUFFERS = 10
Public BUFFER_SIZE As Long  '= 8192
Public Const MAPPER_ID = -1
Public Const GWL_WNDPROC = -4
Public Const WAVE_FORMAT_QUERY = &H1

'callback mode
Public Const CM_WINDOWS = 1
Public Const CM_FUNCTION = 2
Public Const CM_QUERY = 3


Type WAVEHDR
    lpData As Long          ' Address of the waveform buffer.
    dwBufferLength As Long  ' Length, in bytes, of the buffer.
    dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
    ' data is in the buffer.
    
    dwUser As Long          ' User data.
    dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
    dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
    lpNext As Long          ' Not used
    reserved As Long        ' Not used
End Type

Type WAVEFORMAT
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Public Type WAVEINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
End Type

'  definiciones para el campo dwFormat de WAVEINCAPS y WAVEOUTCAPS
Public Const WAVE_INVALIDFORMAT = &H0            '  formato no válido
Public Const WAVE_FORMAT_1M08 = &H1              '  11,025 KHz, Mono,   8 bits
Public Const WAVE_FORMAT_1S08 = &H2              '  11,025 KHz, Estéreo, 8 bitss
Public Const WAVE_FORMAT_1M16 = &H4              '  11,025 KHz, Mono,   16 bits
Public Const WAVE_FORMAT_1S16 = &H8              '  11,025 KHz, Estéreo, 16 bits
Public Const WAVE_FORMAT_2M08 = &H10             '  22,05  KHz, Mono,   8 bits
Public Const WAVE_FORMAT_2S08 = &H20             '  22,05  KHz, Estéreo, 8 bits
Public Const WAVE_FORMAT_2M16 = &H40             '  22,05  KHz, Mono,   16 bits
Public Const WAVE_FORMAT_2S16 = &H80             '  22,05  KHz, Estéreo, 16 bits
Public Const WAVE_FORMAT_4M08 = &H100            '  44,1   KHz, Mono,   8 bits
Public Const WAVE_FORMAT_4S08 = &H200            '  44,1   KHz, Estéreo, 8 bits
Public Const WAVE_FORMAT_4M16 = &H400            '  44,1   KHz, Mono,   16 bits
Public Const WAVE_FORMAT_4S16 = &H800            '  44,1   KHz, Estéreo, 16 bits


Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long

' Global Memory Flags
'Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub CopyStringFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal a As String, p As Any, ByVal cb As Long)
Public Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Public Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)

'Private Const GWL_WNDPROC       As Long = -4
Private Const CW_USEDEFAULT     As Long = &H80000000

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByRef lParam As WAVEHDR) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal ExStyle As Long, ByVal ClassName As String, ByVal WindowName As String, ByVal Style As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal WndParent As Long, ByVal Menu As Long, ByVal Instance As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long


'------------
Type CRITICAL_SECTION
    dummy As Long
End Type

Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)





Dim hWaveIn As Long
Dim wformat As WAVEFORMAT
Dim hmem(NUM_BUFFERS) As Long
Dim inHdr(NUM_BUFFERS) As WAVEHDR
Dim fProcessing As Boolean
Dim fPausing As Boolean
Dim lpPrevWndProc As Long
Dim mvarCallBackMode As Integer
Public largoSamples As Long
Dim msg As String * 200
Dim hWndCallBack As Long
Dim waveCriticalSection As CRITICAL_SECTION
Dim initWaveCriticalSection As Long

Public Property Get isPausing() As Boolean
    isPausing = fPausing
End Property

Public Property Get isProcessing() As Boolean
    isProcessing = fProcessing
End Property

Public Property Get nAvgBytesPerSec() As Long
    nAvgBytesPerSec = wformat.nAvgBytesPerSec
End Property

'start audio input from soundcard
Function StartInput(Optional lBuffer As Long) As Boolean
    Dim i As Long
    Dim rc As Long
    
    If mvarCallBackMode <= 0 Then
        MsgBox "Not initialized", vbCritical
        Exit Function
    End If
    
    If fProcessing Then
        StartInput = True
        Exit Function
    End If
    
    BUFFER_SIZE = (wformat.nSamplesPerSec * wformat.nBlockAlign * wformat.nChannels * 0.1) - ((wformat.nSamplesPerSec * wformat.nBlockAlign * wformat.nChannels * 0.1) Mod (wformat.nBlockAlign))
    BUFFER_SIZE = BUFFER_SIZE * 2
    
    If lBuffer > 0 Then BUFFER_SIZE = lBuffer
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next
    
    
    If mvarCallBackMode = CM_FUNCTION Then
        rc = waveInOpen(hWaveIn, DEVICE_ID, wformat, AddressOf WaveProc, 0, CALLBACK_FUNCTION)
        
        'ok. aqui viene la parte importante. para prevenir deadlock en la funcion WaveInProc
        'se hace uso de uan Seccion Critica
        Call InitializeCriticalSection(waveCriticalSection)
        'Se registra valor inicial de Critical Section
        initWaveCriticalSection = waveCriticalSection.dummy
        
    ElseIf mvarCallBackMode = CM_WINDOWS Then
        rc = waveInOpen(hWaveIn, DEVICE_ID, wformat, hWndCallBack, 0, CALLBACK_WINDOW)
    End If
    
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg, vbCritical
        StartInput = False
        Exit Function
    End If
    
    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical
        End If
    Next
    
    For i = 0 To NUM_BUFFERS - 1
        addData inHdr(i)
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical
        End If
    Next
    
    largoSamples = (wformat.wBitsPerSample / 8) * wformat.nChannels
    If largoSamples = 0 Then Beep: largoSamples = 1
    
    
    fProcessing = True
    rc = waveInStart(hWaveIn)
    StartInput = True
    
End Function

Sub addData(iHdr As WAVEHDR)
    Dim sBuff  As String
    Dim rc As Long
    
    rc = waveInAddBuffer(hWaveIn, iHdr, Len(iHdr))
    sBuff = Space(BUFFER_SIZE)
    CopyMemory ByVal sBuff, ByVal iHdr.lpData, BUFFER_SIZE
End Sub

' Stop receiving audio input on the soundcard
Sub StopInput()
    Dim iRet As Long
    Dim i As Long
    
    fProcessing = False
    iRet = waveInReset(hWaveIn)
    iRet = waveInStop(hWaveIn)
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    
    If mvarCallBackMode = CM_FUNCTION Then
        Do
            DoEvents
            frmGogoLive.Caption = waveCriticalSection.dummy
        Loop While (waveCriticalSection.dummy > initWaveCriticalSection)
        Call DeleteCriticalSection(waveCriticalSection)
        
    ElseIf mvarCallBackMode = CM_WINDOWS Then
    
    'destruir la ventana previa si habia una
        If hWndCallBack <> 0 Then
            Call SetWindowLong(hWndCallBack, GWL_WNDPROC, lpPrevWndProc)
            Call DestroyWindow(hWndCallBack)
        End If
    End If
    
    iRet = waveInClose(hWaveIn)
    
End Sub

' Initialize input soundcard service
Public Function Initialize(ByVal DeviceID As Long, Frequency As Long, ByVal Stereo As Boolean, ByVal nBits As Long, Mode As Long) As Boolean
    Dim rc As Long
    Dim msg As String
    
    If (Mode <> CM_FUNCTION) And (Mode <> CM_WINDOWS) And (Mode <> CM_FUNCTION) And (Mode <> CM_QUERY) Then
        'error in callback mode
        Exit Function
    End If
    
    'Set The WAV wformat
    wformat.wFormatTag = 1
    If Stereo Then
        wformat.nChannels = 2
    Else
        wformat.nChannels = 1
    End If
    wformat.wBitsPerSample = nBits
    wformat.nSamplesPerSec = Frequency
    wformat.nBlockAlign = wformat.nChannels * wformat.wBitsPerSample / 8
    wformat.nAvgBytesPerSec = wformat.nSamplesPerSec * wformat.nBlockAlign
    wformat.cbSize = Len(wformat)
    
    If Mode = CM_QUERY Then
        'Determianr si el formato es soportado por el dispositivo
        rc = waveInOpen(0&, DeviceID, wformat, 0&, 0&, WAVE_FORMAT_QUERY)
        
        If rc = WAVERR_BADFORMAT Then
            Exit Function
        ElseIf rc <> 0 Then
            waveInGetErrorText rc, msg, Len(msg)
            'MsgBox msg, vbCritical
            Exit Function
        Else
            Initialize = True
            Exit Function
        End If
        
    ElseIf Mode = CM_WINDOWS Then
        'hWnd = hwndIn
        
        'destruir la ventana previa si habia una
        If hWndCallBack <> 0 Then
            Call SetWindowLong(hWndCallBack, GWL_WNDPROC, lpPrevWndProc)
            Call DestroyWindow(hWndCallBack)
        End If
        'Exit Function
        'crear una ventana que procese mensajes
        hWndCallBack = CreateWindowEx(0, "STATIC", vbNullString, 0, CW_USEDEFAULT, 0, 0, 0, 0, 0, App.hInstance, 0)
        'asociamos nuestra funcion que procesa mensajes a la ventana
        lpPrevWndProc = SetWindowLong(hWndCallBack, GWL_WNDPROC, AddressOf WindowProc)
        
    End If
    
    mvarCallBackMode = Mode
    Initialize = True
End Function

'procedure for windows callback mode
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByRef wavhdr As WAVEHDR) As Long
    Dim i As Integer
    Dim rc As Long
    
    If uMsg = MM_WIM_DATA Then
        ' Process sound buffer if Processing
        If (fProcessing) Then
            For i = 0 To (NUM_BUFFERS - 1)
                If inHdr(i).dwFlags And WHDR_DONE Then
                    
                    If Not (fPausing) Then
                        frmGogoLive.callBackWave inHdr(i).lpData, BUFFER_SIZE   '<--------  sub necesaria en el form
                    End If
                    
                    rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
                    If rc <> 0 Then
                        MsgBox "Failed (WaveInAddBuffer)", vbCritical
                    End If
                    
                End If
            Next i
        End If
    ElseIf uMsg = MM_WIM_OPEN Or uMsg = MM_WIM_CLOSE Then
        '
    Else
        'WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, wavhdr)
    End If
    
    
End Function

'procedure for function callback mode
Sub WaveProc(ByVal hw As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByVal wParam As Long, ByRef wavhdr As WAVEHDR)
    Dim i As Integer
    Dim rc As Long
    
    ' Process sound buffer if Processing
    If (fProcessing) Then
        For i = 0 To (NUM_BUFFERS - 1)
            If inHdr(i).dwFlags And WHDR_DONE Then
                
                'si no existe un pausa procesamos los datos
                If Not (fPausing) Then
                    'Esta es la Seccion Critica. Aqui evitamos que el thread principal
                    'del sistema que llama a WaveProc manipule los datos del buffer
                    'antes que lo procesemos
                    Call EnterCriticalSection(waveCriticalSection)
                    
                    'procesar los datos
                    frmGogoLive.callBackWave inHdr(i).lpData, BUFFER_SIZE   '<--------  sub necesaria en el form
                    
                    'Salimos de la Seccion Critica
                    Call LeaveCriticalSection(waveCriticalSection)
                    
                End If
                                
                rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
                If rc <> 0 Then
                    MsgBox "Failed (waveInAddBuffer)", vbCritical
                End If
                
            End If
        Next i
    End If
End Sub

' Stop receiving audio input on the soundcard
Sub PauseInput()
    fPausing = Not fPausing
End Sub

Public Function GetNameDevice(ByVal DeviceID As Long) As String
    Dim WV As WAVEINCAPS
    
    waveInGetDevCaps DeviceID, WV, Len(WV)
    GetNameDevice = WV.szPname
    
End Function
