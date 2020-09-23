VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''
''    Class written by Fernando Aldea G.    ''
''    e-mail: fernando_aldea@latinmail.com  ''
''    Release January, 2003                 ''
''                                          ''
''    sorry for not translate completly     ''
''    & sorry about my English!               ''
''''''''''''''''''''''''''''''''''''''''''''''

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type


Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lpFlags As Long, lpOutBufferSize As Long, lpInBufferSize As Long, lpMaxInstances As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long



Private hReadPipe As Long
Private hWritePipe As Long
Private mvarnBytesTotal As Long
Private mvarnBytesActual As Long
Private mvarestaListo As Boolean

Public Property Get nBytesTotal() As String
    nBytesTotal = mvarnBytesTotal
End Property

Public Function Crear(nBytes As Long) As Boolean
    Dim sa As SECURITY_ATTRIBUTES
    Dim Ret As Long
    
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&


    Ret = CreatePipe(hReadPipe, hWritePipe, sa, nBytes)
       
    
    If Ret <> 0 Then
        Crear = True
        mvarestaListo = True
        Actualizar_nBytesActual
        'Actualizar_nBytesTotal
        mvarnBytesTotal = nBytes
    Else
        mvarnBytesTotal = 0
        mvarnBytesActual = 0
        Crear = False
        mvarestaListo = False
    End If

End Function

Public Function Destruir()
    GlobalFree hReadPipe
    GlobalFree hWritePipe
    Ret = CloseHandle(hReadPipe)
    Ret = CloseHandle(hWritePipe)
    mvarestaListo = False
End Function

Public Function toRead(ByVal pBuf As Long, nLargo As Long) As Long
    Dim Leidos As Long
        
    
    If ReadFile(hReadPipe, pBuf, nLargo, Leidos, 0&) = 1 Then
        toRead = Leidos
        mvarnBytesActual = mvarnBytesActual - Leidos
    Else
        toRead = 0
    End If
    
    
End Function

Public Function toWrite(ByVal pBuf As Long, nLargo As Long) As Long
    Dim Escritos As Long
    If hWritePipe = 0 Then Beep: MsgBox "error"
    If WriteFile(hWritePipe, ByVal pBuf, nLargo, Escritos, 0&) = 1 Then
        toWrite = Escritos
        mvarnBytesActual = mvarnBytesActual + Escritos
    Else
        toWrite = 0
    End If
End Function

Private Sub Class_Terminate()
    Destruir
End Sub

Public Property Get nBytesActual() As Long
    nBytesActual = mvarnBytesActual
End Property

Public Function Actualizar_nBytesActual()
    Dim Remain As Long
    
    Call PeekNamedPipe(hReadPipe, 0&, 0&, 0&, Remain, 0&)
    
    mvarnBytesActual = Remain
    Actualizar_nBytesActual = Remain
End Function
Public Function Actualizar_nBytesTotal()
    'Dim avRead As Long, avWrite As Long
    
    'Call GetNamedPipeInfo(hReadPipe, 0&, avRead, avWrite, 0&)
    
    'mvarnBytesTotal = avRead + avWrite
    'Actualizar_nBytesTotal = mvarnBytesTotal
    
    
    Actualizar_nBytesTotal = mvarnBytesTotal
    
End Function

Public Property Get estaListo() As Boolean
    estaListo = mvarestaListo
End Property

