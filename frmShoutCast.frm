VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmShoutCast 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmShoutCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''    Module written by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    http://orbita.starmedia.com/gogolive/ ''
''    Release Jan, 2004                     ''
''                                          ''
''    sorry about my English!               ''
''    your comments are welcome             ''
''''''''''''''''''''''''''''''''''''''''''''''

'Const MP3DATA_LENGTH = 40000     'metadata

Public ListenConnected As Boolean  'true si hay alguien escuchando
'Public TitleName As String  ' Max 15 chars
'Dim lMp3Data As Long    'metadata
'Dim lMetaData As Long   'metadata
Dim isHeaderSended As Boolean

Private Sub Form_Unload(Cancel As Integer)
    ListenConnected = False
End Sub

Private Sub sckMain_Close()
    EndServer
    StartServer
End Sub

Private Sub sckMain_ConnectionRequest(ByVal requestID As Long)
    frmGogoLive.sbInfo.Panels(1).Text = "Incoming Listen..."
    
    If Me.sckMain.State <> sckClosed Then Me.sckMain.Close
    Me.sckMain.Accept requestID
    
    
End Sub

Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
    Dim var As String
    
    Me.sckMain.GetData var, vbString, bytesTotal
    Me.Caption = "timer:" & Timer
    
    If isHeaderSended = False Then
        Dim header As String
        
        'preparar la respuesta en el header
        'header = "ICY 200 OK" & vbCrLf & _
        '         "icy-metaint: " & MP3DATA_LENGTH & vbCrLf & _
        '         "icy-br: " & mp3Kbps & vbCrLf & vbCrLf
        '
                 
        header = "ICY 200 OK" & vbCrLf & _
                 "icy-name:" & TITLE_NAME & vbCrLf & _
                 "icy-br:" & mp3Kbps & vbCrLf & vbCrLf
        
        'enviar header
        Me.sckMain.SendData header
        
        isHeaderSended = True
        ListenConnected = True
        
        frmGogoLive.sbInfo.Panels(1).Text = "Listener: " & Me.sckMain.RemoteHostIP
    End If
    
End Sub

Public Sub SenDataToListen(ByVal hBuffer As Long, ByVal Largo As Long)
    Dim sBuffer As String
    'Dim pos As Long
    
    sBuffer = Space(Largo)
    Call CopyMemory(ByVal sBuffer, ByVal hBuffer, Largo)
    
    'esta parte es para implementar la incropracion de Metadata durante la trasmision
    'If lMp3Data + Largo > (MP3DATA_LENGTH - 1) Then
    '    pos = (MP3DATA_LENGTH - 1) - lMp3Data
    '    lMp3Data = Len(sBuffer) - pos
    '    sBuffer = Left(sBuffer, pos) & MetadataPrepare() & Mid(sBuffer, pos + 1)
    'ElseIf lMp3Data + Largo = (MP3DATA_LENGTH - 1) Then
    '    pos = (MP3DATA_LENGTH - 1) - lMp3Data
    '    lMp3Data = Len(sBuffer) - pos
    '    sBuffer = Left(sBuffer, pos) & MetadataPrepare()
    'Else
    '    lMp3Data = lMp3Data + Largo
    'End If
    
    
    If Me.sckMain.State <> sckClosed Then Me.sckMain.SendData sBuffer
    
End Sub


Public Sub StartServer()
    Me.sckMain.Close
    Me.sckMain.Protocol = sckTCPProtocol
    
    Me.sckMain.LocalPort = PORT_LISTEN
    Me.sckMain.Listen
    
    ListenConnected = False
    isHeaderSended = False
    
    'mostrar configuracion en el form principal
    frmGogoLive.lblOutput.Caption = Me.sckMain.LocalIP & ":" & Me.sckMain.LocalPort
    frmGogoLive.sbInfo.Panels(1).Text = "Waiting for Listener..."

End Sub

Public Sub EndServer()
    Me.sckMain.Close
    ListenConnected = False
    isHeaderSended = False
    
    'mostrar configuracion en el form principal
    frmGogoLive.lblOutput.Caption = "No service started"
    frmGogoLive.sbInfo.Panels(1).Text = ""
    
End Sub


'Private Function MetadataPrepare() As String
'    Dim mStr As String
'
'
'    mStr = "StreamTitle='title of the song'"
'    mStr = Chr((Len(mStr) + 1) / 16) & mStr
'
'    MetadataPrepare = mStr
'End Function

