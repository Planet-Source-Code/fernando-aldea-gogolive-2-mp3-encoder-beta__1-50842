Attribute VB_Name = "modGogo"
''''''''''''''''''''''''''''''''''''''''''''''
''    Module written by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    http://orbita.starmedia.com/gogolive/ ''
''    Release Jan, 2004                     ''
''                                          ''
''    sorry about my English!               ''
''    your comments are welcome             ''
''''''''''''''''''''''''''''''''''''''''''''''


' Api. Global Memory Flags
  Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Const ULONG_MAX = 255


' Configuration
Public Const MC_INPUTFILE = 1
Public Const MC_INPDEV_FILE = 0              ' input device is file
Public Const MC_INPDEV_STDIO = 1             '                 stdin
Public Const MC_INPDEV_USERFUNC = 2          '       defined by user
Public Const MC_OUTPUTFILE = 2
    Public Const MC_OUTDEV_FILE = 0              ' output device is file
    Public Const MC_OUTDEV_STDOUT = 1            '                  stdout
    Public Const MC_OUTDEV_USERFUNC = 2          '        defined by user
    Public Const MC_OUTDEV_USERFUNC_WITHVBRTAG = 3       '       defined by user
Public Const MC_ENCODEMODE = 3
    Public Const MC_MODE_MONO = 0                ' mono
    Public Const MC_MODE_STEREO = 1              ' stereo
    Public Const MC_MODE_JOINT = 2               ' joint-stereo
    Public Const MC_MODE_MSSTEREO = 3            ' mid/side stereo
    Public Const MC_MODE_DUALCHANNEL = 4         ' dual channel
Public Const MC_BITRATE = 4
Public Const MC_INPFREQ = 5
Public Const MC_OUTFREQ = 6
Public Const MC_STARTOFFSET = 7
Public Const MC_USEPSY = 8
Public Const MC_USELPF16 = 9
Public Const MC_USEMMX = 10                      ' MMX
Public Const MC_USE3DNOW = 11                    ' 3DNow!
Public Const MC_USEKNI = 12                      ' SSE=KNI
Public Const MC_USEE3DNOW = 13                   ' Enhanced 3D Now!
Public Const MC_USESPC1 = 14                     ' special switch for debug
Public Const MC_USESPC2 = 15                     ' special switch for debug
Public Const MC_ADDTAG = 16
Public Const MC_EMPHASIS = 17
Public Const MC_EMP_NONE = 0                 ' no empahsis
Public Const MC_EMP_5015MS = 1               ' 50/15ms
Public Const MC_EMP_CCITT = 3                ' CCITT
Public Const MC_VBR = 18
Public Const MC_CPU = 19
Public Const MC_BYTE_SWAP = 20
Public Const MC_8BIT_PCM = 21
Public Const MC_MONO_PCM = 22
Public Const MC_TOWNS_SND = 23
Public Const MC_THREAD_PRIORITY = 24
Public Const MC_READTHREAD_PRIORITY = 25
Public Const MC_OUTPUT_FORMAT = 26
Public Const MC_OUTPUT_NORMAL = 0            ' mp3+TAG=see MC_ADDTAG
Public Const MC_OUTPUT_RIFF_WAVE = 1         ' RIFF/WAVE
Public Const MC_OUTPUT_RIFF_RMP = 2          ' RIFF/RMP
Public Const MC_RIFF_INFO = 27
Public Const MC_VERIFY = 28
Public Const MC_OUTPUTDIR = 29
Public Const MC_VBRBITRATE = 30
Public Const MC_ENHANCEDFILTER = 31
Public Const MC_MSTHRESHOLD = 32

'Language
Public Const MC_LANG = 33
Public Const MC_MAXFILELENGTH = 34
Public Const MC_MAXFLEN_IGNORE = ULONG_MAX
Public Const MC_MAXFLEN_WAVEHEADER = ULONG_MAX - 1
Public Const MC_OUTSTREAM_BUFFERD = 35
Public Const MC_OBUFFER_ENABLE = 1
Public Const MC_OBUFFER_DISABLE = 0

'Errors
Public Const ME_NOERR = 0                        ' return normally
Public Const ME_EMPTYSTREAM = 1                  ' stream becomes empty
Public Const ME_HALTED = 2                       ' stopped by user
Public Const ME_MOREDATA = 3
Public Const ME_INTERNALERROR = 10               ' internal error
Public Const ME_PARAMERROR = 11                  ' parameters error
Public Const ME_NOFPU = 12                       ' no FPU
Public Const ME_INFILE_NOFOUND = 13              ' can't open input file
Public Const ME_OUTFILE_NOFOUND = 14             ' can't open output file
Public Const ME_FREQERROR = 15                   ' frequency is not good
Public Const ME_BITRATEERROR = 16                ' bitrate is not good
Public Const ME_WAVETYPE_ERR = 17                ' WAV format is not good
Public Const ME_CANNOT_SEEK = 18                 ' can't seek
Public Const ME_BITRATE_ERR = 19                 ' only for compatibility
Public Const ME_BADMODEORLAYER = 20              ' mode/layer not good
Public Const ME_NOMEMORY = 21                    ' fail to allocate memory
Public Const ME_CANNOT_SET_SCOPE = 22            ' thread error
Public Const ME_CANNOT_CREATE_THREAD = 23        ' fail to create thear
Public Const ME_WRITEERROR = 24                  ' lock of capacity of disk


' getting configuration
Public Const MG_INPUTFILE = 1                    ' name of input file
Public Const MG_OUTPUTFILE = 2                   ' name of output file
Public Const MG_ENCODEMODE = 3                   ' type of encoding
Public Const MG_BITRATE = 4                      ' bitrate
Public Const MG_INPFREQ = 5                      ' input frequency
Public Const MG_OUTFREQ = 6                      ' output frequency
Public Const MG_STARTOFFSET = 7                  ' offset of input PCM
Public Const MG_USEPSY = 8                       ' psycho-acoustics
Public Const MG_USEMMX = 9                       ' MMX
Public Const MG_USE3DNOW = 10                    ' 3DNow!
Public Const MG_USEKNI = 11                      ' SSE=KNI
Public Const MG_USEE3DNOW = 12                   ' Enhanced 3DNow!

Public Const MG_USESPC1 = 13                     ' special switch for debug
Public Const MG_USESPC2 = 14                     ' special switch for debug
Public Const MG_COUNT_FRAME = 15                 ' amount of frame
Public Const MG_NUM_OF_SAMPLES = 16              ' number of sample for 1 frame
Public Const MG_MPEG_VERSION = 17                ' MPEG VERSION
Public Const MG_READTHREAD_PRIORITY = 18         ' thread priority to read for BeOS




Enum t_lang
    tLANG_UNKNOWN
    tLANG_JAPANESE_SJIS
    tLANG_JAPANESE_EUC
    tLANG_ENGLISH
    tLANG_GERMAN
    tLANG_SPANISH
End Enum

Type MCP_INPDEV_USERFUNC
    pUserFunc As Long   ' pointer to user-function for call-back or MPGE_NULL_FUNC if none
    nSize As Long       ' size of file or MC_INPDEV_MEMORY_NOSIZE if unknown
    nBit As Long        ' nBit = 8 or 16
    nFreq As Long       'input frequency
    nChn As Long        'number of channel(1 or 2)
End Type


  Declare Function MPGE_closeCoderVB Lib "gogo.dll" () As Long
  Declare Function MPGE_detectConfigureVB Lib "gogo.dll" () As Long
  Declare Function MPGE_endCoderVB Lib "gogo.dll" () As Long
  Declare Function MPGE_getConfigureVB Lib "gogo.dll" (ByVal Mode As Long, para1 As Any) As Long
  Declare Function MPGE_getConfigureVB2 Lib "gogo.dll" Alias "MPGE_getConfigureVB" (ByVal Mode As Long, ByRef para1 As String) As Long
  Declare Function MPGE_getUnitStatesVB Lib "gogo.dll" (unit As Long) As Long
  Declare Function MPGE_getVersionVB Lib "gogo.dll" (pNum As Long, pStr As String) As Long
  Declare Function MPGE_initializeWorkVB Lib "gogo.dll" () As Long
  Declare Function MPGE_processFrameVB Lib "gogo.dll" () As Long
  Declare Function MPGE_setConfigureVB Lib "gogo.dll" (ByVal Mode As Long, ByVal dwPara1 As Long, ByVal dwPara2 As String) As Long
  Declare Function MPGE_setConfigureVB2 Lib "gogo.dll" Alias "MPGE_setConfigureVB" (ByVal Mode As Long, ByVal dwPara1 As Long, dwPara2 As MCP_INPDEV_USERFUNC) As Long
  Declare Function MPGE_setConfigureVB3 Lib "gogo.dll" Alias "MPGE_setConfigureVB" (ByVal Mode As Long, ByVal dwPara1 As Long, dwPara2 As Long) As Long
  Declare Function MPGE_processTrack Lib "gogo.dll" (ByRef frameNum As Integer) As Long
  Declare Function MPGE_processTrack2 Lib "gogo.dll" Alias "MPGE_processTrack" (ByVal frameNum As Long) As Long
 

Public Const ENCODING = 1
Public Const OPENED = 2
Public Const CLOSED = 3

Function GetError(ByVal n As Long) As String

    Select Case n
        Case ME_NOERR
            GetError = " return normally"
        Case ME_EMPTYSTREAM
            GetError = " stream becomes empty"
        Case ME_HALTED
            GetError = " stopped by user"
        Case ME_MOREDATA
        Case ME_INTERNALERROR
            GetError = " internal error"
        Case ME_PARAMERROR
            GetError = " parameters error"
        Case ME_NOFPU
            GetError = " no FPU"
        Case ME_INFILE_NOFOUND
            GetError = " open input file"
        Case ME_OUTFILE_NOFOUND
            GetError = " open output file"
        Case ME_FREQERROR
            GetError = " frequency is not good"
        Case ME_BITRATEERROR
            GetError = " bitrate is not good"
        Case ME_WAVETYPE_ERR
            GetError = " WAV format is not good"
        Case ME_CANNOT_SEEK
            GetError = "  seek"
        Case ME_BITRATE_ERR
            GetError = " only for compatibility"
        Case ME_BADMODEORLAYER
            GetError = " mode/layer not good"
        Case ME_NOMEMORY
            GetError = " fail to allocate memory"
        Case ME_CANNOT_SET_SCOPE
            GetError = " thread error"
        Case ME_CANNOT_CREATE_THREAD
            GetError = " fail to create thear"
        Case ME_WRITEERROR
            GetError = " lock of capacity of disk"
    End Select
    
    GetError = GetError & "(" & n & ")"
    
End Function


'esta rutina la llama gogo.dll cuando tiene los datos de salida listos
Function InputCallbackFunction(ByVal hBuf As Long, ByVal Largo As Long) As Long
            
    InputCallbackFunction = frmGogoLive.callBackMp3(hBuf, Largo)
    
End Function


'esta rutina la llama gogo.dll cuando tiene los datos de salida listos
Function OutputCallbackFunction(ByVal hBuf As Long, ByVal Length As Long) As Long
    
    'si el modo de salida es shoutcast (transmision sobre TCP/IP)
    If (Mode And MODE_OUTPUT_SHOUTCAST) = MODE_OUTPUT_SHOUTCAST Then
        'verificar que existe una conexion
        If frmShoutCast.ListenConnected Then
            'enviar datos al Receptor
            frmShoutCast.SenDataToListen hBuf, Length
        End If
    ElseIf (Mode And MODE_OUTPUT_WAVE) = MODE_OUTPUT_WAVE Then
        'Pronto
    End If
    
    OutputCallbackFunction = ME_NOERR
    
End Function
