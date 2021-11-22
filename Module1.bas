Attribute VB_Name = "Module1"
Type WAVEFORMAT


    FormatTag As Integer
        channels As Integer
        SamplesPerSec As Long
        AvgBytesPerSec As Long
        BlockAling As Integer
        End Type


Type PCMSPECIFIC
    SampBit As Integer
    End Type


Type WFORMAT
    wftName As String * 4
    wftType As String * 4
    wftSize As Long
    wft As WAVEFORMAT
    pcm As PCMSPECIFIC
    End Type


Type RFORMAT
    Name As String * 4
    FileSize As Long
    End Type


Type WavHead
    RiffChk As RFORMAT
    Fchnk As WFORMAT
    End Type


Type DatHead
    DataString As String * 4
    DataLength As Long
    End Type
    Global voice(30000) As Integer


Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long


