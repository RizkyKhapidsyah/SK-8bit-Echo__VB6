VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Make WAV ECHO"
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   2745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Make Echo"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim XNBytes As String * 1
    Dim w As WavHead
    Dim d As DatHead
    Close
    MousePointer = 11
    'The "source.wav" file can be any 8-bit
    '     wav file that you want to
    'add reverb effect too
    Open "c:\source.wav" For Binary As #1
    'The "result.wav" file will be created
    Open "c:\result.wav" For Binary As #2
    Get #1, , w 'Get the "RIFF" tag that is In the beginning of all wav files
    'this makes sure we have an authentic wa
    '     v file
    Get #1, , d 'Also get the format of the wav file (e.g. 8bit mono etc...)


    If w.Fchnk.pcm.SampBit <> 8 Then
        MsgBox "VBWAVE can process only sounds sampled at 8 bits/sample"
        'If the WAV file is NOT 8-bit we must ex
        '     it now since
        'this code is only for 8-bit wav files.
        Exit Sub
    End If
    Put #2, , w 'Transfer the original header and format of the source wav file
    Put #2, , d 'into the result wav file. this info remains unchanged
    NSamples = 3000


    For i = 1 To NSamples
        Get #1, , XNBytes
        voice(i) = Asc(XNBytes) - 128
        Put #2, , XNBytes
    Next
    FLen = LOF(1)


    While Not EOF(1)
        i = i + 1
        Get #1, , XNBytes
        Sample = Asc(XNBytes) - 128
        vecho = 0.7 * Sample + 0.25 * voice((i - 1000) Mod NSamples) + 0.2 * voice((i - 2000) Mod NSamples) + 0.15 * voice((i - NSamples) Mod NSamples)
        voice(i Mod NSamples) = Sample
        vecho = vecho + 128
        If vecho > 255 Then vecho = 255
        If vecho < 0 Then vecho = 0
        XNBytes = Chr$(vecho)
        Put #2, , XNBytes
    Wend
    Close #1
    Close #2
    MousePointer = 0
End Sub

