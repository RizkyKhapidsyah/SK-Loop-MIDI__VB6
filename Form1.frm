VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "MIDI with Loop"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Resume"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Loop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loop Off"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2520
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


Dim mssg As String * 255


   Public Function GetShortName(ByVal sLongFileName As String) As String
       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
   End Function
Private Sub Check1_Click()


 If Check1.Value = 1 Then
   Timer1.Enabled = True
 ElseIf Check1.Value = 0 Then
   Timer1.Enabled = False
   Label1.Caption = "Loop Off"
 End If
 
 
End Sub

Private Sub Command1_Click()
  Dim MInfo As String
  Dim ShortName As String
Screen.MousePointer = 11

CommonDialog1.CancelError = True
On Error GoTo EH1

CommonDialog1.Filter = "Sequence (*.mid)|*.mid"
CommonDialog1.Flags = &H80000 Or &H1000
CommonDialog1.ShowOpen

'#####################################################
'IMPORTANT!!!!!!!! all MCI Commands cannot see LONG FileNames!
'Therefore we must convert it to Short Name Format (This applies to MIDI, AVI and WAVs too)
  ShortName = GetShortName(CommonDialog1.filename)
'#####################################################

i = mciSendString("close mid1", 0&, 0, 0)
i = mciSendString("open " & ShortName & " type sequencer alias mid1", 0&, 0, 0)
i = mciSendString("status mid1 length", mssg, 255, 0)
  MInfo = "Length = " & Str(mssg) & " milliseconds" & vbCrLf
  WAVlength = Str(mssg)
i = mciSendString("status mid1 mode", mssg, 255, 0)
  MInfo = MInfo & "Mode = " & CStr(mssg) & vbCrLf
  
Text1.Text = MInfo
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Check1.Enabled = True
Screen.MousePointer = 0

Exit Sub

EH1:

Screen.MousePointer = 0
If Err = 32755 Then Err.Clear: Exit Sub
MsgBox Err.Description, vbExclamation, "ERR #" & Err
End Sub

Private Sub Command2_Click()
 i = mciSendString("play mid1 from 0", 0&, 0, 0)
End Sub


Private Sub Command3_Click()
 i = mciSendString("pause mid1", 0&, 0, 0)
End Sub


Private Sub Command4_Click()
 i = mciSendString("stop mid1", 0&, 0, 0)
End Sub



Private Sub Command5_Click()
 i = mciSendString("resume mid1", 0&, 0, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 i = mciSendString("close mid1", 0&, 0, 0)


End Sub


Private Sub Timer1_Timer()
Dim WAVpos As String

i = mciSendString("status mid1 position", mssg, 255, 0)
WAVpos = Str(mssg)

If WAVpos = WAVlength Then
 i = mciSendString("play mid1 from 0", 0&, 0, 0)
End If

Label1.Caption = WAVpos
End Sub



