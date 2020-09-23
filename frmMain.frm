VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   735
   ClientLeft      =   2400
   ClientTop       =   1185
   ClientWidth     =   6075
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   0
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   2
      Top             =   780
      Width           =   6075
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00A99781&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E7E4E0&
      Height          =   750
      Left            =   0
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E7E4E0&
         Height          =   255
         Left            =   5880
         TabIndex        =   3
         Top             =   -60
         Width           =   195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E7E4E0&
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   4200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type CounterInfo
    hCounter As Long
    strName As String
End Type

Dim hQuery As Long
Dim Counters(0 To 99) As CounterInfo
Dim currentCounterIdx As Long
Dim iPerformanceDetail As PERF_DETAIL
Dim newPos, oldPos

'drag window
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2




Public Sub AddCounter(strCounterName As String, hQuery As Long)
    Dim pdhStatus As PDH_STATUS
    Dim hCounter As Long
    
    pdhStatus = PdhVbAddCounter(hQuery, strCounterName, hCounter)
    Counters(currentCounterIdx).hCounter = hCounter
    Counters(currentCounterIdx).strName = strCounterName
    currentCounterIdx = currentCounterIdx + 1
End Sub

Private Sub UpdateValues()
    Dim dblCounterValue As Double
    Dim pdhStatus As Long
    Dim strInfo As String
    Dim i As Long
        
    PdhCollectQueryData (hQuery)
    
    i = 0  'Only one counter but you can add more

    dblCounterValue = _
            PdhVbGetDoubleCounterValue(Counters(i).hCounter, pdhStatus)
        
        'Some error checking, make sure the query went through
        If (pdhStatus = PDH_CSTATUS_VALID_DATA) _
        Or (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
        strInfo = "CPU Usage: %" & Format$(dblCounterValue, "0")
        If newPos = -1 Then
            newPos = 0
        Else
            newPos = dblCounterValue
        End If
        If oldPos = -1 Then oldPos = 0
        pic2.Picture = pic1.Image
        pic1.Cls
        pic1.PaintPicture pic2.Image, -1, 0
        'MsgBox Int(50 - (oldPos / 2))
        pic1.Line (pic1.Width - 2, Int(50 - (oldPos / 2)))-(pic1.Width - 1, Int(50 - (newPos / 2)))
        oldPos = newPos
        'Me.Caption = Format$(dblCounterValue, "0") & "% - CPU Status"
        End If
        
    Label1 = strInfo
End Sub

Private Sub Form_Load()
    oldPos = -1
    newPos = -1
    
    Dim pdhStatus As PDH_STATUS
    
    pdhStatus = PdhOpenQuery(0, 1, hQuery)
    If pdhStatus <> ERROR_SUCCESS Then
        MsgBox "OpenQuery failed"
        End
    End If
    ' Add the processor time query
    AddCounter "\Processor(0)\% Processor Time", hQuery
    UpdateValues    ' Force an immediate display of the counter values
    Timer1.Enabled = True
    xwidth = Me.ScaleWidth
    xheight = Me.ScaleHeight
    SetWindowPos Me.hwnd, -1, 0, 0, xwidth, xheight, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    PdhCloseQuery (hQuery)
End Sub


Private Sub pb1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.BorderStyle <> 0 Then
Let Me.BorderStyle = 0
Else
Let Me.BorderStyle = 2

End If
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Timer1_Timer()
    ' fires once per second, can be changed.
    UpdateValues
End Sub
Private Sub form_resize()
'Let pb1.Width = Me.Width
'Let pb1.Height = Me.Height
End Sub
