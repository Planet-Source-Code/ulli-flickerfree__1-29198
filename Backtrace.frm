VERSION 5.00
Begin VB.Form Bounce 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Screen Vert Freq Sync Exanple"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   722
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox ckSync 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00FFFFFF&
      Caption         =   "Syncronize Movement  to Vertical Scan Frequency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   2475
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3165
      TabIndex        =   0
      Top             =   195
      Width           =   840
   End
   Begin VB.Image img 
      Height          =   1425
      Left            =   1500
      Picture         =   "Backtrace.frx":0000
      Top             =   2655
      Width           =   1320
   End
End
Attribute VB_Name = "Bounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'example how to do screen painting while the screen does its vertical retrace

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Done As Boolean
Private HorizPos
Private VertPos
Private BorderRight
Private BorderBottom
Private HorizIncr
Private VertIncr
Private Const BorderTop As Long = 5
Private MachineCode(0 To 15) As Byte

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    HorizIncr = 2
    VertIncr = 1
    HorizPos = 200
    VertPos = 200
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'set up m/c code
    'need not push AX or DX - they are always free; they usually have the return code
    MachineCode(0) = &HBA  'mov dx,3da        ;video port
    MachineCode(1) = &HDA
    MachineCode(2) = &H3

WaitForScan:
    MachineCode(3) = &HEC  'in al,dx          ;read port

    MachineCode(4) = &H24  'and al, 8         ;check retrace bit
    MachineCode(5) = &H8

    MachineCode(6) = &H75  'jnz WaitForScan   ;is already in retrace - wait for vertical scan
    MachineCode(7) = &HFB  '                  ;exiting now might not yield the full retrace time

WaitForRetrace:
    MachineCode(8) = &HEC  'in al,dx          ;read port

    MachineCode(9) = &H24  'and al, 8         ;check retrace bit
    MachineCode(10) = &H8

    MachineCode(11) = &H74 'jz WaitForRetrace ;is in vertical scan - wait for retrace
    MachineCode(12) = &HFB
    
    MachineCode(13) = &HC2 'ret 16            ;vert retrace has just begun - return
    MachineCode(14) = &H10
    MachineCode(15) = &H0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With img
        .Visible = True
        Show
        Do
            HorizPos = HorizPos + HorizIncr
            If HorizPos <= 0 Or HorizPos >= BorderRight Then
                HorizIncr = -HorizIncr
            End If
            VertPos = VertPos + VertIncr
            If VertPos <= BorderTop Or VertPos >= BorderBottom Then
                VertIncr = -VertIncr
            End If
            If ckSync = vbChecked Then
                'wait for full sized vert retrace - exec m/c code
                CallWindowProc VarPtr(MachineCode(0)), 0&, 0&, 0&, 0&
              Else 'NOT CKSYNC...
                'simulate vert scan freq about 95 Hz on my video - plain wait
                Sleep 10
            End If
            .Move HorizPos, VertPos 'move to new posn
            DoEvents 'stay reactive
        Loop Until Done
    End With 'Img
    
    Unload Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Done = True

End Sub

Private Sub Form_Resize()
 
    BorderRight = ScaleWidth - img.Width
    BorderBottom = ScaleHeight - img.Height
    HorizPos = 0
    VertPos = BorderTop
    VertIncr = Abs(VertIncr)
    HorizIncr = Abs(HorizIncr)

End Sub

':) Ulli's VB Code Formatter V2.5.12 (25.11.2001 10:52:57) 17 + 98 = 115 Lines
