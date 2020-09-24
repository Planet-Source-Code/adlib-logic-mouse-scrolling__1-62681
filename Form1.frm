VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   720
      ScaleHeight     =   1515
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   360
      Width           =   2655
      Begin VB.Image Image1 
         Height          =   4605
         Left            =   0
         Picture         =   "Form1.frx":0000
         Top             =   0
         Width           =   2685
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   600
      TabIndex        =   0
      Top             =   3720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   50
   End
   Begin VB.Label Label3 
      Caption         =   "This grid also has scrolling abilities.  The TopRow property is changed each time the scrolling event occurs."
      Height          =   1095
      Left            =   3480
      TabIndex        =   4
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Or click this MSFLEXGIRD"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "This picture box has scrolling abilities.  (In this example an Image control is moved each time the scrolling event occurs.)"
      Height          =   975
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Enable Scrolling for the picturebox and the grid
Private Sub Form_Load()
AddScrollness MSFlexGrid1.hwnd
AddScrollness Picture1.hwnd
End Sub
'End the scrolling ability for the picturebox and grid
Private Sub Form_Unload(Cancel As Integer)
RemoveScrollness MSFlexGrid1.hwnd
RemoveScrollness Picture1.hwnd
End Sub

Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Get the scrolling value
Dim Scroll As Integer
Scroll = GetScrollMovement(MSFlexGrid1.hwnd)
    
    If Scroll = 1 Then '1 = scroll up
        If MSFlexGrid1.TopRow <> 1 Then _
        MSFlexGrid1.TopRow = MSFlexGrid1.TopRow - 1
        
    ElseIf Scroll = -1 Then '-1 = scroll down
        MSFlexGrid1.TopRow = MSFlexGrid1.TopRow + 1
        
    Else '0 = the event was fired for a proper MouseMove
    
    '<<DEAL WITH THE OTHER MOUSEMOVE stuff here>>
    End If
    

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Get the scrolling value
Dim Scroll As Integer
Scroll = GetScrollMovement(Picture1.hwnd)

    If Scroll = 1 Then 'Scroll up
        Image1.Top = Image1.Top + 500
    ElseIf Scroll = -1 Then 'Scroll down
        Image1.Top = Image1.Top - 500
    Else 'The event was fired for a proper MouseMove
    
    '<<DEAL WITH THE OTHER MOUSEMOVE stuff here>>
    End If
End Sub

Private Sub Image1_Click()

'This is neccessary in this example, because the picturebox deals with
'the scrolling of its contents: the image1
Picture1.SetFocus
End Sub

