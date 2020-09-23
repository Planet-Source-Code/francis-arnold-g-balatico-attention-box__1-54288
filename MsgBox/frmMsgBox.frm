VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3750
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMsgBox.frx":0000
   ScaleHeight     =   2550
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNoScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H00696969&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   405
      Width           =   3495
   End
   Begin VB.PictureBox picB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1890
      ScaleHeight     =   255
      ScaleWidth      =   870
      TabIndex        =   2
      Top             =   2205
      Width           =   870
   End
   Begin VB.PictureBox picB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2805
      ScaleHeight     =   255
      ScaleWidth      =   870
      TabIndex        =   1
      Top             =   2205
      Width           =   870
   End
   Begin VB.TextBox txtScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H00696969&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   405
      Width           =   3495
   End
   Begin VB.Label lblDrag 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3765
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'variables used to ensure a flickerless image swapping
Private curIndex As Byte
Private onButton As Boolean

'API declarations for dragging form
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

'procedure to drag a no-titlebar form
Private Sub FormDrag(frmName As Form)
    ReleaseCapture
    Call SendMessage(frmName.hWnd, &HA1, 2, 0&)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            'checks the layout used, and returns the OK/Yes value
            If attBoxLayout = 0 Then
                    attContext = attOK
                    Unload Me
             
            ElseIf attBoxLayout = 1 Then
                    attContext = attOK
                    Unload Me
                
                
            ElseIf attBoxLayout = 2 Then
                    attContext = attYes
                    Unload Me
                End If
            
            
        Case vbKeyEscape
           
            'checks the layout used, and returns the Cancel/No value
            If attBoxLayout = 0 Then
                    attContext = attOK
                    Unload Me
             
            ElseIf attBoxLayout = 1 Then
                    attContext = attCancel
                    Unload Me
                           
            ElseIf attBoxLayout = 2 Then
                    attContext = attNo
                    Unload Me
            End If
        
    End Select
End Sub

Private Sub Form_Load()
    Beep
End Sub

'swaps the buttons to unhighlighted state
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Swap2Orig
End Sub

'uses the procedure to enable form movement
Private Sub lblDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

'swaps the buttons to unhighlighted state
Private Sub lblDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Swap2Orig
End Sub
'swaps the button to the pressed state
Private Sub picB_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If attBoxLayout = 0 Then
        Set picB(0).Picture = picOk3
    ElseIf attBoxLayout = 1 Then
        If index = 0 Then
            Set picB(0).Picture = picCancel3
        Else
            Set picB(1).Picture = picOk3
        End If
    ElseIf attBoxLayout = 2 Then
        If index = 0 Then
            Set picB(0).Picture = picNo3
        Else
            Set picB(1).Picture = picYes3
        End If
    End If
End Sub
'swaps button to the highlighted state
Private Sub picB_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not curIndex = index Then
        onButton = False
    End If
    
    If onButton = True Then Exit Sub
    
    curIndex = index
    onButton = True
    
    If attBoxLayout = 0 Then
        Set picB(0).Picture = picOk2
    ElseIf attBoxLayout = 1 Then
        If index = 0 Then
            Set picB(1).Picture = picOk
            Set picB(0).Picture = picCancel2
        Else
            Set picB(1).Picture = picOk2
            Set picB(0).Picture = picCancel
        End If
    ElseIf attBoxLayout = 2 Then
        If index = 0 Then
            Set picB(1).Picture = picYes
            Set picB(0).Picture = picNo2
        Else
            Set picB(1).Picture = picYes2
        Set picB(0).Picture = picNo
        End If
    End If
End Sub
'executes when the user releases a pressed button
Private Sub picB_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ensures that the cursor is still over the button otherwise cancel
    If X < 0 Or X > picB(index).Width Then
        Call Swap2Hilyt(index)
        Exit Sub
    End If
        
    If Y < 0 Or Y > picB(index).Height Then
        Call Swap2Hilyt(index)
        Exit Sub
    End If
    
    'checks the layout used, the value pressed and returns it
    If attBoxLayout = 0 Then
        Set picB(0).Picture = picOk2
            attContext = attOK
            Unload Me
     
    ElseIf attBoxLayout = 1 Then
        If index = 0 Then
            Set picB(0).Picture = picCancel2
                attContext = attCancel
                Unload Me
        Else
            Set picB(1).Picture = picOk2
                attContext = attOK
                Unload Me
        End If
        
    ElseIf attBoxLayout = 2 Then
        If index = 0 Then
            Set picB(0).Picture = picNo2
                attContext = attNo
                Unload Me
        Else
            Set picB(1).Picture = picYes2
                attContext = attYes
                Unload Me
        End If
        
    End If
End Sub
'prevents textbox to receive focus
Private Sub txtNoScroll_GotFocus()
    picB(0).SetFocus
End Sub
'swaps buttons to unhighlighted state
Private Sub txtNoScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Swap2Orig
End Sub
'prevents textbox to receive focus
Private Sub txtScroll_GotFocus()
    picB(0).SetFocus
End Sub
'swaps buttons to unhighlighted state
Private Sub txtScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Swap2Orig
End Sub
'procedure used to swap the buttons to the highlighted unpressed state
Private Sub Swap2Hilyt(index As Integer)
    If attBoxLayout = 0 Then
        Set picB(0).Picture = picOk2
    ElseIf attBoxLayout = 1 Then
        If index = 0 Then
            Set picB(0).Picture = picCancel2
        Else
            Set picB(1).Picture = picOk2
        End If
    ElseIf attBoxLayout = 2 Then
        If index = 0 Then
            Set picB(0).Picture = picNo2
        Else
            Set picB(1).Picture = picYes2
        End If
    End If
End Sub
'procedure used to swap the buttons to the unhighlighted unpressed state
Private Sub Swap2Orig()
    If onButton = False Then Exit Sub
    
    onButton = False
    curIndex = 2
    
    If attBoxLayout = 0 Then
        Set picB(0).Picture = picOk
    ElseIf attBoxLayout = 1 Then
        
            Set picB(1).Picture = picOk
            Set picB(0).Picture = picCancel
        
    ElseIf attBoxLayout = 2 Then
        
            Set picB(1).Picture = picYes
            Set picB(0).Picture = picNo
        
    End If
    
End Sub
