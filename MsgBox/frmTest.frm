VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attention Box Demo"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      Caption         =   "With scrollbars"
      Height          =   450
      Index           =   4
      Left            =   2970
      TabIndex        =   5
      Top             =   390
      Width           =   1650
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Exit"
      Height          =   450
      Index           =   3
      Left            =   2970
      TabIndex        =   4
      Top             =   870
      Width           =   1650
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "YesNo Style"
      Height          =   450
      Index           =   2
      Left            =   1260
      TabIndex        =   3
      Top             =   1350
      Width           =   1650
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "OkCancel Style"
      Height          =   450
      Index           =   1
      Left            =   1260
      TabIndex        =   2
      Top             =   870
      Width           =   1650
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "OkOnly Style"
      Height          =   450
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   390
      Width           =   1650
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00696969&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00696969&
      Height          =   1440
      Left            =   60
      Top             =   375
      Width           =   1110
   End
   Begin VB.Label lblSwitch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Switchboard to test Attention Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   2865
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTest_Click(index As Integer)
    Select Case index
        Case 0
            AttBox "This is the OkOnly Style" & vbCrLf + vbCrLf & "Hope you like it. Please don't forget to vote.", 0
        Case 1
            AttBox "This is the OkCancel Style" & vbCrLf + vbCrLf & "Hope you like it. Please don't forget to vote.", 1
        Case 2
            AttBox "This is the YesNo Style" & vbCrLf + vbCrLf & "Hope you like it. Please don't forget to vote.", 2
        Case 3
            If AttBox("Are you sure you want to close Attention Box Demo??", attOkCancel) = attOK Then
                Unload Me
            Else
                Exit Sub
            End If
        Case 4
            AttBox "This demonstrates the attention box's ability to hold an unlimited length " & _
                        "of string by accomodating it through the use of vertical scrollbars. " & _
                         "This is an automatic feature and the user need not problem about it.  This " & _
                         "feature makes it so that the attention box has uniform dimensions making it more " & _
                         "appealing than a normal message box." & vbCrLf + vbCrLf & "Hope you like it. Please don't forget to vote.", 0
    End Select
End Sub

Private Sub Form_Load()
    LoadAttImages
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadAttImages
End Sub
