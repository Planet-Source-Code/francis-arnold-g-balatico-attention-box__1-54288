Attribute VB_Name = "AttMod"
Option Explicit

'Declare picture variables to hold the images in the resource file
Public picCancel As Picture
Public picCancel2 As Picture
Public picCancel3 As Picture

Public picOk As Picture
Public picOk2 As Picture
Public picOk3 As Picture

Public picYes As Picture
Public picYes2 As Picture
Public picYes3 As Picture

Public picNo As Picture
Public picNo2 As Picture
Public picNo3 As Picture

'Declaration of variable identifiers
Public attContext As AttValue
Public attBoxLayout As AttType

'Enumerations of values

Public Enum AttValue
    attOK = 0
    attCancel = 1
    attYes = 3
    attNo = 4
End Enum

Public Enum AttType
    attOkOnly = 0
    attOkCancel = 1
    attYesNo = 2
End Enum

Public testMsg As String
Public testType As AttType

'API declaration for counting the lines of the textbox
Private Declare Function Sendmessageaslong Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const EM_GETLINECOUNT = 186


'procedure to load the images onto the picture variables
'allows for faster image swapping
Public Sub LoadAttImages()
    Set picCancel = LoadResPicture("CANCEL", vbResBitmap)
    Set picCancel2 = LoadResPicture("CANCEL2", vbResBitmap)
    Set picCancel3 = LoadResPicture("CANCEL3", vbResBitmap)
    
    Set picOk = LoadResPicture("OK", vbResBitmap)
    Set picOk2 = LoadResPicture("OK2", vbResBitmap)
    Set picOk3 = LoadResPicture("OK3", vbResBitmap)
    
    Set picYes = LoadResPicture("YES", vbResBitmap)
    Set picYes2 = LoadResPicture("YES2", vbResBitmap)
    Set picYes3 = LoadResPicture("YES3", vbResBitmap)
    
    Set picNo = LoadResPicture("NO", vbResBitmap)
    Set picNo2 = LoadResPicture("NO2", vbResBitmap)
    Set picNo3 = LoadResPicture("NO3", vbResBitmap)
End Sub
'procedure to unload the images
Public Sub UnloadAttImages()
    Set picCancel = Nothing
    Set picCancel2 = Nothing
    Set picCancel3 = Nothing
    
    Set picOk = Nothing
    Set picOk2 = Nothing
    Set picOk3 = Nothing
    
    Set picYes = Nothing
    Set picYes2 = Nothing
    Set picYes3 = Nothing
    
    Set picNo = Nothing
    Set picNo2 = Nothing
    Set picNo3 = Nothing
End Sub
'procedure to count lines in a textbox
Public Function LineCount(msgTxt As TextBox) As Long
    LineCount = Sendmessageaslong(msgTxt.hWnd, EM_GETLINECOUNT, 0, 0)
End Function

'calling procedure of attention box
Public Function AttBox(Optional ByRef attString As String, Optional ByRef attBoxType As AttType) As AttValue
    With frmMsgBox
           .txtNoScroll.Text = attString
            
           If LineCount(.txtNoScroll) > 8 Then
                .txtNoScroll.Visible = False
                .txtScroll.Text = attString
                .txtScroll.Visible = True
           Else
                .txtNoScroll.Visible = True
                .txtScroll.Visible = False
                .txtNoScroll.Text = attString
           End If
        
        Select Case attBoxType
            Case 0
                .picB(1).Visible = False
                
                Set .picB(0).Picture = picOk
                
                attBoxLayout = 0
                
                .Show vbModal
                
                AttBox = attContext
                
            Case 1
                .picB(1).Visible = True
                
                Set .picB(0).Picture = picCancel
                Set .picB(1).Picture = picOk
                
                attBoxLayout = 1
                
                .Show vbModal
                
                AttBox = attContext

            Case 2
                .picB(1).Visible = True
                
                Set .picB(0).Picture = picNo
                Set .picB(1).Picture = picYes
                
                attBoxLayout = 2
                
                .Show vbModal
                
                AttBox = attContext
        End Select
    End With
End Function
