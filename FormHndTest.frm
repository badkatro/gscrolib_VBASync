VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormHndTest 
   OleObjectBlob   =   "FormHndTest.frx":0000
   Caption         =   "FormHndTest"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   18
End
Attribute VB_Name = "FormHndTest"
Attribute VB_Base = "0{803758AA-1EE3-43CE-8296-F4EE373AF2E7}{3C17FF6B-88E3-4F2E-BD8A-AB908D4AA71E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim rr As Long
Public NewCountDownTime As Integer
Public pCounter As Integer
Public nthAhMsg As Integer      ' Position counter: no 1 is positioned at x, y, no 2 at x, 2y, etc (stacked)

Sub Label1_Click()

Unload Me

End Sub
Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

If NewCountDownTime = 0 Then
    NewCountDownTime = 2
ElseIf NewCountDownTime = -1 Then
    GoTo unloadDirectly
End If

'Label1.ControlTipText = "Exiting in " & NewCountDownTime & " seconds!"
If pCounter <= 1 Then
    'Debug.Print "pCounter = " & pCounter & vbCr
    pCounter = pCounter + 1
    With Me.Label2
        
        If Right(.Caption, 2) <> "s." Then
            .Caption = Split(Me.Label2.Caption, ":")(0) & "  " & NewCountDownTime & " s."
        End If
        
        .Visible = True
        
        
        .ZOrder (fmtop)
        .Top = Me.Height - 15
    End With
    
    Me.Repaint
    Call gscrolib.PauseForSeconds(CInt(NewCountDownTime))

unloadDirectly:
    Unload Me
    
Else
    'Debug.Print "pCounter = " & pCounter & vbCr
End If

End Sub


Sub UserForm_Activate()

Dim fSty As Long
Dim lp As POINTAPI

pCounter = 0
'rr = GetActiveWindow
rr = FindWindow("ThunderDFrame", "FormHndTest")
'Debug.Print "Form had handle " & rr

fSty = GetWindowLong(rr, GWL_STYLE)

SetWindowLong rr, GWL_STYLE, CLng(fSty And (Not &HC00000))

' Position userform at mouse cursor
'If GetCursorPos(lp) <> 0 Then
'    Me.Left = PixelsToPoints(lp.x, False)
'    Me.top = PixelsToPoints(lp.y, True)
'End If

Dim nh As Integer
nh = Me.Height

Dim ntop As Single, nleft As Single

nleft = 40
If nthAhMsg > 0 Then      ' Stackem on top of others when calling multiple times
    
    ' Position userform left edge
    If nthAhMsg > 0 Then
        ntop = ((nthAhMsg - 1) * nh) + 95
    Else
        ntop = 95
    End If
Else
    ntop = 95
End If

Me.Left = nleft
Me.Top = ntop


Me.Label2.ZOrder (fmtop)
Me.Label2.ForeColor = wdColorDarkBlue

'If CInt(Me.Width) <> 220 Then
    'Me.Label1.Width = Me.Width - 17
'End If
'
'If Me.Height <> 120 Then
    'Me.Label1.Height = Me.Height - 50
'End If

Application.ScreenRefresh

End Sub
Sub UserForm_Click()

Unload Me

End Sub


Private Sub UserForm_Resize()

'Me.Label1.Width = Me.Width - 16
'Me.Label1.Height = Me.Height - 28

End Sub
