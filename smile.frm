VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Picture         =   "smile.frx":0000
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)



Private Function fMakeATranspArea(AreaType As String, pCordinate() As Long) As Boolean


Const RGN_DIFF = 4

Dim lOriginalForm As Long
Dim ltheHole As Long
Dim lNewForm As Long
Dim lFwidth As Single
Dim lFHeight As Single
Dim lborder_width As Single
Dim ltitle_height As Single


 On Error GoTo Trap
    lFwidth = ScaleX(Width, vbTwips, vbPixels)
    lFHeight = ScaleY(Height, vbTwips, vbPixels)
    lOriginalForm = CreateRectRgn(0, 0, lFwidth, lFHeight)
    
    lborder_width = (lFHeight - ScaleWidth) / 2
    ltitle_height = lFHeight - lborder_width - ScaleHeight

Select Case AreaType
  
    Case "Elliptic"
 
            ltheHole = CreateEllipticRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))

    Case "RectAngle"
   
            ltheHole = CreateRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))

      
    Case "RoundRect"
             
               ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(5), pCordinate(6))

    Case "Circle"
               ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(3), pCordinate(4))
    
    Case Else
           MsgBox "Unknown Shape!"
           Exit Function

       End Select

    lNewForm = CreateRectRgn(0, 0, 0, 0)
    CombineRgn lNewForm, lOriginalForm, _
        ltheHole, 1
    
    SetWindowRgn hwnd, lNewForm, True
    Me.Refresh
    fMakeATranspArea = False
Exit Function

Trap:



End Function

Private Sub Form_Load()

Dim lParam(1 To 6) As Long

lParam(1) = 1
lParam(2) = 1
lParam(3) = 50
lParam(4) = 50
lParam(5) = 50
lParam(6) = 50


Call fMakeATranspArea("Circle", lParam())


rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Randomize
Form1.Top = (Rnd * Screen.Height) - 750

Form1.Left = (Rnd * Screen.Width) - 750

If Form1.Left <= 0 Then Form1.Left = 0
If Form1.Top <= 0 Then Form1.Top = 0

End Sub

Private Sub Timer1_Timer()
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
End Sub
