VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   ScaleHeight     =   3915
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   2880
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "To use, move your cursor to the edges of the form and click and drag."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CPos As POINTAPI
Dim ResizeDir As Integer

Dim CurX, CurY As Long

Private Sub cmdExit_Click()

    End

End Sub


Private Sub Form_Load()

    'Snapping distance for outline
    SnapDistance = 100

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
            If x > 0 And x < 50 And y > 50 And y < (Me.Height - 50) Then 'Middle left
                tmrResize.Enabled = True
                StartResize Me.Left, Me.Top, Me.Width, Me.Height
                ResizeDir = 2
                Screen.MousePointer = 9
                Else
                If x > (Me.Width - 50) And x < Me.Width And y > 50 And y < (Me.Height - 50) Then 'Middle Right
                    tmrResize.Enabled = True
                    StartResize Me.Left, Me.Top, Me.Width, Me.Height
                    ResizeDir = 6
                    Screen.MousePointer = 9
                    Else
                    If y > 0 And y < 50 And x > 50 And x < (Me.Width - 50) Then 'Middle Top
                        tmrResize.Enabled = True
                        StartResize Me.Left, Me.Top, Me.Width, Me.Height
                        ResizeDir = 8
                        Screen.MousePointer = 7
                        Else
                        If y > Me.Height - 50 And y < Me.Height And x > 50 And x < (Me.Width - 50) Then 'Middle Bottom
                            tmrResize.Enabled = True
                            StartResize Me.Left, Me.Top, Me.Width, Me.Height
                            ResizeDir = 4
                            Screen.MousePointer = 7
                            Else
                            If x > 0 And x < 50 And y > 0 And y < 50 Then 'Top Left
                                tmrResize.Enabled = True
                                StartResize Me.Left, Me.Top, Me.Width, Me.Height
                                ResizeDir = 1
                                Screen.MousePointer = 8
                                Else
                                If x > (Me.Width - 50) And x < Me.Width And y > 0 And y < 50 Then 'Top Right
                                    tmrResize.Enabled = True
                                    StartResize Me.Left, Me.Top, Me.Width, Me.Height
                                    ResizeDir = 7
                                    Screen.MousePointer = 6
                                    Else
                                    If y > (Me.Height - 50) And y < (Me.Height) And x > 0 And x < 50 Then 'Bottom Left
                                        tmrResize.Enabled = True
                                        StartResize Me.Left, Me.Top, Me.Width, Me.Height
                                        ResizeDir = 3
                                        Screen.MousePointer = 6
                                        Else
                                        If y > (Me.Height - 50) And y < (Me.Height) And x > (Me.Width - 50) And x < Me.Width Then  'Bottom Right
                                            tmrResize.Enabled = True
                                            StartResize Me.Left, Me.Top, Me.Width, Me.Height
                                            ResizeDir = 5
                                            Screen.MousePointer = 8
                                            Else
                                            MousePointer = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Display correct resize cursor
    
        If x > 0 And x < 50 And y > 50 And y < (Me.Height - 50) Then 'Middle left
            MousePointer = 9
            Else
            If x > (Me.Width - 50) And x < Me.Width And y > 50 And y < (Me.Height - 50) Then 'Middle Right
                MousePointer = 9
                Else
                If y > 0 And y < 50 And x > 50 And x < (Me.Width - 50) Then 'Middle Top
                    MousePointer = 7
                    Else
                    If y > Me.Height - 50 And y < Me.Height And x > 50 And x < (Me.Width - 50) Then 'Middle Bottom
                        MousePointer = 7
                        Else
                        If x > 0 And x < 50 And y > 0 And y < 50 Then 'Top Left
                            MousePointer = 8
                            Else
                            If x > (Me.Width - 50) And x < Me.Width And y > 0 And y < 50 Then 'Top Right
                                MousePointer = 6
                                Else
                                If y > (Me.Height - 50) And y < (Me.Height) And x > 0 And x < 50 Then 'Bottom Left
                                    MousePointer = 6
                                    Else
                                    If y > (Me.Height - 50) And y < (Me.Height) And x > (Me.Width - 50) And x < Me.Width Then  'Bottom Right
                                        MousePointer = 8
                                        Else
                                        MousePointer = 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    MousePointer = 0
    Screen.MousePointer = 0
    
    EndResize
    tmrResize.Enabled = False
    
    If EndLeft <> "" Then Me.Left = EndLeft
    If EndTop <> "" Then Me.Top = EndTop
    If EndWidth <> 0 Then Me.Width = EndWidth
    If EndHeight <> 0 Then Me.Height = EndHeight

End Sub


Private Sub tmrResize_Timer()

    If ResizeDir <> 0 Then

        GetCursorPos CPos
        CurX = (CPos.x * 15)
        CurY = (CPos.y * 15)
        
        Select Case ResizeDir
        
            Case 1  'Top left
                If CurX < SnapDistance And CurY > SnapDistance Then  'Snap Left
                    If CurY < (Me.Top - SnapDistance) Or CurY > (Me.Top + SnapDistance) Then
                        StartResize 0, (CurY - 8), (Me.Left + Me.Width), Me.Height + Me.Top - CurY
                        Else
                        StartResize 0, Me.Top, (Me.Left + Me.Width), Me.Height
                    End If
                    Else
                    If CurY < SnapDistance And CurX > SnapDistance Then  'Snap Top
                        If CurX < (Me.Left - SnapDistance) Or CurX > (Me.Left + SnapDistance) Then
                            StartResize (CurX - 8), 0, Me.Left + Me.Width - CurX + 8, (Me.Height + Me.Top) - 8
                            Else
                            StartResize Me.Left, 0, Me.Width, (Me.Height + Me.Top) - 8
                        End If
                        Else
                        If CurX < SnapDistance And CurY < SnapDistance Then  'Snap Left + Top
                            StartResize 0, 0, (Me.Left + Me.Width), (Me.Height + Me.Top) - 8
                            Else
                            If CurX > (Me.Left - SnapDistance) And CurX < (Me.Left + SnapDistance) And CurY > (Me.Top - SnapDistance) And CurY < (Me.Top + SnapDistance) Then  'Snap to Form Left + Top
                                StartResize Me.Left, Me.Top, Me.Width, Me.Height
                                Else
                                If CurX > (Me.Left - SnapDistance) And CurX < (Me.Left + SnapDistance) Or CurY < (Me.Top - SnapDistance) And CurY > (Me.Top + SnapDistance) Then  'Snap to Form Left
                                    StartResize Me.Left, CurY - 8, Me.Width, Me.Height + Me.Top - CurY
                                    Else
                                    If CurY > (Me.Top - SnapDistance) And CurY < (Me.Top + SnapDistance) Or CurX < (Me.Left - SnapDistance) And CurX > (Me.Left + SnapDistance) Then  'Snap to Form Top
                                        StartResize CurX - 8, Me.Top, (Me.Left + Me.Width - CurX) + 8, Me.Height
                                        Else
                                        'Resize
                                        StartResize (CurX - 8), (CurY - 8), (Me.Left + Me.Width - CurX) + 8, Me.Top + Me.Height - CurY
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
        
            Case 2  'Middle Left
                If CurX < SnapDistance Then
                    StartResize 0, Me.Top, Me.Left + Me.Width, Me.Height
                    Else
                    If CurX > (Me.Left - SnapDistance) And CurX < Me.Left + SnapDistance Then  'Snap to Form
                        StartResize Me.Left, Me.Top, Me.Width, Me.Height
                        Else
                        StartResize CurX - 8, Me.Top, (Me.Left + Me.Width - CurX) + 8, Me.Height
                    End If
                End If
                
            Case 3  'Bottom Left
                If CurX < SnapDistance And CurY < (Screen.Height - SnapDistance) Then  'Snap Left
                    If CurY > ((Me.Top + Me.Height) - SnapDistance) And CurY < ((Me.Top + Me.Height) + SnapDistance) Then 'Snap Left + Form Bottom
                        StartResize 0, Me.Top, (Me.Left + Me.Width), Me.Height
                        Else
                        StartResize 0, Me.Top, (Me.Left + Me.Width), CurY - Me.Top
                    End If
                    Else
                    If CurX > SnapDistance And CurY > (Screen.Height - SnapDistance) Then  'Snap Bottom
                        If CurX > (Me.Left - SnapDistance) And CurX < (Me.Left + SnapDistance) Then  'Snap Bottom + Form left
                            StartResize Me.Left, Me.Top, Me.Width, (Screen.Height - Me.Top)
                            Else
                            StartResize (CurX - 8), Me.Top, Me.Left + Me.Width - CurX + 8, (Screen.Height - Me.Top)
                        End If
                        Else
                        If CurX < SnapDistance And CurY > (Screen.Height - SnapDistance) Then  'Snap Left + Bottom
                            StartResize 0, Me.Top, (Me.Left + Me.Width), (Screen.Height - Me.Top)
                            Else
                            If CurX > (Me.Left - SnapDistance) And CurX < (Me.Left + SnapDistance) And CurY > ((Me.Top + Me.Height) - SnapDistance) And CurY < ((Me.Top + Me.Height) + SnapDistance) Then  'Snap to Form
                                StartResize Me.Left, Me.Top, Me.Width, Me.Height
                                Else
                                If CurX > (Me.Left - SnapDistance) And CurX < (Me.Left + SnapDistance) Or CurY < ((Me.Top + Me.Height) - SnapDistance) And CurY > ((Me.Top + Me.Height) + SnapDistance) Then   'Snap to Form Left
                                    StartResize Me.Left, Me.Top, Me.Width, (CurY + 15) - Me.Top
                                    Else
                                    If CurY > ((Me.Top + Me.Height) - SnapDistance) And CurY < ((Me.Top + Me.Height) + SnapDistance) Or CurX < (Me.Left - SnapDistance) And CurX > (Me.Left + SnapDistance) Then    'Snap to Form Bottom
                                        StartResize (CurX - 15), Me.Top, Me.Left + Me.Width - CurX + 15, Me.Height
                                        Else
                                        'Resize
                                        StartResize (CurX - 15), Me.Top, Me.Left + Me.Width - CurX + 15, (CurY + 15) - Me.Top
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            
            Case 4  'Middle Bottom
                If CurY > (Screen.Height - SnapDistance) Then
                    StartResize Me.Left, Me.Top, Me.Width, (Screen.Height - Me.Top)
                    Else
                    If CurY > ((Me.Top + Me.Height) - SnapDistance) And CurY < ((Me.Top + Me.Height) + SnapDistance) Then  'Snap to form
                        StartResize Me.Left, Me.Top, Me.Width, Me.Height
                        Else
                        StartResize Me.Left, Me.Top, Me.Width, (CurY + 15) - Me.Top
                    End If
                End If
                
            Case 5  'Bottom Right
                If CurX > (Screen.Width - SnapDistance) And CurY < (Screen.Height - SnapDistance) Then  'Snap Right
                    If CurY > ((Me.Top + Me.Height) - SnapDistance) And CurY < ((Me.Top + Me.Height) + SnapDistance) Then  'Snap Right + Form Bottom
                        StartResize Me.Left, Me.Top, (Screen.Width - Me.Left), Me.Height
                        Else
                        StartResize Me.Left, Me.Top, (Screen.Width - Me.Left), (CurY + 15) - Me.Top
                    End If
                    Else
                    If CurX < (Screen.Width - SnapDistance) And CurY > (Screen.Height - SnapDistance) Then  'Snap Bottom
                        If CurX > ((Me.Left + Me.Width) - SnapDistance) And CurX < ((Me.Left + Me.Width) + SnapDistance) Then  ' Snap Bottom + Form Right
                            StartResize Me.Left, Me.Top, Me.Width, (Screen.Height - Me.Top)
                            Else
                            StartResize Me.Left, Me.Top, (CurX + 15) - Me.Left, (Screen.Height - Me.Top)
                        End If
                        Else
                        If CurX > (Screen.Width - SnapDistance) And CurY > (Screen.Height - SnapDistance) Then  'Snap Right + Bottom
                            StartResize Me.Left, Me.Top, (Screen.Width - Me.Left), (Screen.Height - Me.Top)
                            Else
                            If CurX > ((Me.Left + Me.Width) - SnapDistance) And CurX < ((Me.Left + Me.Width) + SnapDistance) And CurY > ((Me.Top + Me.Height) - SnapDistance) And CurY < ((Me.Top + Me.Height) + SnapDistance) Then 'Snap to Form
                                StartResize Me.Left, Me.Top, Me.Width, Me.Height
                                Else
                                If CurY > ((Me.Top + Me.Height) - SnapDistance) And CurY < ((Me.Top + Me.Height) + SnapDistance) Or CurX < ((Me.Left + Me.Width) - SnapDistance) And CurX > ((Me.Left + Me.Width) + SnapDistance) Then 'Snap to Form Bottom
                                    StartResize Me.Left, Me.Top, (CurX + 15) - Me.Left, Me.Height
                                    Else
                                    If CurX > ((Me.Left + Me.Width) - SnapDistance) And CurX < ((Me.Left + Me.Width) + SnapDistance) Or CurY < ((Me.Top + Me.Height) - SnapDistance) And CurY > ((Me.Top + Me.Height) + SnapDistance) Then 'Snap to Form Right
                                        StartResize Me.Left, Me.Top, Me.Width, (CurY + 15) - Me.Top
                                        Else
                                        'Resize
                                        StartResize Me.Left, Me.Top, (CurX + 15) - Me.Left, (CurY + 15) - Me.Top
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            Case 6  'Middle Right
                If CurX > (Screen.Width - SnapDistance) Then
                    StartResize Me.Left, Me.Top, ((Screen.Width - Me.Left)), Me.Height
                    Else
                    If CurX > ((Me.Left + Me.Width) - SnapDistance) And CurX < ((Me.Left + Me.Width) + SnapDistance) Then  'Snap to Form
                        StartResize Me.Left, Me.Top, Me.Width, Me.Height
                        Else
                        StartResize Me.Left, Me.Top, (CurX + 15) - Me.Left, Me.Height
                    End If
                End If
                
            Case 7  'Top Right
                If CurX < (Screen.Width - SnapDistance) And CurY < SnapDistance Then  'Snap Top
                    If CurX > ((Me.Left + Me.Width) - SnapDistance) And CurX < ((Me.Left + Me.Width) + SnapDistance) Then  'Snap Top + Form Right
                        StartResize Me.Left, 0, Me.Width, (Me.Height + Me.Top) - 8
                        Else
                        StartResize Me.Left, 0, (CurX + 15) - Me.Left, (Me.Height + Me.Top) - 8
                    End If
                    Else
                    If CurX > (Screen.Width - SnapDistance) And CurY > SnapDistance Then  'Snap Right
                        If CurY > (Me.Top - SnapDistance) And CurY < (Me.Top + SnapDistance) Then  'Snap Right + Form Top
                            StartResize Me.Left, Me.Top, (Screen.Width - Me.Left), Me.Height + 8
                            Else
                            StartResize Me.Left, (CurY - 8), (Screen.Width - Me.Left), Me.Height + Me.Top - CurY
                        End If
                        Else
                        If CurX > (Screen.Width - SnapDistance) And CurY < SnapDistance Then  'Snap Top + Right
                            StartResize Me.Left, 0, (Me.Left + Me.Width), (Me.Height + Me.Top) - 8
                            Else
                            If CurX > ((Me.Left + Me.Width) - SnapDistance) And CurX < ((Me.Left + Me.Width) + SnapDistance) And CurY > (Me.Top - SnapDistance) And CurY < (Me.Top + SnapDistance) Then   'Snap to Form
                                StartResize Me.Left, Me.Top, Me.Width, Me.Height - 8
                                Else
                                If CurX > ((Me.Left + Me.Width) - SnapDistance) And CurX < ((Me.Left + Me.Width) + SnapDistance) Or CurY < (Me.Top - SnapDistance) And CurY > (Me.Top + SnapDistance) Then  'Snap to Form Right
                                    StartResize Me.Left, (CurY - 8), Me.Width, Me.Top + Me.Height - CurY - 8
                                    Else
                                    If CurY > (Me.Top - SnapDistance) And CurY < (Me.Top + SnapDistance) Or CurX < ((Me.Left + Me.Width) - SnapDistance) And CurX > ((Me.Left + Me.Width) + SnapDistance) Then   'Snap to Top
                                        StartResize Me.Left, Me.Top, (CurX + 8) - Me.Left, Me.Height - 8
                                        Else
                                        StartResize Me.Left, (CurY - 8), (CurX + 8) - Me.Left, Me.Top + Me.Height - CurY
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            
            Case 8  'Middle Top
                If CurY < SnapDistance Then
                    StartResize Me.Left, 0, Me.Width, (Me.Height + Me.Top) - 8
                    Else
                    If CurY > (Me.Top - SnapDistance) And CurY < (Me.Top + SnapDistance) Then  'Snap to Form
                        StartResize Me.Left, Me.Top, Me.Width, Me.Height - 8
                        Else
                        StartResize Me.Left, (CurY - 8), Me.Width, Me.Height + Me.Top - CurY
                    End If
                End If
        
        End Select
        
    End If

End Sub


