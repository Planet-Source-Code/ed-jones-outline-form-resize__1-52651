Attribute VB_Name = "modResize"
'/////////////////////////////////////////////////////////////////
'///////////// OUTLINE FORM RESIZE by Edward Jones ///////////////
'/////////////   Copyright (C) Edward Jones, 2004  ///////////////
'/////////////////////////////////////////////////////////////////
'Title:    Outline Form Resize
'Author:   Edward Jones
'Date:     25/03/2004
'Version:  1.0
'Website:  http://www.coldfusionuk.biz
'E-Mail:   edwardjones@coldfusionuk.biz

'This code will allow you to resize forms using outline dragging.
'This is usefull for people like me who want to skin their app
'and have the problem of flickering graphics.  This code works
'best if the forms borderstyle is set to "0 - None"

'Settings:
'You can customise the snapping distance of the outline in the
'Form Load sub for frmMain.

'Feel free to use this code in your application but please keep
'these comments and copyright information.  Also if anyone updates
'the code please let me know via the e-mail address above.  If you
'like this code and find it usefull it would be nice to know =)
'/////////////////////////////////////////////////////////////////



'Cursor Pos API
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public EndTop, EndLeft, EndWidth, EndHeight As Long

Public SnapDistance As Integer
Public Sub EndResize()

    frmTop.Hide
    frmLeft.Hide
    frmBottom.Hide
    frmRight.Hide
    
End Sub

Public Sub StartResize(Optional Left, Optional Top, Optional Width, Optional Height As Long)
    
    frmTop.Hide
    frmTop.Top = Top
    frmTop.Width = Width
    frmTop.Left = Left
    frmTop.Height = 10
    frmTop.Show
    
    frmLeft.Hide
    frmLeft.Top = Top
    frmLeft.Left = Left
    frmLeft.Height = Height
    frmLeft.Width = 10
    frmLeft.Show
    
    frmBottom.Hide
    frmBottom.Top = Top + Height
    frmBottom.Left = Left
    frmBottom.Height = 10
    frmBottom.Width = Width
    frmBottom.Show
    
    frmRight.Hide
    frmRight.Top = Top
    frmRight.Left = Left + Width
    frmRight.Height = Height
    frmRight.Width = 10
    frmRight.Show
    
    EndLeft = Left
    EndTop = Top
    EndWidth = Width
    EndHeight = Height

End Sub

