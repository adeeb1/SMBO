Attribute VB_Name = "modText"
Option Explicit

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Sub SetFont(ByRef GlobalFont As Long, ByVal Font As String, ByVal Size As Byte)
    GlobalFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hDC As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long, ByVal GlobalFont As Long)
    Call SelectObject(hDC, GlobalFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, x + 1, y + 0, Text, Len(Text))
    Call TextOut(hDC, x + 0, y + 1, Text, Len(Text))
    Call TextOut(hDC, x - 1, y - 0, Text, Len(Text))
    Call TextOut(hDC, x - 0, y - 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, x, y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
    Dim i As Integer
    Dim Msg2() As String
    
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text)
    frmMirage.txtChat.SelText = vbNewLine
    
    Msg2 = Split(Msg, " ")
    
    For i = 0 To UBound(Msg2)
        If Mid$(Msg2(i), 1, 7) = "http://" Or Mid$(Msg2(i), 1, 4) = "www." Then
            frmMirage.txtChat.SelColor = QBColor(BRIGHTBLUE)
            frmMirage.txtChat.SelUnderline = True
        Else
            frmMirage.txtChat.SelColor = QBColor(Color)
            frmMirage.txtChat.SelUnderline = False
        End If
            frmMirage.txtChat.SelText = Msg2(i) & " "
    Next i
    
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text) - 1

    If frmMirage.chkAutoScroll.Value = Unchecked Then
        frmMirage.txtChat.SelStart = frmMirage.txtChat.SelStart
    End If
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text & (Msg & vbNewLine)
    Else
        Txt.Text = Txt.Text & Msg
    End If

    Txt.SelStart = Len(Txt.Text)
End Sub


