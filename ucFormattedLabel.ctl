VERSION 5.00
Begin VB.UserControl ucFormattedLabel 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   ClipControls    =   0   'False
   MaskColor       =   &H00000000&
   PropertyPages   =   "ucFormattedLabel.ctx":0000
   ScaleHeight     =   4335
   ScaleWidth      =   6330
   ToolboxBitmap   =   "ucFormattedLabel.ctx":000E
End
Attribute VB_Name = "ucFormattedLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Type rsTextItem
    Text As String
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
    Strikethru As Boolean
    NewLine As Boolean
    Color As Long
End Type

Private Enum rsTag
    rsNone = 1
    rsBold = 2
    rsItalic = 3
    rsUnderline = 4
    rsStrikethru = 5
    rsColor = 6
    rsCR = 7
End Enum

Private Type rsNextTag
    Tag As rsTag
    value As String
    Start As Integer
    End As Integer
End Type

Private TextItems() As rsTextItem

Const cBoldStart = "<b>"
Const cBoldStop = "</b>"
Const cItalicStart = "<i>"
Const cItalicStop = "</i>"
Const cUnderlineStart = "<u>"
Const cUnderlineStop = "</u>"
Const cStrikethruStart = "<s>"
Const cStrikethruStop = "</s>"
Const cBreakeLine = "<br>"
Const cColorStop = "</c>"

Dim m_Autosize As Boolean
Dim m_Caption As String
Dim m_WordWrap As Boolean
'*********************************************************
' USER_CONTROL METHODS
'*********************************************************
Private Sub UserControl_InitProperties()

    Set UserControl.Font = Ambient.Font
    m_Caption = Ambient.DisplayName

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_Autosize = PropBag.ReadProperty("Autosize", False)
    m_WordWrap = PropBag.ReadProperty("WordWrap", False)
    
End Sub

Private Sub UserControl_Resize()
    
    If UserControl.ScaleWidth < 150 Then UserControl.Width = 150
    If UserControl.ScaleHeight < 150 Then UserControl.Height = 150
    
    Refresh
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Autosize", m_Autosize, False)
    Call PropBag.WriteProperty("WordWrap", m_WordWrap, False)
    
End Sub
'*********************************************************
' PROPERTIES
'*********************************************************
Public Property Get BackColor() As OLE_COLOR
    
    BackColor = UserControl.BackColor

End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    
    Refresh
    
End Property
Public Property Get ForeColor() As OLE_COLOR
    
    ForeColor = UserControl.ForeColor

End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    
    Refresh
    
End Property
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "Caption"

    Caption = m_Caption
    
End Property
Public Property Let Caption(ByVal New_Caption As String)

    m_Caption = Trim(New_Caption)
    PropertyChanged "Caption"
    
    Refresh
    
End Property
Public Property Get Autosize() As Boolean

    Autosize = m_Autosize
    
End Property
Public Property Let Autosize(ByVal New_Autosize As Boolean)

    m_Autosize = New_Autosize
    PropertyChanged "Autosize"
    
    Refresh
    
End Property
Public Property Get WordWrap() As Boolean

    WordWrap = m_WordWrap
    
End Property
Public Property Let WordWrap(ByVal New_WordWrap As Boolean)

    m_WordWrap = New_WordWrap
    PropertyChanged "WordWrap"
    
    Refresh
    
End Property
Public Property Get Font() As Font
    
    Set Font = UserControl.Font
    
End Property
Public Property Set Font(ByVal New_Font As Font)
    
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    
    Refresh
    
End Property
'*********************************************************
' CUSTOM PROCEDURES
'*********************************************************
Private Sub Refresh()
    
    Dim bTemp As Boolean
    Dim nI As Integer
    Dim sTemp As String
    Dim TempTextItem As rsTextItem

    UserControl.AutoRedraw = True
    UserControl.Cls

        With TempTextItem
            .Text = m_Caption
            .Bold = UserControl.Font.Bold
            .Italic = UserControl.Font.Italic
            .Underline = UserControl.Font.Underline
            .Strikethru = UserControl.Font.Strikethrough
            .NewLine = True
            .Color = UserControl.ForeColor
        End With

        bTemp = CreateItems()
        
        For nI = 1 To UBound(TextItems)
            UserControl.FontBold = TextItems(nI).Bold
            UserControl.FontItalic = TextItems(nI).Italic
            UserControl.FontUnderline = TextItems(nI).Underline
            UserControl.FontStrikethru = TextItems(nI).Strikethru
            UserControl.ForeColor = TextItems(nI).Color
            
            If nI < UBound(TextItems) Then
                If TextItems(nI + 1).NewLine = True Then
                    UserControl.Print TextItems(nI).Text
                
                Else
                    UserControl.Print TextItems(nI).Text;
                
                End If
                
            Else
                UserControl.Print TextItems(nI).Text;
            End If
            
        Next nI

        With TempTextItem
            m_Caption = .Text
            UserControl.Font.Bold = .Bold
            UserControl.Font.Italic = .Italic
            UserControl.Font.Underline = .Underline
            UserControl.Font.Strikethrough = .Strikethru
            UserControl.ForeColor = .Color
        End With

        Select Case m_Autosize
        Case True
            UserControl.Height = CLng(UserControl.TextHeight(GetTestString()))
        
        End Select
        
    UserControl.AutoRedraw = False

End Sub
Private Function CreateItems() As Boolean

    ReDim TextItems(1 To 1)
    
    Dim NextTag As rsNextTag
    Dim nI As Integer: nI = 1
    Dim sTemp As String
    Dim lForeColor As Long
    Dim bTemp As Boolean
    
    lForeColor = IIf((UserControl.ForeColor And &H80000000), GetSysColor(UserControl.ForeColor And &HFFFFFF), UserControl.ForeColor)

    With TextItems(1)
        .Text = m_Caption
        .Bold = UserControl.Font.Bold
        .Italic = UserControl.Font.Italic
        .Underline = UserControl.Font.Underline
        .Strikethru = UserControl.Font.Strikethrough
        .NewLine = True
        .Color = lForeColor
    End With
    
    NextTag = GetNextTag(nI)
    nI = nI + 1
    
    While NextTag.Tag <> rsNone
                
        sTemp = TextItems(nI - 1).Text
        
        If (Trim(Left(sTemp, NextTag.Start - 1)) = "") And (NextTag.Tag <> rsCR) Then
            
            If nI <> 2 Then
                TextItems(nI - 2).Text = TextItems(nI - 2).Text & Left(sTemp, NextTag.Start - 1)
            End If
            
            TextItems(nI - 1).Text = Right(sTemp, Len(sTemp) - NextTag.End)

            Select Case NextTag.Tag
            Case rsBold
                TextItems(nI - 1).Bold = CBool(NextTag.value)
        
            Case rsItalic
                TextItems(nI - 1).Italic = CBool(NextTag.value)
        
            Case rsUnderline
                TextItems(nI - 1).Underline = CBool(NextTag.value)

            Case rsStrikethru
                TextItems(nI - 1).Strikethru = CBool(NextTag.value)

            Case rsCR
                TextItems(nI - 1).NewLine = True
        
            Case rsColor
                TextItems(nI - 1).Color = IIf((NextTag.value = "-1"), lForeColor, CLng(NextTag.value))
        
            End Select

            NextTag = GetNextTag(nI - 1)

        Else
            TextItems(nI - 1).Text = Left(sTemp, NextTag.Start - 1)
            
            If Trim(Right(sTemp, Len(sTemp) - NextTag.End)) = "" Then
                NextTag = GetNextTag(nI - 1)
                nI = nI + 1
            
            Else
            
                            ReDim Preserve TextItems(1 To nI)
                
                With TextItems(nI)
                    .Text = Right(sTemp, Len(sTemp) - NextTag.End)
                    .Bold = TextItems(nI - 1).Bold
                    .Italic = TextItems(nI - 1).Italic
                    .Underline = TextItems(nI - 1).Underline
                    .Strikethru = TextItems(nI - 1).Strikethru
                    .Color = TextItems(nI - 1).Color
                    .NewLine = False
                End With
                
                Select Case NextTag.Tag
                Case rsBold
                    TextItems(nI).Bold = CBool(NextTag.value)
                
                Case rsItalic
                    TextItems(nI).Italic = CBool(NextTag.value)
                
                Case rsUnderline
                    TextItems(nI).Underline = CBool(NextTag.value)
        
                Case rsStrikethru
                    TextItems(nI).Strikethru = CBool(NextTag.value)
        
                Case rsCR
                    TextItems(nI).NewLine = True
                
                Case rsColor
                    TextItems(nI).Color = IIf((NextTag.value = "-1"), lForeColor, CLng(NextTag.value))
                
                End Select
                
                NextTag = GetNextTag(nI)
                nI = nI + 1
            
            End If
            
        End If
        
    Wend
    
    If m_WordWrap = True Then
        bTemp = WrapTextItems
    End If

    CreateItems = True
    
End Function
Private Function GetNextTag(Index As Integer) As rsNextTag
    
    Dim sString As String
    Dim NextTag As rsNextTag
    Dim sTag As String
    
    sString = TextItems(Index).Text
    
    With NextTag
        .Tag = rsNone
        .Start = 0
        .End = 0
        .value = ""
    
        Do
            .Start = InStr(.Start + 1, sString, "<")
            
            If .Start = 0 Then Exit Do
            
            .End = InStr(.Start, sString, ">")
            
            If .End = 0 Then Exit Do
        
            sTag = Mid(sString, .Start, .End - .Start + 1)
            sTag = LCase(sTag)
            
            Select Case sTag
            Case cBoldStart, cBoldStop
                .Tag = rsBold
                .value = IIf((sTag = cBoldStart), "True", "False")
                Exit Do
            
            Case cItalicStart, cItalicStop
                .Tag = rsItalic
                .value = IIf((sTag = cItalicStart), "True", "False")
                Exit Do
            
            Case cUnderlineStart, cUnderlineStop
                .Tag = rsUnderline
                .value = IIf((sTag = cUnderlineStart), "True", "False")
                Exit Do
            
            Case cStrikethruStart, cStrikethruStop
                .Tag = rsStrikethru
                .value = IIf((sTag = cStrikethruStart), "True", "False")
                Exit Do
            
            Case cBreakeLine
                .Tag = rsCR
                .value = ""
                Exit Do
            
            Case cColorStop
                .Tag = rsColor
                .value = "-1"
                
                Exit Do
    
            Case Else
                If sTag Like "<c=*>" = True Then
                    If IsNumeric(Mid(sTag, 4, Len(sTag) - 4)) = True Then
                        .Tag = rsColor
                        .value = Mid(sTag, 4, Len(sTag) - 4)
                        
                        Exit Do
                    End If
                    
                End If
                
            End Select
            
        Loop
    End With

    GetNextTag = NextTag
    
End Function
Private Function GetTestString() As String

    Dim sTemp As String
    Dim nI As Integer
    
    sTemp = ""
    
    For nI = 2 To UBound(TextItems)
        
        If TextItems(nI).NewLine = True Then sTemp = sTemp & "ABCD" & vbCrLf

    Next nI

    GetTestString = sTemp
    
End Function
Private Function WrapTextItems() As Boolean

    Dim nI As Integer: nI = 1
    Dim fWidth As Single
    Dim fTempWidth As Single
    Dim bTemp As Boolean
    Dim vTemp As Variant
    
    fWidth = UserControl.ScaleWidth
    
    Do
        UserControl.FontBold = TextItems(nI).Bold
        UserControl.FontItalic = TextItems(nI).Italic
    
        If TextItems(nI).NewLine = True Then fTempWidth = 0
        
        fTempWidth = fTempWidth + UserControl.TextWidth(TextItems(nI).Text)
        
        If fTempWidth > fWidth Then
            
            vTemp = WrapItem(nI, (fWidth) - (fTempWidth - UserControl.TextWidth(TextItems(nI).Text)))
            
            If Trim(vTemp(1)) <> "" Then
                If (vTemp(3) = "novyradek") And (TextItems(nI).NewLine = False) Then
                    TextItems(nI).NewLine = True
                    
                ElseIf (vTemp(3) = "novyradek") And (TextItems(nI).NewLine = True) Then
                    bTemp = InsertTextItem(nI)
                    
                    TextItems(nI).Text = vTemp(1)
                        
                    TextItems(nI + 1).Text = vTemp(2)
                    TextItems(nI + 1).NewLine = True
                    
                    nI = nI + 1
                
                ElseIf vTemp(3) = "" Then
                    bTemp = InsertTextItem(nI)
                    
                    TextItems(nI).Text = vTemp(1)
                        
                    TextItems(nI + 1).Text = vTemp(2)
                    TextItems(nI + 1).NewLine = True
                    
                    nI = nI + 1
                 
                End If
                                   
            ElseIf Trim(vTemp(1)) = "" Then
                If TextItems(nI).NewLine = True Then
                    
                    If nI + 1 = UBound(TextItems) + 1 Then Exit Do
                    
                    TextItems(nI + 1).NewLine = True
                    
                    nI = nI + 1
                    
                Else
                    TextItems(nI).NewLine = True
                
                End If
                
            End If
        
        Else
            nI = nI + 1
            
        End If
        
        If nI = UBound(TextItems) + 1 Then Exit Do
                
    Loop
   
    WrapTextItems = True
    
End Function
Private Function InsertTextItem(nIndex As Integer) As Boolean

    Dim nI As Integer
    
    ReDim Preserve TextItems(1 To UBound(TextItems) + 1)
    
    For nI = UBound(TextItems) To nIndex + 1 Step -1
        TextItems(nI).Bold = TextItems(nI - 1).Bold
        TextItems(nI).Color = TextItems(nI - 1).Color
        TextItems(nI).Italic = TextItems(nI - 1).Italic
        TextItems(nI).NewLine = TextItems(nI - 1).NewLine
        TextItems(nI).Strikethru = TextItems(nI - 1).Strikethru
        TextItems(nI).Text = TextItems(nI - 1).Text
        TextItems(nI).Underline = TextItems(nI - 1).Underline
    Next nI
    
    InsertTextItem = True

End Function
Private Function WrapItem(nIndex As Integer, fWidth As Long) As Variant

    Dim sRetVal(1 To 2) As String
    Dim sTemp As String
    Dim lPozice(1 To 2) As Long
    Dim sLine(1 To 3) As String

    sTemp = TextItems(nIndex).Text

    lPozice(1) = IIf((InStr(1, sTemp, " ") = 0), Len(sTemp) + 1, InStr(1, sTemp, " "))
    lPozice(2) = IIf((InStr(lPozice(1) + 1, sTemp, " ") = 0), Len(sTemp) + 1, InStr(lPozice(1) + 1, sTemp, " "))

    While (lPozice(1) <> Len(sTemp) + 1) ' And (lPozice(2) <> Len(sTemp) + 1)

        sLine(1) = Left(sTemp, lPozice(1) - 1)
        sLine(2) = Left(sTemp, lPozice(2) - 1)

        If (UserControl.TextWidth(sLine(1)) <= fWidth) And (UserControl.TextWidth(sLine(2)) > fWidth) Then
            sLine(1) = RTrim(sLine(1))
            sLine(2) = LTrim(Right(sTemp, Len(sTemp) - lPozice(1)))
            sLine(3) = ""
            
            WrapItem = sLine
        
            Exit Function
        End If

        lPozice(1) = lPozice(2)
        lPozice(2) = IIf((InStr(lPozice(1) + 1, sTemp, " ") = 0), Len(sTemp) + 1, InStr(lPozice(1) + 1, sTemp, " "))

    Wend

    lPozice(1) = InStr(1, sTemp, " ")
    'lPozice(2) = InStr(lPozice(1) + 1, sTemp, " ")

    If lPozice(1) <> 0 Then
        sLine(1) = Left(sTemp, lPozice(1) - 1)
        sLine(2) = LTrim(Right(sTemp, Len(sTemp) - lPozice(1)))
        sLine(3) = "novyradek"
    Else
        sLine(1) = ""
        sLine(2) = ""
        sLine(3) = "novyradek"
    End If
    
    WrapItem = sLine

End Function
