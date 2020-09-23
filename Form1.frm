VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Image to RTF/HTML Converter 3000"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Save RTF"
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      ToolTipText     =   "Load Picture"
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Convert 2 HTML"
      Height          =   315
      Left            =   2640
      TabIndex        =   16
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Strikethrough"
      Height          =   255
      Left            =   8160
      TabIndex        =   15
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Underline"
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   8520
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Itallic"
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   8520
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bold"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   8520
      Width           =   855
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   135
      Left            =   7560
      TabIndex        =   9
      Top             =   8040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      _Version        =   393216
      LargeChange     =   1
      Min             =   3
      SelStart        =   4
      Value           =   4
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Text            =   "Hello"
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Text"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   6
      Text            =   "#"
      Top             =   7920
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Symbol"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   7920
      Width           =   975
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   7455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   13150
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Color Points"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   7920
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   7560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Clear"
      Height          =   255
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   111
      TabIndex        =   0
      Top             =   7560
      Width           =   1695
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   0
         Stretch         =   -1  'True
         ToolTipText     =   "Click here to load a picture!"
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   8160
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Font Size"
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   7800
      Width           =   1815
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is just a quick little app i through together for no real reason whatsoever.
'Feel free to do with it as you please.
'Take note that the RTF to HTML conversion is by Joseph Huntly.
' and the RGB conversion is by another author as well (i'm afraid i don't remember who right at the moment.)
'there isn't an enormous amount of use for this, but its neat to play around with.
'maybe someone can come up with a 'practical use' for it (heh..right)
' i put very little time into the GUI, cuz frankly i don't care that much about this app.
'if you want to actually spice it up or do the resizing, (or add error handling for that matter) go for it.
' have fun, -K. Michael Farris


Private Sub Check1_Click()
    If Check1.Value = 1 Then
        text1.Font.Bold = True
    Else
        text1.Font.Bold = False
    End If

End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        text1.Font.Italic = True
    Else
        text1.Font.Italic = False
    End If

End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        text1.Font.Underline = True
    Else
        text1.Font.Underline = False
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        text1.Font.Strikethrough = True
    Else
        text1.Font.Strikethrough = False
    End If
End Sub



Private Sub Command1_Click()
cd1.Filter = "Image Files (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif|All files (*.*)|*.*"

    cd1.ShowOpen
    Image1.Picture = LoadPicture(cd1.FileName, , , Image1.Width, Image1.Height)
End Sub

Private Sub Command2_Click()
    text1.Text = ""

End Sub

Private Sub Command3_Click()
'This colors the picture and uses the numbers from its Red value to determine the character used.
    ProgressBar1.Max = Picture1.ScaleHeight
    For y = 1 To Picture1.ScaleHeight
        ProgressBar1.Value = y
        For x = 1 To Picture1.ScaleWidth
            DoEvents
            'this is based on the ColorPicker code, I'm not sure who the author was, but thanks.
            'this determines the RGB values of the pixel, these are needed to work with the Rtf Selcolor property.
            texr = Picture1.Point(x, y) And 255
            texg = (Picture1.Point(x, y) And 65280) / 256

            texb = (Picture1.Point(x, y) And 16711680) / 65535

            text1.SelStart = Len(text1.Text) + 1
            text1.SelLength = 1
            text1.SelColor = RGB(texr, texg, texb)
            text1.SelText = Chr(Asc(texr))
            DoEvents
        ' just a note, this would be faster if the drawing was done 'invisibly' as in a function, and then sent to the RTF box.
        Next x
        'Joseph's RTF to HTML code didn't add line breaks, which is vital to the picture coming out right.  So i just added them here.
        'they are invisible due to the coloring
        text1.SelText = "<br>" & vbCrLf
        DoEvents
    Next y
End Sub

Private Sub Command4_Click()
'this uses the selected symbol or letter to 'paint' the picture
    ProgressBar1.Max = Picture1.ScaleHeight
    For y = 1 To Picture1.ScaleHeight
        ProgressBar1.Value = y
        For x = 1 To Picture1.ScaleWidth
            DoEvents
            texr = Picture1.Point(x, y) And 255
            texg = (Picture1.Point(x, y) And 65280) / 256

            texb = (Picture1.Point(x, y) And 16711680) / 65535

            text1.SelStart = Len(text1.Text) + 1
            text1.SelLength = 1
            text1.SelColor = RGB(texr, texg, texb)
            text1.SelText = Text2.Text
            DoEvents

        Next x
        text1.SelText = "<br>" & vbCrLf
        DoEvents
    Next y
End Sub

Private Sub Command5_Click()
'this uses the defined text to 'paint' the picture
    ProgressBar1.Max = Picture1.ScaleHeight
    For y = 1 To Picture1.ScaleHeight
        ProgressBar1.Value = y

        For x = 1 To Picture1.ScaleWidth
            m = m + 1
            If m = Len(Text3.Text) + 1 Then
                m = 1
            End If
            DoEvents
            texr = Picture1.Point(x, y) And 255
            texg = (Picture1.Point(x, y) And 65280) / 256

            texb = (Picture1.Point(x, y) And 16711680) / 65535

            text1.SelStart = Len(text1.Text) + 1
            text1.SelLength = 1
            text1.SelColor = RGB(texr, texg, texb)
            text1.SelText = Mid(Text3.Text, m, 1)
            DoEvents

        Next x

        text1.SelText = "<br>" & vbCrLf
        DoEvents
    Next y
End Sub

Private Sub Command6_Click()
    frmHTML.Show
    frmHTML.text1 = RichToHTML(text1, 0, Len(text1.Text))

End Sub

Private Sub Command7_Click()
cd1.Filter = "RichText Format (*.rtf)|*.rtf"

cd1.ShowSave
text1.SaveFile cd1.FileName

End Sub

Private Sub Form_Load()
    Image1.Width = Picture1.ScaleWidth
    Image1.Height = Picture1.ScaleHeight
End Sub

Private Sub Form_Resize()
    text1.Width = Me.Width - 500
End Sub

Private Sub Form_Terminate()
Unload frmHTML

End Sub

Private Sub Image1_Click()
cd1.Filter = "Image Files (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif|All files (*.*)|*.*"

    cd1.ShowOpen
    Image1.Picture = LoadPicture(cd1.FileName, , , Image1.Width, Image1.Height)

End Sub

Private Sub Slider1_Click()
    text1.Font.Size = Slider1.Value
    Label2.Caption = Slider1.Value
End Sub
Function RichToHTML(rtbRichTextBox As RichTextLib.RichTextBox, Optional lngStartPosition As Long, Optional lngEndPosition As Long) As String

    '**********************************************************
    '*            Rich To HTML by Joseph Huntley              *
    '*               joseph_huntley@email.com                 *
    '*                http://joseph.vr9.com                   *
    '**********************************************************
    '*   You may use this code freely as long as credit is    *
    '* given to the author, and the header remains intact.    *
    '**********************************************************

    '--------------------- The Arguments -----------------------
    'rtbRichTextBox     - The rich textbox control to convert.
    'lngStartPosition   - The character position to start from.
    'lngEndPosition     - The character position to end at.
    '-----------------------------------------------------------
    'Returns:     The rich text converted to HTML.

    'Description: Converts rich text to HTML.

    Dim blnBold As Boolean, blnUnderline As Boolean, blnStrikeThru As Boolean
    Dim blnItalic As Boolean, strLastFont As String, lngLastFontColor As Long
    Dim strHTML As String, lngColor As Long, lngRed As Long, lngGreen As Long
    Dim lngBlue As Long, lngCurText As Long, strHex As String, intLastAlignment As Integer

    Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2

    'check for lngStartPosition ad lngEndPosition

    If IsMissing(lngStartPosition&) Then lngStartPosition& = 0
    If IsMissing(lngEndPosition&) Then lngEndPosition& = Len(rtbRichTextBox.Text)

    lngLastFontColor& = -1                                 'no color
    frmHTML.pb2.Max = lngEndPosition&

    For lngCurText& = lngStartPosition& To lngEndPosition&
        rtbRichTextBox.SelStart = lngCurText&
        rtbRichTextBox.SelLength = 1
        frmHTML.pb2.Value = lngCurText&
        DoEvents
        If intLastAlignment% <> rtbRichTextBox.SelAlignment Then
            intLastAlignment% = rtbRichTextBox.SelAlignment

            Select Case rtbRichTextBox.SelAlignment
                Case AlignLeft: strHTML$ = strHTML$ & "<p align=left>"
                Case AlignRight: strHTML$ = strHTML$ & "<p align=right>"
                Case AlignCenter: strHTML$ = strHTML$ & "<p align=center>"
            End Select

        End If

        If blnBold <> rtbRichTextBox.SelBold Then
            If rtbRichTextBox.SelBold = True Then
                strHTML$ = strHTML$ & "<b>"
            Else
                strHTML$ = strHTML$ & "</b>"
            End If
            blnBold = rtbRichTextBox.SelBold
        End If

        If blnUnderline <> rtbRichTextBox.SelUnderline Then
            If rtbRichTextBox.SelUnderline = True Then
                strHTML$ = strHTML$ & "<u>"
            Else
                strHTML$ = strHTML$ & "</u>"
            End If
            blnUnderline = rtbRichTextBox.SelUnderline
        End If


        If blnItalic <> rtbRichTextBox.SelItalic Then
            If rtbRichTextBox.SelItalic = True Then
                strHTML$ = strHTML$ & "<i>"
            Else
                strHTML$ = strHTML$ & "</i>"
            End If
            blnItalic = rtbRichTextBox.SelItalic
        End If


        If blnStrikeThru <> rtbRichTextBox.SelStrikeThru Then
            If rtbRichTextBox.SelStrikeThru = True Then
                strHTML$ = strHTML$ & "<s>"
            Else
                strHTML$ = strHTML$ & "</s>"
            End If
            blnStrikeThru = rtbRichTextBox.SelStrikeThru
        End If

        If strLastFont$ <> rtbRichTextBox.SelFontName Then
            strLastFont$ = rtbRichTextBox.SelFontName
            'added a change here to Joseph's original code to include a font size.
            strHTML$ = strHTML$ + "<font face=""" & strLastFont$ & """" + " size=" + """" + "-5" + """" + "> "
        End If

        If lngLastFontColor& <> rtbRichTextBox.SelColor Then
            lngLastFontColor& = rtbRichTextBox.SelColor

            ''Get hexidecimal value of color
            strHex$ = Hex(rtbRichTextBox.SelColor)
            strHex$ = String$(6 - Len(strHex$), "0") & strHex$
            strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)

            strHTML$ = strHTML$ + "<font color=#" & strHex$ & ">"
        End If

        strHTML$ = strHTML$ + rtbRichTextBox.SelText
        DoEvents
    Next lngCurText&
    'adding teletype seems to tighten the picture up a bit in HTML
    'you can change or delete the <h> tag to change the size of the picture even more
    
    RichToHTML = "<h6><tt>" + strHTML$ + "</tt></h6>"

End Function
