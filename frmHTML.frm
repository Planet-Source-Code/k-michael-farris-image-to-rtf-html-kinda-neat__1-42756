VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHTML 
   Caption         =   "HTML Conversion - Please wait"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   Icon            =   "frmHTML.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7935
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   7560
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1440
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7560
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmHTML.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    cd1.Filter = "HTML/HTM (*.html;*.htm)|*.html;*.htm"

    cd1.ShowSave
'standard variable name here...freek
    freek = FreeFile
    Open cd1.FileName For Output As freek
    Print #freek, text1.Text
    Close freek

End Sub
