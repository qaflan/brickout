VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "D¸zenle"
   ClientHeight    =   2940
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3375
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command2 
         Caption         =   "›«—”Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "T¸rkÁe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Iptal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Tamam"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "http://vbgaming.blogfa.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   720
      TabIndex        =   9
      Top             =   2640
      Width           =   2025
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "http://www.TakClick.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   720
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "yazar:  Cavad Nur˛i"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Toplarin sayi:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    End
End Sub

Private Sub Command1_Click()
    Me.Caption = "D¸zenle"
    Label1.Caption = "Toplarin sayi:"
    OKButton.Caption = "&Tamam"
    CancelButton.Caption = "Iptal"
    Label2.Caption = "yazar:  Cavad Nur˛i"
    Label1.Left = 480
    OKButton.Left = 480
    CancelButton.Left = 1800
    Text1.Left = 1800
    Form1.gLang = "Turkish"
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Me.Caption = " ‰ŸÌ„« "
    Label1.Caption = " ⁄œ«œ  ÊÅ"
    OKButton.Caption = "„Ê«›ﬁ"
    CancelButton.Caption = "«‰’—«›"
    Label2.Caption = " Ê”ÿ: ÃÊ«œ ‰Ê—Ì"
    Label1.Left = 1800
    OKButton.Left = 1800
    CancelButton.Left = 480
    Text1.Left = 480
    Form1.gLang = "Farsi"
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Command2_Click
End Sub

Private Sub OKButton_Click()
    On Error Resume Next
    If IsNumeric(Text1.Text) = False Then
        MsgBox IIf(Form1.gLang = "Turkish", "L¸tfen bir eded yaziniz", "·ÿ›« ⁄œœ Ê«—œ ﬂ‰Ìœ"), vbExclamation
    ElseIf Val(Text1.Text) > 4 Then
        MsgBox IIf(Form1.gLang = "Turkish", "Toplarin sayi 4 den Áox ola bilmez", " ⁄œ«œ  ÊÅÂ« ‰„Ì  Ê«‰œ »Ì‘ «“ 4  « »«‘œ"), vbExclamation
    Else
        Form1.BallCount = Val(Text1.Text)
        Load Form1
        Form1.Show
        Unload Me
    End If
End Sub
