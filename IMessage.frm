VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form IMessage 
   ClientHeight    =   4065
   ClientLeft      =   7440
   ClientTop       =   4185
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5580
   Begin VB.CommandButton cmdNewBuddy 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      MaskColor       =   &H00FF00FF&
      Picture         =   "IMessage.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   " Add a new buddy  "
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin RichTextLib.RichTextBox showmsg 
      Height          =   1815
      Left            =   60
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3201
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"IMessage.frx":0314
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox typemsg 
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   2520
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"IMessage.frx":0390
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3840
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":040C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":051E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":0630
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":0742
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":0B94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbFonts 
      Height          =   330
      Left            =   60
      TabIndex        =   3
      Top             =   2160
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Decreace Font Size"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Increase Font Size"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.ComboBox cmbFonts 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "&Font"
      Begin VB.Menu mnuFontBold 
         Caption         =   "Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFontItalic 
         Caption         =   "Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFontUnderline 
         Caption         =   "Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFontSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontSize 
         Caption         =   "Font Size"
         Begin VB.Menu mnuFontPT 
            Caption         =   "8 pt"
            Index           =   0
         End
         Begin VB.Menu mnuFontPT 
            Caption         =   "10 pt"
            Index           =   1
         End
         Begin VB.Menu mnuFontPT 
            Caption         =   "12 pt"
            Index           =   2
         End
         Begin VB.Menu mnuFontPT 
            Caption         =   "14 pt"
            Index           =   3
         End
      End
      Begin VB.Menu mnuSpellDefine 
         Caption         =   "Spellcheck"
         Begin VB.Menu mnuDefine 
            Caption         =   "Define"
         End
         Begin VB.Menu mnuSpell 
            Caption         =   "Spellcheck"
         End
      End
   End
   Begin VB.Menu mnuSpellDefine2 
      Caption         =   "Spellcheck"
      Begin VB.Menu mnuDefine2 
         Caption         =   "Define"
      End
      Begin VB.Menu mnuSpell2 
         Caption         =   "Spellcheck"
      End
   End
End
Attribute VB_Name = "IMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dWord As String

Private Sub cmbFonts_Click()
    typemsg.SelFontName = cmbFonts.Text
    typemsg.SetFocus
End Sub

Private Sub cmdSend_Click()
typemsg.Text = Replace(typemsg.Text, vbCrLf, "")
If typemsg.Text <> "" And Len(typemsg) > 0 Then
    showmsg.SelStart = Len(showmsg.Text)
    showmsg.SelBold = True
    showmsg.SelColor = vbRed
    showmsg.SelText = Client.Caption & ": "
    
    
    showmsg.SelStart = Len(showmsg.Text)
    showmsg.SelBold = False
    showmsg.SelColor = vbBlack
    showmsg.SelText = typemsg.Text & vbCrLf
  
    PlaySound ("sounds/imsend.wav")
    Client.Winsock1.SendData ".msg " & Client.Caption & " " & Word(Me.Caption, 1) & " ..//.. " & typemsg.Text
    Client.WaitFor (".msgOK")
    typemsg.Text = ""
End If
typemsg.Text = Replace(typemsg.Text, vbCrLf, "")
typemsg.SetFocus
End Sub

Private Sub Form_Load()
Dim i As Integer
    mnuFont.Visible = False
    mnuSpellDefine2.Visible = False
    'bAllowScroll = True
    'Call SetHook(showmsg.hwnd, True)
    For i = 1 To Screen.FontCount
        cmbFonts.AddItem Screen.Fonts(i)
    Next i
    cmbFonts.RemoveItem (0)
    cmbFonts.SelText = "Verdana"
End Sub

'Private Sub Form_Unload(Cancel As Integer)
    'Call SetHook(showmsg.hwnd, False)
'End Sub

Private Sub mnuFontBold_Click()

If typemsg.SelBold = True Then
   typemsg.SelBold = False
   tbFonts.Buttons(1).Value = tbrUnpressed
Else
   typemsg.SelBold = True
   tbFonts.Buttons(1).Value = tbrPressed
End If

End Sub

Private Sub mnuFontItalic_Click()

If typemsg.SelItalic = True Then
   typemsg.SelItalic = False
   tbFonts.Buttons(2).Value = tbrUnpressed
Else
   typemsg.SelItalic = True
   tbFonts.Buttons(2).Value = tbrPressed
End If

End Sub

Private Sub mnuFontUnderline_Click()

If typemsg.SelUnderline = True Then
   typemsg.SelUnderline = False
   tbFonts.Buttons(3).Value = tbrUnpressed
Else
   typemsg.SelUnderline = True
   tbFonts.Buttons(3).Value = tbrPressed
End If

End Sub

Private Sub mnuFontPT_Click(Index As Integer)
   typemsg.SelFontSize = Word(mnuFontPT(Index).Caption, 1)
End Sub

Private Sub showmsg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
       If showmsg.SelText <> "" And is_chars(Replace(Trim(showmsg.SelText), vbCrLf, "")) = True Then
            dWord = Replace(Trim(showmsg.SelText), vbCrLf, "")
        Else
            dWord = "- Please highlight a word."
        End If
        mnuSpell2.Caption = "Spellcheck " & dWord
        mnuDefine2.Caption = "Define " & dWord
        PopupMenu mnuSpellDefine2
End If
End Sub

Private Function is_chars(x As String) As Boolean
Dim i As Integer
Dim flag As Integer
For i = 1 To Len(x)
    If (Asc(UCase(Mid(x, i, 1))) >= vbKeyA And Asc(UCase(Mid(x, i, 1))) <= vbKeyZ) Or Mid(x, i, 1) = " " Then
        flag = 0
    Else
        flag = 1
        Exit For
        is_chars = False
    End If
Next i

If flag = 0 Then
    is_chars = True
End If
    
End Function

Private Sub tbFonts_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
    Case 1
        mnuFontBold_Click
    Case 2
        mnuFontItalic_Click
    Case 3
        mnuFontUnderline_Click
    Case 5
        If typemsg.SelFontSize > 8 Then
            typemsg.SelFontSize = typemsg.SelFontSize - 2
        End If
    Case 6
        If typemsg.SelFontSize < 14 Then
            typemsg.SelFontSize = typemsg.SelFontSize + 2
        End If
End Select

End Sub

Private Sub typemsg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    typemsg_Click
End If
End Sub

Private Sub typemsg_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call cmdSend_Click
End If
End Sub

Private Sub typemsg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo HandleError
If Button = vbRightButton Then
        If typemsg.SelText <> "" And is_chars(Replace(Trim(typemsg.SelText), vbCrLf, "")) = True Then
            dWord = Replace(Trim(typemsg.SelText), vbCrLf, "")
        Else
            dWord = "- Please highlight a word."
        End If
        mnuSpell.Caption = "Spellcheck " & dWord
        mnuDefine.Caption = "Define " & dWord
    
    If typemsg.SelBold = True Then
       mnuFontBold.Checked = True
    Else
       mnuFontBold.Checked = False
    End If
    
    If typemsg.SelItalic = True Then
       mnuFontItalic.Checked = True
    Else
       mnuFontItalic.Checked = False
    End If
    
    If typemsg.SelUnderline = True Then
       mnuFontUnderline.Checked = True
    Else
       mnuFontUnderline.Checked = False
    End If
    PopupMenu mnuFont
End If

HandleError:
    'MsgBox Err.Number & " - " & Err.Description, vbOKOnly
End Sub

Private Sub typemsg_Click()
If cmbFonts.Text <> typemsg.SelFontName Then
    cmbFonts.Text = ""
    cmbFonts.SelText = typemsg.SelFontName
End If

    If typemsg.SelBold = True Then
       tbFonts.Buttons(1).Value = tbrPressed
    Else
       tbFonts.Buttons(1).Value = tbrUnpressed
       
    End If
    
    If typemsg.SelItalic = True Then
       tbFonts.Buttons(2).Value = tbrPressed
    Else
       tbFonts.Buttons(2).Value = tbrUnpressed
    End If
    
    If typemsg.SelUnderline = True Then
       tbFonts.Buttons(3).Value = tbrPressed
    Else
       tbFonts.Buttons(3).Value = tbrUnpressed
    End If
End Sub

Private Sub mnuSpell_Click()
    Call DefineSpell(".spell")
End Sub

Private Sub mnuDefine_Click()
    Call DefineSpell(".define")
End Sub

Private Sub mnuSpell2_Click()
    Call DefineSpell(".spell")
End Sub

Private Sub mnuDefine2_Click()
    Call DefineSpell(".define")
End Sub

Private Function DefineSpell(wtd As String)
If Word(dWord, 1) <> "-" Then
    'MsgBox "Gonna DO IT", vbOKOnly
    Client.Winsock1.SendData wtd & " " & Word(Me.Caption, 1) & " " & dWord
    Client.WaitFor (wtd)
End If
End Function
