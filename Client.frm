VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Client 
   Caption         =   "Sign On"
   ClientHeight    =   5520
   ClientLeft      =   13845
   ClientTop       =   3315
   ClientWidth     =   2610
   Icon            =   "Client.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleMode       =   0  'User
   ScaleWidth      =   2615
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   4800
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame LoginFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      TabIndex        =   3
      Top             =   -120
      Width           =   2775
      Begin VB.CommandButton cmdSignOn 
         Caption         =   "Sign On"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   4200
         Width           =   735
      End
      Begin VB.ComboBox cmbUsername 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Client.frx":000C
         Left            =   120
         List            =   "Client.frx":000E
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         ToolTipText     =   "Click Here To Enter Your Screen Name"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label lblmain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   90
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "ScreenName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   2520
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label3 
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   5400
         Width           =   1095
      End
   End
   Begin VB.Frame MainFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Visible         =   0   'False
      Width           =   2775
      Begin TabDlg.SSTab SSTab 
         Height          =   4215
         Left            =   30
         TabIndex        =   8
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   7435
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabMaxWidth     =   2117
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Online"
         TabPicture(0)   =   "Client.frx":0010
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "TreeView1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "runlog"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "List Setup"
         TabPicture(1)   =   "Client.frx":002C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TreeView2"
         Tab(1).Control(1)=   "cmdDelBuddy"
         Tab(1).Control(2)=   "cmdNewBuddy"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin RichTextLib.RichTextBox runlog 
            Height          =   735
            Left            =   120
            TabIndex        =   15
            Top             =   3360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1296
            _Version        =   393217
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"Client.frx":0048
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdNewBuddy 
            Height          =   495
            Left            =   -74880
            MaskColor       =   &H00FF00FF&
            Picture         =   "Client.frx":00C4
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   " Add a new buddy  "
            Top             =   3600
            UseMaskColor    =   -1  'True
            Width           =   735
         End
         Begin VB.CommandButton cmdDelBuddy 
            Caption         =   "Remove"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74040
            TabIndex        =   11
            Top             =   3720
            Width           =   855
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   2895
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   5106
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   3
            HotTracking     =   -1  'True
            ImageList       =   "ImageList1"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   10
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   5318
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LineStyle       =   1
            Style           =   6
            HotTracking     =   -1  'True
            ImageList       =   "ImageList1"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1680
         Top             =   4800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Client.frx":03D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Client.frx":04EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Client.frx":0A3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Client.frx":0B50
               Key             =   "down"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Client.frx":0C62
               Key             =   "right"
            EndProperty
         EndProperty
      End
      Begin MCI.MMControl MMControl1 
         Height          =   375
         Left            =   -120
         TabIndex        =   14
         Top             =   4680
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   661
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Online"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   645
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileLogOut 
         Caption         =   "&Log Out"
      End
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileStatus 
         Caption         =   "My &Status"
         Begin VB.Menu mnuStatusOnline 
            Caption         =   "&Online"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuStatusAway 
            Caption         =   "&Away"
         End
      End
      Begin VB.Menu mnuStatusSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strIncoming As String
Dim Start As Integer
Dim oldLabel As String

Private Sub cmdSignOn_Click()
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Winsock1.RemotePort = 1008
    'Winsock1.RemoteHost = "216.77.225.246" 'put your IP here and comment out the one below
    Winsock1.RemoteHost = "127.0.0.1"       'to allow people to connect to your IP
    Winsock1.Connect
    
Do Until Winsock1.State = sckConnected
    DoEvents: DoEvents: DoEvents: DoEvents
    If Winsock1.State = sckError Then
        MsgBox "Problem connecting!"
        Exit Sub
    End If
Loop
    Winsock1.SendData (".login" & " " & LCase(cmbUsername.Text) & " " & LCase(txtPassword.Text))
End Sub

Private Sub lblStatus_Click()
    PopupMenu mnuFileStatus
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
    ' Close the device.
    MMControl1.Command = "Close"
End Sub

Private Sub Form_Load()
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.TimeFormat = mciFormatMilliseconds
    Me.BorderStyle = 1
    lblmain.Caption = "Instant Messanger" & vbCrLf & "by: Karim Bardai" & vbCrLf & "http://www.fataldesigns.com/"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Winsock1.State <> sckClosed Then Winsock1.Close
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.Height < 4000 Then
        Me.Height = 4000
        Exit Sub
    End If
    
    MainFrame.Width = Me.ScaleWidth
    MainFrame.Height = Me.ScaleHeight
    
    SSTab.Width = MainFrame.Width - 65
    SSTab.Height = MainFrame.Height - 1000
    
    TreeView1.Width = SSTab.Width - 260
    TreeView1.Height = SSTab.Height - 1350
    TreeView2.Width = TreeView1.Width
    TreeView2.Height = TreeView1.Height
    
    runlog.Width = TreeView1.Width
    runlog.Top = TreeView1.Top + TreeView1.Height + 25
    
    cmdNewBuddy.Top = SSTab.Height - 625
    cmdDelBuddy.Top = SSTab.Height - 525
    
    Shape1.Width = Me.ScaleWidth
    
    lblStatus.Left = Me.ScaleWidth - lblStatus.Width - 120
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Bold = False
    Node.Image = "right"
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
    Node.Bold = True
    Node.Image = "down"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Expanded = False Then
    Node.Expanded = True
Else
    Node.Expanded = False
End If
If Not Node.Parent Is Nothing Then
    If Node.Image <> 2 Then
            Dim test As Integer
            FormIsLoaded (TreeView1.SelectedItem)
    End If
End If
End Sub

Private Function FormIsLoaded(frm As String)
Dim FormNbr As Integer
Dim flag As Integer
flag = 0
For FormNbr = 0 To Forms.Count - 1
  If LCase(Word(Forms(FormNbr).Caption, 1)) = LCase(frm) Then
        Forms(FormNbr).SetFocus
        Exit Function
  Else
    flag = 1
  End If
Next FormNbr

If flag = 1 Then
    Dim NewIMessage As New IMessage
    NewIMessage.Show ownerform:=Me
    NewIMessage.Caption = frm & " - Instant Message"
End If
End Function

Private Function GetFormNumber(frm As String) As Integer
Dim FormNbr As Integer
For FormNbr = 0 To Forms.Count - 1
  If LCase(Word(Forms(FormNbr).Caption, 1)) = LCase(frm) Then
        GetFormNumber = FormNbr
        Exit Function
 End If
Next FormNbr
End Function

Private Sub TreeView2_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Bold = False
    cmdDelBuddy.Enabled = False
End Sub

Private Sub TreeView2_Expand(ByVal Node As MSComctlLib.Node)
    Node.Bold = True
End Sub

Private Sub TreeView2_BeforeLabelEdit(Cancel As Integer)
    oldLabel = TreeView2.SelectedItem.Text
End Sub

Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Text <> "Buddies" Then
        cmdDelBuddy.Enabled = True
    Else
        cmdDelBuddy.Enabled = False
    End If
End Sub

Private Sub TreeView2_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim tn As Node
If oldLabel <> "Buddies" Then
    If UCase(NewString) <> UCase(oldLabel) Then
        If Correct_Screenname(NewString) = True Then
            If check_for_duplicate(NewString) = True Then
                TreeView2.SelectedItem.Key = NewString
                For Each tn In TreeView1.Nodes
                    If UCase(tn.Key) = UCase(oldLabel) Then
                        tn.Key = NewString
                        tn.Text = NewString
                        Winsock1.SendData ".updateBuddy " & LCase(Me.Caption) & " " & oldLabel & " " & NewString
                        WaitFor (".statusUpdate")
                        Exit For
                    End If
                Next
            Else
                MsgBox "A buddy with the user name " & UCase(NewString) & " already exists.", vbOKOnly + vbCritical
                Cancel = 1
            End If
            Else
                Cancel = 1
            End If
        Else
            Cancel = 1
        End If
    Else
        Cancel = 1
End If
End Sub

Private Function check_for_duplicate(user As String) As Boolean
Dim tn As Node
Dim flag As Integer
    For Each tn In TreeView1.Nodes
        If UCase(tn.Key) = UCase(user) Then
            flag = 1
            check_for_duplicate = False
            Exit For
        Else
            flag = 0
        End If
    Next
If flag = 0 Then
    check_for_duplicate = True
End If
End Function

Private Sub cmdDelBuddy_Click()
Dim reply As String
Dim tn As Node
    reply = MsgBox("Are you sure you want to delete the following buddy from your list?" & vbCrLf & TreeView2.SelectedItem.Key, vbYesNo + vbCritical)
    If reply = vbYes Then
       Winsock1.SendData ".delBuddy " & LCase(Me.Caption) & " " & TreeView2.SelectedItem.Key
       WaitFor (".statusUpdate")
        TreeView1.Nodes.Remove TreeView2.SelectedItem.Key
        TreeView2.Nodes.Remove TreeView2.SelectedItem.Key
       cmdDelBuddy.Enabled = False
       Call Online_Offline_Text
    End If
End Sub

Private Sub cmdNewBuddy_Click()
Dim newbuddy As String
    cmdNewBuddy.Enabled = False
    newbuddy = InputBox("Enter user name:", cmdNewBuddy.ToolTipText)
    If StrPtr(newbuddy) = 0 Then
        cmdNewBuddy.Enabled = True
    Else
        If Correct_Screenname(newbuddy) = True Then
            If check_for_duplicate(newbuddy) = True Then
                TreeView1.Nodes.Add "Offline", tvwChild, newbuddy, newbuddy
                TreeView2.Nodes.Add "Buddies", tvwChild, newbuddy, newbuddy
                Winsock1.SendData ".newBuddy " & LCase(Me.Caption) & " " & newbuddy
                WaitFor (".statusUpdate")
                cmdNewBuddy.Enabled = True
            Else
               MsgBox "A buddy with the user name " & UCase(newbuddy) & " already exists.", vbOKOnly + vbCritical
               Call cmdNewBuddy_Click
            End If
        Else
            Call cmdNewBuddy_Click
        End If
    End If
End Sub

Private Function Correct_Screenname(screenname As String) As Boolean
Dim i As Integer
Dim flag As Integer
If LCase(screenname) <> LCase(Me.Caption) And Len(screenname) >= 5 And Len(screenname) <= 15 And Not IsNumeric(Left(screenname, 1)) Then
For i = 1 To Len(screenname)
    If (Asc(Mid(screenname, i, 1)) >= vbKey1 And Asc(Mid(screenname, i, 1)) <= vbKey9) Then
        flag = 0
    ElseIf (Asc(UCase(Mid(screenname, i, 1))) >= vbKeyA And Asc(UCase(Mid(screenname, i, 1))) <= vbKeyZ) Then
        flag = 0
    Else
        flag = 1
        MsgBox "A screen name in your list is too short or contains invalid" & vbCrLf & "characters.", vbOKOnly + vbCritical
        Correct_Screenname = False
        Exit For
    End If
Next i
Else
    flag = 1
    MsgBox "A screen name in your list is too short or contains invalid" & vbCrLf & "characters.", vbOKOnly + vbCritical
    Correct_Screenname = False
End If
If flag = 0 Then
    Correct_Screenname = True
End If
End Function

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim i As Long
    Winsock1.GetData strIncoming
   
    If strIncoming = ".badlogin" Then
        MsgBox "The screen name or password you entered is not valid. ", vbOKOnly + vbCritical
        If Winsock1.State <> sckClosed Then
            Winsock1.Close
        End If
    ElseIf strIncoming = ".goodlogin" Then
        Call good_login
        
    ElseIf Word(strIncoming, 1) = ".showonline" And Word(strIncoming, 2) <> "0" Then
        Call Show_Online_buddies(strIncoming)
        
    ElseIf Word(strIncoming, 1) = ".statusUpdate" And Word(strIncoming, 2) <> "0" Then
        Call status_update(Word(strIncoming, 3), Word(strIncoming, 4))
        
    ElseIf Word(strIncoming, 1) = ".msg" Then
        Call get_message(Word(strIncoming, 2), strIncoming)
        
    ElseIf Word(strIncoming, 1) = ".define" Then
        Call get_definition(Word(strIncoming, 2), Word(strIncoming, 3), strIncoming)
        
    ElseIf Word(strIncoming, 1) = ".spell" Then
            Call get_spelling(Word(strIncoming, 2), strIncoming)

    End If
End Sub

Private Function get_spelling(buddy As String, msg As String)
Dim formNum As Integer
Dim definition As String
formNum = GetFormNumber(buddy)
If formNum <> 0 Then
    definition = SplitString(msg, "..//..")
    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
    Forms(formNum).showmsg.SelBold = True
    Forms(formNum).showmsg.SelText = "spellcheck: "
    
    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
    Forms(formNum).showmsg.SelBold = False
    Forms(formNum).showmsg.SelText = definition & vbCrLf
End If
End Function

Private Function get_definition(buddy As String, Word As String, msg As String)
Dim formNum As Integer
Dim definition As String
formNum = GetFormNumber(buddy)
If formNum <> 0 Then
    definition = SplitString(msg, "..//..")
    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
    Forms(formNum).showmsg.SelBold = True
    Forms(formNum).showmsg.SelText = Word & ": "
    
    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
    Forms(formNum).showmsg.SelBold = False
    Forms(formNum).showmsg.SelText = definition & vbCrLf
End If
End Function

Private Function get_message(mfrom As String, msg As String)
Dim i As Long
Dim formNum As Integer
Dim sendMsg As String
    PlaySound ("sounds/imrcv.wav")
    FormIsLoaded (mfrom)
    formNum = GetFormNumber(mfrom)
    sendMsg = SplitString(msg, "..//..")
    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
    Forms(formNum).showmsg.SelBold = True
    Forms(formNum).showmsg.SelColor = vbBlue
    Forms(formNum).showmsg.SelText = mfrom & ": "
    
    Forms(formNum).showmsg.SelStart = Len(Forms(formNum).showmsg.Text)
    Forms(formNum).showmsg.SelBold = False
    Forms(formNum).showmsg.SelColor = vbBlack
    Forms(formNum).showmsg.SelText = sendMsg & vbCrLf
End Function

Private Sub good_login()
        mnuFile.Visible = True
        Me.BorderStyle = vbSizable
        
        LoginFrame.Visible = False
        MainFrame.Visible = True
        
        Me.Caption = cmbUsername.Text
        
        strIncoming = ""
            TreeView1.Nodes.Add , , "Online", "Buddies", "down"
            TreeView1.Nodes.Add , , "Offline", "Offline", "right"
            TreeView2.Nodes.Add , , "Buddies", "Buddies"
            
            TreeView1.Nodes.Item(1).Expanded = True
            TreeView2.Nodes.Item(1).Expanded = True
            
            TreeView1.Nodes.Item(1).Bold = True
            TreeView2.Nodes.Item(1).Bold = True
            
            TreeView1.Nodes.Item(2).ForeColor = vbButtonShadow
            
                        
            Winsock1.SendData ".updateStatus" & " " & "1" & " " & LCase(Me.Caption)
            WaitFor (".statusUpdate")
                       
            Winsock1.SendData ".getonlinebuddies" & " " & LCase(Me.Caption)
            WaitFor (".showonline")
End Sub

Private Function Show_Online_buddies(buddies As String)
    Dim i As Long, oncount As Integer
    Dim status As Integer
    Dim n As String
    oncount = 0
    For i = 3 To Words(strIncoming)
        status = Right(Word(buddies, i), 1)
        If status = 2 Then
            n = "Offline"
        Else
            n = "Online"
        End If
        TreeView1.Nodes.Add n, tvwChild, Left(Word(buddies, i), Len(Word(buddies, i)) - 1), Left(Word(buddies, i), Len(Word(buddies, i)) - 1), status, status
        TreeView2.Nodes.Add "Buddies", tvwChild, Left(Word(buddies, i), Len(Word(buddies, i)) - 1), Left(Word(buddies, i), Len(Word(buddies, i)) - 1)
    Next i
    Call Online_Offline_Text
End Function

Private Function status_update(buddy As String, status As Integer)
    Dim tn As Node
    Dim n As String
    Dim frmNum As Integer

        If status = 2 Then
            n = "Offline"
        Else
            n = "Online"
        End If
    For Each tn In TreeView1.Nodes
        If LCase(tn.Key) = LCase(buddy) Then
            'TreeView1.Nodes.Remove tn.Key
            'TreeView1.Nodes.Add n, tvwChild, buddy, buddy, status, status
            If Word(TreeView1.Nodes(tn.Key).Parent, 1) <> "Buddies" Then
                If status = 1 Then
                    Call PlaySound("sounds\dooropen.wav")
                    runlog.SelStart = Len(runlog.Text)
                    runlog.SelColor = vbBlue
                    runlog.SelText = buddy & " has signed on (" & Time & ")" & vbCrLf
                End If
                If status <> 2 Then
                    frmNum = GetFormNumber(LCase(buddy))
                    If frmNum <> 0 Then
                        Forms(frmNum).cmdSend.Enabled = True
                        Forms(frmNum).typemsg.Enabled = True
                        Forms(frmNum).tbFonts.Enabled = True
                        Forms(frmNum).showmsg.SelStart = Len(Forms(frmNum).showmsg.Text)
                        Forms(frmNum).showmsg.SelColor = vbBlue
                        Forms(frmNum).showmsg.SelText = buddy & " has signed on (" & Time & ")." & vbCrLf
                    End If
                End If
            End If
            If Word(TreeView1.Nodes(tn.Key).Parent, 1) <> "Offline" Then
                If status = 2 Then
                    Call PlaySound("sounds\doorslam.wav")
                    runlog.SelStart = Len(runlog.Text)
                    runlog.SelColor = vbRed
                    runlog.SelText = buddy & " has signed off (" & Time & ")" & vbCrLf
                    frmNum = GetFormNumber(LCase(buddy))
                    If frmNum <> 0 Then
                        Forms(frmNum).cmdSend.Enabled = False
                        Forms(frmNum).typemsg.Enabled = False
                        Forms(frmNum).tbFonts.Enabled = False
                        Forms(frmNum).showmsg.SelStart = Len(Forms(frmNum).showmsg.Text)
                        Forms(frmNum).showmsg.SelColor = vbRed
                        Forms(frmNum).showmsg.SelText = buddy & " has signed off (" & Time & ")." & vbCrLf
                    End If
                End If
            End If
            TreeView1.Nodes.Remove tn.Key
            TreeView1.Nodes.Add n, tvwChild, buddy, buddy, status, status
            Call Online_Offline_Text
            Exit For
        End If
    Next
End Function

Private Function Online_Offline_Text()
Dim tn As Node
Dim oncount
Dim offcount
oncount = 0
offcount = 0
TreeView1.Nodes.Item(1).Selected = True
    oncount = TreeView1.SelectedItem.Children
TreeView1.Nodes.Item(2).Selected = True
    offcount = TreeView1.SelectedItem.Children

    TreeView1.Nodes.Item(1).Text = "Buddies (" & oncount & "/" & oncount + offcount & ")"
    TreeView1.Nodes.Item(2).Text = "Offline (" & offcount & "/" & oncount + offcount & ")"
End Function

Private Sub mnuFileLogOut_Click()
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Me.BorderStyle = 1
    Me.Caption = "Logon"
    Me.Width = 2700
    Me.Height = 6000
    MainFrame.Visible = False
    LoginFrame.Visible = True
    TreeView1.Nodes.Clear
    TreeView2.Nodes.Clear
    mnuFile.Visible = False
    txtPassword.Text = ""
    mnuStatusAway.Checked = False
    mnuStatusOnline.Checked = True
Dim i As Integer
    For i = 0 To Forms.Count - 1
    If Forms.Count <> 1 Then
        Unload Forms(1)
    End If
    Next i
End Sub

Private Sub mnuFileClose_Click()
    End
End Sub

Private Sub mnuStatusOnline_Click()
    If mnuStatusOnline.Checked = False Then
        mnuStatusOnline.Checked = True
        mnuStatusAway.Checked = False
        lblStatus.Caption = "Online"
        Winsock1.SendData ".updateStatus" & " " & "1" & " " & LCase(Me.Caption)
        WaitFor (".statusUpdate")
    End If
End Sub

Private Sub mnuStatusAway_Click()
    If mnuStatusAway.Checked = False Then
        mnuStatusAway.Checked = True
        mnuStatusOnline.Checked = False
        lblStatus.Caption = "Away"
        Winsock1.SendData ".updateStatus" & " " & "3" & " " & LCase(Me.Caption)
        WaitFor (".statusUpdate")
    End If
End Sub

Sub WaitFor(ResponseCode As String)
    Start = 0
    tmrTimeout.Enabled = True
    While Len(strIncoming) = 0
        DoEvents
        If Start > 20 Then
            MsgBox "Service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
            Call mnuFileLogOut_Click
        End If
    Wend
    Start = 0
    While Word(strIncoming, 1) <> ResponseCode
        DoEvents
        If Start > 20 Then
           MsgBox "Service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + strIncoming, 64, MsgTitle
           Exit Sub
           Call mnuFileLogOut_Click
        End If
    Wend
    strIncoming = ""
    tmrTimeout.Enabled = False
End Sub

Private Sub tmrTimeout_Timer()
    Start = Start + 1
End Sub
