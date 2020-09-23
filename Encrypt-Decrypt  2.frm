VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Power Encrypter 1.0"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "Encrypt-Decrypt  2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6600
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RTB5 
      Height          =   1695
      Left            =   7800
      TabIndex        =   39
      Top             =   3840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2990
      _Version        =   393217
      TextRTF         =   $"Encrypt-Decrypt  2.frx":1A7A
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Encrypt-Decrypt  2.frx":1B43
      Left            =   7800
      List            =   "Encrypt-Decrypt  2.frx":1B83
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Encrypt-Decrypt  2.frx":1BC3
      Left            =   9000
      List            =   "Encrypt-Decrypt  2.frx":1C15
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Encrypt-Decrypt  2.frx":1C67
      Left            =   7800
      List            =   "Encrypt-Decrypt  2.frx":1CB9
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox RTB4 
      Height          =   2655
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.ListBox List1 
      Height          =   7275
      ItemData        =   "Encrypt-Decrypt  2.frx":1D0B
      Left            =   7200
      List            =   "Encrypt-Decrypt  2.frx":28CA
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "&Message"
      TabPicture(0)   =   "Encrypt-Decrypt  2.frx":3489
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Encrypt"
      TabPicture(1)   =   "Encrypt-Decrypt  2.frx":34A5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Command5"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Decrypt"
      TabPicture(2)   =   "Encrypt-Decrypt  2.frx":34C1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(2)=   "Command4"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton Command5 
         Caption         =   "Encrypt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         TabIndex        =   34
         Top             =   5520
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Decrypt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         TabIndex        =   33
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Frame Frame5 
         Caption         =   "Security Information:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   5055
         Begin VB.TextBox deKey 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            TabIndex        =   12
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox deRID 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "_"
            TabIndex        =   8
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox deSID 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "_"
            TabIndex        =   9
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox dePassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "_"
            TabIndex        =   10
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox deSWord 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "_"
            TabIndex        =   11
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label Label10 
            Caption         =   "Key:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   32
            Top             =   1725
            Width           =   1935
         End
         Begin VB.Label Label9 
            Caption         =   "Receiver's ID:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   31
            Top             =   285
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Sender's ID:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   30
            Top             =   645
            Width           =   1935
         End
         Begin VB.Label Label7 
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   29
            Top             =   1005
            Width           =   1935
         End
         Begin VB.Label Label6 
            Caption         =   "Secret Word:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   28
            Top             =   1365
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Security Information:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   5055
         Begin VB.TextBox msgSWord 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "_"
            TabIndex        =   4
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox msgPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "_"
            TabIndex        =   3
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox msgRID 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "_"
            TabIndex        =   2
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox msgSID 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "_"
            TabIndex        =   1
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label5 
            Caption         =   "Secret Word:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   26
            Top             =   1365
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   25
            Top             =   1005
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "Receiver's ID:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   24
            Top             =   645
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Sender's ID:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   285
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Decrypted Message:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         TabIndex        =   16
         Top             =   2640
         Width           =   5055
         Begin VB.TextBox RTB3 
            Height          =   2415
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   13
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Encrypted Message:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   5055
         Begin VB.TextBox RTB2 
            Height          =   4215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   7
            Top             =   600
            Width           =   4815
         End
         Begin VB.TextBox enKey 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label Label1 
            Caption         =   "Key:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   21
            Top             =   285
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Type your message here:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   5055
         Begin VB.TextBox RTB1 
            Height          =   3255
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   5
            Top             =   240
            Width           =   4815
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If SSTab1.Tab = 0 Then
    CD1.DialogTitle = "Open Messages"
    CD1.Filter = "Messages Files(*.msf)|*.msf|Text Files(*.txt)|*.txt|Word Files(*.RTF)|*.RTF"
ElseIf SSTab1.Tab = 1 Then
    CD1.DialogTitle = "Open Encrypted Messages"
    CD1.Filter = "Encrypted Messages(*.PEN)|*.PEN"
Else
    Exit Sub
End If
CD1.FileName = ""
CD1.Flags = &H4&
CD1.ShowOpen
If CD1.FileName = "" Then Exit Sub
RTB5.LoadFile CD1.FileName
If SSTab1.Tab = 0 Then
    RTB1.Text = RTB5.Text
ElseIf SSTab1.Tab = 1 Then
    RTB2.Text = RTB5.Text
End If

End Sub

Private Sub Command2_Click()
If SSTab1.Tab = 0 Or SSTab1.Tab = 2 Then
    CD1.DialogTitle = "Save Message"
    CD1.Filter = "Message Files(*.msf)|*.msf"
    a = ".msf"
ElseIf SSTab1.Tab = 1 Then
    CD1.DialogTitle = "Save Encrypted Message"
    CD1.Filter = "Encrypted Messages(*.PEN)|*.PEN"
    a = ".PEN"
End If
CD1.FileName = ""
CD1.Flags = &H4&
CD1.ShowSave
If CD1.FileName = "" Then Exit Sub
If SSTab1.Tab = 0 Then RTB5.Text = RTB1.Text
If SSTab1.Tab = 1 Then RTB5.Text = RTB2.Text
If SSTab1.Tab = 2 Then RTB5.Text = RTB3.Text
RTB5.SaveFile CD1.FileName

End Sub

Private Sub Command4_Click()
If InStr(RTB1.Text, Chr(1)) Then
    SSTab1.Tab = 1
    RTB2.SetFocus
    RTB2.SelStart = InStr(RTB2.Text, Chr(1)) - 1
    RTB2.SelLength = 1
    MsgBox ("Character: ' " & Chr(1) & " '"), vbExclamation, "Invalid Character"
    Exit Sub
End If
If Len(deKey.Text) <> 19 And Len(deKey.Text) <> 14 Then
        MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
        Exit Sub
End If
For qw = 1 To Len(deKey.Text)
    If Mid(deKey.Text, qw, 1) = "." Then
        numdot = numdot + 1
    End If
Next qw
If (Len(deKey.Text) = 19 And numdot <> 3) Or (Len(deKey.Text) = 14 And numdot <> 2) Then
        MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
        Exit Sub
End If
For qw = 5 To Len(deKey.Text) Step 5
    If Mid(deKey.Text, qw, 1) <> "." Then
        MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
        Exit Sub
    End If
Next qw
numdot = 0
RTB4.Text = "QWERTYUIOPASDFGHJKLZXCVBNM"
For qw = 1 To Len(deKey.Text) Step 5
    If InStr(RTB4.Text, Mid(deKey.Text, qw, 1)) = 0 Then
        MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
        Exit Sub
    End If
    RTB4.SelStart = InStr(RTB4.Text, Mid(deKey.Text, qw, 1)) - 1
    RTB4.SelLength = 1
    RTB4.SelText = ""
Next qw
RTB4.Text = ""
qw = 0
qw1 = "1234567890"
For qw = 1 To Len(deKey.Text)
    If InStr(qw1, Mid(deKey.Text, qw, 1)) Then
        qw3 = qw3 + 1
    End If
Next qw
If Len(deKey.Text) = 19 Then
    If qw3 <> 12 Then
            MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
            Exit Sub
    End If
ElseIf Len(deKey.Text) = 14 Then
    If qw3 <> 9 Then
        MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
        Exit Sub
    End If
End If
qw = 0
qw1 = ""
qw3 = 0
If RTB2.Text = "" Then
    MsgBox ("Empty 'Encryption' Textbox"), vbExclamation, "Invalid Value"
    Exit Sub
End If
If Len(deKey.Text) = 19 Then
    a = Right(deKey.Text, 3)
30
    List1.ListIndex = a
    RTB4.Text = String(Len(RTB2.Text), Chr(1))
    Do
        e = e + 1
        d = Int(GenRnd("") * Len(RTB2.Text)) + 1
10
        If Mid(RTB4.Text, d, 1) <> Chr(1) Then
            d = d + 1
            If d > Len(RTB4.Text) Then d = 1
            GoTo 10
        End If
        RTB4.SelStart = d - 1
        RTB4.SelLength = 1
        RTB4.SelText = Mid(RTB2.Text, e, 1)
    Loop Until Len(RTB4.Text) = e
    If Mid(RTB4.Text, 1, InStr(RTB4.Text, " ") - 1) <> deSID.Text Then
        RTB3.Text = ""
        MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
        Exit Sub
    End If
    For newvar1 = InStr(RTB4.Text, " ") + 1 To Len(RTB4.Text)
        If Mid(RTB4.Text, newvar1, 1) = " " Then Exit For
    Next newvar1
    If Mid(RTB4.Text, InStr(RTB4.Text, " ") + 1, newvar1 - (InStr(RTB4.Text, " ") + 1)) <> deRID.Text Then
        RTB3.Text = ""
        MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
        Exit Sub
    End If
    For newvar2 = Len(RTB4.Text) To 1 Step -1
        If Mid(RTB4.Text, newvar2, 1) = " " Then Exit For
    Next newvar2
    If Right(RTB4.Text, Len(RTB4.Text) - newvar2) <> deSWord.Text Then
        RTB3.Text = ""
        MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
        Exit Sub
    End If
    For newvar3 = newvar2 - 1 To 1 Step -1
        If Mid(RTB4.Text, newvar3, 1) = " " Then
            newvar3 = newvar3 + 1
            Exit For
        End If
    Next newvar3
    If Mid(RTB4.Text, newvar3, newvar2 - newvar3) <> dePassword.Text Then
        RTB3.Text = ""
        MsgBox ("One of the information is incorrect."), vbExclamation, "Invalid Information"
        Exit Sub
    End If
    RTB3.Text = Mid(RTB4.Text, newvar1 + 1, (newvar3 - 2) - (newvar1 + 1) + 1)
ElseIf Len(deKey.Text) = 14 Then
    If Mid(deKey.Text, 2, 3) / 2 = Int(Mid(deKey.Text, 2, 3) / 2) Then
        a = Mid(deKey.Text, 7, 3)
    Else
        a = Mid(deKey.Text, 12, 3)
    End If
    GoTo 30
End If
    
End Sub

Private Sub Form_Load()
List1.ListIndex = 0

End Sub

Private Sub RTB1_GotFocus()
For Each a In Controls
    On Error GoTo 10
    a.TabStop = False
10
    Resume 20
20
Next a

End Sub

Private Sub RTB1_LostFocus()
For Each a In Controls
    On Error GoTo 10
    a.TabStop = True
10
    Resume 20
20
Next a

End Sub

Private Sub Command5_Click()
        If msgSID = "" Then
            MsgBox ("Sender's ID is empty."), vbExclamation, "Empty Field"
            SSTab1.Tab = 0
            msgSID.SetFocus
            Exit Sub
        End If
        If msgRID = "" Then
            MsgBox ("Receiver's ID is empty."), vbExclamation, "Empty Field"
            SSTab1.Tab = 0
            msgRID.SetFocus
            Exit Sub
        End If
        If msgPassword = "" Then
            MsgBox ("Password is empty."), vbExclamation, "Empty Field"
            SSTab1.Tab = 0
            msgPassword.SetFocus
            Exit Sub
        End If
        If msgSWord = "" Then
            MsgBox ("Secrect Word is empty."), vbExclamation, "Empty Field"
            SSTab1.Tab = 0
            msgSWord.SetFocus
            Exit Sub
        End If
        If InStr(msgSWord.Text, " ") Or InStr(msgSWord.Text, Chr(1)) Then
            SSTab1.Tab = 0
            msgSWord.SetFocus
            msgSWord.SelStart = InStr(msgSWord.Text, " ") - 1
            msgSWord.SelLength = 1
            MsgBox ("Character: ' " & msgSWord.SelText & " '"), vbExclamation, "Invalid Character"
            Exit Sub
        ElseIf InStr(msgPassword.Text, " ") Or InStr(msgPassword.Text, Chr(1)) Then
            SSTab1.Tab = 0
            msgPassword.SetFocus
            msgPassword.SelStart = InStr(msgPassword.Text, " ") - 1
            msgPassword.SelLength = 1
            MsgBox ("Character: ' " & msgPassword.SelText & " '"), vbExclamation, "Invalid Character"
            Exit Sub
        ElseIf InStr(msgRID.Text, " ") Or InStr(msgRID.Text, Chr(1)) Then
            SSTab1.Tab = 0
            msgRID.SetFocus
            msgRID.SelStart = InStr(msgRID.Text, " ") - 1
            msgRID.SelLength = 1
            MsgBox ("Character: ' " & msgRID.SelText & " '"), vbExclamation, "Invalid Character"
            Exit Sub
        ElseIf InStr(msgSID.Text, " ") Or InStr(msgSID.Text, Chr(1)) Then
            SSTab1.Tab = 0
            msgSID.SetFocus
            msgSID.SelStart = InStr(msgSID.Text, " ") - 1
            msgSID.SelLength = 1
            MsgBox ("Character: ' " & msgSID.SelText & " '"), vbExclamation, "Invalid Character"
            Exit Sub
        End If
        If InStr(RTB1.Text, Chr(1)) Then
            SSTab1.Tab = 0
            RTB1.SetFocus
            RTB1.SelStart = InStr(RTB1.Text, Chr(1)) - 1
            RTB1.SelLength = 1
            MsgBox ("Character: ' " & Chr(1) & " '"), vbExclamation, "Invalid Character"
            Exit Sub
        End If
        Randomize
        a = Int(Rnd * 999)
        List1.ListIndex = a
        RTB4.Text = msgSID.Text & " " & msgRID.Text & " " & RTB1.Text & " " & msgPassword.Text & " " & msgSWord.Text
        If Len(RTB4.Text) / 2 = Int(Len(RTB4.Text) / 2) Then
            e = 4
        Else
            e = 3
        End If
        For d = 1 To e
            f = Int(Rnd * (Combo1.ListCount - 1)) + 1
            g = g & Combo1.List(f)
            Combo1.RemoveItem f
            If d > 1 Then
                If e = 3 Then
                    If Mid(g, 2, 3) / 2 = Int(Mid(g, 2, 3) / 2) And d = 2 Then
                        g1 = a & ""
                        g = g & String$(3 - Len(g1), "0") & g1
                        GoTo 10
                    ElseIf Mid(g, 2, 3) / 2 <> Int(Mid(g, 2, 3) / 2) And d = 3 Then
                        g1 = a & ""
                        g = g & String$(3 - Len(g1), "0") & g1
                        GoTo 10
                    End If
                ElseIf e = 4 Then
                    If d = 4 Then
                        g1 = a & ""
                        g = g & String$(3 - Len(g1), "0") & g1
                        GoTo 10
                    End If
                End If
            End If
            For f1 = 1 To 3
                f = Int(Rnd * (Combo3.ListCount - 1)) + 1
                g = g & Combo3.List(f)
            Next f1
10
            If d <> e Then
                g = g & "."
            End If
        Next d
        Combo1.Clear
        For d = 0 To 25
            Combo1.AddItem Combo2.List(d)
        Next d
        enKey.Text = g
        Do
            numstring = numstring + 1
            b = Int(Val(GenRnd("")) * Len(RTB4.Text)) + 1
20
            If Mid(RTB4.Text, b, 1) = Chr(1) Then
                b = b + 1
                If b > Len(RTB4.Text) Then b = 1
                GoTo 20
            End If
            c = c & Mid(RTB4.Text, b, 1)
            RTB4.SelStart = b - 1
            RTB4.SelLength = 1
            RTB4.SelText = Chr(1)
        Loop While Len(RTB4.Text) <> numstring
        RTB2.Text = c
        
End Sub

