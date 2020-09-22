VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prisco Number Generator"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
   Icon            =   "Prisco Number Generator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FRAgen 
      Caption         =   "Generated Number"
      Height          =   690
      Left            =   75
      TabIndex        =   10
      Top             =   3150
      Width           =   9165
      Begin VB.Label LBLgen 
         Height          =   315
         Left            =   75
         TabIndex        =   11
         Top             =   300
         Visible         =   0   'False
         Width           =   9015
      End
   End
   Begin VB.CommandButton CMDgen 
      Caption         =   "Generate Number"
      Height          =   540
      Left            =   3750
      TabIndex        =   9
      ToolTipText     =   "Generate a number"
      Top             =   2325
      Width           =   1665
   End
   Begin MSComCtl2.UpDown UAD 
      Height          =   315
      Left            =   6375
      TabIndex        =   5
      Top             =   1095
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   5
      OrigLeft        =   6750
      OrigTop         =   1800
      OrigRight       =   7005
      OrigBottom      =   2115
      Max             =   100
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox CHKoptions 
      Caption         =   "No ""Zero"" On First Letter"
      Height          =   390
      Index           =   2
      Left            =   975
      TabIndex        =   3
      ToolTipText     =   "Generated number won't starts from zero"
      Top             =   1500
      Width           =   3165
   End
   Begin VB.Frame FRAcontrol 
      Caption         =   "Process Options"
      Height          =   1290
      Left            =   825
      TabIndex        =   1
      Top             =   750
      Width           =   7440
      Begin VB.ComboBox CMO 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Prisco Number Generator.frx":0442
         Left            =   6525
         List            =   "Prisco Number Generator.frx":044C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   825
         Width           =   765
      End
      Begin VB.CheckBox CHKoptions 
         Caption         =   "Only Use This Type Of Number :"
         Height          =   240
         Index           =   3
         Left            =   3900
         TabIndex        =   7
         ToolTipText     =   "Use only Odd or Even numbers"
         Top             =   825
         Width           =   2640
      End
      Begin VB.TextBox TXTmaxlen 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   5175
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "5"
         Top             =   350
         Width           =   390
      End
      Begin VB.CheckBox CHKoptions 
         Caption         =   "Automatically Load Number To Clipboard"
         Height          =   390
         Index           =   1
         Left            =   150
         TabIndex        =   2
         ToolTipText     =   "Writes automatically generated number into system clipboard"
         Top             =   300
         Width           =   3240
      End
      Begin VB.Label LBLmaxlen 
         AutoSize        =   -1  'True
         Caption         =   "Number Lenght :"
         Height          =   195
         Left            =   3900
         TabIndex        =   6
         Top             =   375
         Width           =   1185
      End
   End
   Begin VB.Label LBLwel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Welcome To Cybersoft® Prsico© Number Generator"
      Height          =   195
      Left            =   2775
      TabIndex        =   0
      Top             =   225
      Width           =   3735
   End
End
Attribute VB_Name = "FRMmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit
Private Sub CHKoptions_Click(Index As Integer)
If Index = 3 Then
    If CHKoptions(3).Value = 0 Then
        CMO.Enabled = False
        CMO.Refresh
        CHKoptions(2).Enabled = True
    ElseIf CHKoptions(3).Value = 1 Then
        CMO.Enabled = True
        If CMO.Text = "Odd" Then
            CHKoptions(2).Enabled = False
        ElseIf CMO.Text = "Even" Then
            CHKoptions(2).Enabled = True
        End If
    End If
End If
End Sub
Private Sub CMDgen_Click()
Dim intloopcon As Integer
Dim intresnum As Integer
Dim intconres As Integer
LBLgen.Visible = False
If Not LBLgen.Caption = "" Then
    LBLgen.Caption = ""
End If
Randomize Timer
intconres = 1
For intloopcon = 1 To TXTmaxlen.Text
    If CHKoptions(2).Value = 0 Then
        If CHKoptions(3).Value = 0 Then
            intresnum = Int((10 - 0 + 0) * Rnd)
            LBLgen.Caption = LBLgen.Caption & intresnum
        ElseIf CHKoptions(3).Value = 1 Then
            If CMO.Text = "Even" Then
j1:                intresnum = Int((10 - 0 + 0) * Rnd)
                If intresnum Mod 2 = 0 Then
                    LBLgen.Caption = LBLgen.Caption & intresnum
                Else
                    GoTo j1
                End If
            ElseIf CMO.Text = "Odd" Then
j2:                intresnum = Int((10 - 0 + 0) * Rnd)
                    If Not intresnum Mod 2 = 0 Then
                        LBLgen.Caption = LBLgen.Caption & intresnum
                    Else
                        GoTo j2
                    End If
            End If
        End If
    Else
        If CHKoptions(3).Value = 0 Then
j5:            intresnum = Int((10 - 0 + 0) * Rnd)
            If Not intresnum = 0 Then
j6:                LBLgen.Caption = LBLgen.Caption & intresnum
            Else
                If intconres = 1 Then
                    GoTo j5
                Else
                    GoTo j6
                End If
            End If
        ElseIf CHKoptions(3).Value = 1 Then
            If CMO.Text = "Even" Then
j3:                intresnum = Int((10 - 0 + 0) * Rnd)
                If intresnum Mod 2 = 0 And Not intresnum = 0 Then
                    LBLgen.Caption = LBLgen.Caption & intresnum
                Else
                    GoTo j3
                End If
            ElseIf CMO.Text = "Odd" Then
j4:                intresnum = Int((10 - 0 + 0) * Rnd)
                    If Not intresnum Mod 2 = 0 Then
                        LBLgen.Caption = LBLgen.Caption & intresnum
                    Else
                        GoTo j4
                    End If
            End If
        End If
    End If
intconres = intconres + 1
Next intloopcon
LBLgen.Visible = True
If CHKoptions(1).Value = 1 Then
    Clipboard.Clear
    Clipboard.SetText (LBLgen.Caption)
End If
End Sub
Private Sub CMO_Click()
If CMO.Text = "Odd" Then
    CHKoptions(2).Enabled = False
Else
    CHKoptions(2).Enabled = True
End If
End Sub
Private Sub Form_Load()
LBLwel.Left = (FRMmain.Width - LBLwel.Width) / 2
CMO.Text = "Even"
End Sub
Private Sub TXTmaxlen_KeyPress(KeyAscii As Integer)
Dim intmsgbox As Integer
Dim inttemp As Integer
Dim inttemp2 As Integer
If KeyAscii = 13 Then
    On Local Error GoTo errhandler
    inttemp2 = TXTmaxlen.Text
    UAD.SetFocus
    If TXTmaxlen.Text > 100 Then
        intmsgbox = MsgBox("Maximum Number Lenght Is 100", vbOKOnly + vbInformation, "Out of range..!!")
        TXTmaxlen.Text = 100
        UAD.SetFocus
    ElseIf TXTmaxlen.Text <= 0 Then
        intmsgbox = MsgBox("Minimum Number Lenght Is 1", vbOKOnly + vbInformation, "Out of range..!!")
        TXTmaxlen.Text = 1
        UAD.SetFocus
    End If
End If
If 1 = 2 Then
errhandler: intmsgbox = MsgBox("You've entered a illegal number format...!!!", vbOKOnly + vbCritical, "Illegal Number Format")
TXTmaxlen.Text = 5
UAD.SetFocus
End If
End Sub
Private Sub UAD_DownClick()
If TXTmaxlen.Text > 1 Then
    UAD.SetFocus
    TXTmaxlen.Text = TXTmaxlen.Text - 1
Else
    Exit Sub
End If
End Sub
Private Sub UAD_UpClick()
If TXTmaxlen.Text < 100 Then
    UAD.SetFocus
    TXTmaxlen.Text = TXTmaxlen.Text + 1
Else
    Exit Sub
End If
End Sub



'*****************************************************************************************************

'I finished this little baby after 3 houres
'It's fast,reiable,nice...and much more...
'8:00 PM    7/21/2002

'Arash Yadegarnia
