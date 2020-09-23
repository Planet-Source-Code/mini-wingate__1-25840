VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWingate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wingate"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmWingate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   3015
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame4 
      Caption         =   "Control"
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   2145
      Width           =   3015
      Begin VB.Frame Frame5 
         Caption         =   "Status"
         Height          =   615
         Left            =   1930
         TabIndex        =   13
         Top             =   120
         Width           =   975
         Begin VB.Shape ShapeLead 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            Shape           =   2  'Oval
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton CmdScan 
         Caption         =   "Scan"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Frame Frame3 
         Caption         =   "Port"
         Height          =   615
         Left            =   960
         TabIndex        =   9
         Top             =   120
         Width           =   975
         Begin VB.TextBox txtPort 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Text            =   "80"
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Wingate List"
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   3015
      Begin VB.ListBox lstIP 
         Height          =   1230
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proxy Found"
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   2985
      Width           =   3015
      Begin VB.ListBox lstWin 
         Height          =   1035
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
   Begin MSWinsockLib.Winsock Sck1 
      Left            =   3240
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFile 
      Height          =   255
      Left            =   2650
      TabIndex        =   2
      Top             =   10
      Width           =   315
   End
   Begin VB.TextBox TxtFile 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   2645
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   20
      TabIndex        =   0
      Top             =   25
      Width           =   375
   End
End
Attribute VB_Name = "frmWingate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

            'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx (:x)
            '  Are you really able to make this better?
            '          Make this program better
            '                if you can!
            '        and be nice SEND me a copie
            '                    OR
            '            post'it back to PSC
            '
            '            LOVE YOU ALL Vicky
            'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx (:x)
            
Option Explicit
Dim GoIP As String
Dim X1, X2, i As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
    CD1.Filter = "All Files (*.*) | *.*"
    CD1.ShowOpen
    TxtFile.Text = CD1.FileName
End Sub

Private Sub CmdScan_Click()
On Error Resume Next

    CmdScan.Enabled = False
    TxtFile.Enabled = False

    lstIP.Clear  'clear the list
    X1 = -1
    X2 = -1
      
      
      Open TxtFile.Text For Input As #1  'open the file
      Do Until EOF(1) = True  'go until the end of the file
        Input #1, GoIP
        X1 = X1 + 1
          If GoIP = "" Then
          Else
            lstIP.AddItem GoIP, X1  'add all the lines into the lstip listbox
          End If
      Loop
      
      Close #1 'close the file


    Call ScanProxy

    TxtFile.Enabled = True
    CmdScan.Enabled = True
End Sub

Private Sub Form_Load()
    TxtFile = App.Path & "\Proxy.txt"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Sck1.Close
    Unload Me
End
End Sub

Private Sub Sck1_Connect()
    ShapeLead.BackColor = &HFF00&
End Sub

Sub ScanProxy()
'This is the scan
      X2 = 0
    PB1.Max = lstIP.ListCount - 1
    PB1.Value = 0
    
    For i = 1 To lstIP.ListCount - 1
         Debug.Print X2 & "  :  " & lstIP.List(X2)
         PB1.Value = PB1.Value + 1
         Sck1.Connect lstIP.List(X2), txtPort
    Do
         Select Case Sck1.State
             Case 7, 8, 9, 0
                 Exit Do
         End Select
         DoEvents
    Loop
    
    If Sck1.State = 7 Then
       lstWin.AddItem lstIP.List(X2)
       Beep
       ShapeLead.BackColor = &HFF00&
    End If
       X2 = X2 + 1
       Sck1.Close
       ShapeLead.BackColor = &HFF&
        Next
        
    MsgBox "Scan Done"
End Sub


