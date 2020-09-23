VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Jays Combo Box"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   315
      Left            =   2220
      TabIndex        =   9
      Top             =   780
      Width           =   675
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Bold"
      Height          =   195
      Left            =   1140
      TabIndex        =   6
      Top             =   840
      Width           =   675
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Normal"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   840
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   420
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Color"
      Height          =   315
      Left            =   2220
      TabIndex        =   3
      Top             =   420
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Text            =   "0"
      Top             =   60
      Width           =   2715
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Item"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   975
   End
   Begin Project1.JaysCombo JaysCombo1 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   1320
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   556
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "jay@alphamedia.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1635
      TabIndex        =   8
      Top             =   4200
      Width           =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0E0FF&
      X1              =   -180
      X2              =   4920
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   -180
      X2              =   4920
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1035
      Left            =   3000
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmJaysCombo.frx":0000
      Height          =   1275
      Left            =   195
      TabIndex        =   7
      Top             =   3240
      Width           =   3420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrColor As OLE_COLOR

Private Sub Command1_Click()
    JaysCombo1.AddItem Text1, Option2, CurrColor
End Sub

Private Sub Command2_Click()
    CommonDialog1.ShowColor
    CurrColor = CommonDialog1.Color
    Shape1.FillColor = CommonDialog1.Color
End Sub

Private Sub Command3_Click()
    JaysCombo1.Clear
End Sub

Private Sub Command4_Click()
    If JaysCombo1.ListCount = 0 Then Exit Sub
    JaysCombo1.RemoveItem Val(InputBox("Enter Combo Index to Remove (0 - " & JaysCombo1.ListCount - 1 & "):"))
End Sub

Private Sub Form_Load()
    CurrColor = vblack
End Sub

