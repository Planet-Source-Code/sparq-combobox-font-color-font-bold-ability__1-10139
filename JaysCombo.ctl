VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl JaysCombo 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   2325
   ScaleWidth      =   2295
   Begin MSComCtl2.FlatScrollBar vBar 
      Height          =   1395
      Left            =   1980
      TabIndex        =   10
      Top             =   720
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   2461
      _Version        =   393216
      Appearance      =   0
      Orientation     =   8323072
   End
   Begin VB.PictureBox btnUp 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2070
      Picture         =   "JaysCombo.ctx":0000
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   30
      Width           =   195
   End
   Begin VB.PictureBox btnDown 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2070
      Picture         =   "JaysCombo.ctx":005E
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   30
      Width           =   195
   End
   Begin VB.TextBox Text1 
      Height          =   310
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   6
      Left            =   60
      TabIndex        =   9
      Top             =   1800
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   5
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   7
      Top             =   1320
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   840
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   1860
   End
   Begin VB.Shape ListBox 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   0
      Top             =   300
      Width           =   2235
   End
End
Attribute VB_Name = "JaysCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim List(999) As String
Dim ListColor(999) As OLE_COLOR
Dim ListBold(999) As Boolean
Dim LabelColor(6) As OLE_COLOR

Dim mTopIndex As Integer
Dim mListCount As Integer
Dim ListState As Integer

Private Sub btnUp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btnUp.Visible = False
    btnDown.Visible = True
    If ListState = 0 Then
        ListState = 1
        DropList
    Else
        ListState = 0
        RemoveList
    End If
End Sub

Public Function Clear()
    For x = 0 To 999
        ListBold(x) = False
        List(x) = ""
        ListColor(x) = vbBlack
    Next x
    Setlabels (0)
End Function

Private Function DropList()
    vBar.Visible = Not (ListCount <= 7)
    If vBar.Visible = True Then
        vBar.Max = mListCount - 7
        vBar.Min = 0
    End If
    Drop = True
    UserControl.Height = 310 + 1530
    ListBox.Top = Text1.Top + (Text1.Height)
    ListBox.Left = Text1.Left
    ListBox.Width = Text1.Width
    ListBox.Height = Height - ((Text1.Top + Text1.Height))
    vBar.Left = Width - vBar.Width
    vBar.Top = ListBox.Top + 15
    vBar.Height = ListBox.Height - 30
    Setlabels (mTopIndex)
End Function

Private Function LoadLabels()
    Dim Top
    Top = 150
    For x = 0 To Label1.Count - 1
        Label1(x).Left = 60

        If vBar.Visible = True Then
            Label1(x).Width = Text1.Width - (290)
        Else
            Label1(x).Width = Text1.Width - 120
        End If
        Label1(x).Top = Top + (210 * (x + 1))
    Next x
End Function

Private Function RemoveList()
    UserControl.Height = 310
    ListState = 0
End Function

Private Sub btnUp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ResetLabels
End Sub

Private Sub btnUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btnUp.Visible = True
    btnDown.Visible = False
End Sub

Private Sub Label1_Click(Index As Integer)
    Text1.FontBold = Label1(Index).FontBold
    Text1.ForeColor = ListColor(mTopIndex + Index)
    Text1 = Right(Label1(Index), Len(Label1(Index)) - 1)
    RemoveList
End Sub

Private Sub ResetLabels()
    For x = 0 To Label1.Count - 1
            Label1(x).BackColor = vbWhite
            Label1(x).ForeColor = LabelColor(x)
    Next x
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    For x = 0 To Label1.Count - 1
        If x = Index Then
            Label1(x).ForeColor = vbWhite
            Label1(x).BackColor = &H800000
        Else
            Label1(x).BackColor = vbWhite
            Label1(x).ForeColor = LabelColor(x)
        End If
    Next x
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ResetLabels
End Sub

Private Sub UserControl_Initialize()
    Height = 310
    ListState = 0
    Setlabels (0)
    
End Sub

Private Sub UserControl_Resize()
    Text1.Height = 310
    Text1.Width = Width
    btnUp.Top = 30
    btnUp.Left = Text1.Width - (btnUp.Width + 30)
    btnDown.Left = btnUp.Left
End Sub

Public Function RemoveItem(Index As Integer)
    If Index >= mListCount Then Exit Function
    If mListCount = 0 Then Exit Function
    
    List(Index) = ""
    For x = Index To 999
        List(x) = List(x + 1)
        ListColor(x) = ListColor(x + 1)
        ListBold(x) = ListBold(x + 1)
        If List(x) = "" And List(x + 1) = "" Then Exit For
    Next x
    mListCount = mListCount - 1
    Setlabels (0)
End Function

Public Function AddItem(Text As String, Optional Bold As Boolean, Optional Color As OLE_COLOR)
  Dim x As Integer
  Dim FirstOpen As Integer
    If Color = 0 Then Color = vbBlack
    RemoveList
    x = 0
    Do Until List(x) = ""
        x = x + 1
    Loop
    List(x) = " " & Text
    ListColor(x) = Color
    ListBold(x) = Bold
    mListCount = x + 1
    Setlabels (0)
End Function

Private Function Setlabels(First As Integer)
    For x = 0 To 6
        If List(x + First) = "" Then
            Label1(x).Visible = False
        Else
            Label1(x) = List(x + First)
            LabelColor(x) = ListColor(x + First)
            Label1(x).ForeColor = LabelColor(x)
            Label1(x).Visible = True
        End If
        Label1(x).FontBold = ListBold(x + First)
    Next x
    mTopIndex = First
End Function

Public Property Get ListCount() As Integer
    ListCount = mListCount
End Property

Public Property Get Item(Index As Integer) As String
    Item = Right(List(Index), Len(List(Index)) - 1)
End Property

Public Property Get Text() As String
    Text = Text1
End Property

Private Sub VBar_Change()
    Setlabels (vBar.Value)
    ResetLabels
End Sub
