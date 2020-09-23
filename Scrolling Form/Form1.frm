VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame framescroll 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   3000
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      LargeChange     =   400
      Left            =   6120
      SmallChange     =   200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2520
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   400
      Left            =   1080
      SmallChange     =   200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4935
      Left            =   2520
      TabIndex        =   3
      Top             =   3120
      Width           =   3375
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   1440
         TabIndex        =   10
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   4680
         TabIndex        =   11
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Resize the form to see scroll bars"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vsval As Integer
Dim hsval As Integer


' this module displays scroll bars on the form when the form
' is resized below its maximum length
Private Sub Form_Resize()
    'the hdisp and vdisp variables calculate the difference between the
    ' current dimensions of the form and the dimenstions of the maximized form
    ' in order to calculate whether scroll bars should be displayed or not
    Dim hdisp As Integer
    Dim vdisp As Integer
    ' h and w variables store space required for displaying horizontal and
    ' vertical scroll bar respectively
    Dim h As Integer
    Dim w As Integer
    ' 12120 is the width of the maximized form. replace this value as required
    hdisp = Me.Width - 12120
    ' 8670 is the height of the maximized form. replace this value as required
    vdisp = Me.Height - 8670
    If hdisp >= 0 And vdisp >= 0 Then
        'if form is resized above the maximized dimensions then
        ' hide both the scrollbars
        VScroll1.Visible = False
        HScroll1.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        ' framescroll must be visible along with both the scroll bars
        ' framescroll is a frame which is shown between then ends of
        ' both the scroll bars to restrict user to click in that area
        ' because some control of the form may be visible in that area
        framescroll.Visible = False
        ' postion the controls to their original locations
        positioncontrols
    ElseIf hdisp >= 0 And vdisp < 0 Then
        ' if height of the resized form falls below the maximized height then
        ' display only the vertical scroll bar
        VScroll1.Visible = True
        HScroll1.Visible = False
        HScroll1.Value = 0
        ' make horizontal space because horizontal scroll bar will not
        ' be visible
        h = 0
        ' change the maximum value of the scroll bar so that scrolling is made
        ' only to the appropriate extent
        HScroll1.Max = -hdisp + VScroll1.Width
        framescroll.Visible = False
    ElseIf hdisp < 0 And vdisp >= 0 Then
        HScroll1.Visible = True
        VScroll1.Visible = False
        VScroll1.Value = 0
        w = 0
        VScroll1.Max = -vdisp + HScroll1.Height
        framescroll.Visible = False
    Else
        HScroll1.Max = -hdisp + VScroll1.Width
        VScroll1.Max = -vdisp + HScroll1.Height
        VScroll1.Visible = True
        HScroll1.Visible = True
        h = HScroll1.Height
        w = VScroll1.Width
        framescroll.Visible = True
    End If
    ' position the scroll bars according to the size of the form
    With VScroll1
        If .Visible Then
            .Top = Me.ScaleTop
            .Left = Me.ScaleWidth - .Width
            .Height = Abs(Me.ScaleHeight - h)
        End If
    End With
    With HScroll1
        If .Visible Then
            .Top = Me.ScaleHeight - .Height
            .Left = Me.ScaleLeft
            .Width = Abs(Me.ScaleWidth - w)
        End If
    End With
    With framescroll
        .Left = VScroll1.Left
        .Top = HScroll1.Top
    End With
End Sub
Private Sub positioncontrols()
    ' position the controls to their original locations
    Frame1.Left = 1320
    Frame1.Top = 120
    Frame2.Left = 2520
    Frame2.Top = 3000
End Sub
Private Sub VScroll1_Change()
    Dim inc As Integer
    ' inc stores the amount of scrolling
    inc = VScroll1.Value - vsval
    vsval = VScroll1.Value
    ' the controls are positioned according to the amount of scrolling
    ' it is better to place controls on frames so that only frame is
    ' required to move instead of moving each control
    ' i have used two frames here
    ' you can even use one frame for all the controls on the form
    ' to make it even easier
    Frame1.Top = Frame1.Top - inc
    Frame2.Top = Frame2.Top - inc
End Sub
Private Sub vScroll1_Scroll()
    VScroll1_Change
End Sub
Private Sub HScroll1_Change()
    Dim inc As Integer
    inc = HScroll1.Value - hsval
    hsval = HScroll1.Value
    Frame1.Left = Frame1.Left - inc
    Frame2.Left = Frame2.Left - inc
End Sub
Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub
