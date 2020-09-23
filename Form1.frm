VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HexSweeper"
   ClientHeight    =   2910
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   194
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   420
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   435
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   266
      TabIndex        =   0
      Top             =   1200
      Width           =   4050
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   3840
      TabIndex        =   4
      Top             =   30
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   30
      Width           =   540
   End
   Begin VB.Menu menuGame 
      Caption         =   "&Game"
      Begin VB.Menu menuNew 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu menuSep 
         Caption         =   "-"
      End
      Begin VB.Menu menuLevel 
         Caption         =   "&Beginner"
         Index           =   0
      End
      Begin VB.Menu menuLevel 
         Caption         =   "&Intermediate"
         Index           =   1
      End
      Begin VB.Menu menuLevel 
         Caption         =   "&Expert"
         Index           =   2
      End
      Begin VB.Menu menuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu menuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''HEXSWEEPER
'''By Paul Bahlawan April 26, 2003
'''
'''To do: -High scores
'''       -'Custom' level
'''       -'?' marks (maybe)

Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim Field(29, 15) As Integer
Dim Overlay(29, 15) As Integer
Dim Alive As Boolean
Dim LevelX As Integer, LevelY As Integer
Dim TotalMine As Integer
Dim Progress As Integer
Dim Flags As Integer
Dim ScoreType As Boolean
Dim StartTime As Single

Private Sub Form_Load()
Dim X As Integer, Y As Integer, Z As Integer
Randomize
Picture1.Top = -50  'Hide this off screen
ScoreType = True
menuLevel_Click 1   'Set level and start a new game
End Sub

Private Sub menuLevel_Click(Index As Integer)
menuLevel(0).Checked = False
menuLevel(1).Checked = False
menuLevel(2).Checked = False
menuLevel(Index).Checked = True
Select Case Index
    Case 0
        LevelX = 8
        LevelY = 8
        TotalMine = 8
    Case 1
        LevelX = 16
        LevelY = 16
        TotalMine = 35
    Case 2
        LevelX = 30
        LevelY = 16
        TotalMine = 75
End Select
'Set window size for selected level
Picture2.Cls
Picture2.Width = LevelX * 18 + 14
Picture2.Height = LevelY * 17 + 12
Picture2.ScaleMode = 1
Form1.Width = Picture2.ScaleWidth + 367
Form1.Height = Picture2.ScaleHeight + 1310
Picture2.ScaleMode = 3
'Position button & clock
Command1.Left = Picture2.Width / 2 - 6
Label2.Left = Picture2.Width - 35
New_Game
End Sub

Private Sub menuNew_Click()
New_Game
End Sub

Private Sub menuExit_Click()
End
End Sub

Private Sub menuAbout_Click()
Dim Temp$
Temp$ = "  HexSweeper" & vbCr & _
        "  by Paul Bahlawan.  April 2003  (build " & App.Revision & ")  "
MsgBox Temp$, , "About HexSweeper"
End Sub

Private Sub Command1_Click()
New_Game
Picture2.SetFocus  'Just to get rid of that dotted box
End Sub

Private Sub New_Game()
Dim X As Integer, Y As Integer, Z As Integer
Timer1.Enabled = False
'Clear/Create the play field
For Y = 0 To LevelY - 1
    For X = 0 To LevelX - 1
        Draw X, Y, 9
        Field(X, Y) = 0
    Next
Next
Picture2.Refresh

'Place mines randomly
For Z = 1 To TotalMine
New_Mine:
    X = Int(Rnd(1) * LevelX)
    Y = Int(Rnd(1) * LevelY)
    If Field(X, Y) = 7 Then GoTo New_Mine
    Field(X, Y) = 7
    'Place numbers around the new mine
    Mine_Count X - 1, Y
    Mine_Count X + 1, Y
    Mine_Count X, Y - 1
    Mine_Count X, Y + 1
    If Y / 2 = Int(Y / 2) Then  'Even row...
        Mine_Count X + 1, Y + 1
        Mine_Count X + 1, Y - 1

    Else                        '...or Odd row
        Mine_Count X - 1, Y - 1
        Mine_Count X - 1, Y + 1
    End If
Next Z
Progress = LevelX * LevelY - TotalMine
Flags = TotalMine
Score
Command1.BackColor = &HC0C0C0
Command1.Caption = "K"
Label2.Caption = "0"
Alive = True
End Sub

Private Sub Mine_Count(X As Integer, Y As Integer)
If X >= 0 And X < LevelX And Y >= 0 And Y < LevelY Then
    If Field(X, Y) <> 7 Then
        Field(X, Y) = Field(X, Y) + 1
    End If
End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cx As Integer, Cy As Integer
If Alive = False Then Exit Sub
If Timer1.Enabled = False Then  'Start the CLOCK
    StartTime = Timer
    Timer1.Enabled = True
End If

'Figure out which cell was clicked
Cy = Int((Y - 3) / 17)
If Cy / 2 = Int(Cy / 2) Then 'Make adjustment for even row...
    Cx = Int((X - 9) / 18)
Else                         '...or for odd row.
    Cx = Int(X / 18)
End If
If Cx < 0 Or Cx >= LevelX Or Cy < 0 Or Cy >= LevelY Then Exit Sub

Select Case Button
    Case 1  '(left button) Uncover the selected cell
        Uncover_Cell Cx, Cy
        If Progress = 0 Then  '!! Winner !!!
            Winner
        End If
    Case 2  '(right button) Toggle flag
        If Overlay(Cx, Cy) = 9 Then  'Flag the cell
            Draw Cx, Cy, 10
            Flags = Flags - 1
        Else
            If Overlay(Cx, Cy) = 10 Then  'Unflag the cell
            Draw Cx, Cy, 9
            Flags = Flags + 1
            End If
        End If
End Select
Picture2.Refresh
Score
End Sub

Private Sub Uncover_Cell(X As Integer, Y As Integer)
If X < 0 Or X >= LevelX Or Y < 0 Or Y >= LevelY Then Exit Sub
If Overlay(X, Y) <> 9 Then Exit Sub
Select Case Field(X, Y)
    Case 0  'Empty cell so uncover it + surounding cells
        Progress = Progress - 1
        Draw X, Y, 0
        Uncover_Cell X - 1, Y  'Yes, this Sub calls it self!
        Uncover_Cell X + 1, Y
        Uncover_Cell X, Y - 1
        Uncover_Cell X, Y + 1
        If Y / 2 = Int(Y / 2) Then  'Even row
            Uncover_Cell X + 1, Y + 1
            Uncover_Cell X + 1, Y - 1
        Else                         'Odd row
            Uncover_Cell X - 1, Y - 1
            Uncover_Cell X - 1, Y + 1
        End If
    Case 1 To 6  'Numbered Cell
        Progress = Progress - 1
        Draw X, Y, Field(X, Y)
    Case 7  'Oh no, a mine!
        Die
        Draw X, Y, 8
End Select
End Sub

Private Sub Die()
Dim X As Integer, Y As Integer
Timer1.Enabled = False
For Y = 0 To LevelY - 1
    For X = 0 To LevelX - 1
        If Field(X, Y) = 7 And Overlay(X, Y) <> 10 Then  'Uncover Mines
            Draw X, Y, 7
        End If
        If Field(X, Y) <> 7 And Overlay(X, Y) = 10 Then  'X wrong Flags
            Draw X, Y, 12
        End If
    Next X
Next Y
Command1.BackColor = vbRed
Command1.Caption = "L"  'wingding sad face
Alive = False
End Sub

Private Sub Winner()
Dim X As Integer, Y As Integer
Timer1.Enabled = False
For Y = 0 To LevelY - 1 'Flag all remaining Mines
    For X = 0 To LevelX - 1
        If Field(X, Y) = 7 And Overlay(X, Y) = 9 Then
            Draw X, Y, 10
        End If
    Next X
Next Y
Flags = 0
Score
Command1.BackColor = vbGreen
Command1.Caption = "J"  'wingding happy face
Alive = False
End Sub

Private Sub Label1_Click()
ScoreType = Not ScoreType
Score
End Sub

Private Sub Score()
If ScoreType Then  'Traditional (Flags to remaining to be placed)
    Label1.ForeColor = vbRed
    Label1.Caption = Flags
Else               'New style (Cells remaining to be uncovered)
    Label1.ForeColor = vbGreen
    Label1.Caption = Progress
End If
End Sub

'''Transperant BitBlt Sub for drawing cells.
'''Picture1 contains all the graphics,
'''the graphic required is selected with Index.
'''Note: a Picture2.Refresh is needed after calling this Sub to update the display
Private Sub Draw(X As Integer, Y As Integer, Index As Integer)
Dim Offset As Integer
Overlay(X, Y) = Index
If Y / 2 = Int(Y / 2) Then Offset = 9  'Offset for even rows
BitBlt Picture2.hDC, (X * 18 + Offset), (Y * 17), 19, 24, Picture1.hDC, (13 * 19), 0&, vbSrcAnd
BitBlt Picture2.hDC, (X * 18 + Offset), (Y * 17), 19, 24, Picture1.hDC, (Index * 19), 0&, vbSrcPaint
End Sub

Private Sub Timer1_Timer()
Dim Temp$
Temp$ = Str(Int(Timer - StartTime))
Temp$ = Right(Temp$, Len(Temp$) - 1)
If Val(Temp$) >= 999 Then
    Temp$ = "999"
    Timer1.Enabled = False
End If
Label2.Caption = Temp$
End Sub
