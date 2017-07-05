VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mazeman 1.1"
   ClientHeight    =   8400
   ClientLeft      =   1695
   ClientTop       =   1950
   ClientWidth     =   9660
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbNone2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8760
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList il 
      Left            =   8640
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0656
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0768
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":087A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":098C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":110A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":121C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":132E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1440
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1552
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1664
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pbNone 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8760
      Picture         =   "Form1.frx":1776
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pbGoal 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8760
      Picture         =   "Form1.frx":1878
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pbMan 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8760
      Picture         =   "Form1.frx":197A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pbTile 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8760
      Picture         =   "Form1.frx":1A7C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8640
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   6090
      Width           =   9600
      Begin VB.Frame fMazeDitor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mazeditor"
         ForeColor       =   &H80000008&
         Height          =   1380
         Left            =   5040
         TabIndex        =   16
         ToolTipText     =   "Mouse: Left=paint, Right=Erase, Ctrl=Set Start, Shift=Set Goal"
         Top             =   735
         Visible         =   0   'False
         Width           =   4425
         Begin VB.Label lInfo 
            Caption         =   "Maze #1 of 1"
            Height          =   225
            Left            =   105
            TabIndex        =   23
            Top             =   945
            Width           =   4215
         End
         Begin VB.Label lAddNew 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3045
            TabIndex        =   22
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label lBack 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "< Back"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   105
            TabIndex        =   21
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label lNext 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Next >"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1470
            TabIndex        =   20
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label lClear 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3045
            TabIndex        =   19
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label lSave 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   105
            TabIndex        =   18
            Top             =   210
            Width           =   1335
         End
      End
      Begin VB.Label lMazeName 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5040
         TabIndex        =   17
         Top             =   210
         Width           =   2640
      End
      Begin VB.Label lStylish 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "[X]   Easy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3330
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lQuit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3330
         TabIndex        =   12
         Top             =   1890
         Width           =   1335
      End
      Begin VB.Label lHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3330
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lGiveUp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Give Up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3330
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lPlay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Play!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3330
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lScores 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   2370
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "High Score Table"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Time Elapsed"
         Height          =   255
         Left            =   8190
         TabIndex        =   6
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label lTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8295
         TabIndex        =   5
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.Label lHelpScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6090
      Left            =   0
      TabIndex        =   13
      Top             =   0
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   9645
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Ideally I would group all this in TYPEs This looks like typical Basic code.. ugh
'
Const Game_Speed = 0.1   '--- seconds per frame, 1000 x 0.1
Const Tile = 16          '--- picture size/tiles (16x16)

'--- Gameplay variables -----
Dim sPlayerName As String       '--- Player Name
Dim bPlaying As Boolean         '--- Game in progress
Dim iPlayerX As Long, iPlayerY As Long  'Player coordinates in grid
Dim siTime As Single         '--- Time elapsed
Dim bStylish As Boolean


'--------------------------
' Read Help text
Private Sub MazeHelp()
Dim sTxt As String
Dim sHelp As String
    sHelp = ""
    On Error GoTo Crash

    Open App.Path & "\mazehelp.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, sTxt
        sTxt = Trim(sTxt & "")
        sHelp = sHelp & sTxt & vbCrLf
    Wend
    lHelpScreen.Caption = sHelp

Crash:
On Error Resume Next
Close #1
End Sub

Private Function First_File(ByRef sFileAndPath As String, Optional bFolder As Boolean = True) As String
Dim s As String
Dim lAttrib As VbFileAttribute
    lAttrib = vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem
    If bFolder Then lAttrib = lAttrib + vbDirectory
    
    s = Dir(sFileAndPath, lAttrib)
    If s <> "" Then
        While s = "." Or s = ".."
            s = Dir
        Wend
    End If
    First_File = s
End Function


'---------------------------------------
' High Scores
' I will read all high scores and optionally enter the new one
'
Private Sub MazeHighScores(Optional NewEntry As Boolean = False)
'Dim sHighScores As String
'Dim lHighscores() As String
'Dim sTxt As String
'Dim i As Long
'Dim t As Long, v As Single
'Dim bDone As Boolean
'
'    On Error GoTo Crash1
'    sHighScores = "": i = 0: bDone = False
'
'    Open App.Path & "\bestscore.dat" For Input As #1
'    While Not EOF(1)
'        Line Input #1, sTxt
'        sTxt = Trim(sTxt & "")
'
'        '--- check if new high score is better than others ---
'        If NewEntry Then
'            t = InStr(1, sTxt, ":"): If t = 0 Then t = Len(sTxt)
'            v = Val(Mid(sTxt, 1, t - 1))
'            If siTime < v Then       '-- new high score
'                '-- don't forget to change code below
'                ReDim Preserve lHighscores(i)
'                lHighscores(i) = Format(siTime, "######.00") & " : " & sPlayerName
'                sHighScores = sHighScores & lHighscores(i) & vbCrLf
'                bDone = True
'                i = i + 1
'            End If
'        End If
'
'        ReDim Preserve lHighscores(i)
'        lHighscores(i) = sTxt
'        sHighScores = sHighScores & sTxt & vbCrLf
'
'        i = i + 1
'    Wend
'
'Crash1:
'    On Error Resume Next
'    Close #1
'    On Error GoTo Crash2
'
'    '--- no high score? go to bottom...
'    If bDone = False And NewEntry = True Then
'        '-- don't forget to change code above
'        ReDim Preserve lHighscores(i)
'        lHighscores(i) = Format(siTime, "######.00") & " : " & sPlayerName
'        sHighScores = sHighScores & lHighscores(i) & vbCrLf
'    End If
'
'    '--- save it if necessary ----
'    If NewEntry = True Then
'        Open App.Path & "\bestscore.dat" For Output As #1
'        For i = LBound(lHighscores()) To UBound(lHighscores())
'            Print #1, lHighscores(i)
'        Next i
'    End If
'
'Crash2:
'    On Error Resume Next
'    Close #1
'    lScores.Caption = sHighScores
End Sub

Private Sub MazeDraw(ByRef xMaze As tMaze, Optional ClearAll As Boolean = False)
Dim X As Long, Y As Long
Dim xx As Long, yy As Long

'normal routine....
With xMaze
    If Not (bStylish) Then
        For Y = LBound(.bMaze(), 2) To UBound(.bMaze(), 2)
            For X = LBound(.bMaze(), 1) To UBound(.bMaze(), 1)
                xx = X * Tile
                yy = Y * Tile
                If ClearAll Then
                    fMain.PaintPicture pbNone.Picture, xx, yy
                Else
                Select Case .bMaze(X, Y)
                        Case mWall: fMain.PaintPicture pbTile.Picture, xx, yy
                        Case mEnd: fMain.PaintPicture pbGoal.Picture, xx, yy
                        'Case mBegin:
                        Case Else: fMain.PaintPicture pbNone.Picture, xx, yy
                    End Select
                End If
            Next X
        Next Y
        
    Else    '---- slow draw -----
        Dim utile As String * 1, dtile As String * 1, ltile As String * 1, rtile As String * 1
        Dim lWall As Long
        
        For Y = LBound(.bMaze(), 2) To UBound(.bMaze(), 2)
            For X = LBound(.bMaze(), 1) To UBound(.bMaze(), 1)
                xx = X * Tile
                yy = Y * Tile

                If ClearAll Then
                    fMain.PaintPicture pbNone2.Picture, xx, yy
                Else


                    Select Case .bMaze(X, Y)
                        Case mEnd: fMain.PaintPicture pbGoal.Picture, xx, yy
                        Case mWall:
                            '--- VB6 junk can't use bit-fields efficiently, I'll use chars then concatenate
                                                   
                            '--- get tiles above, below, right and left. Assume no tile if out of bounds
                            ltile = " ": rtile = " ": utile = " ": dtile = " "
                            If X > LBound(.bMaze(), 1) Then If .bMaze(X - 1, Y) = mWall Then ltile = "L"
                            If X < UBound(.bMaze(), 1) Then If .bMaze(X + 1, Y) = mWall Then rtile = "R"
                            
                            If Y > LBound(.bMaze(), 2) Then If .bMaze(X, Y - 1) = mWall Then utile = "U"
                            If Y < UBound(.bMaze(), 2) Then If .bMaze(X, Y + 1) = mWall Then dtile = "D"
                            
                            
                            '--- pick..
                    'tLR , tR, tL, tUD, tD, tU,tLRUD,tNone,tLRUDZ,tLD,tLU,tRU,tRD,,tLRD,tLRU,tLUD,rRUD
                    '           case "LRUD"
                            Select Case ltile & rtile & utile & dtile
                                Case "LRUD": lWall = eTiles.tLRUD ' +
                                Case "    ": lWall = eTiles.tNone
                                
                                Case "L   ": lWall = eTiles.tL
                                Case " R  ": lWall = eTiles.tR
                                Case "LR  ": lWall = eTiles.tLR ' -
                                
                                Case "  U ": lWall = eTiles.tU
                                Case "   D": lWall = eTiles.tD
                                Case "  UD": lWall = eTiles.tUD ' |
                                
                                '    "LRUD"
                                Case " RU ": lWall = eTiles.tRU ' |_
                                Case "L U ": lWall = eTiles.tLU
                                Case "L  D": lWall = eTiles.tLD
                                Case " R D": lWall = eTiles.tRD
                                
                                Case "LRU ": lWall = eTiles.tLRU ' T
                                Case "LR D": lWall = eTiles.tLRD
                                Case "L UD": lWall = eTiles.tLUD
                                Case " RUD": lWall = eTiles.tRUD
                            End Select
                            
                            'lWall = 1
                            fMain.PaintPicture il.ListImages(lWall + 1).Picture, xx, yy
                        
                        Case Else:
                            fMain.PaintPicture pbNone2.Picture, xx, yy
                            
                    End Select 'bmaze
                End If 'clearall
            Next X
        Next Y
    End If
    
    'Draw player if Maze was drawn
    If Not ClearAll Then
        fMain.PaintPicture pbMan.Picture, iPlayerX * Tile, iPlayerY * Tile
    End If
End With
End Sub



'-----------------
' Initialize
Private Sub GameInit()
    tTimer.Enabled = False
    bPlaying = False
    iPlayerX = 0: iPlayerY = 0
    iMove = imNone
    iWillMove = imNone
    siTime = 0
    Call MazeDraw(MazeList(MazeNo), True)
    lMazeName.Caption = MazeList(MazeNo).MazeName
    
    iMove = imNone
    iWillMove = imNoChange
    lTime.Caption = "..."
    
    '-- determine player start position, determined at loading
    iPlayerX = MazeList(MazeNo).XStart
    iPlayerY = MazeList(MazeNo).YStart
        
    '-- read high score table ---
    On Error GoTo Crash
    
    Me.SetFocus
    '--
    Exit Sub
Crash:
    '-- no high scores yet? --
End Sub

'--- Pause and avoid showing the maze
Private Sub GamePause(Optional HideMaze As Boolean = True)
    If bPlaying Then
        tTimer.Enabled = False
        If HideMaze Then fMain.Hide
    End If
End Sub

'--- Unpause and show the maze
Private Sub GameUnPause()
    If bPlaying Then
        fMain.Show
        tTimer.Enabled = True
    End If
End Sub



'==========================================
' COMMAND BUTTONS
' I don't know WHY the commandbuttons steal
' the vbKEYArrow commands, it's illogical.
'
' Keypreview doesn't work, so I'm doing an ugly
' hack.. I don't have time to fight with VB now.
'
'----------------------------------------
Private Sub lGiveUp_Click()
    GamePause
    If MsgBox("Are you sure you want to Give Up?", _
       vbQuestion + vbYesNo, App.Title) = vbYes Then
        GameUnPause
        GameInit
    End If
    GameUnPause
End Sub

Private Sub lHelp_Click()
    GamePause False
    lHelpScreen.Visible = True
    MsgBox "Game is paused", vbOKOnly, App.Title
    lHelpScreen.Visible = False
    GameUnPause
End Sub

Private Sub lPlay_Click()
    '--- Manage name ---
    If sPlayerName = "" Then
        sPlayerName = Trim(InputBox("Please enter your name!", App.Title, sPlayerName) & "")
        If sPlayerName = "" Then sPlayerName = "-----"
    End If
    
    Call MazeDraw(MazeList(MazeNo))
    tTimer.Enabled = True
    bPlaying = True
End Sub

Private Sub lQuit_Click()
    GamePause
        
    If MsgBox("Are you sure you want to quit?", _
       vbQuestion + vbYesNo, App.Title) = vbYes Then
              Form_Quit
    Else
        GameUnPause
    End If
End Sub



'=========================================
' INITALIZE and manage Form
'
Private Sub Form_Initialize()
    ReDim MazeList(0)
    bStylish = True
    MazeNo = 1
    
    Call MazeLoad
    Call MazeHelp
    sPlayerName = ""
    MazeHighScores
End Sub

Private Sub Form_Load()
    fMain.AutoRedraw = False
    tTimer.Enabled = False
    tTimer.Interval = Int(1000 * Game_Speed)
    Me.KeyPreview = True
   
    GameInit
End Sub

Private Sub Form_Paint()
        Call MazeDraw(MazeList(MazeNo), Not bPlaying)
End Sub

Private Sub Form_Quit()
    bPlaying = False
    tTimer.Enabled = False
    Unload fMain
End Sub

'================================================
' GAME movement


' Indicate where the player will move next.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft: iWillMove = imLeft: KeyCode = 0
        Case vbKeyLeft, vbKeyNumpad4: iWillMove = imLeft: KeyCode = 0
        Case vbKeyRight, vbKeyNumpad6: iWillMove = imRight: KeyCode = 0
        Case vbKeyUp, vbKeyNumpad8: iWillMove = imUp: KeyCode = 0
        Case vbKeyDown, vbKeyNumpad2: iWillMove = imDown: KeyCode = 0
        Case vbKeyNumpad5, vbKeySpace: iWillMove = imNone: KeyCode = 0
        Case vbKeyEscape: lQuit_Click
        Case vbKeyReturn: If Not bPlaying Then Call lPlay_Click
    End Select
End Sub


Private Sub lStylish_Click()
Dim s As String
    s = lStylish.Caption
    
    bStylish = Not bStylish
    If bStylish Then
        Mid$(s, 2, 1) = "X"
    Else
        Mid$(s, 2, 1) = " "
    End If
    lStylish.Caption = s
    If (bPlaying) Then Call MazeDraw(MazeList(MazeNo), False)
End Sub

Private Sub lTime_Click()
    If Not bPlaying Then Call lPlay_Click
End Sub

'-------------------------------------------------
' The game is synched to the timer, at 1/2 second
'
Private Sub tTimer_Timer()
Dim OldX As Long, OldY As Long
Dim NewX As Long, NewY As Long
    
With MazeList(MazeNo)
    '-- remember player --- I'd use bit shift instead of multiplying... I hate VB6
    OldX = iPlayerX * Tile
    OldY = iPlayerY * Tile
       
    '--- player may change direction only if there's no wall...
    If (iWillMove = imLeft) Then
        If (iPlayerX > LBound(.bMaze(), 1)) And (.bMaze(iPlayerX - 1, iPlayerY) <> mWall) Then
            iMove = imLeft
            iWillMove = imNoChange
        End If
    End If
    If (iWillMove = imRight) Then
        If (iPlayerX < UBound(.bMaze(), 1)) And (.bMaze(iPlayerX + 1, iPlayerY) <> mWall) Then
            iMove = imRight
            iWillMove = imNoChange
        End If
    End If
    If (iWillMove = imUp) Then
        If (iPlayerY > LBound(.bMaze(), 2)) And (.bMaze(iPlayerX, iPlayerY - 1) <> mWall) Then
            iMove = imUp
            iWillMove = imNoChange
        End If
    End If
    If (iWillMove = imDown) Then
        If (iPlayerY < UBound(.bMaze(), 2)) And (.bMaze(iPlayerX, iPlayerY + 1) <> mWall) Then
            iMove = imDown
            iWillMove = imNoChange
        End If
    End If
    
    '--- player is actually moving...
    If (iMove = imLeft) Then
        If (iPlayerX > LBound(.bMaze(), 1)) And (.bMaze(iPlayerX - 1, iPlayerY) <> mWall) Then
            iPlayerX = iPlayerX - 1
        End If
    End If
    If (iMove = imRight) Then
        If (iPlayerX < UBound(.bMaze(), 1)) And (.bMaze(iPlayerX + 1, iPlayerY) <> mWall) Then
            iPlayerX = iPlayerX + 1
        End If
    End If
    If (iMove = imUp) Then
        If (iPlayerY > LBound(.bMaze(), 2)) And (.bMaze(iPlayerX, iPlayerY - 1) <> mWall) Then
            iPlayerY = iPlayerY - 1
        End If
    End If
    If (iMove = imDown) Then
        If (iPlayerY < UBound(.bMaze(), 2)) And (.bMaze(iPlayerX, iPlayerY + 1) <> mWall) Then
            iPlayerY = iPlayerY + 1
        End If
    End If
    
    'new coordinates
     NewX = iPlayerX * Tile
     NewY = iPlayerY * Tile
     
    'if player moved, update pic....
    'I'll draw new position, then erase old. I prefer a ghost effect rather than flickering.
    If (OldX <> NewX) Or (OldY <> NewY) Then
        fMain.PaintPicture pbMan.Picture, NewX, NewY    'draw
        fMain.PaintPicture pbNone.Picture, OldX, OldY   'erase
    End If
    
    '-- update time display ----
    siTime = siTime + Game_Speed
    lTime.Caption = Format(siTime, "###,###.00")
    
    '-- check if player won! ---------
    If .bMaze(iPlayerX, iPlayerY) = mEnd Then
        tTimer.Enabled = False
        MazeHighScores True
        MsgBox "Excellent! You Won!" & vbCrLf & "Time was " & Format(siTime, "###,###.00") & " seconds!", vbOKOnly, App.Title
        MazeNo = MazeNo + 1: If MazeNo > UBound(MazeList) Then MazeNo = 1
        GameInit
    End If
End With
End Sub



'================================================== Mazeditor ==========================================
'--(re)draw
Private Sub MazeDrawNow()
    lInfo.Caption = "Maze " & MazeNo & " of " & UBound(MazeList)
    With MazeList(MazeNo)
        lMazeName.Caption = .MazeName
        iPlayerX = .XStart: iPlayerY = .YStart
    End With
    'Call MazeDraw(MazeList(MazeNo))
    'DoEvents
    Me.Refresh
End Sub

'--activate dblclicking high score table label
Private Sub Label3_DblClick()
    tTimer.Enabled = False
    bStylish = False
    bPlaying = True
    fMazeDitor.Visible = True
    Call MazeDrawNow
End Sub


'--back
Private Sub lBack_Click()
    If MazeNo > 1 Then MazeNo = MazeNo - 1
    Call MazeDrawNow
End Sub
'--next
Private Sub lNext_Click()
    If MazeNo < UBound(MazeList) Then MazeNo = MazeNo + 1
    Call MazeDrawNow
End Sub

'--clear
Private Sub lClear_Click()
    If MsgBox("Clear forever... Are you sure?", vbCritical + vbYesNo) = vbYes Then
        Call Maze_Clear
        Call MazeDrawNow
    End If
End Sub

'--new
Private Sub lAddNew_Click()
    If MsgBox("Want a new, blank Maze?", vbQuestion + vbYesNo) = vbYes Then
        ReDim Preserve MazeList(UBound(MazeList) + 1)
        MazeNo = UBound(MazeList)
        With MazeList(MazeNo)
            ReDim .bMaze(39, 24)
            .iMazeX = 40
            .iMazeY = 25
            .MazeName = "NoName"
            .Comment = ""
        End With
        
        Call Maze_Clear
        Call MazeDrawNow
        
    End If
End Sub

'erase a maze, adds border, and default start and end
Private Sub Maze_Clear()
Dim X As Long, Y As Long
Dim xL As Long, xU As Long  'Lbound, UBound
Dim yL As Long, yU As Long  'Lbound, UBound
   
    With MazeList(MazeNo)
        xL = LBound(.bMaze(), 1): xU = UBound(.bMaze(), 1)
        yL = LBound(.bMaze(), 2): yU = UBound(.bMaze(), 2)

        For Y = yL To yU: .bMaze(xL, Y) = mWall: .bMaze(xU, Y) = mWall: Next Y  'vertical border
        For X = xL To xU: .bMaze(X, yL) = mWall: .bMaze(X, yU) = mWall: Next X  'horizontal border
        
        'clear
        For Y = yL + 1 To yU - 1
            For X = xL + 1 To xU - 1
                .bMaze(X, Y) = mBlank
            Next X
        Next Y
        
        'set player/start and goal defaults
        .XStart = xL + 1: .YStart = yL + 1
        .bMaze(.XStart, .YStart) = mBegin   'start (upper left)
        .bMaze(xU - 1, yU - 1) = mEnd          'end (lower right)
    End With
End Sub

'save
Private Sub lSave_Click()
Dim s As String, t As String, l As String
Dim i As Integer
Dim m As Long, X As Long, Y As Long, xAdd As Long
    s = App.Path & "\mazeman.bak": If First_File(s) <> "" Then Call Kill(s)   'delete old backup, create new
    t = App.Path & "\mazeman.txt": If First_File(t) <> "" Then Name t As s
        
    i = FreeFile
    Open t For Output As #i
    Print #i, ";Mazes for Mazeman. "
    Print #i, ";Mazeditor will overwrite the entire file."
    
    For m = 1 To UBound(MazeList)
        With MazeList(m)
            Print #i, "[MAZENAME]"
            Print #i, .MazeName
            Print #i, "[COMMENTS]"
            If .Comment <> "" Then Print #i, .Comment
            Print #i, "[MAZEDATA]"
            If LBound(.bMaze(), 1) = 0 Then xAdd = 1 Else xAdd = 0      'string is 1 based, array should be 0 based
            
            For Y = LBound(.bMaze(), 2) To UBound(.bMaze(), 2)
                l = String$(UBound(.bMaze(), 1) + xAdd, " ")
                For X = LBound(.bMaze(), 1) To UBound(.bMaze(), 1)
                    Select Case .bMaze(X, Y)
                        Case mWall: Mid$(l, X + xAdd, 1) = "X"
                        Case mBegin: Mid$(l, X + xAdd, 1) = "S"
                        Case mEnd: Mid$(l, X + xAdd, 1) = "G"
                    End Select
                Next X
                Print #i, l
            Next Y
            Print #i, "[END]"
        End With
    Next m
    Close #i
    
    MsgBox "Ok"
End Sub


Private Sub Form_MouseX(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 0 Then Exit Sub
Dim xtile As Long, ytile As Long, xx As Long, yy As Long
Dim xL As Long, yL As Long, xU As Long, yU As Long  'Lbound, Ubound
Dim lDraw As Long, bDraw As Byte
       
    With MazeList(MazeNo)
        xL = LBound(.bMaze(), 1): xU = UBound(.bMaze(), 1)
        yL = LBound(.bMaze(), 2): yU = UBound(.bMaze(), 2)
        
        xtile = X \ Tile: ytile = Y \ Tile

        If fMazeDitor.Visible Then
            'avoid out of bounds if window is larger than grid
            If (xtile < xL) Then xtile = xL Else If (xtile > xU) Then xtile = xU
            If (ytile < yL) Then ytile = yL Else If (ytile > yU) Then ytile = yU
    
            bDraw = mBlank
            If (Button And vbLeftButton) > 0 Then       'paint
                If (Shift And vbShiftMask) > 0 Then
                    '--- erase previous starting man
                    xx = .XStart * Tile: yy = .YStart * Tile
                    .bMaze(.XStart, .YStart) = mBlank
                    fMain.PaintPicture pbNone.Picture, xx, yy
                    '--- set new start
                    bDraw = mBegin
                    .XStart = xtile: .YStart = ytile
                    
                ElseIf (Shift And vbCtrlMask) > 0 Then
                    bDraw = mEnd
                Else
                    bDraw = mWall
                End If
            End If
            
            .bMaze(xtile, ytile) = bDraw
            xx = xtile * Tile
            yy = ytile * Tile
            
            Select Case bDraw
                Case mWall: fMain.PaintPicture pbTile.Picture, xx, yy
                Case mEnd: fMain.PaintPicture pbGoal.Picture, xx, yy
                Case mBegin: fMain.PaintPicture pbMan.Picture, xx, yy
                Case Else: fMain.PaintPicture pbNone.Picture, xx, yy
            End Select
        End If 'fmazeditor
    End With 'mazelist
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'it only sends 1st down...
    Call Form_MouseX(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'it only sends if mouse moves...
    Call Form_MouseX(Button, Shift, X, Y)
End Sub

Private Sub lMazeName_Click()
    If fMazeDitor.Visible Then
        MazeList(MazeNo).MazeName = Trim$(InputBox("Maze name", App.Title, MazeList(MazeNo).MazeName))
        Call MazeDrawNow
    End If
End Sub

