Attribute VB_Name = "mMMan"
Option Explicit

Public Type tMaze
    bMaze() As Byte     '-- Maze
    MazeName As String
    Comment As String
    
    iMazeX As Long
    iMazeY As Long    '-- Maze size
    
    XStart As Long
    YStart As Long  '-- Starting position
End Type

Public MazeList() As tMaze

Public MazeNo As Long  'current maze/level # (starts from 1)

'=========================================================
' eMove constants
Public Enum eMove
    imNone = 0
    imUp
    imDown
    imLeft
    imRight
    imNoChange
End Enum
Public iMove As eMove            '--- Where the player is going
Public iWillMove As eMove        '--- Where will it go if possibe (like in Pacman)

' Maze constants. Maze is a byte array, that's why I'll use byte constants
Public Const mBlank    As Byte = 0
Public Const mWall     As Byte = 1
Public Const mBegin    As Byte = 2
Public Const mEnd      As Byte = 3

'tiles, x,y  xr,xl  yu,yd  z(cross) xy(solid)
Public Enum eTiles
    tLR     'horizontal
    tR      'right neighbor
    tL
    tUD     'vertical
    tD      'bottom neighbor
    tU
    tLRUD   'all, cross +
    tNone   'no neighbors
    tLRUDZ  'solid, in theory if all neighbors are crosses
    
    tLD     'neighbor L + D
    tLU
    tRU
    tRD
    
    tLRD    'neighbor left + right + down (T)
    tLRU
    tLUD
    tRUD
End Enum
'=====================================================================
'-----------------------------------------
' Loads Mazeman.txt from local directory
'
Public Sub MazeLoad()
Dim sLine As String, c As String     '-- Line of text read from file
Dim i As Long, X As Long, Y As Long
Dim b As Byte
Dim iBeginOk As Long, iEndOk As Long
Dim iFile As Integer
On Error GoTo Crash

Dim bNext_Name As Boolean, bNext_Comment As Boolean, bNext_Maze As Boolean

iFile = FreeFile
Open App.Path & "\mazeman.txt" For Input As #iFile


Do
    ReDim Preserve MazeList(UBound(MazeList) + 1)
    With MazeList(UBound(MazeList))
        ReDim .bMaze(39, 25)            '-- hardcoded...
        .iMazeX = 40: .iMazeY = 0   '-- default maze width and height
        iBeginOk = 0: iEndOk = 0    '-- to determine if there's a start and end point
        .XStart = 1: .YStart = 1
        
        Do
ReadNewLine_Reset:
            bNext_Name = False: bNext_Comment = False: bNext_Maze = False
ReadNewLine:
            If EOF(iFile) Then Exit Do
            
            Input #iFile, sLine
            sLine = UCase(Trim(sLine & ""))
                
            If sLine = "" Then GoTo ReadNewLine
            If Left$(sLine, 1) = ";" Then GoTo ReadNewLine
            If sLine = "[MAZENAME]" Then bNext_Name = True: bNext_Comment = False: bNext_Maze = False: GoTo ReadNewLine
            If sLine = "[COMMENTS]" Then bNext_Name = False: bNext_Comment = True: bNext_Maze = False: GoTo ReadNewLine
            If sLine = "[MAZEDATA]" Then bNext_Name = False: bNext_Comment = False: bNext_Maze = True: GoTo ReadNewLine
            
            If sLine <> "[END]" Then
                If bNext_Name Then .MazeName = sLine: GoTo ReadNewLine_Reset
                If bNext_Comment Then .Comment = sLine: GoTo ReadNewLine_Reset
                            
                '-- read the string, convert it to our byte format
                For i = 0 To 39
                    c = Mid(sLine, i + 1, 1)
                    
                    '-- Determine if it's beginning, end, floor or wall ---
                    Select Case c
                        Case " ":
                            b = mBlank
                            
                        Case "S", "B":          '-- Start/Begin
                            b = mBegin
                            iBeginOk = iBeginOk + 1 '-- Ok, begins somewhere
                            .XStart = i: .YStart = .iMazeY     '-- Save start coordinates...
                            
                        Case "E", "F", "G":                     '-- End/Finish/Goal
                            b = mEnd
                            iEndOk = iEndOk + 1
                            
                        Case Else:
                            b = mWall
                    End Select
                    
                    .bMaze(i, .iMazeY) = b
                Next i
                .iMazeY = .iMazeY + 1
                
            '----------------------------------------------------
            Else 's="[END]"
                bNext_Maze = False
                '--- something was loaded?
                If (.iMazeX < 2) Or (.iMazeY < 2) Then
                    MsgBox """Mazeman.txt"" does not have a proper maze!", vbCritical + vbOKOnly, App.Title
                    GoTo Crash2
                End If
                
                '--- check ending and starting ---
                If iBeginOk = 0 Then
                    MsgBox """Mazeman.txt"" problem: No starting point defined!", vbCritical + vbOKOnly, App.Title
                    GoTo Crash2
                End If
                If iEndOk = 0 Then
                    MsgBox """Mazeman.txt"" problem: No goal defined!", vbCritical + vbOKOnly, App.Title
                    GoTo Crash2
                End If
                Exit Do
                
            End If 's<>"[END]"
        Loop
    End With
    If EOF(iFile) Then Exit Do
Loop
    
Close #1
On Error GoTo 0
Exit Sub

Crash:
    MsgBox "Error #" & Err.Number & " reading Mazeman.txt." & vbCrLf & Err.Description, vbCritical + vbOKOnly, App.Title
    Resume Next
Crash2:
'
End Sub

Private Sub Maze_Load_RLE(ByRef s As String, ByRef xMaze As tMaze)

End Sub
