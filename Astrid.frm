VERSION 5.00
Begin VB.Form frmAstar 
   Caption         =   "A* Demo"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   436
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   1560
      Picture         =   "Astrid.frx":0000
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   1080
      Picture         =   "Astrid.frx":1302
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generate Map"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   360
      Picture         =   "Astrid.frx":17F4
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Redraw"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Path"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picWorking 
      AutoRedraw      =   -1  'True
      Height          =   3495
      Left            =   0
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label2 
      Height          =   855
      Left            =   5040
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"Astrid.frx":1CE6
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   6135
   End
End
Attribute VB_Name = "frmAstar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Code by FireClaw 26/10/00
' Description:  Demo of the A* pathing algorithm
' No copyright - feel free to copy this code and use it in any way you feel like.
' Any comments, improvements, and bug-fixes to bigcalm@hotmail.com

' Please please please, if you have no idea what A* is, or what it is used for, then
' please read some of the following resources, before mailing me:
' http://theory.stanford.edu/~amitp/GameProgramming/    - Amit Patel's fantastic notes
'       on A* and other games programming stuff.  Truly superb resource.
' http://www.gameai.com     -   Another good resouce for Artificial Intelligence, including some
'               notes on A*
' http://www.gamedev.net    -   If you want your brain fried by Maths, check out these very useful pages (novice programmers should stay away)
' http://www.ccg.leeds.ac.uk/james/aStar/   - a java representation of the A* algorithm
' http://www.utm.edu/cgi-bin/caldwell/tutor/departments/math/graph/intro - If you are not intending to
'       use a tile based system, but a vector based system, this page introduces graph theory (you will
'       need to reduce your vector objects to a connected graph for A* to work out a path).

' This Project implements the A* algorithm for finding a path through terrain, and the
' algorithm is used in games to help find a path from one place to another on a map.
' Different types of A* are used in other types of game - for example, chess game
' opponents use an A* variant called ID-A* (Iterative Deepening).
' There are many possible implementations of A*, of this is only one.
' If you want to include this into a game environment you will have to make several
' modifications.
' Namely:  1) Use your own MAP class.  (If you used an array to hold the map rather than a collection
'                   the A* algorithm would be much faster).
'               2) If your game is "interactive" instead of "turn taking" you will need to be
'                   able to interrupt path finding, to draw graphics, do AI, etc.  I suggest
'                   giving A* a maximum period of time to work on the problem, start
'                   any unit walking in the general direction (or, get them to scratch their heads),
'                   and give A* a few miliseconds per frame to work out paths.  The reason
'                   behind this is that A* IS COMPUTATIONALLY EXPENSIVE, even though it will always
'                   find a path if there is one.  If there is no path, (e.g. the end is blocked off), then
'                   then A* will search every available square before giving up).  You could even
'                   give A* a maximum search time, in which case, your unit just stays where they
'                   are (and assumes no path if not found in a certain amount of time).
'                   You can also "tune" the A* algorithm by modifying the heuristic estimate of the
'                   distance - by giving an under-estimate, it will find paths quicker, but
'                   these might actually be *better* (less movement needed) paths.
'               3) Adjust the A* to suit your game - if you allow diagonal movement, want to allow
'                  path splicing, etc.  Check out http://theory.stanford.edu/~amitp/GameProgramming/
'                   for possible ways to adapt this A* algorithm to your game (also covers Path Splicing).
'               4) Improve the A* algorithm (use Beam A*, or change the heuristic) - my
'                   algorithm is probably a little too general.  Plus changing the HeapNode Class to having Early-Binding would help
' I would also recommend writing A* in C or C++, as it will be much faster because of
' C's good implementation of heaps ( I've only written this in VB 'cos my C compiler's broken).

' This implementation of A* can cope with different types of terrain (i.e. MoveCost of Map is taken into account), so
' this implementation can find quicker paths if there are Roads, or navigate around Mountains, etc.
' For Route finding software (that finds a road-route from one location to another), A* is also
' used, but instead of considering all roads, it will navigate to a B road, then an A road, then to a motorway,
' and then reverse this if necessary, using several levels of A* pathing to speed up the search.
' This dividing of A* is also used on some "World" map type games, where continents and islands are "precalculated"
' before telling A* to search for a path.

Private Const MaxX As Long = 10  ' These constants would normally be part of Map - i.e. Map.Width, Map.Height
Private Const MaxY As Long = 10
Private Map1 As Map
Private Path As Path
Private CoStart As CoOrdinate
Private CoEnd As CoOrdinate

' Find the path
Private Sub Command1_Click()
Dim StartTime As Long
Dim EndTime As Long
Dim CoOrd As CoOrdinate

    StartTime = timeGetTime
    Set Path = bAStar(Map1, CoStart.X, CoStart.Y, CoEnd.X, CoEnd.Y, 0, 0, MaxX, MaxY)
    EndTime = timeGetTime
    Label2.Caption = "Time taken: " & EndTime - StartTime & " miliseconds"
    If Path Is Nothing Then
        Label2.Caption = Label2.Caption & " (No path found) "
    Else
        Label2.Caption = Label2.Caption & " (Total Cost: " & Path.TotalMoveCost & " )"
    End If
    Command2_Click
End Sub

' Draw map
Private Sub Command2_Click()
    DrawBoard
    FlipBoard
End Sub

' Generate random map
Private Sub Command3_Click()
Dim i As Long
Dim j As Long
Dim k As Long
Dim MoveCost As Long

    Set Map1 = Nothing
    Set Map1 = New Map
    Set Path = Nothing
    Randomize
    For i = 0 To MaxX
        For j = 0 To MaxY
            k = Rnd * 1.5
            If k = 1 Then
                If Rnd > 0.5 Then
                    k = 5 + Rnd * 2
                    Select Case k
                        Case 5
                            MoveCost = 10
                        Case 6
                            MoveCost = 0
                        Case 7
                            MoveCost = 5
                    End Select
                Else
                    MoveCost = 1
                End If
            Else
                MoveCost = 1
            End If
            Map1.Add "R" & i & "C" & j, i, j, MoveCost, k, "R" & i & "C" & j
        Next
    Next
    Map1("R5C5").NodeType = 1
    Map1("R10C10").NodeType = 1
    Set CoStart = Nothing
    Set CoStart = New CoOrdinate
    CoStart.X = 5
    CoStart.Y = 5
    Set CoEnd = Nothing
    Set CoEnd = New CoOrdinate
    CoEnd.X = 10
    CoEnd.Y = 10
    Command2_Click
End Sub

Private Sub Form_Load()
    Command3_Click
    Combo1.AddItem "Normal"
    Combo1.AddItem "Mountain"
    Combo1.AddItem "Road"
    Combo1.AddItem "Forest"
    Combo1.AddItem "Impassable"
    Combo1.ListIndex = 0
End Sub

' Erm, I've realised that somewhere I've reversed co-ordinates
' So rows are columns and vice-versa when it comes to displaying the information
' Oops.
Private Sub DrawBoard()
Dim MN As MapNode
Dim CoOrd As CoOrdinate
Dim LastCoOrd As CoOrdinate

    For Each MN In Map1
        Select Case MN.NodeType
            Case 0
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, picWorking.hDC, 0, 0, BLACKNESS
            Case 1
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, picWorking.hDC, 0, 0, WHITENESS
            Case 2
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, Picture1.hDC, 0, 0, vbSrcCopy
            Case 3
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, Picture1.hDC, 0, 0, WHITENESS
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, Picture1.hDC, 0, 0, vbSrcInvert
            Case 4
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, Picture2.hDC, 0, 0, vbSrcCopy
            Case 5
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, Picture3.hDC, 0, 0, vbSrcCopy
            Case 6
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, Picture3.hDC, 20, 0, vbSrcCopy
            Case 7
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, Picture3.hDC, 20, 20, vbSrcCopy
            Case 8
                BitBlt picWorking.hDC, MN.Y * 20, MN.X * 20, 20, 20, Picture3.hDC, 0, 20, vbSrcCopy
        End Select
    Next
    If Not (Path Is Nothing) Then
        Set LastCoOrd = Nothing
        For Each CoOrd In Path
            If Not (LastCoOrd Is Nothing) Then
                picWorking.Line (LastCoOrd.Y * 20 + 10, LastCoOrd.X * 20 + 10)-(CoOrd.Y * 20 + 10, CoOrd.X * 20 + 10)
            End If
            Set LastCoOrd = CoOrd
        Next
    End If
End Sub

Private Sub FlipBoard()
    BitBlt Me.hDC, 0, 0, picWorking.ScaleWidth, picWorking.ScaleHeight, picWorking.hDC, 0, 0, vbSrcCopy
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MLX As Long
Dim MLY As Long
    MLX = Y \ 20
    MLY = X \ 20
    If MLX < 0 Or MLX > MaxX Then
        Exit Sub
    End If
    If MLY < 0 Or MLY > MaxY Then
        Exit Sub
    End If
    Select Case Combo1.ListIndex
        Case 0
            Map1("R" & MLX & "C" & MLY).NodeType = 1
            Map1("R" & MLX & "C" & MLY).MoveCost = 1
        Case 1
            Map1("R" & MLX & "C" & MLY).NodeType = 5
            Map1("R" & MLX & "C" & MLY).MoveCost = 10
        Case 2
            Map1("R" & MLX & "C" & MLY).NodeType = 6
            Map1("R" & MLX & "C" & MLY).MoveCost = 0
        Case 3
            Map1("R" & MLX & "C" & MLY).NodeType = 7
            Map1("R" & MLX & "C" & MLY).MoveCost = 5
        Case 4
            Map1("R" & MLX & "C" & MLY).NodeType = 0
            Map1("R" & MLX & "C" & MLY).MoveCost = 10
    End Select
    Command2_Click
End Sub
