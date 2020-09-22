Attribute VB_Name = "AStar"
Option Explicit

' Essentially, there is a set of nodes (map locations and a cost) called OPEN and
' another set called CLOSED. Each time through the main loop, you pick out the
' best element from OPEN (where "best" means "the one with the lowest cost"),
' and you look at its neighbors. You then put any unvisited neighbors into the
' OPEN set. The cost of a node is the sum of the current cost of walking from the
' start to that node and the heuristic estimate of the cost from that node to the goal.

Type MarkD
    Cost As Single
    Set As Long ' 0 = None, 1 = Open, 2 = Closed
    Direction As Long ' 0 to 3
End Type

' This Array is used to store information about the nodes that have already been examined
Private MarkData() As MarkD

' Heuristic.
Public Function Cost(pStartX As Long, pStartY As Long, pDestinationX As Long, pDestinationY As Long) As Single
    ' Use Manhattan distance for this demo
    Cost = Abs(pStartX - pDestinationX) + Abs(pStartY - pDestinationY)
    ' Note: If we were allowing simple diagonal movement (i.e. diagonal move costs the same as any other move)
    ' then the formula would be:
    ' Cost = Max(Abs(pStartX - pDestinationX),Abs(pStartY - pDestinationY))
    ' If straight line distance then
    ' Cost = Sqr((pStartX - pDestinationX)^2 + (pStartY - pDestinationY)^2)
End Function

' This is the main A* function
Public Function bAStar(pMap As Map, pStartX As Long, pStartY As Long, pDestinationX As Long, pDestinationY As Long, MinX As Long, MinY As Long, MaxX As Long, MaxY As Long) As Path
Dim OPENSet As New Heap
' Dim CLOSEDSet As New Heap
Dim AddNode As New AStarNode
Dim BestNode As AStarNode
Dim CurrentNode As HeapNode
Dim i As Long
Dim j As Long
Dim k As Long
Dim InsertDirection As Long
Dim BackPath As Path
Dim Path As Path

    ' Initialize
    InsertDirection = 1
    ReDim MarkData(MaxX - MinX, MaxY - MinY)
    For i = MinX To MaxX
        For j = MinY To MaxY
            MarkData(i, j).Cost = 0
            MarkData(i, j).Set = 0
        Next
    Next

    ' Add the initial node to the OPEN set
    AddNode.X = pStartX
    AddNode.Y = pStartY
    AddNode.gval = 0
    AddNode.hval = Cost(pStartX, pStartY, pDestinationX, pDestinationY)
    OPENSet.Add AddNode, AddNode.gval + AddNode.hval, "R" & pStartX & "C" & pStartY
    MarkData(AddNode.X, AddNode.Y).Cost = AddNode.gval + AddNode.hval
    MarkData(AddNode.X, AddNode.Y).Set = 1
    ' Remove the references to pMap().NodeType changes if implementing into a game, and
    ' use returned PATH instead.
    'pMap("R" & AddNode.X & "C" & AddNode.Y).NodeType = 2
    Set AddNode = Nothing
    
    ' Main loop
    Do While OPENSet.Count > 0  ' If OPENSet.Count = 0 then there is no path
        ' Get Best Node, where "Best Node" is "The one that is most likely to lead to the goal"
        Set CurrentNode = OPENSet.GetLeftMostElement
        
        ' Note that in this algorithm, CLOSEDSet is never used - I use the MarkData() array
        ' instead.  BUT if you make changes, then the CLOSEDSet may come in useful.
        ' Therefore, uncomment all CLOSEDSet lines if you need it
        'CLOSEDSet.Add CurrentNode.Item, CurrentNode.Value, CurrentNode.ItemKey
        
        ' check if we've reached the goal
        Set BestNode = CurrentNode.Item
        If BestNode.X = pDestinationX And BestNode.Y = pDestinationY Then
            Exit Do ' We have reached the goal - exit and back-track to build path
        End If
        
        ' Add neighbours of current square into the OPEN set.
        ' Need to do this a different way each time to make the path "look" right
        '           0
        '         1  2
        '           3
        ' If you are allowing diagonal movement, then you would have more possible
        ' directions.
        ' e.g.    1     2     3
        '           4           6
        '           7     8    9
        ' You could always do the directions in a random order - as long as it's not the
        ' same each time (you'd end up with odd looking paths and a skewed heap).
        If InsertDirection = 1 Then
            i = 0
            j = 3
        Else
            i = 3
            j = 0
        End If
        For i = i To j Step InsertDirection
            Set AddNode = New AStarNode
            ' Get neighbour...
            Select Case i
                Case 0 ' Up Neighbour
                    AddNode.Y = BestNode.Y - 1
                    AddNode.X = BestNode.X
                Case 1 ' Left Neighbour
                    AddNode.X = BestNode.X - 1
                    AddNode.Y = BestNode.Y
                Case 2 ' Right Neighbour
                    AddNode.X = BestNode.X + 1
                    AddNode.Y = BestNode.Y
                Case 3 ' Bottom Neighbour
                    AddNode.Y = BestNode.Y + 1
                    AddNode.X = BestNode.X
            End Select
            
            If AddNode.X < MinX Or AddNode.X > MaxX Or AddNode.Y < MinY Or AddNode.Y > MaxY Then
            Else
                If pMap("R" & AddNode.X & "C" & AddNode.Y).NodeType = 0 Then
                Else
                    AddNode.gval = BestNode.gval + pMap("R" & BestNode.X & "C" & BestNode.Y).MoveCost
                    AddNode.hval = Cost(AddNode.X, AddNode.Y, pDestinationX, pDestinationY)
                    If MarkData(AddNode.X, AddNode.Y).Set = 0 Then
                        MarkData(AddNode.X, AddNode.Y).Cost = AddNode.gval + AddNode.hval
                        MarkData(AddNode.X, AddNode.Y).Set = 1
                        MarkData(AddNode.X, AddNode.Y).Direction = ReverseDirection(i)
                        OPENSet.Add AddNode, AddNode.gval + AddNode.hval, "R" & AddNode.X & "C" & AddNode.Y
                        'pMap("R" & AddNode.X & "C" & AddNode.Y).NodeType = 2  ' Remove this
                    Else
                        ' It's already in the OPEN set, so we may need to update it.
                        If MarkData(AddNode.X, AddNode.Y).Cost > AddNode.gval + AddNode.hval And MarkData(AddNode.X, AddNode.Y).Set = 1 Then
                            OPENSet.Delete "R" & AddNode.X & "C" & AddNode.Y, MarkData(AddNode.X, AddNode.Y).Cost
                            MarkData(AddNode.X, AddNode.Y).Cost = AddNode.gval + AddNode.hval
                            MarkData(AddNode.X, AddNode.Y).Direction = ReverseDirection(i)
                            OPENSet.Add AddNode, AddNode.gval + AddNode.hval, "R" & AddNode.X & "C" & AddNode.Y
                        Else
                            ' Do nothing because we don't need to assign a higher cost to an already checked node, or we don't want to recheck a closed node
                        End If
                    End If
                End If
            End If
            Set AddNode = Nothing
        Next
        InsertDirection = InsertDirection * -1
        If OPENSet.Delete(CurrentNode.ItemKey, CurrentNode.Value) = False Then
            Debug.Print "Deletion of best node failed"
        End If
        
        'pMap(CurrentNode.ItemKey).NodeType = 3 ' Remove this
        MarkData(BestNode.X, BestNode.Y).Set = 2
        Set AddNode = Nothing
        Set BestNode = Nothing
        Set CurrentNode = Nothing
    Loop
    If OPENSet.Count > 0 Then
        ' A path is possible
        Set BackPath = New Path
        ' Were going to use the MARK array to construct the path from the end-coordinate back to the start
        ' Because we've stored the DIRECTION that the square was moved into, we can back-track from
        ' the end to the start
        i = pDestinationX
        j = pDestinationY
        k = 1
        BackPath.Add "R" & i & "C" & j, i, j, "R" & i & "C" & j
        'pMap("R" & i & "C" & j).NodeType = 4 ' Remove this
        BackPath.TotalMoveCost = BackPath.TotalMoveCost + pMap("R" & i & "C" & j).MoveCost
        Do
            Select Case MarkData(i, j).Direction
                Case 0 ' Up
                    j = j - 1
                Case 1 ' Left
                    i = i - 1
                Case 2 ' Right
                    i = i + 1
                Case 3 ' Bottom
                    j = j + 1
            End Select
            BackPath.Add "R" & i & "C" & j, i, j, "R" & i & "C" & j
            BackPath.TotalMoveCost = BackPath.TotalMoveCost + pMap("R" & i & "C" & j).MoveCost
            'pMap("R" & i & "C" & j).NodeType = 4 ' Remove this
            k = k + 1
        Loop Until i = pStartX And j = pStartY
        ' Only problem now is that this path is now in reverse order.  So let's reverse it
        Set Path = New Path
        For i = k To 1 Step -1
            Path.Add BackPath(i).Key, BackPath(i).X, BackPath(i).Y
        Next
        Path.TotalMoveCost = BackPath.TotalMoveCost
        Set BackPath = Nothing
    Else
        Set Path = Nothing
    End If
    ' Tidy up
    Set OPENSet = Nothing
    'Set CLOSEDSet = Nothing
    Set bAStar = Path
End Function

' Simply flips the direction.
' 0 = Up, 1 = Left, 2 = Right, 3 = Down
Private Function ReverseDirection(pDir As Long) As Long
    Select Case pDir
        Case 0
            ReverseDirection = 3
        Case 1
            ReverseDirection = 2
        Case 2
            ReverseDirection = 1
        Case 3
            ReverseDirection = 0
    End Select
End Function
