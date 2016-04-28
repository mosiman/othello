Attribute VB_Name = "Othello_Mod_Dillon"
Type Player
    name As String * 25
    points As Integer
    Icon As String * 70
    intRep As Integer
    Avatar As New StdPicture
    isComputer As Integer
End Type

Type Moves
    posX As Integer
    posY As Integer
    points As Integer
End Type

Type HighScore
    name As String * 25
    intTime As Integer
    points As Integer
End Type

Global Const lenHighScore = 29
Global isComputer As Integer ' 0: no computer, 1: easybot, 2: hardbot
Global HighScores(1 To 6) As HighScore
Global Player(1 To 2) As Player


Public Sub searchReplace(theBoard() As Integer, ByVal currPlayer As Integer, theEnd As Boolean, ByVal posX As Integer, ByVal posY As Integer, ByVal direction As Integer, ByVal start As Boolean, Optional ByVal ogX As Integer, Optional ByVal ogY As Integer)
'This procedure starts from an origin and searches through 8 directions for a valid move, "theEnd".
'"theEnd" is reached when the procedure traverses more than one coordinate that contains enemy pieces.
'When theEnd is reached, the enemy pieces are switched to friendly pieces and the score is calculated.
    
    Dim previousX As Integer
    Dim previousY As Integer
    
    If start = True Then
        ogX = posX
        ogY = posY
    End If
    
    previousX = posX
    previousY = posY
    
    
    Select Case direction   'Changes X position depending on direction defined by K.
        Case 1, 2, 8
            posX = posX - 1
        Case 4, 5, 6
            posX = posX + 1
    End Select
    
    Select Case direction   'Changes Y position depending on direction defined by K.
        Case 2, 3, 4
            posY = posY - 1
        Case 8, 7, 6
            posY = posY + 1
    End Select
    
    If posX <> -1 And posY <> -1 And posX <> 10 And posY <> 10 Then ' Check if position is outside board.
        If theBoard(posX, posY) = (currPlayer Mod 2 + 1) And theEnd = False Then 'Ensure distance traversed > 1
                'start = False
                searchReplace theBoard(), currPlayer, theEnd, posX, posY, direction, False, ogX, ogY
        ElseIf theBoard(posX, posY) = currPlayer And theEnd = False Then 'Marks the end of a line.
            If start = False Then
                theEnd = True
            End If
        End If
        
        'If a valid play, switch all pieces in between to current player's.
        
        If theEnd = True Then
            
            If Not start Then
                Player(currPlayer Mod 2 + 1).points = Player(currPlayer Mod 2 + 1).points - 1
                Player(currPlayer).points = Player(currPlayer).points + 1
            ElseIf theBoard(ogX, ogY) = 0 Then
                Player(currPlayer).points = Player(currPlayer).points + 1
            End If
            theBoard(previousX, previousY) = currPlayer

        End If
    End If
    
End Sub

Public Sub boardPlace(board As MSFlexGrid, gridArr() As Integer, currPlayer As Integer, r As Integer, c As Integer)
    Dim X As Integer, Y As Integer
    board.Row = Y
    board.Col = X
    
    gridArr(c, r) = currPlayer
End Sub


Public Sub nextTurn(CPlayer)
    CPlayer = CPlayer Mod 2 + 1
End Sub

Public Sub refreshBoard(gridArr() As Integer, board As MSFlexGrid, isPlayable As Boolean, ByVal currPlayer As Integer)
'This procedure takes the 2D array gridArr() and draws the pieces on the grid as defined by gridArr().

    Dim K As Integer, X As Integer, Z As Integer
    Dim FullStop As Boolean
    
    FullStop = False 'Used to help determine endgame.
    Z = 1
    
    'Loop through entire board.
    
    For K = 0 To 9
        For X = 0 To 9
            If gridArr(X, K) = 1 Then 'If Player 1 is found, set picture.
                board.Col = X
                board.Row = K
                Set board.CellPicture = frmGame.imgP1CellPic.Picture
                
                Do While Z <= 8 And FullStop = False And currPlayer = 1 'Check if there are any valid moves for this player. Do not check further if there is one valid move.

                    If checkEndGame(gridArr(), currPlayer, False, X, K, Z, True, FullStop) = True Then
                        isPlayable = True
                        FullStop = True
                    End If
                    
                    Z = Z + 1
                Loop
                
                Z = 1
            ElseIf gridArr(X, K) = 2 Then 'If Player 2 is found, set picture.
            
                board.Col = X
                board.Row = K
                Set board.CellPicture = frmGame.imgP2CellPic.Picture
            
                Do While Z <= 8 And FullStop = False And currPlayer = 2  'Check if there are any valid moves for this player. Do not check further if there is one valid move.

                    If checkEndGame(gridArr(), currPlayer, False, X, K, Z, True, FullStop) = True Then
                        isPlayable = True
                        FullStop = True
                    End If
                    
                    Z = Z + 1
                Loop
                Z = 1
            End If
            
        Next X
    Next K
    

End Sub

Public Function checkEndGame(theBoard() As Integer, ByVal currPlayer As Integer, theEnd As Boolean, ByVal posX As Integer, ByVal posY As Integer, ByVal direction As Integer, start As Boolean, FullStop As Boolean) As Boolean
'Modified searchReplace to return boolean value for valid move.
'Differs from playablePiece in method of determining valid move.
'checkEndGame searches for a blank (legal) spot from an already valid location.
'playablePiece searches for the current player's piece from a blank spot.
    
    Select Case direction   'Changes X position depending on direction defined by exterior loop.
        Case 1, 2, 8
            posX = posX - 1
        Case 4, 5, 6
            posX = posX + 1
    End Select
    
    Select Case direction   'Changes Y position depending on direction defined by exterior loop.
        Case 2, 3, 4
            posY = posY - 1
        Case 8, 7, 6
            posY = posY + 1
    End Select
    
    If posX <> -1 And posY <> -1 And posX <> 10 And posY <> 10 Then ' Check if position is outside board.
    
        If theBoard(posX, posY) = (currPlayer Mod 2 + 1) And theEnd = False Then 'Ensure distance moved > 1
                checkEndGame = checkEndGame(theBoard(), currPlayer, theEnd, posX, posY, direction, False, FullStop)
        ElseIf theBoard(posX, posY) = 0 And theEnd = False Then ' Valid position found if distance traversed > 1
            If start = False Then
                theEnd = True
            End If
        End If
        
        If theEnd = True Then
            checkEndGame = True
            FullStop = True 'Stop all further searches.
        Else
            checkEndGame = False
        End If
        
   End If
    
End Function

Public Sub ComputerTurn(theBoard() As Integer, currPlayer As Integer, isPlayable As Boolean)
'This procedure contains all subprocedures needed for computer to play.

    Dim K As Integer, X As Integer, Z As Integer, O As Integer
    Dim StopSearch As Boolean 'If true, cease further searching.
    Dim numMoves As Integer
    Dim PossibleMoves(1 To 100) As Moves
    Dim randNum As Integer
    Dim points(1 To 100) As Integer
    
    isPlayable = False
    numMoves = 0
    Z = 1

    'Loop through entire board and build a list of valid positions to play.

    For K = 0 To 9
        For X = 0 To 9
                StopSearch = False
                If theBoard(X, K) = 0 Then
                    Do While Z <= 8 And StopSearch = False 'Check if there are any valid moves for this player.
    
                        If playablePiece(theBoard(), currPlayer, False, X, K, Z, True, StopSearch) = True Then
                            numMoves = numMoves + 1
                            PossibleMoves(numMoves).posX = X
                            PossibleMoves(numMoves).posY = K
                            
                            If (X = 0 And K = 0) Or (X = 9 And K = 9) Or (X = 9 And K = 0) Or (X = 0 And K = 9) Then
                                PossibleMoves(numMoves).points = 100 'Use absurdly high number in order to prioritize corners.
                            Else
                                For O = 1 To 8
                                    PossibleMoves(numMoves).points = grabPoints(theBoard(), currPlayer, False, X, K, O, True, True)
                                Next O
                            End If
                        End If
    
                        Z = Z + 1
                    Loop
    
                    Z = 1
                End If
        Next X
    Next K

    randNum = Int(Rnd() * (numMoves) + 1)
    
    'Check computer's difficulty and place piece accordingly.
    
    If Player(2).isComputer = 1 Then 'Choose a random valid piece.
        For K = 1 To 8
            searchReplace theBoard(), currPlayer, False, PossibleMoves(randNum).posX, PossibleMoves(randNum).posY, K, True
        Next K
    ElseIf Player(2).isComputer = 2 Then 'Chooses the piece that yields the highest possible amount of points, corners prioritized.
        BubbleSortPoints PossibleMoves(), numMoves
        For K = 1 To 8
            searchReplace theBoard(), currPlayer, False, PossibleMoves(1).posX, PossibleMoves(1).posY, K, True
        Next K
    End If
End Sub

Public Function grabPoints(theBoard() As Integer, currPlayer As Integer, theEnd As Boolean, ByVal posX As Integer, ByVal posY As Integer, ByVal direction As Integer, start As Boolean, firstPass As Boolean, Optional ByVal ogX As Integer, Optional ByVal ogY As Integer) As Integer
'This function returns the theoretical amount of points gained if placed at a particular location.

    Dim previousX As Integer
    Dim previousY As Integer
    Dim Temp As Integer
    
    previousX = posX
    previousY = posY

    
    Select Case direction   'Changes X position depending on direction defined by K.
        Case 1, 2, 8
            posX = posX - 1
        Case 4, 5, 6
            posX = posX + 1
    End Select
    
    Select Case direction   'Changes Y position depending on direction defined by K.
        Case 2, 3, 4
            posY = posY - 1
        Case 8, 7, 6
            posY = posY + 1
    End Select
    
    If posX <> -1 And posY <> -1 And posX <> 10 And posY <> 10 Then ' Check if position is outside board.
        If theBoard(posX, posY) = (currPlayer Mod 2 + 1) And theEnd = False Then 'Ensure distance traversed > 1
                Temp = grabPoints(theBoard(), currPlayer, theEnd, posX, posY, direction, False, firstPass)
        ElseIf theBoard(posX, posY) = currPlayer And theEnd = False Then 'Marks the end of a line.
            If start = False Then
                theEnd = True
            End If
        End If
        
        'If a valid play, switch all pieces in between to current player's.
        
        If theEnd = True Then
            
            If Not start Then
                grabPoints = Temp + 1
            ElseIf start = True Then
                If firstPass = True Then
                    grabPoints = Temp + 1
                    firstPass = False
                Else
                    grabPoints = Temp
                End If
            End If

        End If
    End If

End Function

Public Sub updateBoard(theBoard() As Integer, boardForm As Form, isPlayable As Boolean, currPlayer As Integer, gameTimer As Integer, gameWinner As Integer)
'Updates the following components of the board:
'Active player's colour shade, refreshing grid, point labels
'Also checks for endgame.
    
    Const GREEN = &HFF00&
    Const YELLOW = &HFFFF&
    Dim inTopFive As Boolean
    
    inTopFive = False
    refreshBoard theBoard(), boardForm.grdBoard, isPlayable, currPlayer
    boardForm.lblTurn.Caption = "It is " & Trim$(Player(currPlayer).name) & "'s turn."
    boardForm.lblPoints1.Caption = Player(1).points
    boardForm.lblPoints2.Caption = Player(2).points
    
    boardForm.shpRectangle(currPlayer).FillColor = GREEN
    boardForm.shpRectangle(currPlayer Mod 2 + 1).FillColor = YELLOW
    
    'End game by disabling grid, display messages to the user.
    
    If isPlayable = False Then
        boardForm.grdBoard.Enabled = False
        If Player(1).points > Player(2).points Then
            boardForm.lblTurn.Caption = Trim$(Player(1).name) & " wins!"
            gameWinner = 1
            If Player(2).isComputer > 0 Then
                checkScores gameTimer, gameWinner, inTopFive
            End If
        Else
            boardForm.lblTurn.Caption = Trim$(Player(2).name) & " wins!"
            gameWinner = 2
        End If
        
        'Let the user know if they placed in the leaderboards.
        
        If inTopFive Then
            MsgBox "You made it into the leaderboards! Check the high scores for your ranking.", vbInformation, "High Score!"
        End If
    End If
    
End Sub

Public Function playablePiece(theBoard() As Integer, ByVal currPlayer As Integer, theEnd As Boolean, ByVal posX As Integer, ByVal posY As Integer, ByVal direction As Integer, start As Boolean, FullStop As Boolean) As Boolean
'Modified searchReplace to return a boolean value for whether or not the move is valid.
'Differs from playablePiece in method of determining valid move.
'checkEndGame searches for a blank (legal) spot from an already valid location.
'playablePiece searches for the current player's piece from a blank spot.

    Select Case direction   'Changes X position depending on direction defined by exterior loop.
        Case 1, 2, 8
            posX = posX - 1
        Case 4, 5, 6
            posX = posX + 1
    End Select
    
    Select Case direction   'Changes Y position depending on direction defined by exterior loop.
        Case 2, 3, 4
            posY = posY - 1
        Case 8, 7, 6
            posY = posY + 1
    End Select
    
    If posX > -1 And posY > -1 And posX < 10 And posY < 10 Then ' Check if position is outside board.
    
        If theBoard(posX, posY) = (currPlayer Mod 2 + 1) And theEnd = False Then 'Ensure distance moved > 1
            playablePiece = playablePiece(theBoard(), currPlayer, theEnd, posX, posY, direction, False, FullStop)
        ElseIf theBoard(posX, posY) = currPlayer And theEnd = False Then ' Valid position found if distance traversed > 1
            If start = False Then
                theEnd = True
            End If
        End If
        
        If theEnd = True Then
            playablePiece = True
            FullStop = True 'Stop all further searches.
        Else
            playablePiece = False
        End If
        
    End If
    
End Function
Public Sub BubbleSortPoints(Arr() As Moves, NumRec As Integer)
    
    'Sorts possible points in ascending order.
    
    Dim K As Integer
    Dim N As Integer
    Dim Count As Integer
    Dim Temp As Moves
    
    For K = (NumRec - 1) To 1 Step -1
        For N = 1 To K
            If Arr(K).points < Arr(K + 1).points Then
                    
                Temp = Arr(K)
                Arr(K) = Arr(K + 1)
                Arr(K + 1) = Temp
                
                
            End If
        Next N
    Next K
    
End Sub

Public Sub checkScores(gameTimer As Integer, gameWinner As Integer, inTopFive As Boolean)
    Const NUMRECS = 6
    Dim playerStats As HighScore
    Dim clearStats As HighScore
    Dim K As Integer
    Dim done As Boolean
    
    K = 0
    inTopFive = False
    
    'Initialize Highscore variable.
    playerStats.intTime = gameTimer
    playerStats.name = Player(gameWinner).name
    playerStats.points = Player(gameWinner).points
    
    ReadFileREC HighScores(), App.Path & "\highscore.rec", lenHighScore, NUMRECS
    
    If playerStats.intTime < HighScores(5).intTime Or HighScores(5).intTime = 32767 Then 'Or HighScores(5).intTime = -1 Then 'Check for -1 because record may not be fully populated
        inTopFive = True
        HighScores(6) = playerStats
        
        bubbleSortScore HighScores(), NUMRECS
        SaveFile App.Path & "\highscore.rec", HighScores(), NUMRECS, lenHighScore
    
    End If

    
    
    
End Sub

Public Sub bubbleSortScore(Arr() As HighScore, NumRec As Integer)
    'Sorts leaderboard scores in ascending order.
    
    Dim K As Integer
    Dim N As Integer
    Dim Count As Integer
    Dim Temp As HighScore
    
    For K = (NumRec - 1) To 1 Step -1
        For N = 1 To K
            If Arr(N).intTime > Arr(N + 1).intTime Then
                    
                Temp = Arr(N)
                Arr(N) = Arr(N + 1)
                Arr(N + 1) = Temp
                
                
            End If
        Next N
    Next K
End Sub
