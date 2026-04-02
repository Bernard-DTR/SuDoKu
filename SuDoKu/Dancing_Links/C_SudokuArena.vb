' Implements sudoku puzzle solver with dancing links.
' Copyright (c) 2006 Miran Uhan
Namespace DancingLink

  Friend Class C_SudokuArena
    Inherits B_DancingArena

    Friend Enum SudokuStore As Integer
      None = 0
      First = 1
      All = 2
    End Enum

    Private _solutions As Integer = 0
    Private ReadOnly _solutionList As New Collection
    Private ReadOnly _size As Integer = 0
    Private _store As SudokuStore = SudokuStore.All

    ''' <summary>
    ''' Create solver with given task matrix.
    ''' </summary>
    ''' <param name="puzzle">Task matrix.</param>
    ''' <param name="boxRows">Number of rows in box.</param>
    ''' <param name="boxCols">Number of columns in box.</param>
    Friend Sub New(
                 ByVal puzzle(,) As Int32,
                 Optional ByVal boxRows As Integer = 0,
                 Optional ByVal boxCols As Integer = 0)
      'Sudoku has 4 constraints
      '  - in each cell must be exactly one number
      '  - in each row must be all numbers
      '  - in each column must be all numbers
      '  - in each box must be all numbers

      MyBase.New(4 * puzzle.Length)
      _solutions = 0
      _size = puzzle.GetLength(0)
      If ((boxRows = 0) And (boxCols = 0)) Then
        boxRows = CInt(Math.Round(Math.Sqrt(_size)))
        boxCols = boxRows
      ElseIf ((boxRows = 0) And (boxCols <> 0)) Then
        boxRows = _size \ boxCols
      ElseIf ((boxRows <> 0) And (boxCols = 0)) Then
        boxCols = _size \ boxRows
      End If
      Dim positions(3) As Integer
      Dim known As New Collection
      For row As Integer = 0 To _size - 1
        For col As Integer = 0 To _size - 1
          Dim boxRow As Integer = row \ boxRows
          Dim boxCol As Integer = col \ boxCols
          For digit As Integer = 0 To _size - 1
            Dim isGiven As Boolean = (puzzle(row, col) = (digit + 1))
            positions(0) = 1 + (row * _size + col)
            positions(1) = 1 + puzzle.Length + (row * _size + digit)
            positions(2) = 1 + 2 * puzzle.Length + (col * _size + digit)
            positions(3) = 1 + 3 * puzzle.Length +
                                 ((boxRow * boxRows + boxCol) * _size + digit)
            Dim newRow As D_DancingNode = AddRow(positions)
            If (isGiven = True) Then
              known.Add(newRow)
            End If
          Next digit
        Next col
      Next row
      RemoveKnown(known)
    End Sub

    ''' <summary>
    ''' Return size of puzzle, i.e. number of rows and columns in puzzle.
    ''' </summary>
    Friend ReadOnly Property Size() As Integer
      Get
        Return _size
      End Get
    End Property

    ''' <summary>
    ''' Set or get way of storing solutions.
    ''' </summary>
    Friend Property StoreSolutions() As SudokuStore
      Get
        Return _store
      End Get
      Set(ByVal value As SudokuStore)
        _store = value
      End Set
    End Property

    ''' <summary>
    ''' Return first solution found.
    ''' </summary>
    Friend ReadOnly Property Solution() As Array
      Get
        Return CType(_solutionList.Item(1), Array)
      End Get
    End Property

    ''' <summary>
    ''' Return i-th solution.
    ''' </summary>
    ''' <param name="index"></param>
    Friend ReadOnly Property Solution(ByVal index As Integer) As Array
      Get
        Return CType(_solutionList.Item(index), Array)
      End Get
    End Property

    ''' <summary>
    ''' Return number of solutions found.
    ''' </summary>
    Friend ReadOnly Property Solutions() As Integer
      Get
        Return _solutions
      End Get
    End Property

    ''' <summary>
    ''' Create solution matrix.
    ''' </summary>
    ''' <param name="rows"></param>
    Friend Overrides Sub HandleSolution(ByVal rows() As D_DancingNode)
      Dim _solution(_size - 1, _size - 1) As Integer
      For i As Integer = 0 To rows.Length() - 1
        If (IsNothing(rows(i)) = False) Then
          Dim value As Integer = rows(i).Row() - 1
          Dim digit, row, col As Integer
          digit = value Mod _size + 1
          value \= _size
          col = value Mod _size
          value \= _size
          row = value Mod _size
          _solution(row, col) = digit
        End If
      Next i
      _solutions += 1
      Select Case _store
        Case SudokuStore.All : _solutionList.Add(_solution)
        'Case SudokuStore.All : If _solutionList.Count < 9  Then _solutionList.Add(_solution)
        Case SudokuStore.First : If _solutionList.Count = 0 Then _solutionList.Add(_solution)
        Case SudokuStore.None
      End Select
    End Sub
  End Class
End Namespace