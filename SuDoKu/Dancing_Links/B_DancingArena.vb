' Copyright (c) 2006 Miran Uhan
Namespace DancingLink

  Friend MustInherit Class B_DancingArena
    'Friend       Elément accessible à partir du même assembly, mais pas à partir de l’extérieur de l’assembly
    'MustInherit  La classe est destinée à être utilisée comme classe de base uniquement

    ''' <summary>
    ''' Nœud racine avec ligne = 0 et colonne = 0, pointant vers la liste d'en-tête de colonne.
    ''' </summary>
    Private ReadOnly _root As E_DancingColumn

    ''' <summary>
    ''' Nombre total de rangées/colonnes dans la matrice. Ne joue aucun rôle pendant le solutionnement; uniquement pour info.
    ''' </summary>
    Private _rows As Integer = 0
    Private ReadOnly _columns As Integer = 0

    ''' <summary>
    ''' Index of first solution row. Used when some solution rows are given.
    ''' Index de la première ligne de solution. Utilisé lorsque certaines lignes de solution sont données
    ''' </summary>
    Private _initial As Integer = 0
    ''' <summary>
    ''' Array holding solution rows with doubly linked lists.
    ''' Array size is defined when creating new arena and
    ''' equals to number of primary columns.
    ''' </summary>
    Private ReadOnly _solutionRows() As D_DancingNode
    ''' <summary>
    ''' Array for pointers to all column headers. Used for easier
    ''' scan over all the columns.
    ''' </summary>
    Private ReadOnly _headerColumns() As E_DancingColumn
    ''' <summary>
    ''' Counter of node updates during solving. Used for evaluation
    ''' of algorithm speed and has no role in solving.
    ''' </summary>
    Private _updates As Long = 0

    ''' <summary>
    ''' Create header with circular doubly linked list of
    ''' column headers for the matrix and root element
    ''' pointing to the leftmost and rightmost column header.
    ''' </summary>
    ''' <param name="primary">
    ''' Number of primary columns in matrix.
    ''' </param>
    ''' <param name="secondary">
    ''' Number of secondary columns in matrix.
    ''' </param>
    Friend Sub New(ByVal primary As Integer, ByVal secondary As Integer)
      _root = New E_DancingColumn(0)
      _rows = 0
      _columns = primary + secondary
      _initial = 0
      _updates = 0
      'Only primary columns form solution.
      ReDim _solutionRows(primary - 1)
      ReDim _headerColumns(primary + secondary - 1)
      For i As Integer = 0 To primary - 1
        _solutionRows(i) = Nothing
        _headerColumns(i) = New E_DancingColumn(i + 1) With {
          .Right = Nothing
        }
        If (i > 0) Then
          'Connecting column to it's neighbours.
          _headerColumns(i).Left = _headerColumns(i - 1)
          _headerColumns(i - 1).Right = _headerColumns(i)
        End If
      Next i
      'Connecting first and last primary column to root node.
      _headerColumns(0).Left = _root
      _root.Right = _headerColumns(0)
      _headerColumns(primary - 1).Right = _root
      _root.Left = _headerColumns(primary - 1)
      'Adding self referential secondary columns.
      If secondary > 0 Then
        For i As Integer = 0 To secondary - 1
          _headerColumns(primary + i) = New E_DancingColumn(primary + i + 1)
        Next i
      End If
    End Sub

    ''' <summary>
    ''' Create header with circular doubly linked list of
    ''' column headers for the matrix and root element
    ''' pointing to the leftmost column header.
    ''' </summary>
    ''' <param name="columns">
    ''' Number of primary columns in matrix.
    ''' </param>
    Friend Sub New(ByVal columns As Integer)
      Me.New(columns, 0)
    End Sub

    ''' <summary>
    ''' Get root element.
    ''' </summary>
    Friend ReadOnly Property Root() As E_DancingColumn
      Get
        Return _root
      End Get
    End Property

    ''' <summary>
    ''' Get leftmost column header.
    ''' </summary>
    Friend ReadOnly Property FirstColumn() As E_DancingColumn
      Get
        Return CType(_root.Right(), E_DancingColumn)
      End Get
    End Property

    ''' <summary>
    ''' Get rightmost column header.
    ''' </summary>
    Friend ReadOnly Property LastColumn() As E_DancingColumn
      Get
        Return CType(_root.Left(), E_DancingColumn)
      End Get
    End Property

    ''' <summary>
    ''' Return total number of rows in matrix.
    ''' </summary>
    Friend ReadOnly Property Rows() As Integer
      Get
        Return _rows
      End Get
    End Property

    ''' <summary>
    ''' Get total number of columns in matrix.
    ''' </summary>
    Friend ReadOnly Property Columns() As Integer
      Get
        Return _columns
      End Get
    End Property

    ''' <summary>
    ''' Get total number of updates, ie number of times program
    ''' covered or uncovered some node during solving.
    ''' </summary>
    Friend ReadOnly Property Updates() As Long
      Get
        Return _updates
      End Get
    End Property

    ''' <summary>
    ''' Implements Knuth's algorithm for covering the column.
    ''' </summary>
    ''' <param name="column"></param>
    Friend Sub CoverColumn(ByVal column As E_DancingColumn)
      'Console.WriteLine($"Couvre colonne {column.Column}")
      'Exclude column header node from the list.
      _updates += 1
      column.Left().Right = column.Right()
      column.Right().Left = column.Left()
      'Now do for each row in excluded column...
      Dim row As D_DancingNode = column.Lower()
      While (Equals(row, column) = False)
        '... cover all nodes in a row...
        Dim col As D_DancingNode = row.Right()
        While (Equals(col, row) = False)
          '... by excluding nodes from their columns.
          _updates += 1
          col.Upper().Lower = col.Lower()
          col.Lower().Upper = col.Upper()
          col.Header().DecRows()
          col = col.Right()
        End While
        row = row.Lower()
      End While
    End Sub

    ''' <summary>
    ''' Implements Knuth's algorithm for uncovering the column.
    ''' </summary>
    ''' <param name="column"></param>
    Friend Sub UncoverColumn(ByVal column As E_DancingColumn)
      'Console.WriteLine($"Découvre colonne {column.Column}")
      'For each row in excluded column...
      Dim row As D_DancingNode = column.Upper()
      While (Equals(row, column) = False)
        '... uncover all nodes in a row...
        Dim col As D_DancingNode = row.Left()
        While (Equals(col, row) = False)
          '... by connecting nodes to their columns.
          _updates += 1
          col.Upper().Lower = col
          col.Lower().Upper = col
          col.Header().IncRows()
          col = col.Left()
        End While
        row = row.Upper()
      End While
      'Return column header node to the list.
      _updates += 1
      column.Left().Right = column
      column.Right().Left = column
    End Sub

    ''' <summary>
    ''' Add row of circular doubly linked nodes defined with
    ''' their column positions.
    ''' </summary>
    ''' <param name="positions"></param>
    Friend Function AddRow(ByVal positions() As Integer) As D_DancingNode
      Dim result As D_DancingNode = Nothing
      If positions.Length() > 0 Then
        Dim found As Boolean 'result of searching of column
        Dim thisNode As D_DancingNode
        Dim prevNode As D_DancingNode = Nothing
        _rows += 1 'counter of number of rows in matrix
        For i As Integer = 0 To positions.Length() - 1
          If ((IsNothing(positions(i)) = False) And (positions(i) > 0)) Then
            'Create new node.
            'Connect new node to the left node in the same row.
            thisNode = New D_DancingNode(_rows, positions(i)) With {
              .Left = prevNode,
              .Right = Nothing
            }
            If (IsNothing(prevNode) = False) Then
              'Not the first node in row,
              prevNode.Right = thisNode
            Else
              'This is first node in this row.
              result = thisNode
            End If
            'Search for column with corresponding column number.
            found = False
            For Each col As E_DancingColumn In _headerColumns
              If col.Column = positions(i) Then
                'Column header found. Remember this header.
                thisNode.Header = col
                'Add new node to column we found. Column class will
                'provide connections to upper and lower nodes.
                col.AddNode(thisNode)
                found = True
              End If
            Next col
            If (found = False) Then
              Console.WriteLine("Can't find header for {0}", positions(i))
            End If
            'Remember new node is last added.
            prevNode = thisNode
            'Make circular connection.
            result.Left = prevNode
            prevNode.Right = result
          End If
        Next i
      End If
      Return result
    End Function

    ''' <summary>
    ''' Select next column to cover. Finds column with minimum rows.
    ''' </summary>
    Friend Function NextColumn() As E_DancingColumn
      Dim result As E_DancingColumn = CType(_root.Right(), E_DancingColumn)
      Dim minRows As Integer = result.Rows()
      Dim scaner As E_DancingColumn = CType(_root.Left(), E_DancingColumn)
      Do While (Equals(scaner, _root.Right()) = False)
        If scaner.Rows() < minRows Then
          result = scaner
          minRows = scaner.Rows()
        End If
        scaner = CType(scaner.Left(), E_DancingColumn)
      Loop
      Return result
    End Function

    ''' <summary>
    ''' Implements Knuth's DLX algorithm.
    ''' </summary>
    ''' <param name="index"></param>
    Friend Sub SolveRecurse(ByVal index As Integer)
      If (Equals(_root, _root.Right()) = True) Then
        'No more columns, we found one of solutions.
        HandleSolution(_solutionRows)
      Else
        'Select next column using some selection algorithm.
        Dim nextCol As E_DancingColumn = NextColumn()
        'Exclude selected column from the matrix.
        CoverColumn(nextCol)
        'Try for each row in selected column...
        Dim row As D_DancingNode = nextCol.Lower()
        While (Equals(row, nextCol) = False)
          'Add row to solution array.
          _solutionRows(index) = row
          ' TRACE ICI
          'TracePlacement(row)
          '... exclude all columns covered by this row ...
          Dim col As D_DancingNode = row.Right()
          While (Equals(col, row) = False)
            CoverColumn(col.Header())
            col = col.Right()
          End While
          '... and try to solve reduced matrix.
          SolveRecurse(index + 1)
          'Now restore all columns covered by this row ...
          col = row.Left()
          While (Equals(col, row) = False)
            UncoverColumn(col.Header())
            col = col.Left()
          End While
          '... and remove row from solution array.
          _solutionRows(index) = Nothing
          row = row.Lower()
        End While
        'Return excluded column back to list.
        UncoverColumn(nextCol)
      End If
    End Sub

    ''' <summary>
    ''' Start recursive solving with initial step.
    ''' </summary>
    Friend Sub Solve()
      _updates = 0
      SolveRecurse(_initial)
    End Sub

    ''' <summary>
    ''' Remove known rows (partial solution known or initial position).
    ''' Global counter of solution rows is used so partial adding of
    ''' solution rows is possible. Care should be taken not to add
    ''' the same solution row twice as space in solution rows array will
    ''' be ocupied and then all solutions will not fit into array
    ''' resulting in overflow.
    ''' </summary>
    ''' <param name="solutions"></param>
    Friend Sub RemoveKnown(ByVal solutions As Collection)
      Dim row As D_DancingNode
      For Each row In solutions
        _solutionRows(_initial) = row
        _initial += 1
        CoverColumn(row.Header())
        Dim col As D_DancingNode = row.Right()
        While (Equals(col, row) = False)
          CoverColumn(col.Header())
          col = col.Right()
        End While
      Next row
    End Sub

    ''' <summary>
    ''' Childs of this class must implement algorithm for handling result.
    ''' </summary>
    ''' <param name="rows"></param>
    Friend MustOverride Sub HandleSolution(ByVal rows() As D_DancingNode)

    ''' <summary>
    ''' For testing purpose only, shows all columns and nodes with connections.
    ''' </summary>
    Friend Sub ShowState()
      Dim col As E_DancingColumn = FirstColumn()
      While (Equals(col, Root()) = False)
        Console.WriteLine("C : {0}", col.ToString())
        Dim row As D_DancingNode = col.Lower()
        While (Equals(row, col) = False)
          Console.WriteLine("   R : {0}", row.ToString())
          row = row.Lower()
        End While
        col = CType(col.Right(), E_DancingColumn)
      End While
    End Sub

    Private Function DecodeDLXRow(node As D_DancingNode) As (Integer, Integer, Integer)
      Dim v As Integer = node.Row - 1
      Dim digit As Integer = v Mod 9 + 1
      v \= 9
      Dim col As Integer = v Mod 9
      v \= 9
      Dim row As Integer = v Mod 9
      Return (row, col, digit)
    End Function

    Private Sub TracePlacement(node As D_DancingNode)
      ' Déclaration des variables
      Dim r As Integer
      Dim c As Integer
      Dim v As Integer

      ' Récupération du tuple
      Dim decoded As (Integer, Integer, Integer)
      decoded = DecodeDLXRow(node)
      r = decoded.Item1
      c = decoded.Item2
      v = decoded.Item3

      Dim i As Integer = c + r * 9
      Console.WriteLine($"DLX place {v} en r{r + 1}c{c + 1}")

      ' Mise à jour de votre tableau U
      U(i, 2) = v.ToString()

      ' Mise à jour des candidats
      For k As Integer = 0 To 80
        If U(k, 3).Contains(v.ToString()) Then
          U(k, 3) = U(k, 3).Replace(v.ToString(), " ")
        End If
      Next
    End Sub

  End Class
End Namespace