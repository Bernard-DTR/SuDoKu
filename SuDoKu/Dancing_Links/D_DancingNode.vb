Option Strict On
Option Explicit On

' Class implementing matrix "1" element as data object.
' Copyright (c) 2006 Miran Uhan
Namespace DancingLink

  Friend Class D_DancingNode

    ''' <summary>
    ''' Pointer to the left node.
    ''' </summary>
    Private _left As D_DancingNode
    ''' <summary>
    ''' Pointer to the right node.
    ''' </summary>
    Private _right As D_DancingNode
    ''' <summary>
    ''' Pointer to the upper node.
    ''' </summary>
    Private _upper As D_DancingNode
    ''' <summary>
    ''' Pointer to the lower node.
    ''' </summary>
    Private _lower As D_DancingNode
    ''' <summary>
    ''' Pointer to the column header.
    ''' </summary>
    Private _header As E_DancingColumn
    ''' <summary>
    ''' Row number of this node. Only used for creating name,
    ''' has no role in solving.
    ''' </summary>
    Private _row As Integer
    ''' <summary>
    ''' Column number of this node. Only used for creating name,
    ''' has no role in solving.
    ''' </summary>
    Private _column As Integer

    ''' <summary>
    ''' Create self-referential node.
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="column"></param>
    Friend Sub New(ByVal row As Integer, ByVal column As Integer)
      _left = Me
      _right = Me
      _upper = Me
      _lower = Me
      _header = Nothing
      _row = row
      _column = column
    End Sub

    ''' <summary>
    ''' Get or set node being left to this node.
    ''' </summary>
    Friend Property Left() As D_DancingNode
      Get
        Return _left
      End Get
      Set(ByVal value As D_DancingNode)
        _left = value
      End Set
    End Property

    ''' <summary>
    ''' Get or set node being right to this node.
    ''' </summary>
    Friend Property Right() As D_DancingNode
      Get
        Return _right
      End Get
      Set(ByVal value As D_DancingNode)
        _right = value
      End Set
    End Property

    ''' <summary>
    ''' Get or set node being upper to this node.
    ''' </summary>
    Friend Property Upper() As D_DancingNode
      Get
        Return _upper
      End Get
      Set(ByVal value As D_DancingNode)
        _upper = value
      End Set
    End Property

    ''' <summary>
    ''' Get or set node being lower to this node.
    ''' </summary>
    Friend Property Lower() As D_DancingNode
      Get
        Return _lower
      End Get
      Set(ByVal value As D_DancingNode)
        _lower = value
      End Set
    End Property

    ''' <summary>
    ''' Get or set header node for this node.
    ''' </summary>
    Friend Property Header() As E_DancingColumn
      Get
        Return _header
      End Get
      Set(ByVal value As E_DancingColumn)
        _header = value
      End Set
    End Property

    ''' <summary>
    ''' Get or set row number for this node.
    ''' </summary>
    Friend Property Row() As Integer
      Get
        Return _row
      End Get
      Set(ByVal value As Integer)
        _row = value
      End Set
    End Property

    ''' <summary>
    ''' Get or set column number for this node.
    ''' </summary>
    Friend Property Column() As Integer
      Get
        Return _column
      End Get
      Set(ByVal value As Integer)
        _column = value
      End Set
    End Property

    ''' <summary>
    ''' For testing purpose only. Verifies weather node fits into
    ''' this row and this column.
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="column"></param>
    Friend Function Verify(ByVal row As Integer,
           ByVal column As Integer) As Boolean
      Return (row = _row) And (column = _column)
    End Function

    ''' <summary>
    ''' Return name of node in form "Row r#, column c#"
    ''' or "NULL" if node is not set.
    ''' </summary>
    ''' <param name="node"></param>
    Friend Function Name(ByVal node As D_DancingNode) As String
      If (IsNothing(node) = True) Then
        Return "NULL"
      Else
        Return node.Name()
      End If
    End Function

    ''' <summary>
    ''' Return name of this node in form "Row r#, column c#"
    ''' </summary>
    Friend Function Name() As String
      Return "row" & _row & ", column" & _column
    End Function

    ''' <summary>
    ''' Return string with this node name and descriptors
    ''' with names of all neighbour nodes and header.
    ''' </summary>
    Public Overrides Function ToString() As String
      Return "Node(" & Name() &
                   "), left(" & Name(Left()) &
                   "), right(" & Name(Right()) &
                   "), upper(" & Name(Upper()) &
                   "), lower(" & Name(Lower()) &
                   "), header(" & Name(Header()) & ")"
    End Function

  End Class

End Namespace