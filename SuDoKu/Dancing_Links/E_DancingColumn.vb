' Class implementing column header.
' Copyright (c) 2006 Miran Uhan
Namespace DancingLink

  Friend Class E_DancingColumn
    Inherits D_DancingNode

    ''' <summary>
    ''' Counter of nodes in this column.
    ''' </summary>
    Private _rows As Integer

    ''' <summary>
    ''' Create self-referential node with row number 0.
    ''' </summary>
    ''' <param name="column"></param>
    Friend Sub New(ByVal column As Integer)
      MyBase.New(0, column)
    End Sub

    ''' <summary>
    ''' Get number of nonzero rows in this column.
    ''' </summary>
    Friend ReadOnly Property Rows() As Integer
      Get
        Return _rows
      End Get
    End Property

    ''' <summary>
    ''' Increase number of rows for 1.
    ''' </summary>
    Friend Sub IncRows()
      _rows += 1
    End Sub

    ''' <summary>
    ''' Decrease number of rows for 1.
    ''' </summary>
    Friend Sub DecRows()
      _rows -= 1
    End Sub

    ''' <summary>
    ''' Return string with this node name and descriptors
    ''' with names of all neighbours and header and number
    ''' of rows in this column.
    ''' </summary>
    Public Overrides Function ToString() As String
      Return MyBase.ToString() & ", rows " & _rows
    End Function

    ''' <summary>
    ''' Add new row node to the end of column. Sets links
    ''' of this node, previously last node and column
    ''' header to form circular list.
    ''' </summary>
    ''' <param name="node"></param>
    Friend Sub AddNode(ByVal node As D_DancingNode)
      Dim last As D_DancingNode = Me.Upper()
      node.Upper = last
      node.Lower = Me
      last.Lower = node
      Me.Upper = node
      _rows += 1
    End Sub
  End Class
End Namespace