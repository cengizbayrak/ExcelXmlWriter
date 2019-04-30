Imports ExcelXmlWriter
Imports System.Collections
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter.Schemas
	Public NotInheritable Class RowsetRowCollection
		Inherits CollectionBase
		Implements IWriter
		Public Default ReadOnly Property Item(index As Integer) As RowsetRow
			Get
				Return DirectCast(MyBase.InnerList(index), RowsetRow)
			End Get
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(row As RowsetRow) As Integer
			Return MyBase.InnerList.Add(row)
		End Function

		Public Function Add() As RowsetRow
			Dim rowsetRow As New RowsetRow()
			Me.Add(rowsetRow)
			Return rowsetRow
		End Function

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

		Public Function Contains(item As RowsetRow) As Boolean
			Return MyBase.InnerList.Contains(item)
		End Function

		Public Sub CopyTo(array As RowsetRow(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As RowsetRow) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Sub Insert(index As Integer, item As RowsetRow)
			MyBase.InnerList.Insert(index, item)
		End Sub

		Public Sub Remove(item As RowsetRow)
			MyBase.InnerList.Remove(item)
		End Sub
	End Class
End Namespace
