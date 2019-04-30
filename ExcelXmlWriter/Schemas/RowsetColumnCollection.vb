Imports ExcelXmlWriter
Imports System.Collections
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter.Schemas
	Public NotInheritable Class RowsetColumnCollection
		Inherits CollectionBase
		Implements IWriter
		Public Default ReadOnly Property Item(index As Integer) As RowsetColumn
			Get
				Return DirectCast(MyBase.InnerList(index), RowsetColumn)
			End Get
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(column As RowsetColumn) As Integer
			Return MyBase.InnerList.Add(column)
		End Function

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

		Public Function Contains(item As RowsetColumn) As Boolean
			Return MyBase.InnerList.Contains(item)
		End Function

		Public Sub CopyTo(array As RowsetColumn(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As RowsetColumn) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Sub Insert(index As Integer, item As RowsetColumn)
			MyBase.InnerList.Insert(index, item)
		End Sub

		Public Sub Remove(item As RowsetColumn)
			MyBase.InnerList.Remove(item)
		End Sub
	End Class
End Namespace
