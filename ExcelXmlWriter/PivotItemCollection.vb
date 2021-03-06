Imports System.Collections
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PivotItemCollection
		Inherits CollectionBase
		Implements IWriter
		Public Default ReadOnly Property Item(index As Integer) As PivotItem
			Get
				Return DirectCast(MyBase.InnerList(index), PivotItem)
			End Get
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(pivotItem As PivotItem) As Integer
			Return MyBase.InnerList.Add(pivotItem)
		End Function

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

		Public Function Contains(item As PivotItem) As Boolean
			Return MyBase.InnerList.Contains(item)
		End Function

		Public Sub CopyTo(array As PivotItem(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As PivotItem) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Sub Insert(index As Integer, item As PivotItem)
			MyBase.InnerList.Insert(index, item)
		End Sub

		Public Sub Remove(item As PivotItem)
			MyBase.InnerList.Remove(item)
		End Sub
	End Class
End Namespace
