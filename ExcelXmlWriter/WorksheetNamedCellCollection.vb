Imports System.Collections
Imports System.Reflection

Namespace ExcelXmlWriter
	Public NotInheritable Class WorksheetNamedCellCollection
		Inherits CollectionBase
		Public Default ReadOnly Property Item(index As Integer) As String
			Get
				Return DirectCast(MyBase.InnerList(index), String)
			End Get
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(name As String) As Integer
			If name Is Nothing Then
				Throw New ArgumentNullException("name")
			End If
			Return MyBase.InnerList.Add(name)
		End Function

		Public Function Contains(item As String) As Boolean
			Return MyBase.InnerList.Contains(item)
		End Function

		Public Sub CopyTo(array As String(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As String) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Sub Insert(index As Integer, item As String)
			MyBase.InnerList.Insert(index, item)
		End Sub

		Public Sub Remove(item As String)
			MyBase.InnerList.Remove(item)
		End Sub
	End Class
End Namespace
