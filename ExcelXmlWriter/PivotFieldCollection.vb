Imports System.Collections
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PivotFieldCollection
		Inherits CollectionBase
		Implements IWriter
		Public Default ReadOnly Property Item(index As Integer) As PivotField
			Get
				Return DirectCast(MyBase.InnerList(index), PivotField)
			End Get
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(pivotField As PivotField) As Integer
			Return MyBase.InnerList.Add(pivotField)
		End Function

		Public Function Add() As PivotField
			Dim pivotField As New PivotField()
			Me.Add(pivotField)
			Return pivotField
		End Function

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

		Public Function Contains(item As PivotField) As Boolean
			Return MyBase.InnerList.Contains(item)
		End Function

		Public Sub CopyTo(array As PivotField(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As PivotField) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Sub Insert(index As Integer, field As PivotField)
			MyBase.InnerList.Insert(index, field)
		End Sub

		Public Sub Remove(field As PivotField)
			MyBase.InnerList.Remove(field)
		End Sub
	End Class
End Namespace
