Imports System.Collections
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PTLineItemCollection
		Inherits CollectionBase
		Implements IWriter
		Public Default ReadOnly Property Item(index As Integer) As PTLineItem
			Get
				Return DirectCast(MyBase.InnerList(index), PTLineItem)
			End Get
		End Property

		Friend Sub New()
		End Sub

		Public Function Add(pTLineItem As PTLineItem) As Integer
			Return MyBase.InnerList.Add(pTLineItem)
		End Function

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "PTLineItems", "urn:schemas-microsoft-com:office:excel")
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
            writer.WriteEndElement()
        End Sub

		Public Function Contains(item As PTLineItem) As Boolean
			Return MyBase.InnerList.Contains(item)
		End Function

		Public Sub CopyTo(array As PTLineItem(), index As Integer)
			MyBase.InnerList.CopyTo(array, index)
		End Sub

		Public Function IndexOf(item As PTLineItem) As Integer
			Return MyBase.InnerList.IndexOf(item)
		End Function

		Public Sub Insert(index As Integer, item As PTLineItem)
			MyBase.InnerList.Insert(index, item)
		End Sub

		Public Sub Remove(item As PTLineItem)
			MyBase.InnerList.Remove(item)
		End Sub
	End Class
End Namespace
