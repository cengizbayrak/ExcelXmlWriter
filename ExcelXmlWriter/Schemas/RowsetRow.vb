Imports ExcelXmlWriter
Imports System.Xml

Namespace ExcelXmlWriter.Schemas
	Public NotInheritable Class RowsetRow
		Implements IWriter
		Private _columns As RowsetColumnCollection

		Public ReadOnly Property Columns() As RowsetColumnCollection
			Get
				If Me._columns Is Nothing Then
					Me._columns = New RowsetColumnCollection()
				End If
				Return Me._columns
			End Get
		End Property

		Public Sub New()
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("z", "row", "#RowsetSchema")
            If Me._columns IsNot Nothing Then
                DirectCast(Me._columns, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
