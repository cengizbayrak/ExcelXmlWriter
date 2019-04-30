Imports ExcelXmlWriter
Imports System.Xml

Namespace ExcelXmlWriter.Schemas
	Public NotInheritable Class Schema
		Implements IWriter
		Private _types As SchemaTypeCollection

		Public ReadOnly Property Types() As SchemaTypeCollection
			Get
				If Me._types Is Nothing Then
					Me._types = New SchemaTypeCollection()
				End If
				Return Me._types
			End Get
		End Property

		Public Sub New()
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Schema", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
            If Me._types IsNot Nothing Then
                DirectCast(Me._types, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
