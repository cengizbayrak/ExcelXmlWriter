Imports ExcelXmlWriter
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter.Schemas
	Public NotInheritable Class ElementType
		Inherits SchemaType
		Implements IWriter
		Private _name As String

		Private _content As SchemaContent

		Private _attributes As AttributeCollection

		Public ReadOnly Property Attributes() As AttributeCollection
			Get
				If Me._attributes Is Nothing Then
					Me._attributes = New AttributeCollection()
				End If
				Return Me._attributes
			End Get
		End Property

		Public Property Content() As SchemaContent
			Get
				Return Me._content
			End Get
			Set
				Me._content = value
			End Set
		End Property

		Public Property Name() As String
			Get
				Return Me._name
			End Get
			Set
				Me._name = value
			End Set
		End Property

		Public Sub New()
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "ElementType", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
            If Me._name IsNot Nothing Then
                writer.WriteAttributeString("s", "name", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882", Me._name)
            End If
            If Me._content <> SchemaContent.NotSet Then
                writer.WriteAttributeString("s", "content", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882", Me._content.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._attributes IsNot Nothing Then
                DirectCast(Me._attributes, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
