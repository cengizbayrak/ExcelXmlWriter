Imports Bc
Imports System.Xml

Namespace ExcelXmlWriter.Schemas
	Public NotInheritable Class Attribute
		Implements IWriter
		Private _type As String

		Public Property Type() As String
			Get
				Return Me._type
			End Get
			Set
				Me._type = value
			End Set
		End Property

		Public Sub New()
		End Sub

		Public Sub New(type As String)
			Me._type = type
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "attribute", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
            If Me._type IsNot Nothing Then
                writer.WriteAttributeString("s", "type", "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882", Me._type)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
