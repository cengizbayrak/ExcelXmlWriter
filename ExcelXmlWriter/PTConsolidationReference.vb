Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PTConsolidationReference
		Implements IWriter
		Private _fileName As String

		Private _reference As String

		Public Property FileName() As String
			Get
				Return Me._fileName
			End Get
			Set
				Me._fileName = value
			End Set
		End Property

		Public Property Reference() As String
			Get
				Return Me._reference
			End Get
			Set
				Me._reference = value
			End Set
		End Property

		Friend Sub New()
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "ConsolidationReference", "urn:schemas-microsoft-com:office:excel")
            If Me._fileName IsNot Nothing Then
                writer.WriteElementString("FileName", "urn:schemas-microsoft-com:office:excel", Me._fileName)
            End If
            If Me._reference IsNot Nothing Then
                writer.WriteElementString("Reference", "urn:schemas-microsoft-com:office:excel", Me._reference)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
