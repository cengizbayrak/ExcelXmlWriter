Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PivotItem
		Implements IWriter
		Private _name As String

		Private _hideDetail As Boolean

		Public Property HideDetail() As Boolean
			Get
				Return Me._hideDetail
			End Get
			Set
				Me._hideDetail = value
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

		Public Sub New(name As String)
			Me._name = name
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "PivotItem", "urn:schemas-microsoft-com:office:excel")
            If Me._name IsNot Nothing Then
                writer.WriteElementString("Name", "urn:schemas-microsoft-com:office:excel", Me._name)
            End If
            If Me._hideDetail Then
                writer.WriteElementString("HideDetail", "urn:schemas-microsoft-com:office:excel", "")
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
