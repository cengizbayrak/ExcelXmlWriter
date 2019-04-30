Imports System.CodeDom
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class WorksheetPageFooter
		Implements IWriter
		Implements IReader
		Implements ICodeWriter
		Private _margin As Single = 0.5F

		Private _data As String

		Public Property Data() As String
			Get
				Return Me._data
			End Get
			Set
				Me._data = value
			End Set
		End Property

		Public Property Margin() As Single
			Get
				Return Me._margin
			End Get
			Set
				Me._margin = value
			End Set
		End Property

		Friend Sub New()
		End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._data IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Data", Me._data)
            End If
            If CDbl(Me._margin) <> 0.5 Then
                Util.AddAssignment(method, targetObject, "Margin", Me._margin)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetPageFooter.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._margin = Util.GetAttribute(element, "Margin", "urn:schemas-microsoft-com:office:excel", 0.5F)
            Me._data = Util.GetAttribute(element, "Data", "urn:schemas-microsoft-com:office:excel")
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "Footer", "urn:schemas-microsoft-com:office:excel")
            If Me._data IsNot Nothing Then
                writer.WriteAttributeString("Data", "urn:schemas-microsoft-com:office:excel", Me._data)
            End If
            If CDbl(Me._margin) <> 0.5 Then
                writer.WriteAttributeString("Margin", "urn:schemas-microsoft-com:office:excel", Me._margin.ToString(CultureInfo.InvariantCulture))
            End If
            writer.WriteEndElement()
        End Sub

		Friend Shared Function IsElement(element As XmlElement) As Boolean
			Return Util.IsElement(element, "Footer", "urn:schemas-microsoft-com:office:excel")
		End Function
	End Class
End Namespace
