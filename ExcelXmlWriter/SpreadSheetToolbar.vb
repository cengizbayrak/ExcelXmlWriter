Imports System.CodeDom
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class SpreadSheetToolbar
		Implements IWriter
		Implements IReader
		Implements ICodeWriter
		Private _hidden As Boolean

		Public Property Hidden() As Boolean
			Get
				Return Me._hidden
			End Get
			Set
				Me._hidden = value
			End Set
		End Property

		Friend Sub New()
		End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._hidden Then
                Util.AddAssignment(method, targetObject, "Hidden", Me._hidden)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not SpreadSheetToolbar.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._hidden = element.GetAttribute("Hidden", "urn:schemas-microsoft-com:office:spreadsheet") = "1"
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("c", "Toolbar", "urn:schemas-microsoft-com:office:component:spreadsheet")
            If Me._hidden Then
                writer.WriteAttributeString("s", "Hidden", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            writer.WriteEndElement()
        End Sub

		Friend Shared Function IsElement(element As XmlElement) As Boolean
			Return Util.IsElement(element, "Toolbar", "urn:schemas-microsoft-com:office:component:spreadsheet")
		End Function
	End Class
End Namespace
