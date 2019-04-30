Imports System.CodeDom
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetAutoFilter
        Implements IReader
        Implements IWriter
        Implements ICodeWriter
        Private _range As String

        Public Property Range() As String
            Get
                Return Me._range
            End Get
            Set(value As String)
                Me._range = value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._range IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Range", Me._range)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetAutoFilter.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._range = Util.GetAttribute(element, "Range", "urn:schemas-microsoft-com:office:excel")
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "AutoFilter", "urn:schemas-microsoft-com:office:excel")
            If Me._range IsNot Nothing Then
                writer.WriteAttributeString("Range", "urn:schemas-microsoft-com:office:excel", Me._range)
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "AutoFilter", "urn:schemas-microsoft-com:office:excel")
        End Function
    End Class
End Namespace
