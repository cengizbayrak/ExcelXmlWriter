Imports System.CodeDom
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetStyleInterior
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _color As String

        Private _pattern As StyleInteriorPattern

        Public Property Color() As String
            Get
                Return Me._color
            End Get
            Set(value As String)
                Me._color = value
            End Set
        End Property

        Public Property Pattern() As StyleInteriorPattern
            Get
                Return Me._pattern
            End Get
            Set(value As StyleInteriorPattern)
                Me._pattern = value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo

            If Me._color IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Color", Me._color)
            End If
            If Me._pattern <> StyleInteriorPattern.NotSet Then
                Util.AddAssignment(method, targetObject, "Pattern", CType(Me._pattern, [Enum]))
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetStyleInterior.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._color = Util.GetAttribute(element, "Color", "urn:schemas-microsoft-com:office:spreadsheet")
            Dim attribute As String = element.GetAttribute("Pattern", "urn:schemas-microsoft-com:office:spreadsheet")
            If attribute IsNot Nothing AndAlso attribute.Length <> 0 Then
                'this._pattern = (int)((StyleInteriorPattern)Enum.Parse(typeof(StyleInteriorPattern), attribute, true));
                Me._pattern = CType([Enum].Parse(GetType(StyleInteriorPattern), attribute, True), StyleInteriorPattern)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Interior", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._color IsNot Nothing Then
                writer.WriteAttributeString("Color", "urn:schemas-microsoft-com:office:spreadsheet", Me._color)
            End If
            If Me._pattern <> StyleInteriorPattern.NotSet Then
                writer.WriteAttributeString("Pattern", "urn:schemas-microsoft-com:office:spreadsheet", Me._pattern.ToString(CultureInfo.InvariantCulture))
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Interior", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
