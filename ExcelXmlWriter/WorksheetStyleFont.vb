Imports System.CodeDom
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetStyleFont
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _fontName As String

        Private _size As Integer

        Private _bold As Boolean

        Private _underline As UnderlineStyle

        Private _italic As Boolean

        Private _strikethrough As Boolean

        Private _color As String

        Public Property Bold() As Boolean
            Get
                Return Me._bold
            End Get
            Set(value As Boolean)
                Me._bold = Value
            End Set
        End Property

        Public Property Color() As String
            Get
                Return Me._color
            End Get
            Set(value As String)
                Me._color = Value
            End Set
        End Property

        Public Property FontName() As String
            Get
                Return Me._fontName
            End Get
            Set(value As String)
                Me._fontName = Value
            End Set
        End Property

        Public Property Italic() As Boolean
            Get
                Return Me._italic
            End Get
            Set(value As Boolean)
                Me._italic = Value
            End Set
        End Property

        Public Property Size() As Integer
            Get
                Return Me._size
            End Get
            Set(value As Integer)
                Me._size = Value
            End Set
        End Property

        Public Property Strikethrough() As Boolean
            Get
                Return Me._strikethrough
            End Get
            Set(value As Boolean)
                Me._strikethrough = Value
            End Set
        End Property

        Public Property Underline() As UnderlineStyle
            Get
                Return Me._underline
            End Get
            Set(value As UnderlineStyle)
                Me._underline = Value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._bold Then
                Util.AddAssignment(method, targetObject, "Bold", Me._bold)
            End If
            If Me._italic Then
                Util.AddAssignment(method, targetObject, "Italic", Me._italic)
            End If
            If Me._underline <> UnderlineStyle.None Then
                Util.AddAssignment(method, targetObject, "Underline", CType(Me._underline, [Enum]))
            End If
            If Me._strikethrough Then
                Util.AddAssignment(method, targetObject, "StrikeThrough", Me._strikethrough)
            End If
            If Me._fontName IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "FontName", Me._fontName)
            End If
            If Me._size <> 0 Then
                Util.AddAssignment(method, targetObject, "Size", Me._size)
            End If
            If Me._color IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Color", Me._color)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetStyleFont.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._bold = element.GetAttribute("Bold", "urn:schemas-microsoft-com:office:spreadsheet") = "1"
            Me._italic = element.GetAttribute("Italic", "urn:schemas-microsoft-com:office:spreadsheet") = "1"
            Dim attribute As String = element.GetAttribute("Underline", "urn:schemas-microsoft-com:office:spreadsheet")
            If attribute IsNot Nothing AndAlso attribute.Length <> 0 Then
                Me._underline = CType([Enum].Parse(GetType(UnderlineStyle), attribute, True), UnderlineStyle)
            End If
            Me._strikethrough = element.GetAttribute("StrikeThrough", "urn:schemas-microsoft-com:office:spreadsheet") = "1"
            Me._fontName = Util.GetAttribute(element, "FontName", "urn:schemas-microsoft-com:office:spreadsheet")
            Me._size = Util.GetAttribute(element, "Size", "urn:schemas-microsoft-com:office:spreadsheet", 0)
            Me._color = Util.GetAttribute(element, "Color", "urn:schemas-microsoft-com:office:spreadsheet")
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Font", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._bold Then
                writer.WriteAttributeString("Bold", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            If Me._italic Then
                writer.WriteAttributeString("Italic", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            If Me._underline <> UnderlineStyle.None Then
                writer.WriteAttributeString("Underline", "urn:schemas-microsoft-com:office:spreadsheet", Me._underline.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._strikethrough Then
                writer.WriteAttributeString("StrikeThrough", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            If Me._fontName IsNot Nothing Then
                writer.WriteAttributeString("FontName", "urn:schemas-microsoft-com:office:spreadsheet", Me._fontName.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._size <> 0 Then
                writer.WriteAttributeString("Size", "urn:schemas-microsoft-com:office:spreadsheet", Me._size.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._color IsNot Nothing Then
                writer.WriteAttributeString("Color", "urn:schemas-microsoft-com:office:spreadsheet", Me._color)
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Font", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
