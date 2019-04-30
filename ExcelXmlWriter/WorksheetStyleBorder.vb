Imports System.CodeDom
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetStyleBorder
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _position As StylePosition

        Private _lineStyle As LineStyleOption

        Private _weight As Integer = -1

        Private _color As String

        Public Property Color() As String
            Get
                Return Me._color
            End Get
            Set(value As String)
                Me._color = value
            End Set
        End Property

        Public Property LineStyle() As LineStyleOption
            Get
                Return Me._lineStyle
            End Get
            Set(value As LineStyleOption)
                Me._lineStyle = value
            End Set
        End Property

        Public Property Position() As StylePosition
            Get
                Return Me._position
            End Get
            Set(value As StylePosition)
                Me._position = value
            End Set
        End Property

        Public Property Weight() As Integer
            Get
                Return Me._weight
            End Get
            Set(value As Integer)
                If value >= 3 Then
                    Me._weight = 3
                    Return
                End If
                If value <= 0 Then
                    Me._weight = 0
                    Return
                End If
                Me._weight = value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._position <> StylePosition.NotSet Then
                Util.AddAssignment(method, targetObject, "Position", CType(Me._position, [Enum]))
            End If
            If Me._weight <> -1 Then
                Util.AddAssignment(method, targetObject, "Weight", Me._weight)
            End If
            If Me._color IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Color", Me._color)
            End If
            If Me._lineStyle <> LineStyleOption.NotSet Then
                Util.AddAssignment(method, targetObject, "LineStyle", CType(Me._lineStyle, [Enum]))
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetStyleBorder.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._color = Util.GetAttribute(element, "Color", "urn:schemas-microsoft-com:office:spreadsheet")
            Dim attribute As String = element.GetAttribute("Position", "urn:schemas-microsoft-com:office:spreadsheet")
            If attribute IsNot Nothing AndAlso attribute.Length <> 0 Then
                Me._position = CType([Enum].Parse(GetType(StylePosition), attribute, True), StylePosition)
            End If
            Me._weight = Util.GetAttribute(element, "Weight", "urn:schemas-microsoft-com:office:spreadsheet", -1)
            Me._color = Util.GetAttribute(element, "Color", "urn:schemas-microsoft-com:office:spreadsheet")
            attribute = element.GetAttribute("LineStyle", "urn:schemas-microsoft-com:office:spreadsheet")
            If attribute IsNot Nothing AndAlso attribute.Length <> 0 Then
                Me._lineStyle = CType([Enum].Parse(GetType(LineStyleOption), attribute, True), LineStyleOption)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Border", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._color IsNot Nothing Then
                writer.WriteAttributeString("s", "Color", "urn:schemas-microsoft-com:office:spreadsheet", Me._color)
            End If
            If Me._position <> StylePosition.NotSet Then
                writer.WriteAttributeString("s", "Position", "urn:schemas-microsoft-com:office:spreadsheet", Me._position.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._lineStyle <> LineStyleOption.NotSet Then
                writer.WriteAttributeString("s", "LineStyle", "urn:schemas-microsoft-com:office:spreadsheet", Me._lineStyle.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._weight <> -1 Then
                writer.WriteAttributeString("s", "Weight", "urn:schemas-microsoft-com:office:spreadsheet", Me._weight.ToString(CultureInfo.InvariantCulture))
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Border", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function

        Friend Function IsSpecial() As Integer
            If Me._position = StylePosition.NotSet OrElse Me._weight = -1 OrElse Me._lineStyle = LineStyleOption.NotSet Then
                Return 2
            End If
            If Me._color Is Nothing Then
                Return 0
            End If
            Return 1
        End Function
    End Class
End Namespace
