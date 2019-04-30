Imports System.CodeDom
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetStyleAlignment
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _horizontal As StyleHorizontalAlignment

        Private _vertical As StyleVerticalAlignment

        Private _indent As Integer

        Private _rotate As Integer

        Private _shrinkToFit As Boolean

        Private _verticalText As Boolean

        Private _wrapText As Boolean

        Private _readingOrder As StyleReadingOrder

        Public Property Horizontal() As StyleHorizontalAlignment
            Get
                Return Me._horizontal
            End Get
            Set(value As StyleHorizontalAlignment)
                Me._horizontal = value
            End Set
        End Property

        Public Property Indent() As Integer
            Get
                Return Me._indent
            End Get
            Set(value As Integer)
                Me._indent = value
            End Set
        End Property

        Public Property ReadingOrder() As StyleReadingOrder
            Get
                Return Me._readingOrder
            End Get
            Set(value As StyleReadingOrder)
                Me._readingOrder = value
            End Set
        End Property

        Public Property Rotate() As Integer
            Get
                Return Me._rotate
            End Get
            Set(value As Integer)
                Me._rotate = value
            End Set
        End Property

        Public Property ShrinkToFit() As Boolean
            Get
                Return Me._shrinkToFit
            End Get
            Set(value As Boolean)
                Me._shrinkToFit = value
            End Set
        End Property

        Public Property Vertical() As StyleVerticalAlignment
            Get
                Return Me._vertical
            End Get
            Set(value As StyleVerticalAlignment)
                Me._vertical = value
            End Set
        End Property

        Public Property VerticalText() As Boolean
            Get
                Return Me._verticalText
            End Get
            Set(value As Boolean)
                Me._verticalText = value
            End Set
        End Property

        Public Property WrapText() As Boolean
            Get
                Return Me._wrapText
            End Get
            Set(value As Boolean)
                Me._wrapText = value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._horizontal <> StyleHorizontalAlignment.Automatic Then
                Util.AddAssignment(method, targetObject, "Horizontal", CType(Me._horizontal, [Enum]))
            End If
            If Me._indent <> 0 Then
                Util.AddAssignment(method, targetObject, "Indent", Me._indent)
            End If
            If Me._rotate <> 0 Then
                Util.AddAssignment(method, targetObject, "Rotate", Me._rotate)
            End If
            If Me._shrinkToFit Then
                Util.AddAssignment(method, targetObject, "ShrinkToFit", Me._shrinkToFit)
            End If
            If Me._vertical <> StyleVerticalAlignment.Automatic Then
                Util.AddAssignment(method, targetObject, "Vertical", CType(Me._vertical, [Enum]))
            End If
            If Me._verticalText Then
                Util.AddAssignment(method, targetObject, "VerticalText", Me._verticalText)
            End If
            If Me._wrapText Then
                Util.AddAssignment(method, targetObject, "WrapText", Me._wrapText)
            End If
            If Me._readingOrder <> StyleReadingOrder.NotSet Then
                Util.AddAssignment(method, targetObject, "ReadingOrder", CType(Me._readingOrder, [Enum]))
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetStyleAlignment.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Dim attribute As String = element.GetAttribute("Horizontal", "urn:schemas-microsoft-com:office:spreadsheet")
            If attribute IsNot Nothing AndAlso attribute.Length <> 0 Then
                Me._horizontal = CType([Enum].Parse(GetType(StyleHorizontalAlignment), attribute, True), StyleHorizontalAlignment)
            End If
            Me._indent = Util.GetAttribute(element, "Indent", "urn:schemas-microsoft-com:office:spreadsheet", 0)
            Me._rotate = Util.GetAttribute(element, "Rotate", "urn:schemas-microsoft-com:office:spreadsheet", 0)
            Me._shrinkToFit = Util.GetAttribute(element, "ShrinkToFit", "urn:schemas-microsoft-com:office:spreadsheet", False)
            attribute = element.GetAttribute("Vertical", "urn:schemas-microsoft-com:office:spreadsheet")
            If attribute IsNot Nothing AndAlso attribute.Length <> 0 Then
                Me._vertical = CType([Enum].Parse(GetType(StyleVerticalAlignment), attribute, True), StyleVerticalAlignment)
            End If
            attribute = element.GetAttribute("ReadingOrder", "urn:schemas-microsoft-com:office:spreadsheet")
            If attribute IsNot Nothing AndAlso attribute.Length <> 0 Then
                Me._readingOrder = CType([Enum].Parse(GetType(StyleReadingOrder), attribute, True), StyleReadingOrder)
            End If
            Me._verticalText = element.GetAttribute("VerticalText", "urn:schemas-microsoft-com:office:spreadsheet") = "1"
            Me._wrapText = element.GetAttribute("WrapText", "urn:schemas-microsoft-com:office:spreadsheet") = "1"
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Alignment", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._horizontal <> StyleHorizontalAlignment.Automatic Then
                writer.WriteAttributeString("Horizontal", "urn:schemas-microsoft-com:office:spreadsheet", Me._horizontal.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._indent <> 0 Then
                writer.WriteAttributeString("Indent", "urn:schemas-microsoft-com:office:spreadsheet", Me._indent.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._rotate <> 0 Then
                writer.WriteAttributeString("Rotate", "urn:schemas-microsoft-com:office:spreadsheet", Me._rotate.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._shrinkToFit Then
                writer.WriteAttributeString("ShrinkToFit", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            If Me._vertical <> StyleVerticalAlignment.Automatic Then
                writer.WriteAttributeString("Vertical", "urn:schemas-microsoft-com:office:spreadsheet", Me._vertical.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._verticalText Then
                writer.WriteAttributeString("VerticalText", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            If Me._wrapText Then
                writer.WriteAttributeString("WrapText", "urn:schemas-microsoft-com:office:spreadsheet", "1")
            End If
            If Me._readingOrder <> StyleReadingOrder.NotSet Then
                writer.WriteAttributeString("ReadingOrder", "urn:schemas-microsoft-com:office:spreadsheet", Me._readingOrder.ToString(CultureInfo.InvariantCulture))
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Alignment", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
