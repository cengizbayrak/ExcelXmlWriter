Imports System.CodeDom
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetPageMargins
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _bottom As Single = 1.0F

        Private _left As Single = 0.75F

        Private _right As Single = 0.75F

        Private _top As Single = 1.0F

        Public Property Bottom() As Single
            Get
                Return Me._bottom
            End Get
            Set(value As Single)
                Me._bottom = value
            End Set
        End Property

        Public Property Left() As Single
            Get
                Return Me._left
            End Get
            Set(value As Single)
                Me._left = value
            End Set
        End Property

        Public Property Right() As Single
            Get
                Return Me._right
            End Get
            Set(value As Single)
                Me._right = value
            End Set
        End Property

        Public Property Top() As Single
            Get
                Return Me._top
            End Get
            Set(value As Single)
                Me._top = value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            Util.AddAssignment(method, targetObject, "Bottom", Me._bottom)
            Util.AddAssignment(method, targetObject, "Left", Me._left)
            Util.AddAssignment(method, targetObject, "Right", Me._right)
            Util.AddAssignment(method, targetObject, "Top", Me._top)
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetPageMargins.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._bottom = Util.GetAttribute(element, "Bottom", "urn:schemas-microsoft-com:office:excel", 1.0F)
            Me._left = Util.GetAttribute(element, "Left", "urn:schemas-microsoft-com:office:excel", 0.75F)
            Me._right = Util.GetAttribute(element, "Right", "urn:schemas-microsoft-com:office:excel", 0.75F)
            Me._top = Util.GetAttribute(element, "Top", "urn:schemas-microsoft-com:office:excel", 1.0F)
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "PageMargins", "urn:schemas-microsoft-com:office:excel")
            writer.WriteAttributeString("Bottom", "urn:schemas-microsoft-com:office:excel", Me._bottom.ToString(CultureInfo.InvariantCulture))
            writer.WriteAttributeString("Left", "urn:schemas-microsoft-com:office:excel", Me._left.ToString(CultureInfo.InvariantCulture))
            writer.WriteAttributeString("Right", "urn:schemas-microsoft-com:office:excel", Me._right.ToString(CultureInfo.InvariantCulture))
            writer.WriteAttributeString("Top", "urn:schemas-microsoft-com:office:excel", Me._top.ToString(CultureInfo.InvariantCulture))
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "PageMargins", "urn:schemas-microsoft-com:office:excel")
        End Function
    End Class
End Namespace
