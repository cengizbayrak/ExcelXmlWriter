Imports System.CodeDom
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetPageLayout
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _orientation As ExcelXmlWriter.Orientation

        Private _centerHorizontal As Boolean

        Private _centerVertical As Boolean

        Public Property CenterHorizontal() As Boolean
            Get
                Return Me._centerHorizontal
            End Get
            Set(value As Boolean)
                Me._centerHorizontal = Value
            End Set
        End Property

        Public Property CenterVertical() As Boolean
            Get
                Return Me._centerVertical
            End Get
            Set(value As Boolean)
                Me._centerVertical = Value
            End Set
        End Property

        Public Property Orientation() As ExcelXmlWriter.Orientation
            Get
                Return Me._orientation
            End Get
            Set(value As ExcelXmlWriter.Orientation)
                Me._orientation = value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._orientation <> ExcelXmlWriter.Orientation.NotSet Then
                Util.AddAssignment(method, targetObject, "Orientation", CType(Me._orientation, [Enum]))
            End If
            If Me._centerHorizontal Then
                Util.AddAssignment(method, targetObject, "CenterHorizontal", True)
            End If
            If Me._centerVertical Then
                Util.AddAssignment(method, targetObject, "CenterVertical", True)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetPageLayout.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Dim attribute As String = Util.GetAttribute(element, "Orientation", "urn:schemas-microsoft-com:office:excel")
            If attribute IsNot Nothing AndAlso attribute.Length > 0 Then
                Me._orientation = CType([Enum].Parse(GetType(ExcelXmlWriter.Orientation), attribute), ExcelXmlWriter.Orientation)
            End If
            Me._centerHorizontal = Util.GetAttribute(element, "CenterHorizontal", "urn:schemas-microsoft-com:office:excel", False)
            Me._centerVertical = Util.GetAttribute(element, "CenterVertical", "urn:schemas-microsoft-com:office:excel", False)
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "Layout", "urn:schemas-microsoft-com:office:excel")
            If Me._orientation <> ExcelXmlWriter.Orientation.NotSet Then
                writer.WriteAttributeString("Orientation", "urn:schemas-microsoft-com:office:excel", Me._orientation.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._centerHorizontal Then
                writer.WriteAttributeString("CenterHorizontal", "urn:schemas-microsoft-com:office:excel", "1")
            End If
            If Me._centerVertical Then
                writer.WriteAttributeString("CenterVertical", "urn:schemas-microsoft-com:office:excel", "1")
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Layout", "urn:schemas-microsoft-com:office:excel")
        End Function
    End Class
End Namespace
