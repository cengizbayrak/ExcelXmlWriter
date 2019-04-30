Imports System.CodeDom
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetCellData
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _type As DataType = DataType.String

        Private _text As String

        Private _parent As IReader

        Friend ReadOnly Property IsSimple() As Boolean
            Get
                If Me._type <> DataType.NotSet AndAlso Me._text IsNot Nothing Then
                    Return True
                End If
                Return False
            End Get
        End Property

        Public Property Text() As String
            Get
                Return Me._text
            End Get
            Set(value As String)
                Me._text = value
            End Set
        End Property

        Public Property Type() As DataType
            Get
                Return Me._type
            End Get
            Set(value As DataType)
                Me._type = value
            End Set
        End Property

        Friend Sub New(parent As IReader)
            Me._parent = parent
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._type <> DataType.NotSet AndAlso Not (TypeOf Me._parent Is WorksheetComment) Then
                Util.AddAssignment(method, targetObject, "Type", CType(Me._type, [Enum]))
            End If
            If Me._text IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Text", Me._text)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetCellData.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Dim attribute As String = element.GetAttribute("Type", "urn:schemas-microsoft-com:office:spreadsheet")
            If attribute IsNot Nothing AndAlso attribute.Length > 0 Then
                Me._type = CType([Enum].Parse(GetType(DataType), attribute), DataType)
            End If
            If Not element.IsEmpty Then
                Me._text = element.InnerText
            End If
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Data", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._type <> DataType.NotSet AndAlso Not (TypeOf Me._parent Is WorksheetComment) Then
                writer.WriteAttributeString("s", "Type", "urn:schemas-microsoft-com:office:spreadsheet", Me._type.ToString(CultureInfo.InvariantCulture))
            End If
            writer.WriteString(Me._text)
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Data", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
