Imports System.CodeDom
Imports System.Collections
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetStyle
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _id As String

        Private _name As String

        Private _parent As String

        Private _numberFormat As String = "General"

        Private _font As WorksheetStyleFont

        Private _interior As WorksheetStyleInterior

        Private _alignment As WorksheetStyleAlignment

        Private _borders As WorksheetStyleBorderCollection

        Friend _book As ExcelXmlWriter.Workbook

        Public ReadOnly Property Alignment() As WorksheetStyleAlignment
            Get
                If Me._alignment Is Nothing Then
                    Me._alignment = New WorksheetStyleAlignment()
                End If
                Return Me._alignment
            End Get
        End Property

        Public ReadOnly Property Borders() As WorksheetStyleBorderCollection
            Get
                If Me._borders Is Nothing Then
                    Me._borders = New WorksheetStyleBorderCollection(Me)
                End If
                Return Me._borders
            End Get
        End Property

        Public ReadOnly Property Font() As WorksheetStyleFont
            Get
                If Me._font Is Nothing Then
                    Me._font = New WorksheetStyleFont()
                End If
                Return Me._font
            End Get
        End Property

        Public ReadOnly Property ID() As String
            Get
                Return Me._id
            End Get
        End Property

        Public ReadOnly Property Interior() As WorksheetStyleInterior
            Get
                If Me._interior Is Nothing Then
                    Me._interior = New WorksheetStyleInterior()
                End If
                Return Me._interior
            End Get
        End Property

        Public Property Name() As String
            Get
                Return Me._name
            End Get
            Set(value As String)
                Me._name = value
            End Set
        End Property

        Public Property NumberFormat() As String
            Get
                Return Me._numberFormat
            End Get
            Set(value As String)
                Me._numberFormat = value
            End Set
        End Property

        Public Property Parent() As String
            Get
                Return Me._parent
            End Get
            Set(value As String)
                Me._parent = value
            End Set
        End Property

        Public ReadOnly Property Workbook() As ExcelXmlWriter.Workbook
            Get
                Return Me._book
            End Get
        End Property

        Public Sub New(id As String)
            Me._id = id
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._name IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Name", Me._name)
            End If
            If Me._parent IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Parent", Me._parent)
            End If
            If Me._font IsNot Nothing Then
                Util.Traverse(type, Me._font, method, targetObject, "Font")
            End If
            If Me._interior IsNot Nothing Then
                Util.Traverse(type, Me._interior, method, targetObject, "Interior")
            End If
            If Me._alignment IsNot Nothing Then
                Util.Traverse(type, Me._alignment, method, targetObject, "Alignment")
            End If
            If Me._borders IsNot Nothing Then
                Util.Traverse(type, Me._borders, method, targetObject, "Borders")
            End If
            If Me._numberFormat <> "General" Then
                Util.AddAssignment(method, targetObject, "NumberFormat", Me._numberFormat)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetStyle.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            Me._id = Util.GetAttribute(element, "ID", "urn:schemas-microsoft-com:office:spreadsheet")
            Me._name = Util.GetAttribute(element, "Name", "urn:schemas-microsoft-com:office:spreadsheet")
            Me._parent = Util.GetAttribute(element, "Parent", "urn:schemas-microsoft-com:office:spreadsheet")
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If WorksheetStyleFont.IsElement(xmlElement) Then
                    DirectCast(Me.Font, IReader).ReadXml(xmlElement)
                ElseIf WorksheetStyleInterior.IsElement(xmlElement) Then
                    DirectCast(Me.Interior, IReader).ReadXml(xmlElement)
                ElseIf WorksheetStyleAlignment.IsElement(xmlElement) Then
                    DirectCast(Me.Alignment, IReader).ReadXml(xmlElement)
                ElseIf Not WorksheetStyleBorderCollection.IsElement(xmlElement) Then
                    If Not Util.IsElement(xmlElement, "NumberFormat", "urn:schemas-microsoft-com:office:spreadsheet") Then
                        Continue For
                    End If
                    Dim attribute As String = xmlElement.GetAttribute("Format", "urn:schemas-microsoft-com:office:spreadsheet")
                    If attribute Is Nothing OrElse attribute.Length <= 0 Then
                        Continue For
                    End If
                    Me._numberFormat = attribute
                Else
                    DirectCast(Me.Borders, IReader).ReadXml(xmlElement)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("Style", "urn:schemas-microsoft-com:office:spreadsheet")
            If Me._id IsNot Nothing Then
                writer.WriteAttributeString("ID", "urn:schemas-microsoft-com:office:spreadsheet", Me._id)
            End If
            If Me._name IsNot Nothing Then
                writer.WriteAttributeString("Name", "urn:schemas-microsoft-com:office:spreadsheet", Me._name)
            End If
            If Me._parent IsNot Nothing Then
                writer.WriteAttributeString("Parent", "urn:schemas-microsoft-com:office:spreadsheet", Me._parent)
            End If
            If Me._alignment IsNot Nothing Then
                DirectCast(Me._alignment, IWriter).WriteXml(writer)
            End If
            If Me._borders IsNot Nothing Then
                DirectCast(Me._borders, IWriter).WriteXml(writer)
            End If
            If Me._font IsNot Nothing Then
                DirectCast(Me._font, IWriter).WriteXml(writer)
            End If
            If Me._interior IsNot Nothing Then
                DirectCast(Me._interior, IWriter).WriteXml(writer)
            End If
            If Me._numberFormat <> "General" Then
                writer.WriteStartElement("NumberFormat", "urn:schemas-microsoft-com:office:spreadsheet")
                writer.WriteAttributeString("s", "Format", "urn:schemas-microsoft-com:office:spreadsheet", Me._numberFormat)
                writer.WriteEndElement()
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Style", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function
    End Class
End Namespace
