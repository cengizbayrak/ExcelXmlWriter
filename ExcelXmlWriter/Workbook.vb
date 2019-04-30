Imports System.CodeDom
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class Workbook
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _documentProperties As DocumentProperties

        Private _excelWorkbook As ExcelXmlWriter.ExcelWorkbook

        Private _styles As WorksheetStyleCollection

        Private _pivotCache As ExcelXmlWriter.PivotCache

        Private _worksheets As WorksheetCollection

        Private _generateExcelProcessingInstruction As Boolean = True

        Private _names As WorksheetNamedRangeCollection

        Private _spreadSheetComponentOptions As ExcelXmlWriter.SpreadSheetComponentOptions

        Public ReadOnly Property ExcelWorkbook() As ExcelXmlWriter.ExcelWorkbook
            Get
                If Me._excelWorkbook Is Nothing Then
                    Me._excelWorkbook = New ExcelXmlWriter.ExcelWorkbook()
                End If
                Return Me._excelWorkbook
            End Get
        End Property

        Public Property GenerateExcelProcessingInstruction() As Boolean
            Get
                Return Me._generateExcelProcessingInstruction
            End Get
            Set(value As Boolean)
                Me._generateExcelProcessingInstruction = value
            End Set
        End Property

        Public ReadOnly Property Names() As WorksheetNamedRangeCollection
            Get
                If Me._names Is Nothing Then
                    Me._names = New WorksheetNamedRangeCollection()
                End If
                Return Me._names
            End Get
        End Property

        '[EditorBrowsable(,)]    // JustDecompile was unable to locate the assembly where attribute parameters types are defined. Generating parameters values is impossible.
        Public ReadOnly Property PivotCache() As ExcelXmlWriter.PivotCache
            Get
                If Me._pivotCache Is Nothing Then
                    Me._pivotCache = New ExcelXmlWriter.PivotCache()
                End If
                Return Me._pivotCache
            End Get
        End Property

        Public ReadOnly Property Properties() As DocumentProperties
            Get
                If Me._documentProperties Is Nothing Then
                    Me._documentProperties = New DocumentProperties()
                End If
                Return Me._documentProperties
            End Get
        End Property

        Public ReadOnly Property SpreadSheetComponentOptions() As ExcelXmlWriter.SpreadSheetComponentOptions
            Get
                If Me._spreadSheetComponentOptions Is Nothing Then
                    Me._spreadSheetComponentOptions = New ExcelXmlWriter.SpreadSheetComponentOptions()
                End If
                Return Me._spreadSheetComponentOptions
            End Get
        End Property

        Public ReadOnly Property Styles() As WorksheetStyleCollection
            Get
                If Me._styles Is Nothing Then
                    Me._styles = New WorksheetStyleCollection(Me)
                End If
                Return Me._styles
            End Get
        End Property

        Public ReadOnly Property Worksheets() As WorksheetCollection
            Get
                If Me._worksheets Is Nothing Then
                    Me._worksheets = New WorksheetCollection()
                End If
                Return Me._worksheets
            End Get
        End Property

        Public Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._documentProperties IsNot Nothing Then
                Util.Traverse(type, Me._documentProperties, method, targetObject, "Properties")
            End If
            If Me._spreadSheetComponentOptions IsNot Nothing Then
                Util.Traverse(type, Me._spreadSheetComponentOptions, method, targetObject, "SpreadSheetComponentOptions")
            End If
            If Me._excelWorkbook IsNot Nothing Then
                Util.Traverse(type, Me._excelWorkbook, method, targetObject, "ExcelWorkbook")
            End If
            If Me._styles IsNot Nothing Then
                Util.Traverse(type, Me._styles, method, targetObject, "Styles")
            End If
            If Me._names IsNot Nothing Then
                Util.Traverse(type, Me._names, method, targetObject, "Names")
            End If
            If Me._worksheets IsNot Nothing Then
                Util.Traverse(type, Me._worksheets, method, targetObject, "Worksheets")
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not Workbook.IsElement(element) Then
                Throw New InvalidOperationException(String.Concat("The specified Xml is not a valid Workbook." & vbLf & "Element Name:", element.NamespaceURI))
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If DocumentProperties.IsElement(xmlElement) Then
                    DirectCast(Me.Properties, IReader).ReadXml(xmlElement)
                ElseIf ExcelXmlWriter.SpreadSheetComponentOptions.IsElement(xmlElement) Then
                    DirectCast(Me.SpreadSheetComponentOptions, IReader).ReadXml(xmlElement)
                ElseIf ExcelXmlWriter.ExcelWorkbook.IsElement(xmlElement) Then
                    DirectCast(Me.ExcelWorkbook, IReader).ReadXml(xmlElement)
                ElseIf WorksheetStyleCollection.IsElement(xmlElement) Then
                    DirectCast(Me.Styles, IReader).ReadXml(xmlElement)
                ElseIf Not WorksheetNamedRangeCollection.IsElement(xmlElement) Then
                    If Not Worksheet.IsElement(xmlElement) Then
                        Continue For
                    End If
                    Dim worksheet__1 As New Worksheet(Nothing)
                    DirectCast(worksheet__1, IReader).ReadXml(xmlElement)
                    Me.Worksheets.Add(worksheet__1)
                Else
                    DirectCast(Me.Names, IReader).ReadXml(xmlElement)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Workbook", "urn:schemas-microsoft-com:office:spreadsheet")
            writer.WriteAttributeString("xmlns", "x", Nothing, "urn:schemas-microsoft-com:office:excel")
            writer.WriteAttributeString("xmlns", "o", Nothing, "urn:schemas-microsoft-com:office:office")
            If Me._documentProperties IsNot Nothing Then
                DirectCast(Me._documentProperties, IWriter).WriteXml(writer)
            End If
            If Me._spreadSheetComponentOptions IsNot Nothing Then
                DirectCast(Me._spreadSheetComponentOptions, IWriter).WriteXml(writer)
            End If
            If Me._excelWorkbook IsNot Nothing Then
                DirectCast(Me._excelWorkbook, IWriter).WriteXml(writer)
            End If
            If Me._styles IsNot Nothing Then
                DirectCast(Me._styles, IWriter).WriteXml(writer)
            End If
            If Me._names IsNot Nothing Then
                DirectCast(Me._names, IWriter).WriteXml(writer)
            End If
            If Me._worksheets IsNot Nothing Then
                DirectCast(Me._worksheets, IWriter).WriteXml(writer)
            End If
            If Me._pivotCache IsNot Nothing Then
                DirectCast(Me._pivotCache, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Workbook", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function

        Public Sub Load(filename As String)
            Dim fileStream As FileStream = Nothing
            Try
                fileStream = New FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Me.Load(fileStream)
            Finally
                If fileStream IsNot Nothing Then
                    fileStream.Close()
                End If
            End Try
        End Sub

        Public Sub Load(stream As Stream)
            If stream Is Nothing Then
                Throw New ArgumentNullException("stream")
            End If
            If stream.Position >= stream.Length Then
                stream.Position = CLng(0)
            End If
            Dim xmlDocument As New XmlDocument()
            xmlDocument.Load(stream)
            DirectCast(Me, IReader).ReadXml(xmlDocument.DocumentElement)
        End Sub

        Public Sub Save(filename As String)
            Dim fileStream As FileStream = Nothing
            Try
                fileStream = New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None)
                Me.Save(fileStream)
            Finally
                If fileStream IsNot Nothing Then
                    fileStream.Close()
                End If
            End Try
        End Sub

        Public Sub Save(stream As Stream)
            If Me.Worksheets.Count = 0 Then
                Me.Worksheets.Add("Sheet 1")
            End If
            Dim xmlTextWriter As New XmlTextWriter(stream, Encoding.UTF8)
            xmlTextWriter.Namespaces = True
            xmlTextWriter.Formatting = Formatting.Indented
            xmlTextWriter.WriteProcessingInstruction("xml", "version='1.0'")
            If Me._generateExcelProcessingInstruction Then
                xmlTextWriter.WriteProcessingInstruction("mso-application", "progid='Excel.Sheet'")
            End If
            DirectCast(Me, IWriter).WriteXml(xmlTextWriter)
            xmlTextWriter.Flush()
        End Sub
    End Class
End Namespace
