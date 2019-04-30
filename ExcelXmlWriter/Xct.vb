Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class Xct
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _sheetIndex As Integer = -2147483648

        Private _operands As CrnCollection

        Public ReadOnly Property Operands() As CrnCollection
            Get
                If Me._operands Is Nothing Then
                    Me._operands = New CrnCollection()
                End If
                Return Me._operands
            End Get
        End Property

        Public Property SheetIndex() As Integer
            Get
                If Me._sheetIndex = -2147483648 Then
                    Return 0
                End If
                Return Me._sheetIndex
            End Get
            Set(value As Integer)
                If value < 0 Then
                    Throw New ArgumentException("Invalid range, > 0")
                End If
                Me._sheetIndex = value
            End Set
        End Property

        Public Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._sheetIndex <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "SheetIndex", Me._sheetIndex)
            End If
            If Me._operands IsNot Nothing Then
                DirectCast(Me._operands, ICodeWriter).WriteTo(type, method, targetObject)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not Xct.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If Not Util.IsElement(xmlElement, "SheetIndex", "urn:schemas-microsoft-com:office:excel") Then
                    If Not Crn.IsElement(xmlElement) Then
                        Continue For
                    End If
                    Dim crn__1 As New Crn()
                    DirectCast(crn__1, IReader).ReadXml(xmlElement)
                    Me.Operands.Add(crn__1)
                Else
                    Me._sheetIndex = Integer.Parse(xmlElement.InnerText)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "Xct", "urn:schemas-microsoft-com:office:excel")
            If Me._operands Is Nothing Then
                writer.WriteElementString("Count", "urn:schemas-microsoft-com:office:excel", "0")
            Else
                Dim count As Integer = Me._operands.Count
                writer.WriteElementString("Count", "urn:schemas-microsoft-com:office:excel", count.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._sheetIndex <> -2147483648 Then
                writer.WriteElementString("SheetIndex", "urn:schemas-microsoft-com:office:excel", Me._sheetIndex.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._operands IsNot Nothing Then
                DirectCast(Me._operands, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Xct", "urn:schemas-microsoft-com:office:excel")
        End Function
    End Class
End Namespace
