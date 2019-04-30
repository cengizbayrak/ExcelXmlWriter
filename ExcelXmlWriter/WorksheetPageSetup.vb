Imports System.CodeDom
Imports System.Collections
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class WorksheetPageSetup
		Implements IWriter
		Implements IReader
		Implements ICodeWriter
		Private _layout As WorksheetPageLayout

		Private _pageMargins As WorksheetPageMargins

		Private _header As WorksheetPageHeader

		Private _footer As WorksheetPageFooter

		Public ReadOnly Property Footer() As WorksheetPageFooter
			Get
				If Me._footer Is Nothing Then
					Me._footer = New WorksheetPageFooter()
				End If
				Return Me._footer
			End Get
		End Property

		Public ReadOnly Property Header() As WorksheetPageHeader
			Get
				If Me._header Is Nothing Then
					Me._header = New WorksheetPageHeader()
				End If
				Return Me._header
			End Get
		End Property

		Public ReadOnly Property Layout() As WorksheetPageLayout
			Get
				If Me._layout Is Nothing Then
					Me._layout = New WorksheetPageLayout()
				End If
				Return Me._layout
			End Get
		End Property

		Public ReadOnly Property PageMargins() As WorksheetPageMargins
			Get
				If Me._pageMargins Is Nothing Then
					Me._pageMargins = New WorksheetPageMargins()
				End If
				Return Me._pageMargins
			End Get
		End Property

		Friend Sub New()
		End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._layout IsNot Nothing Then
                Util.Traverse(type, Me._layout, method, targetObject, "Layout")
            End If
            If Me._header IsNot Nothing Then
                Util.Traverse(type, Me._header, method, targetObject, "Header")
            End If
            If Me._footer IsNot Nothing Then
                Util.Traverse(type, Me._footer, method, targetObject, "Footer")
            End If
            If Me._pageMargins IsNot Nothing Then
                Util.Traverse(type, Me._pageMargins, method, targetObject, "PageMargins")
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetPageSetup.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If WorksheetPageLayout.IsElement(xmlElement) Then
                    DirectCast(Me.Layout, IReader).ReadXml(xmlElement)
                ElseIf WorksheetPageHeader.IsElement(xmlElement) Then
                    DirectCast(Me.Header, IReader).ReadXml(xmlElement)
                ElseIf Not WorksheetPageFooter.IsElement(xmlElement) Then
                    If Not WorksheetPageMargins.IsElement(xmlElement) Then
                        Continue For
                    End If
                    DirectCast(Me.PageMargins, IReader).ReadXml(xmlElement)
                Else
                    DirectCast(Me.Footer, IReader).ReadXml(xmlElement)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "PageSetup", "urn:schemas-microsoft-com:office:excel")
            If Me._layout IsNot Nothing Then
                DirectCast(Me._layout, IWriter).WriteXml(writer)
            End If
            If Me._header IsNot Nothing Then
                DirectCast(Me._header, IWriter).WriteXml(writer)
            End If
            If Me._footer IsNot Nothing Then
                DirectCast(Me._footer, IWriter).WriteXml(writer)
            End If
            If Me._pageMargins IsNot Nothing Then
                DirectCast(Me._pageMargins, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub

		Friend Shared Function IsElement(element As XmlElement) As Boolean
			Return Util.IsElement(element, "PageSetup", "urn:schemas-microsoft-com:office:excel")
		End Function
	End Class
End Namespace
