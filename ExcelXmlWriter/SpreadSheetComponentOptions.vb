Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class SpreadSheetComponentOptions
		Implements IWriter
		Implements IReader
		Implements ICodeWriter
		Private _nextSheetNumber As Integer = -2147483648

		Private _doNotEnableResize As Boolean

		Private _preventPropBrowser As Boolean

		Private _maxHeight As String

		Private _maxWidth As String

		Private _spreadsheetAutoFit As Boolean

		Private _toolbar As SpreadSheetToolbar

		Public Property DoNotEnableResize() As Boolean
			Get
				Return Me._doNotEnableResize
			End Get
			Set
				Me._doNotEnableResize = value
			End Set
		End Property

		Public Property MaxHeight() As String
			Get
				Return Me._maxHeight
			End Get
			Set
				Me._maxHeight = value
			End Set
		End Property

		Public Property MaxWidth() As String
			Get
				Return Me._maxWidth
			End Get
			Set
				Me._maxWidth = value
			End Set
		End Property

		Public Property NextSheetNumber() As Integer
			Get
				Return Me._nextSheetNumber
			End Get
			Set
				Me._nextSheetNumber = value
			End Set
		End Property

		Public Property PreventPropBrowser() As Boolean
			Get
				Return Me._preventPropBrowser
			End Get
			Set
				Me._preventPropBrowser = value
			End Set
		End Property

		Public Property SpreadsheetAutoFit() As Boolean
			Get
				Return Me._spreadsheetAutoFit
			End Get
			Set
				Me._spreadsheetAutoFit = value
			End Set
		End Property

		Public ReadOnly Property Toolbar() As SpreadSheetToolbar
			Get
				If Me._toolbar Is Nothing Then
					Me._toolbar = New SpreadSheetToolbar()
				End If
				Return Me._toolbar
			End Get
		End Property

		Friend Sub New()
		End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._toolbar IsNot Nothing Then
                DirectCast(Me._toolbar, ICodeWriter).WriteTo(type, method, New CodePropertyReferenceExpression(targetObject, "Toolbar"))
            End If
            If Me._nextSheetNumber <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "NextSheetNumber", Me._nextSheetNumber)
            End If
            If Me._spreadsheetAutoFit Then
                Util.AddAssignment(method, targetObject, "SpreadsheetAutoFit", Me._spreadsheetAutoFit)
            End If
            If Me._doNotEnableResize Then
                Util.AddAssignment(method, targetObject, "DoNotEnableResize", Me._doNotEnableResize)
            End If
            If Me._preventPropBrowser Then
                Util.AddAssignment(method, targetObject, "PreventPropBrowser", Me._preventPropBrowser)
            End If
            If Me._maxHeight IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "MaxHeight", Me._maxHeight)
            End If
            If Me._maxWidth IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "MaxWidth", Me._maxWidth)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not SpreadSheetComponentOptions.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If SpreadSheetToolbar.IsElement(xmlElement) Then
                    DirectCast(Me.Toolbar, IReader).ReadXml(xmlElement)
                ElseIf Util.IsElement(xmlElement, "NextSheetNumber", "urn:schemas-microsoft-com:office:component:spreadsheet") Then
                    Me._nextSheetNumber = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "SpreadsheetAutoFit", "urn:schemas-microsoft-com:office:component:spreadsheet") Then
                    Me._spreadsheetAutoFit = True
                ElseIf Util.IsElement(xmlElement, "DoNotEnableResize", "urn:schemas-microsoft-com:office:component:spreadsheet") Then
                    Me._doNotEnableResize = True
                ElseIf Util.IsElement(xmlElement, "PreventPropBrowser", "urn:schemas-microsoft-com:office:component:spreadsheet") Then
                    Me._preventPropBrowser = True
                ElseIf Not Util.IsElement(xmlElement, "MaxHeight", "urn:schemas-microsoft-com:office:component:spreadsheet") Then
                    If Not Util.IsElement(xmlElement, "MaxWidth", "urn:schemas-microsoft-com:office:component:spreadsheet") Then
                        Continue For
                    End If
                    Me._maxWidth = xmlElement.InnerText
                Else
                    Me._maxHeight = xmlElement.InnerText
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("c", "ComponentOptions", "urn:schemas-microsoft-com:office:component:spreadsheet")
            If Me._toolbar IsNot Nothing Then
                DirectCast(Me._toolbar, IWriter).WriteXml(writer)
            End If
            If Me._nextSheetNumber <> -2147483648 Then
                writer.WriteElementString("NextSheetNumber", "urn:schemas-microsoft-com:office:component:spreadsheet", Me._nextSheetNumber.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._spreadsheetAutoFit Then
                writer.WriteElementString("SpreadsheetAutoFit", "urn:schemas-microsoft-com:office:component:spreadsheet", "")
            End If
            If Me._doNotEnableResize Then
                writer.WriteElementString("DoNotEnableResize", "urn:schemas-microsoft-com:office:component:spreadsheet", "")
            End If
            If Me._preventPropBrowser Then
                writer.WriteElementString("PreventPropBrowser", "urn:schemas-microsoft-com:office:component:spreadsheet", "")
            End If
            If Me._maxHeight IsNot Nothing Then
                writer.WriteElementString("MaxHeight", "urn:schemas-microsoft-com:office:component:spreadsheet", Me._maxHeight)
            End If
            If Me._maxWidth IsNot Nothing Then
                writer.WriteElementString("MaxWidth", "urn:schemas-microsoft-com:office:component:spreadsheet", Me._maxWidth)
            End If
            writer.WriteEndElement()
        End Sub

		Friend Shared Function IsElement(element As XmlElement) As Boolean
			Return Util.IsElement(element, "ComponentOptions", "urn:schemas-microsoft-com:office:component:spreadsheet")
		End Function
	End Class
End Namespace
