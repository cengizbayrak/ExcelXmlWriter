Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class ExcelWorkbook
		Implements IWriter
		Implements IReader
		Implements ICodeWriter
		Private _windowHeight As Integer = -2147483648

		Private _windowWidth As Integer = -2147483648

		Private _windowTopX As Integer = -2147483648

		Private _windowTopY As Integer = -2147483648

		Private _activeSheet As Integer = -2147483648

		Private _hideWorkbookTabs As Boolean

		Private _protectStructure As Boolean

		Private _protectWindows As Boolean

		Private _links As ExcelLinksCollection

		Public Property ActiveSheetIndex() As Integer
			Get
				Return Me._activeSheet
			End Get
			Set
				Me._activeSheet = value
			End Set
		End Property

		Public Property HideWorkbookTabs() As Boolean
			Get
				Return Me._hideWorkbookTabs
			End Get
			Set
				Me._hideWorkbookTabs = value
			End Set
		End Property

		Public ReadOnly Property Links() As ExcelLinksCollection
			Get
				If Me._links Is Nothing Then
					Me._links = New ExcelLinksCollection()
				End If
				Return Me._links
			End Get
		End Property

		Public Property ProtectStructure() As Boolean
			Get
				Return Me._protectStructure
			End Get
			Set
				Me._protectStructure = value
			End Set
		End Property

		Public Property ProtectWindows() As Boolean
			Get
				Return Me._protectWindows
			End Get
			Set
				Me._protectWindows = value
			End Set
		End Property

		Public Property WindowHeight() As Integer
			Get
				Return Me._windowHeight
			End Get
			Set
				Me._windowHeight = value
			End Set
		End Property

		Public Property WindowTopX() As Integer
			Get
				Return Me._windowTopX
			End Get
			Set
				Me._windowTopX = value
			End Set
		End Property

		Public Property WindowTopY() As Integer
			Get
				Return Me._windowTopY
			End Get
			Set
				Me._windowTopY = value
			End Set
		End Property

		Public Property WindowWidth() As Integer
			Get
				Return Me._windowWidth
			End Get
			Set
				Me._windowWidth = value
			End Set
		End Property

		Friend Sub New()
			CrnCollection.GlobalCounter = 0
		End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._links IsNot Nothing Then
                DirectCast(Me._links, ICodeWriter).WriteTo(type, method, targetObject)
            End If
            If Me._hideWorkbookTabs Then
                Util.AddAssignment(method, targetObject, "HideWorkbookTabs", Me._hideWorkbookTabs)
            End If
            If Me._windowHeight <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "WindowHeight", Me._windowHeight)
            End If
            If Me._windowWidth <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "WindowWidth", Me._windowWidth)
            End If
            If Me._windowTopX <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "WindowTopX", Me._windowTopX)
            End If
            If Me._windowTopY <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "WindowTopY", Me._windowTopY)
            End If
            If Me._activeSheet <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "ActiveSheetIndex", Me._activeSheet)
            End If
            Util.AddAssignment(method, targetObject, "ProtectWindows", Me._protectWindows)
            Util.AddAssignment(method, targetObject, "ProtectStructure", Me._protectStructure)
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not ExcelWorkbook.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If Util.IsElement(xmlElement, "HideWorkbookTabs", "urn:schemas-microsoft-com:office:excel") Then
                    Me._hideWorkbookTabs = True
                ElseIf Util.IsElement(xmlElement, "WindowHeight", "urn:schemas-microsoft-com:office:excel") Then
                    Me._windowHeight = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "WindowTopX", "urn:schemas-microsoft-com:office:excel") Then
                    Me._windowTopX = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "WindowTopY", "urn:schemas-microsoft-com:office:excel") Then
                    Me._windowTopY = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "WindowWidth", "urn:schemas-microsoft-com:office:excel") Then
                    Me._windowWidth = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Not Util.IsElement(xmlElement, "ActiveSheet", "urn:schemas-microsoft-com:office:excel") Then
                    If Not SupBook.IsElement(xmlElement) Then
                        Continue For
                    End If
                    Dim supBook__1 As New SupBook()
                    DirectCast(supBook__1, IReader).ReadXml(xmlElement)
                    Me.Links.Add(supBook__1)
                Else
                    Me._activeSheet = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "ExcelWorkbook", "urn:schemas-microsoft-com:office:excel")
            If Me._links IsNot Nothing Then
                DirectCast(Me._links, IWriter).WriteXml(writer)
            End If
            If Me._hideWorkbookTabs Then
                writer.WriteElementString("HideWorkbookTabs", "urn:schemas-microsoft-com:office:excel", "")
            End If
            If Me._windowHeight <> -2147483648 Then
                writer.WriteElementString("WindowHeight", "urn:schemas-microsoft-com:office:excel", Me._windowHeight.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._windowTopX <> -2147483648 Then
                writer.WriteElementString("WindowTopX", "urn:schemas-microsoft-com:office:excel", Me._windowTopX.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._windowTopY <> -2147483648 Then
                writer.WriteElementString("WindowTopY", "urn:schemas-microsoft-com:office:excel", Me._windowTopY.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._windowWidth <> -2147483648 Then
                writer.WriteElementString("WindowWidth", "urn:schemas-microsoft-com:office:excel", Me._windowWidth.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._activeSheet <> -2147483648 Then
                writer.WriteElementString("ActiveSheet", "urn:schemas-microsoft-com:office:excel", Me._activeSheet.ToString(CultureInfo.InvariantCulture))
            End If
            Util.WriteElementString(writer, "ProtectStructure", "urn:schemas-microsoft-com:office:excel", Me._protectStructure)
            Util.WriteElementString(writer, "ProtectWindows", "urn:schemas-microsoft-com:office:excel", Me._protectWindows)
            writer.WriteEndElement()
        End Sub

		Friend Shared Function IsElement(element As XmlElement) As Boolean
			Return Util.IsElement(element, "ExcelWorkbook", "urn:schemas-microsoft-com:office:excel")
		End Function
	End Class
End Namespace
