Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetOptions
        Implements IWriter
        Implements IReader
        Implements ICodeWriter

        Private _topRowVisible As Integer = -2147483648

        Private _splitHorizontal As Integer = -2147483648

        Private _splitVertical As Integer = -2147483648

        Private _topRowBottomPane As Integer = -2147483648

        Private _leftColumnRightPane As Integer = -2147483648

        Private _activePane As Integer = -2147483648

        Private _viewableRange As String

        Private _gridLineColor As String

        Private _pageSetup As WorksheetPageSetup

        Private _print As WorksheetPrintOptions

        Private _protectObjects As Boolean

        Private _protectScenarios As Boolean

        Private _selected As Boolean

        Private _freezePanes As Boolean

        Private _fitToPage As Boolean

        Public Property ActivePane() As Integer
            Get
                Return Me._activePane
            End Get
            Set(value As Integer)
                Me._activePane = Value
            End Set
        End Property

        Public Property FitToPage() As Boolean
            Get
                Return Me._fitToPage
            End Get
            Set(value As Boolean)
                Me._fitToPage = Value
            End Set
        End Property

        Public Property FreezePanes() As Boolean
            Get
                Return Me._freezePanes
            End Get
            Set(value As Boolean)
                Me._freezePanes = Value
            End Set
        End Property

        Public Property GridLineColor() As String
            Get
                Return Me._gridLineColor
            End Get
            Set(value As String)
                Me._gridLineColor = Value
            End Set
        End Property

        Public Property LeftColumnRightPane() As Integer
            Get
                Return Me._leftColumnRightPane
            End Get
            Set(value As Integer)
                Me._leftColumnRightPane = Value
            End Set
        End Property

        Public ReadOnly Property PageSetup() As WorksheetPageSetup
            Get
                If Me._pageSetup Is Nothing Then
                    Me._pageSetup = New WorksheetPageSetup()
                End If
                Return Me._pageSetup
            End Get
        End Property

        Public ReadOnly Property Print() As WorksheetPrintOptions
            Get
                If Me._print Is Nothing Then
                    Me._print = New WorksheetPrintOptions()
                End If
                Return Me._print
            End Get
        End Property

        Public Property ProtectObjects() As Boolean
            Get
                Return Me._protectObjects
            End Get
            Set(value As Boolean)
                Me._protectObjects = Value
            End Set
        End Property

        Public Property ProtectScenarios() As Boolean
            Get
                Return Me._protectScenarios
            End Get
            Set(value As Boolean)
                Me._protectScenarios = Value
            End Set
        End Property

        Public Property Selected() As Boolean
            Get
                Return Me._selected
            End Get
            Set(value As Boolean)
                Me._selected = Value
            End Set
        End Property

        Public Property SplitHorizontal() As Integer
            Get
                Return Me._splitHorizontal
            End Get
            Set(value As Integer)
                Me._splitHorizontal = Value
            End Set
        End Property

        Public Property SplitVertical() As Integer
            Get
                Return Me._splitVertical
            End Get
            Set(value As Integer)
                Me._splitVertical = Value
            End Set
        End Property

        Public Property TopRowBottomPane() As Integer
            Get
                Return Me._topRowBottomPane
            End Get
            Set(value As Integer)
                Me._topRowBottomPane = Value
            End Set
        End Property

        Public Property TopRowVisible() As Integer
            Get
                Return Me._topRowVisible
            End Get
            Set(value As Integer)
                Me._topRowVisible = Value
            End Set
        End Property

        Public Property ViewableRange() As String
            Get
                Return Me._viewableRange
            End Get
            Set(value As String)
                Me._viewableRange = Value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            Util.AddComment(method, "Options")
            If Me._selected Then
                Util.AddAssignment(method, targetObject, "Selected", Me._selected)
            End If
            If Me._topRowVisible <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "TopRowVisible", Me._topRowVisible)
            End If
            If Me._freezePanes Then
                Util.AddAssignment(method, targetObject, "FreezePanes", Me._freezePanes)
            End If
            If Me._fitToPage Then
                Util.AddAssignment(method, targetObject, "FitToPage", Me._fitToPage)
            End If
            If Me._splitHorizontal <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "SplitHorizontal", Me._splitHorizontal)
            End If
            If Me._topRowBottomPane <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "TopRowBottomPane", Me._topRowBottomPane)
            End If
            If Me._splitVertical <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "SplitVertical", Me._splitVertical)
            End If
            If Me._leftColumnRightPane <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "LeftColumnRightPane", Me._leftColumnRightPane)
            End If
            If Me._activePane <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "ActivePane", Me._activePane)
            End If
            If Me._viewableRange IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "ViewableRange", Me._viewableRange)
            End If
            If Me._gridLineColor IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "GridlineColor", Me._gridLineColor)
            End If
            Util.AddAssignment(method, targetObject, "ProtectObjects", Me._protectObjects)
            Util.AddAssignment(method, targetObject, "ProtectScenarios", Me._protectScenarios)
            If Me._pageSetup IsNot Nothing Then
                Util.Traverse(type, Me._pageSetup, method, targetObject, "PageSetup")
            End If
            If Me._print IsNot Nothing Then
                Util.Traverse(type, Me._print, method, targetObject, "Print")
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetOptions.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If Util.IsElement(xmlElement, "Selected", "urn:schemas-microsoft-com:office:excel") Then
                    Me._selected = True
                ElseIf Util.IsElement(xmlElement, "TopRowVisible", "urn:schemas-microsoft-com:office:excel") Then
                    Me._topRowVisible = Util.GetAttribute(xmlElement, "TopRowVisible", "urn:schemas-microsoft-com:office:excel", Int32.MinValue)
                ElseIf Util.IsElement(xmlElement, "FreezePanes", "urn:schemas-microsoft-com:office:excel") Then
                    Me._freezePanes = True
                ElseIf Util.IsElement(xmlElement, "SplitHorizontal", "urn:schemas-microsoft-com:office:excel") Then
                    Me._splitHorizontal = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "TopRowBottomPane", "urn:schemas-microsoft-com:office:excel") Then
                    Me._topRowBottomPane = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "SplitVertical", "urn:schemas-microsoft-com:office:excel") Then
                    Me._splitVertical = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "LeftColumnRightPane", "urn:schemas-microsoft-com:office:excel") Then
                    Me._leftColumnRightPane = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "ActivePane", "urn:schemas-microsoft-com:office:excel") Then
                    Me._activePane = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "ViewableRange", "urn:schemas-microsoft-com:office:excel") Then
                    Me._viewableRange = xmlElement.InnerText
                ElseIf Util.IsElement(xmlElement, "GridlineColor", "urn:schemas-microsoft-com:office:excel") Then
                    Me._gridLineColor = xmlElement.InnerText
                ElseIf Util.IsElement(xmlElement, "ProtectObjects", "urn:schemas-microsoft-com:office:excel") Then
                    Me._protectObjects = Boolean.Parse(xmlElement.InnerText)
                ElseIf Util.IsElement(xmlElement, "ProtectScenarios", "urn:schemas-microsoft-com:office:excel") Then
                    Me._protectScenarios = Boolean.Parse(xmlElement.InnerText)
                ElseIf Util.IsElement(xmlElement, "FitToPage", "urn:schemas-microsoft-com:office:excel") Then
                    Me._fitToPage = True
                ElseIf Not WorksheetPageSetup.IsElement(xmlElement) Then
                    If Not WorksheetPrintOptions.IsElement(xmlElement) Then
                        Continue For
                    End If
                    DirectCast(Me.Print, IReader).ReadXml(xmlElement)
                Else
                    DirectCast(Me.PageSetup, IReader).ReadXml(xmlElement)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "WorksheetOptions", "urn:schemas-microsoft-com:office:excel")
            If Me._selected Then
                writer.WriteElementString("Selected", "urn:schemas-microsoft-com:office:excel", "")
            End If
            If Me._fitToPage Then
                writer.WriteElementString("FitToPage", "urn:schemas-microsoft-com:office:excel", "")
            End If
            If Me._topRowVisible <> -2147483648 Then
                writer.WriteElementString("TopRowVisible", "urn:schemas-microsoft-com:office:excel", Me._topRowVisible.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._freezePanes Then
                writer.WriteElementString("FreezePanes", "urn:schemas-microsoft-com:office:excel", "")
            End If
            If Me._splitHorizontal <> -2147483648 Then
                writer.WriteElementString("SplitHorizontal", "urn:schemas-microsoft-com:office:excel", Me._splitHorizontal.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._topRowBottomPane <> -2147483648 Then
                writer.WriteElementString("TopRowBottomPane", "urn:schemas-microsoft-com:office:excel", Me._topRowBottomPane.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._splitVertical <> -2147483648 Then
                writer.WriteElementString("SplitVertical", "urn:schemas-microsoft-com:office:excel", Me._splitVertical.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._leftColumnRightPane <> -2147483648 Then
                writer.WriteElementString("LeftColumnRightPane", "urn:schemas-microsoft-com:office:excel", Me._leftColumnRightPane.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._activePane <> -2147483648 Then
                writer.WriteElementString("ActivePane", "urn:schemas-microsoft-com:office:excel", Me._activePane.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._viewableRange IsNot Nothing Then
                writer.WriteElementString("ViewableRange", "urn:schemas-microsoft-com:office:excel", Me._viewableRange)
            End If
            If Me._gridLineColor IsNot Nothing Then
                writer.WriteElementString("GridlineColor", "urn:schemas-microsoft-com:office:excel", Me._gridLineColor)
            End If
            writer.WriteElementString("ProtectObjects", "urn:schemas-microsoft-com:office:excel", Me._protectObjects.ToString(CultureInfo.InvariantCulture))
            writer.WriteElementString("ProtectScenarios", "urn:schemas-microsoft-com:office:excel", Me._protectScenarios.ToString(CultureInfo.InvariantCulture))
            If Me._pageSetup IsNot Nothing Then
                DirectCast(Me._pageSetup, IWriter).WriteXml(writer)
            End If
            If Me._print IsNot Nothing Then
                DirectCast(Me._print, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "WorksheetOptions", "urn:schemas-microsoft-com:office:excel")
        End Function
    End Class
End Namespace
