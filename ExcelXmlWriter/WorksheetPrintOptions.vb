Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetPrintOptions
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _gridLines As Boolean

        Private _blackAndWhite As Boolean

        Private _draftQuality As Boolean

        Private _validPrinterInfo As Boolean

        Private _paperSizeIndex As Integer = -2147483648

        Private _horizontalResolution As Integer = 600

        Private _verticalResolution As Integer = 600

        Private _rowColHeadings As Boolean

        Private _scale As Integer = 100

        Private _fitWidth As Integer = 1

        Private _fitHeight As Integer = 1

        Private _leftToRight As Boolean

        Private _commentsLayout As PrintCommentsLayout

        Private _printErrors As PrintErrorsOption

        Public Property BlackAndWhite() As Boolean
            Get
                Return Me._blackAndWhite
            End Get
            Set(value As Boolean)
                Me._blackAndWhite = Value
            End Set
        End Property

        Public Property CommentsLayout() As PrintCommentsLayout
            Get
                Return Me._commentsLayout
            End Get
            Set(value As PrintCommentsLayout)
                Me._commentsLayout = Value
            End Set
        End Property

        Public Property DraftQuality() As Boolean
            Get
                Return Me._draftQuality
            End Get
            Set(value As Boolean)
                Me._draftQuality = Value
            End Set
        End Property

        Public Property FitHeight() As Integer
            Get
                Return Me._fitHeight
            End Get
            Set(value As Integer)
                Me._fitHeight = Value
            End Set
        End Property

        Public Property FitWidth() As Integer
            Get
                Return Me._fitWidth
            End Get
            Set(value As Integer)
                Me._fitWidth = Value
            End Set
        End Property

        Public Property GridLines() As Boolean
            Get
                Return Me._gridLines
            End Get
            Set(value As Boolean)
                Me._gridLines = Value
            End Set
        End Property

        Public Property HorizontalResolution() As Integer
            Get
                Return Me._horizontalResolution
            End Get
            Set(value As Integer)
                Me._horizontalResolution = Value
            End Set
        End Property

        Public Property LeftToRight() As Boolean
            Get
                Return Me._leftToRight
            End Get
            Set(value As Boolean)
                Me._leftToRight = Value
            End Set
        End Property

        Public Property PaperSizeIndex() As Integer
            Get
                If Me._paperSizeIndex = -2147483648 Then
                    Return 0
                End If
                Return Me._paperSizeIndex
            End Get
            Set(value As Integer)
                Me._paperSizeIndex = Value
            End Set
        End Property

        Public Property PrintErrors() As PrintErrorsOption
            Get
                Return Me._printErrors
            End Get
            Set(value As PrintErrorsOption)
                Me._printErrors = Value
            End Set
        End Property

        Public Property RowColHeadings() As Boolean
            Get
                Return Me._rowColHeadings
            End Get
            Set(value As Boolean)
                Me._rowColHeadings = Value
            End Set
        End Property

        Public Property Scale() As Integer
            Get
                Return Me._scale
            End Get
            Set(value As Integer)
                Me._scale = Value
            End Set
        End Property

        Public Property ValidPrinterInfo() As Boolean
            Get
                Return Me._validPrinterInfo
            End Get
            Set(value As Boolean)
                Me._validPrinterInfo = Value
            End Set
        End Property

        Public Property VerticalResolution() As Integer
            Get
                Return Me._verticalResolution
            End Get
            Set(value As Integer)
                Me._verticalResolution = Value
            End Set
        End Property

        Friend Sub New()
        End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._paperSizeIndex <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "PaperSizeIndex", Me._paperSizeIndex)
            End If
            If Me._horizontalResolution <> 600 Then
                Util.AddAssignment(method, targetObject, "HorizontalResolution", Me._horizontalResolution)
            End If
            If Me._verticalResolution <> 600 Then
                Util.AddAssignment(method, targetObject, "VerticalResolution", Me._verticalResolution)
            End If
            If Me._blackAndWhite Then
                Util.AddAssignment(method, targetObject, "BlackAndWhite", True)
            End If
            If Me._draftQuality Then
                Util.AddAssignment(method, targetObject, "DraftQuality", True)
            End If
            If Me._gridLines Then
                Util.AddAssignment(method, targetObject, "Gridlines", True)
            End If
            If Me._scale <> 100 Then
                Util.AddAssignment(method, targetObject, "Scale", Me._scale)
            End If
            If Me._fitWidth <> 1 Then
                Util.AddAssignment(method, targetObject, "FitWidth", Me._fitWidth)
            End If
            If Me._fitHeight <> 1 Then
                Util.AddAssignment(method, targetObject, "FitHeight", Me._fitHeight)
            End If
            If Me._leftToRight Then
                Util.AddAssignment(method, targetObject, "LeftToRight", True)
            End If
            If Me._rowColHeadings Then
                Util.AddAssignment(method, targetObject, "RowColHeadings", True)
            End If
            If Me._printErrors <> PrintErrorsOption.Displayed Then
                Util.AddAssignment(method, targetObject, "PrintErrors", CType(Me._printErrors, [Enum]))
            End If
            If Me._commentsLayout <> PrintCommentsLayout.None Then
                Util.AddAssignment(method, targetObject, "CommentsLayout", CType(Me._commentsLayout, [Enum]))
            End If
            If Me._validPrinterInfo Then
                Util.AddAssignment(method, targetObject, "ValidPrinterInfo", True)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetPrintOptions.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If Util.IsElement(xmlElement, "Gridlines", "urn:schemas-microsoft-com:office:excel") Then
                    Me._gridLines = True
                ElseIf Util.IsElement(xmlElement, "BlackAndWhite", "urn:schemas-microsoft-com:office:excel") Then
                    Me._blackAndWhite = True
                ElseIf Util.IsElement(xmlElement, "DraftQuality", "urn:schemas-microsoft-com:office:excel") Then
                    Me._draftQuality = True
                ElseIf Util.IsElement(xmlElement, "ValidPrinterInfo", "urn:schemas-microsoft-com:office:excel") Then
                    Me._validPrinterInfo = True
                ElseIf Util.IsElement(xmlElement, "PaperSizeIndex", "urn:schemas-microsoft-com:office:excel") Then
                    Me._paperSizeIndex = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "HorizontalResolution", "urn:schemas-microsoft-com:office:excel") Then
                    Me._horizontalResolution = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "VerticalResolution", "urn:schemas-microsoft-com:office:excel") Then
                    Me._verticalResolution = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "RowColHeadings", "urn:schemas-microsoft-com:office:excel") Then
                    Me._rowColHeadings = True
                ElseIf Util.IsElement(xmlElement, "LeftToRight", "urn:schemas-microsoft-com:office:excel") Then
                    Me._leftToRight = True
                ElseIf Util.IsElement(xmlElement, "Scale", "urn:schemas-microsoft-com:office:excel") Then
                    Me._scale = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "FitWidth", "urn:schemas-microsoft-com:office:excel") Then
                    Me._fitWidth = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Util.IsElement(xmlElement, "FitHeight", "urn:schemas-microsoft-com:office:excel") Then
                    Me._fitHeight = Integer.Parse(xmlElement.InnerText, CultureInfo.InvariantCulture)
                ElseIf Not Util.IsElement(xmlElement, "PrintErrors", "urn:schemas-microsoft-com:office:excel") Then
                    If Not Util.IsElement(xmlElement, "CommentsLayout", "urn:schemas-microsoft-com:office:excel") Then
                        Continue For
                    End If
                    Me._commentsLayout = CType([Enum].Parse(GetType(PrintCommentsLayout), xmlElement.InnerText), PrintCommentsLayout)
                Else
                    Me._printErrors = CType([Enum].Parse(GetType(PrintErrorsOption), xmlElement.InnerText), PrintErrorsOption)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "Print", "urn:schemas-microsoft-com:office:excel")
            If Me._paperSizeIndex <> -2147483648 Then
                writer.WriteElementString("PaperSizeIndex", "urn:schemas-microsoft-com:office:excel", Me._paperSizeIndex.ToString(CultureInfo.InvariantCulture))
            End If
            writer.WriteElementString("HorizontalResolution", "urn:schemas-microsoft-com:office:excel", Me._horizontalResolution.ToString(CultureInfo.InvariantCulture))
            writer.WriteElementString("VerticalResolution", "urn:schemas-microsoft-com:office:excel", Me._verticalResolution.ToString(CultureInfo.InvariantCulture))
            If Me._blackAndWhite Then
                writer.WriteElementString("BlackAndWhite", "urn:schemas-microsoft-com:office:excel", "")
            End If
            If Me._draftQuality Then
                writer.WriteElementString("DraftQuality", "urn:schemas-microsoft-com:office:excel", "")
            End If
            If Me._gridLines Then
                writer.WriteElementString("Gridlines", "urn:schemas-microsoft-com:office:excel", "")
            End If
            If Me._scale <> 100 Then
                writer.WriteElementString("Scale", "urn:schemas-microsoft-com:office:excel", Me._scale.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._fitWidth <> 1 Then
                writer.WriteElementString("FitWidth", "urn:schemas-microsoft-com:office:excel", Me._fitWidth.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._fitHeight <> 1 Then
                writer.WriteElementString("FitHeight", "urn:schemas-microsoft-com:office:excel", Me._fitHeight.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._leftToRight Then
                writer.WriteElementString("LeftToRight", "urn:schemas-microsoft-com:office:excel", "")
            End If
            If Me._rowColHeadings Then
                writer.WriteElementString("RowColHeadings", "urn:schemas-microsoft-com:office:excel", "")
            End If
            If Me._printErrors <> PrintErrorsOption.Displayed Then
                writer.WriteElementString("PrintErrors", "urn:schemas-microsoft-com:office:excel", Me._printErrors.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._commentsLayout <> PrintCommentsLayout.None Then
                writer.WriteElementString("CommentsLayout", "urn:schemas-microsoft-com:office:excel", Me._commentsLayout.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._validPrinterInfo Then
                writer.WriteElementString("ValidPrinterInfo", "urn:schemas-microsoft-com:office:excel", "")
            End If
            writer.WriteEndElement()
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Print", "urn:schemas-microsoft-com:office:excel")
        End Function
    End Class
End Namespace
