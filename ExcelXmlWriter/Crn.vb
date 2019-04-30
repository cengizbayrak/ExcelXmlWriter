Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class Crn
		Implements IWriter
		Implements IReader
		Implements ICodeWriter
		Private _row As String

		Private _text As String

		Private _numbers As NumberCollection

		Private _colFirst As Integer = -2147483648

		Private _colLast As Integer = -2147483648

		Public Property ColFirst() As Integer
			Get
				If Me._colFirst = -2147483648 Then
					Return 0
				End If
				Return Me._colFirst
			End Get
			Set
				If value < 0 Then
					Throw New ArgumentException("Invalid range, > 0")
				End If
				Me._colFirst = value
			End Set
		End Property

		Public Property ColLast() As Integer
			Get
				If Me._colLast = -2147483648 Then
					Return 0
				End If
				Return Me._colLast
			End Get
			Set
				If value < 0 Then
					Throw New ArgumentException("Invalid range, > 0")
				End If
				Me._colLast = value
			End Set
		End Property

		Public ReadOnly Property Numbers() As NumberCollection
			Get
				If Me._numbers Is Nothing Then
					Me._numbers = New NumberCollection()
				End If
				Return Me._numbers
			End Get
		End Property

		Public Property Row() As String
			Get
				Return Me._row
			End Get
			Set
				Me._row = value
			End Set
		End Property

		Public Property Text() As String
			Get
				Return Me._text
			End Get
			Set
				Me._text = value
			End Set
		End Property

		Public Sub New()
		End Sub

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            If Me._row IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Row", Me._row)
            End If
            If Me._colFirst <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "ColFirst", Me._colFirst)
            End If
            If Me._colLast <> -2147483648 Then
                Util.AddAssignment(method, targetObject, "ColLast", Me._colLast)
            End If
            If Me._numbers IsNot Nothing Then
                DirectCast(Me._numbers, ICodeWriter).WriteTo(type, method, targetObject)
            End If
            If Me._text IsNot Nothing Then
                Util.AddAssignment(method, targetObject, "Text", Me._text)
            End If
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not Crn.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing Then
                    Continue For
                End If
                If Util.IsElement(xmlElement, "Row", "urn:schemas-microsoft-com:office:excel") Then
                    Me._row = xmlElement.InnerText
                ElseIf Util.IsElement(xmlElement, "ColFirst", "urn:schemas-microsoft-com:office:excel") Then
                    Me._colFirst = Integer.Parse(xmlElement.InnerText)
                ElseIf Util.IsElement(xmlElement, "ColLast", "urn:schemas-microsoft-com:office:excel") Then
                    Me._colLast = Integer.Parse(xmlElement.InnerText)
                ElseIf Not Util.IsElement(xmlElement, "Number", "urn:schemas-microsoft-com:office:excel") Then
                    If Not Util.IsElement(xmlElement, "Text", "urn:schemas-microsoft-com:office:excel") Then
                        Continue For
                    End If
                    Me._text = xmlElement.InnerText
                Else
                    Me.Numbers.Add(xmlElement.InnerText)
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "Crn", "urn:schemas-microsoft-com:office:excel")
            If Me._row IsNot Nothing Then
                writer.WriteElementString("Row", "urn:schemas-microsoft-com:office:excel", Me._row)
            End If
            If Me._colFirst <> -2147483648 Then
                writer.WriteElementString("ColFirst", "urn:schemas-microsoft-com:office:excel", Me._colFirst.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._colLast <> -2147483648 Then
                writer.WriteElementString("ColLast", "urn:schemas-microsoft-com:office:excel", Me._colLast.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._numbers IsNot Nothing Then
                DirectCast(Me._numbers, IWriter).WriteXml(writer)
            End If
            If Me._text IsNot Nothing Then
                writer.WriteElementString("Text", "urn:schemas-microsoft-com:office:excel", Me._text)
            End If
            writer.WriteEndElement()
        End Sub

		Friend Shared Function IsElement(element As XmlElement) As Boolean
			Return Util.IsElement(element, "Crn", "urn:schemas-microsoft-com:office:excel")
		End Function
	End Class
End Namespace
