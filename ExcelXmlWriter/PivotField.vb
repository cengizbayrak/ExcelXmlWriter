Imports System.Globalization
Imports System.Xml

Namespace ExcelXmlWriter
	Public NotInheritable Class PivotField
		Implements IWriter
		Private _name As String

		Private _dataField As String

		Private _position As Integer = -2147483648

		Private _orientation As PivotFieldOrientation = PivotFieldOrientation.NotSet

		Private _function As PTFunction

        Private _dataType As ExcelXmlWriter.DataType

		Private _pivotItems As PivotItemCollection

		Private _parentField As String

		Public Property DataField() As String
			Get
				Return Me._dataField
			End Get
			Set
				Me._dataField = value
			End Set
		End Property

        Public Property DataType() As ExcelXmlWriter.DataType
            Get
                Return Me._dataType
            End Get
            Set(value As ExcelXmlWriter.DataType)
                Me._dataType = Value
            End Set
        End Property

		Public Property [Function]() As PTFunction
			Get
				Return Me._function
			End Get
			Set
				Me._function = value
			End Set
		End Property

		Public Property Name() As String
			Get
				Return Me._name
			End Get
			Set
				Me._name = value
			End Set
		End Property

		Public Property Orientation() As PivotFieldOrientation
			Get
				Return Me._orientation
			End Get
			Set
				Me._orientation = value
			End Set
		End Property

		Public Property ParentField() As String
			Get
				Return Me._parentField
			End Get
			Set
				Me._parentField = value
			End Set
		End Property

		Public ReadOnly Property PivotItems() As PivotItemCollection
			Get
				If Me._pivotItems Is Nothing Then
					Me._pivotItems = New PivotItemCollection()
				End If
				Return Me._pivotItems
			End Get
		End Property

		Public Property Position() As Integer
			Get
				Return Me._position
			End Get
			Set
				Me._position = value
			End Set
		End Property

		Public Sub New()
		End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("x", "PivotField", "urn:schemas-microsoft-com:office:excel")
            If Me._dataField IsNot Nothing Then
                writer.WriteElementString("DataField", "urn:schemas-microsoft-com:office:excel", Me._dataField)
            End If
            If Me._name IsNot Nothing Then
                writer.WriteElementString("Name", "urn:schemas-microsoft-com:office:excel", Me._name)
            End If
            If Me._parentField IsNot Nothing Then
                writer.WriteElementString("ParentField", "urn:schemas-microsoft-com:office:excel", Me._parentField)
            End If
            If Me._dataType <> ExcelXmlWriter.DataType.NotSet Then
                writer.WriteElementString("DataType", "urn:schemas-microsoft-com:office:excel", Me._dataType.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._function <> PTFunction.NotSet Then
                writer.WriteElementString("Function", "urn:schemas-microsoft-com:office:excel", Me._function.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._position <> -2147483648 Then
                writer.WriteElementString("Position", "urn:schemas-microsoft-com:office:excel", Me._position.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._orientation <> PivotFieldOrientation.NotSet Then
                writer.WriteElementString("Orientation", "urn:schemas-microsoft-com:office:excel", Me._orientation.ToString(CultureInfo.InvariantCulture))
            End If
            If Me._pivotItems IsNot Nothing Then
                DirectCast(Me._pivotItems, IWriter).WriteXml(writer)
            End If
            writer.WriteEndElement()
        End Sub
	End Class
End Namespace
