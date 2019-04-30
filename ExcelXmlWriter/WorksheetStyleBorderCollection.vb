Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetStyleBorderCollection
        Inherits CollectionBase
        Implements IWriter
        Implements IReader
        Implements ICodeWriter
        Private _style As WorksheetStyle

        Default Public Property Item(index As Integer) As WorksheetStyleBorder
            Get
                Return DirectCast(MyBase.InnerList(index), WorksheetStyleBorder)
            End Get
            Set(value As WorksheetStyleBorder)
                MyBase.InnerList(index) = value
            End Set
        End Property

        Friend Sub New(style As WorksheetStyle)
            Me._style = style
        End Sub

        Public Function Add(border As WorksheetStyleBorder) As Integer
            Return MyBase.InnerList.Add(border)
        End Function

        Public Function Add() As WorksheetStyleBorder
            Dim worksheetStyleBorder As New WorksheetStyleBorder()
            Me.Add(worksheetStyleBorder)
            Return worksheetStyleBorder
        End Function

        Public Function Add(position As StylePosition, lineStyle As LineStyleOption) As WorksheetStyleBorder
            Return Me.Add(position, lineStyle, 0, Nothing)
        End Function

        Public Function Add(position As StylePosition, lineStyle As LineStyleOption, weight As Integer) As WorksheetStyleBorder
            Return Me.Add(position, lineStyle, weight, Nothing)
        End Function

        Public Function Add(position As StylePosition, lineStyle As LineStyleOption, weight As Integer, color As String) As WorksheetStyleBorder
            Dim worksheetStyleBorder As New WorksheetStyleBorder() With {
                .Position = position,
                .LineStyle = lineStyle,
                .Weight = weight,
                .Color = color
            }
            Me.Add(worksheetStyleBorder)
            Return worksheetStyleBorder
        End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            Dim rightExpressionForValue As CodeExpression()
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Dim ıtem As WorksheetStyleBorder = Me(i)
                Dim ınt32 As Integer = ıtem.IsSpecial()
                If ınt32 = 0 Then
                    Dim statements As CodeStatementCollection = method.Statements
                    rightExpressionForValue = New CodeExpression() {Util.GetRightExpressionForValue(ıtem.Position, GetType(StylePosition)), Util.GetRightExpressionForValue(ıtem.LineStyle, GetType(LineStyleOption)), New CodePrimitiveExpression(DirectCast(ıtem.Weight, Object))}
                    statements.Add(New CodeMethodInvokeExpression(targetObject, "Add", rightExpressionForValue))
                ElseIf ınt32 <> 1 Then
                    Dim str As String = String.Concat(Me._style.ID, "Border", i.ToString(CultureInfo.InvariantCulture))
                    Dim codeVariableDeclarationStatement As New CodeVariableDeclarationStatement(GetType(WorksheetStyleBorder), str, New CodeMethodInvokeExpression(targetObject, "Add", New CodeExpression(-1) {}))
                    method.Statements.Add(codeVariableDeclarationStatement)
                    DirectCast(ıtem, ICodeWriter).WriteTo(type, method, New CodeVariableReferenceExpression(str))
                Else
                    Dim codeStatementCollection As CodeStatementCollection = method.Statements
                    rightExpressionForValue = New CodeExpression() {Util.GetRightExpressionForValue(ıtem.Position, GetType(StylePosition)), Util.GetRightExpressionForValue(ıtem.LineStyle, GetType(LineStyleOption)), New CodePrimitiveExpression(DirectCast(ıtem.Weight, Object)), New CodePrimitiveExpression(ıtem.Color)}
                    codeStatementCollection.Add(New CodeMethodInvokeExpression(targetObject, "Add", rightExpressionForValue))
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IReader_ReadXml(element As XmlElement) Implements ExcelXmlWriter.IReader.ReadXml
            If Not WorksheetStyleBorderCollection.IsElement(element) Then
                Throw New ArgumentException("Invalid element", "element")
            End If
            For Each childNode As XmlNode In element.ChildNodes
                Dim xmlElement As XmlElement = TryCast(childNode, XmlElement)
                If xmlElement Is Nothing OrElse Not WorksheetStyleBorder.IsElement(xmlElement) Then
                    Continue For
                End If
                Dim worksheetStyleBorder__1 As New WorksheetStyleBorder()
                DirectCast(worksheetStyleBorder__1, IReader).ReadXml(xmlElement)
                Me.Add(worksheetStyleBorder__1)
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            writer.WriteStartElement("s", "Borders", "urn:schemas-microsoft-com:office:spreadsheet")
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
            writer.WriteEndElement()
        End Sub

        Public Function Contains(item As WorksheetStyleBorder) As Boolean
            Return MyBase.InnerList.Contains(item)
        End Function

        Public Sub CopyTo(array As WorksheetStyleBorder(), index As Integer)
            MyBase.InnerList.CopyTo(array, index)
        End Sub

        Public Function IndexOf(item As WorksheetStyleBorder) As Integer
            Return MyBase.InnerList.IndexOf(item)
        End Function

        Public Sub Insert(index As Integer, border As WorksheetStyleBorder)
            MyBase.InnerList.Insert(index, border)
        End Sub

        Friend Shared Function IsElement(element As XmlElement) As Boolean
            Return Util.IsElement(element, "Borders", "urn:schemas-microsoft-com:office:spreadsheet")
        End Function

        Public Sub Remove(border As WorksheetStyleBorder)
            MyBase.InnerList.Remove(border)
        End Sub

        Public Function ToArray() As Object()
            Return MyBase.InnerList.ToArray()
        End Function
    End Class
End Namespace
