Imports System.CodeDom
Imports System.Collections
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetCellCollection
        Inherits CollectionBase
        Implements IWriter
        Implements ICodeWriter
        Private _row As WorksheetRow

        Default Public Property Item(index As Integer) As WorksheetCell
            Get
                Return DirectCast(MyBase.InnerList(index), WorksheetCell)
            End Get
            Set(value As WorksheetCell)
                If value IsNot Nothing Then
                    value._row = Me._row
                End If
                MyBase.InnerList(index) = value
            End Set
        End Property

        Friend Sub New(row As WorksheetRow)
            If row Is Nothing Then
                Throw New ArgumentNullException("row")
            End If
            Me._row = row
        End Sub

        Public Function Add(cell As WorksheetCell) As Integer
            If cell Is Nothing Then
                Throw New ArgumentNullException("cell")
            End If
            cell._row = Me._row
            Return MyBase.InnerList.Add(cell)
        End Function

        Public Function Add() As WorksheetCell
            Dim worksheetCell As New WorksheetCell()
            Me.Add(worksheetCell)
            Return worksheetCell
        End Function

        Public Function Add(text As String) As WorksheetCell
            Dim worksheetCell As New WorksheetCell(text)
            Me.Add(worksheetCell)
            Return worksheetCell
        End Function

        Public Function Add(text As String, dataType As DataType, styleID As String) As WorksheetCell
            Dim worksheetCell As New WorksheetCell(text, dataType, styleID)
            Me.Add(worksheetCell)
            Return worksheetCell
        End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Dim ıtem As WorksheetCell = Me(i)
                If Not ıtem.IsSimple Then
                    Dim str As String = "cell"
                    If Worksheet.cellDeclaration Is Nothing Then
                        Worksheet.cellDeclaration = New CodeVariableDeclarationStatement(GetType(WorksheetCell), "cell")
                        method.Statements.Add(Worksheet.cellDeclaration)
                    End If
                    Dim codeAssignStatement As New CodeAssignStatement(New CodeVariableReferenceExpression("cell"), New CodeMethodInvokeExpression(targetObject, "Add", New CodeExpression(-1) {}))
                    method.Statements.Add(codeAssignStatement)
                    DirectCast(ıtem, ICodeWriter).WriteTo(type, method, New CodeVariableReferenceExpression(str))
                Else
                    Dim statements As CodeStatementCollection = method.Statements
                    Dim codePrimitiveExpression As CodeExpression() = New CodeExpression() {New CodePrimitiveExpression(ıtem.Data.Text), Util.GetRightExpressionForValue(ıtem.Data.Type, GetType(DataType)), New CodePrimitiveExpression(ıtem.StyleID)}
                    statements.Add(New CodeMethodInvokeExpression(targetObject, "Add", codePrimitiveExpression))
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

        Public Function Contains(item As WorksheetCell) As Boolean
            Return MyBase.InnerList.Contains(item)
        End Function

        Public Sub CopyTo(array As WorksheetCell(), index As Integer)
            MyBase.InnerList.CopyTo(array, index)
        End Sub

        Public Function IndexOf(item As WorksheetCell) As Integer
            Return MyBase.InnerList.IndexOf(item)
        End Function

        Public Sub Insert(index As Integer, item As WorksheetCell)
            MyBase.InnerList.Insert(index, item)
        End Sub

        Public Sub Remove(item As WorksheetCell)
            MyBase.InnerList.Remove(item)
        End Sub

        Public Function ToArray() As Object()
            Return MyBase.InnerList.ToArray()
        End Function
    End Class
End Namespace
