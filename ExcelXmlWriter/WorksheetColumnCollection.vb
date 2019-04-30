Imports System.CodeDom
Imports System.Collections
Imports System.Globalization
Imports System.Reflection
Imports System.Xml

Namespace ExcelXmlWriter
    Public NotInheritable Class WorksheetColumnCollection
        Inherits CollectionBase
        Implements IWriter
        Implements ICodeWriter
        Private _table As WorksheetTable

        Default Public ReadOnly Property Item(index As Integer) As WorksheetColumn
            Get
                Return DirectCast(MyBase.InnerList(index), WorksheetColumn)
            End Get
        End Property

        Friend Sub New(table As WorksheetTable)
            If table Is Nothing Then
                Throw New ArgumentNullException("table")
            End If
            Me._table = table
        End Sub

        Public Function Add(column As WorksheetColumn) As Integer
            If column Is Nothing Then
                Throw New ArgumentNullException("column")
            End If
            column._table = Me._table
            Return MyBase.InnerList.Add(column)
        End Function

        Public Function Add() As WorksheetColumn
            Dim worksheetColumn As New WorksheetColumn()
            Me.Add(worksheetColumn)
            Return worksheetColumn
        End Function

        Public Function Add(width As Integer) As WorksheetColumn
            Dim worksheetColumn As New WorksheetColumn() With { _
                .Width = width
            }
            Me.Add(worksheetColumn)
            Return worksheetColumn
        End Function

        Private Sub ExcelXmlWriter_ICodeWriter_WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression) Implements ExcelXmlWriter.ICodeWriter.WriteTo
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                Dim ıtem As WorksheetColumn = Me(i)
                If Not ıtem.IsSimple Then
                    Dim str As String = String.Concat("column", i.ToString(CultureInfo.InvariantCulture))
                    Dim codeVariableDeclarationStatement As New CodeVariableDeclarationStatement(GetType(WorksheetColumn), str, New CodeMethodInvokeExpression(targetObject, "Add", New CodeExpression(-1) {}))
                    method.Statements.Add(codeVariableDeclarationStatement)
                    DirectCast(ıtem, ICodeWriter).WriteTo(type, method, New CodeVariableReferenceExpression(str))
                Else
                    Dim statements As CodeStatementCollection = method.Statements
                    Dim codePrimitiveExpression As CodeExpression() = New CodeExpression() {New CodePrimitiveExpression(DirectCast(ıtem.Width, Object))}
                    statements.Add(New CodeMethodInvokeExpression(targetObject, "Add", codePrimitiveExpression))
                End If
            Next
        End Sub

        Private Sub ExcelXmlWriter_IWriter_WriteXml(writer As XmlWriter) Implements ExcelXmlWriter.IWriter.WriteXml
            For i As Integer = 0 To MyBase.InnerList.Count - 1
                DirectCast(MyBase.InnerList(i), IWriter).WriteXml(writer)
            Next
        End Sub

        Public Function Contains(item As WorksheetColumn) As Boolean
            Return MyBase.InnerList.Contains(item)
        End Function

        Public Sub CopyTo(array As WorksheetColumn(), index As Integer)
            MyBase.InnerList.CopyTo(array, index)
        End Sub

        Public Function IndexOf(item As WorksheetColumn) As Integer
            Return MyBase.InnerList.IndexOf(item)
        End Function

        Public Sub Insert(index As Integer, item As WorksheetColumn)
            MyBase.InnerList.Insert(index, item)
        End Sub

        Public Sub Remove(item As WorksheetColumn)
            MyBase.InnerList.Remove(item)
        End Sub
    End Class
End Namespace
