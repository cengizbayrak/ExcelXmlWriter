Imports System.CodeDom

Namespace ExcelXmlWriter
	Public Interface ICodeWriter
		Sub WriteTo(type As CodeTypeDeclaration, method As CodeMemberMethod, targetObject As CodeExpression)
	End Interface
End Namespace
