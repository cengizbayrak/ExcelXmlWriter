Imports System.CodeDom
Imports System.Collections
Imports System.ComponentModel
Imports System.ComponentModel.Design.Serialization
Imports System.Globalization
Imports System.Reflection
Imports System.Text
Imports System.Xml

Namespace ExcelXmlWriter
    Friend NotInheritable Class Util
        Private Sub New()
        End Sub

        Public Shared Sub AddAssignment(method As CodeMemberMethod, targetObject As CodeExpression, [property] As String, value As String)
            method.Statements.Add(Util.GetAssignment(targetObject, [property], value))
        End Sub

        Public Shared Sub AddAssignment(method As CodeMemberMethod, targetObject As CodeExpression, [property] As String, value As Single)
            method.Statements.Add(New CodeAssignStatement(New CodePropertyReferenceExpression(targetObject, [property]), Util.GetRightExpressionForValue(value, GetType(Single))))
        End Sub

        Public Shared Sub AddAssignment(method As CodeMemberMethod, targetObject As CodeExpression, [property] As String, enumValue As [Enum])
            method.Statements.Add(New CodeAssignStatement(New CodePropertyReferenceExpression(targetObject, [property]), Util.GetRightExpressionForValue(enumValue, enumValue.[GetType]())))
        End Sub

        Public Shared Sub AddAssignment(method As CodeMemberMethod, targetObject As CodeExpression, [property] As String, value As Integer)
            method.Statements.Add(Util.GetAssignment(targetObject, [property], value))
        End Sub

        Public Shared Sub AddAssignment(method As CodeMemberMethod, targetObject As CodeExpression, [property] As String, value As Boolean)
            method.Statements.Add(Util.GetAssignment(targetObject, [property], value))
        End Sub

        Public Shared Sub AddAssignment(method As CodeMemberMethod, targetObject As CodeExpression, [property] As String, value As DateTime)
            method.Statements.Add(Util.GetAssignment(targetObject, [property], value))
        End Sub

        Public Shared Sub AddComment(method As CodeMemberMethod, text As String)
            Util.AddCommentSeparator(method)
            method.Statements.Add(New CodeCommentStatement(String.Concat(" ", text)))
            Util.AddCommentSeparator(method)
        End Sub

        Public Shared Sub AddCommentSeparator(method As CodeMemberMethod)
            method.Statements.Add(New CodeCommentStatement("-----------------------------------------------"))
        End Sub

        Friend Shared Function CreateSafeName(strName As String, prefix As String) As String
            Dim stringBuilder As New StringBuilder()
            For i As Integer = 0 To strName.Length - 1
                Dim chr As Char = strName(i)
                If chr >= "A"c AndAlso chr <= "Z"c OrElse chr >= "a"c AndAlso chr <= "z"c OrElse chr = "_"c OrElse chr >= "0"c AndAlso chr <= "9"c Then
                    stringBuilder.Append(chr)
                End If
            Next
            If stringBuilder.Length = 0 Then
                Return prefix
            End If
            If stringBuilder(0) < "0"c OrElse stringBuilder(0) > "9"c Then
                Return stringBuilder.ToString()
            End If
            Return String.Concat(prefix, stringBuilder.ToString())
        End Function

        Private Shared Function GetAssignment(targetObject As CodeExpression, [property] As String, value As String) As CodeAssignStatement
            Return New CodeAssignStatement(New CodePropertyReferenceExpression(targetObject, [property]), Util.GetRightExpressionForValue(value, GetType(String)))
        End Function

        Private Shared Function GetAssignment(targetObject As CodeExpression, [property] As String, value As DateTime) As CodeAssignStatement
            Return New CodeAssignStatement(New CodePropertyReferenceExpression(targetObject, [property]), Util.GetRightExpressionForValue(value, GetType(DateTime)))
        End Function

        Private Shared Function GetAssignment(targetObject As CodeExpression, [property] As String, value As Boolean) As CodeAssignStatement
            Return New CodeAssignStatement(New CodePropertyReferenceExpression(targetObject, [property]), Util.GetRightExpressionForValue(value, GetType(Boolean)))
        End Function

        Private Shared Function GetAssignment(targetObject As CodeExpression, [property] As String, value As Integer) As CodeAssignStatement
            Return New CodeAssignStatement(New CodePropertyReferenceExpression(targetObject, [property]), Util.GetRightExpressionForValue(value, GetType(Integer)))
        End Function

        Public Shared Function GetAttribute(element As XmlElement, name As String, ns As String, defaultValue As Integer) As Integer
            Dim attribute As String = element.GetAttribute(name, ns)
            If attribute Is Nothing OrElse attribute.Length = 0 Then
                Return defaultValue
            End If
            Dim ınt32 As Integer = defaultValue
            Try
                Dim ınt321 As Integer = attribute.IndexOf("."c)
                If ınt321 <> -1 Then
                    attribute = attribute.Substring(0, ınt321)
                End If
                ınt32 = Integer.Parse(attribute, CultureInfo.InvariantCulture)
            Catch
            End Try
            Return ınt32
        End Function

        Public Shared Function GetAttribute(element As XmlElement, name As String, ns As String, defaultValue As Boolean) As Boolean
            Dim attribute As String = element.GetAttribute(name, ns)
            If attribute Is Nothing OrElse attribute.Length = 0 Then
                Return defaultValue
            End If
            If attribute = "1" Then
                Return True
            End If
            Return False
        End Function

        Public Shared Function GetAttribute(element As XmlElement, name As String, ns As String, defaultValue As Single) As Single
            Dim attribute As String = element.GetAttribute(name, ns)
            If attribute Is Nothing OrElse attribute.Length = 0 Then
                Return defaultValue
            End If
            Dim [single] As Single = defaultValue
            Try
                [single] = Single.Parse(attribute, CultureInfo.InvariantCulture)
            Catch
            End Try
            Return [single]
        End Function

        Public Shared Function GetAttribute(element As XmlElement, name As String, ns As String) As String
            Dim attributeNode As XmlAttribute = element.GetAttributeNode(name, ns)
            If attributeNode Is Nothing Then
                Return Nothing
            End If
            Return attributeNode.Value
        End Function

        Public Shared Function GetRightExpressionForValue(value As Object, valueType As Type) As CodeExpression
            Dim codeExpressionArray As CodeExpression()
            Dim j As Integer
            If TypeOf value Is DBNull Then
                If DirectCast(valueType, Object) Is DirectCast(GetType(Decimal), Object) Then
                    Return New CodePrimitiveExpression(Nothing)
                End If
                Return New CodePrimitiveExpression(Nothing)
            End If
            If DirectCast(valueType, Object) Is DirectCast(GetType(String), Object) AndAlso TypeOf value Is String Then
                Return New CodePrimitiveExpression(DirectCast(value, String))
            End If
            If valueType.IsPrimitive Then
                If DirectCast(valueType, Object) Is DirectCast(value.[GetType](), Object) Then
                    Return New CodePrimitiveExpression(value)
                End If
                Dim converter As TypeConverter = TypeDescriptor.GetConverter(valueType)
                If Not converter.CanConvertFrom(value.[GetType]()) Then
                    Return New CodePrimitiveExpression(value)
                End If
                Return New CodePrimitiveExpression(converter.ConvertFrom(value))
            End If
            If valueType.IsArray Then
                Dim array As Array = DirectCast(value, Array)
                Dim codeArrayCreateExpression As New CodeArrayCreateExpression()
                codeArrayCreateExpression.CreateType = New CodeTypeReference(valueType.GetElementType())
                If array IsNot Nothing Then
                    For Each obj As Object In array
                        codeArrayCreateExpression.Initializers.Add(Util.GetRightExpressionForValue(obj, valueType.GetElementType()))
                    Next
                End If
                Return codeArrayCreateExpression
            End If
            Dim typeConverter As TypeConverter = Nothing
            typeConverter = TypeDescriptor.GetConverter(valueType)
            If valueType.IsEnum AndAlso TypeOf value Is String Then
                value = typeConverter.ConvertFromString(value.ToString())
            End If
            If typeConverter IsNot Nothing Then
                Dim ınstanceDescriptor As InstanceDescriptor = Nothing
                If typeConverter.CanConvertTo(GetType(InstanceDescriptor)) Then
                    ınstanceDescriptor = DirectCast(typeConverter.ConvertTo(value, GetType(InstanceDescriptor)), InstanceDescriptor)
                End If
                If ınstanceDescriptor IsNot Nothing Then
                    If TypeOf ınstanceDescriptor.MemberInfo Is FieldInfo Then
                        Dim codeFieldReferenceExpression As New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(ınstanceDescriptor.MemberInfo.DeclaringType.FullName), ınstanceDescriptor.MemberInfo.Name)
                        Return codeFieldReferenceExpression
                    End If
                    If TypeOf ınstanceDescriptor.MemberInfo Is PropertyInfo Then
                        Dim codePropertyReferenceExpression As New CodePropertyReferenceExpression(New CodeTypeReferenceExpression(ınstanceDescriptor.MemberInfo.DeclaringType.FullName), ınstanceDescriptor.MemberInfo.Name)
                        Return codePropertyReferenceExpression
                    End If
                    Dim objArray As Object() = New Object(ınstanceDescriptor.Arguments.Count - 1) {}
                    ınstanceDescriptor.Arguments.CopyTo(objArray, 0)
                    Dim rightExpressionForValue As CodeExpression() = New CodeExpression(CInt(objArray.Length) - 1) {}
                    If TypeOf ınstanceDescriptor.MemberInfo Is MethodInfo Then
                        Dim parameters As ParameterInfo() = DirectCast(ınstanceDescriptor.MemberInfo, MethodInfo).GetParameters()
                        For i As Integer = 0 To CInt(objArray.Length) - 1
                            rightExpressionForValue(i) = Util.GetRightExpressionForValue(objArray(i), parameters(i).ParameterType)
                        Next
                        Dim codeMethodInvokeExpression As New CodeMethodInvokeExpression(New CodeTypeReferenceExpression(ınstanceDescriptor.MemberInfo.DeclaringType.FullName), ınstanceDescriptor.MemberInfo.Name, New CodeExpression(-1) {})
                        codeExpressionArray = rightExpressionForValue
                        For j = 0 To CInt(codeExpressionArray.Length) - 1
                            Dim codeExpression As CodeExpression = codeExpressionArray(j)
                            codeMethodInvokeExpression.Parameters.Add(codeExpression)
                        Next
                        Return codeMethodInvokeExpression
                    End If
                    If TypeOf ınstanceDescriptor.MemberInfo Is ConstructorInfo Then
                        Dim parameterInfoArray As ParameterInfo() = DirectCast(ınstanceDescriptor.MemberInfo, ConstructorInfo).GetParameters()
                        For k As Integer = 0 To CInt(objArray.Length) - 1
                            rightExpressionForValue(k) = Util.GetRightExpressionForValue(objArray(k), parameterInfoArray(k).ParameterType)
                        Next
                        Dim codeObjectCreateExpression As New CodeObjectCreateExpression(ınstanceDescriptor.MemberInfo.DeclaringType.FullName, New CodeExpression(-1) {})
                        codeExpressionArray = rightExpressionForValue
                        For j = 0 To CInt(codeExpressionArray.Length) - 1
                            Dim codeExpression1 As CodeExpression = codeExpressionArray(j)
                            codeObjectCreateExpression.Parameters.Add(codeExpression1)
                        Next
                        Return codeObjectCreateExpression
                    End If
                End If
            End If
            Return Nothing
        End Function

        Public Shared Function IsElement(element As XmlElement, name As String, ns As String) As Boolean
            If element Is Nothing OrElse Not (element.LocalName = name) Then
                Return False
            End If
            Return element.NamespaceURI = ns
        End Function

        Public Shared Sub Traverse(type As CodeTypeDeclaration, o As ICodeWriter, method As CodeMemberMethod, targetObject As CodeExpression, propertyName As String)
            o.WriteTo(type, method, New CodePropertyReferenceExpression(targetObject, propertyName))
        End Sub

        Public Shared Sub WriteElementString(writer As XmlWriter, elementName As String, prefix As String, value As Boolean)
            If value Then
                writer.WriteElementString(elementName, prefix, "True")
                Return
            End If
            writer.WriteElementString(elementName, prefix, "False")
        End Sub
    End Class
End Namespace
