' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public License
' as published by the Free Software Foundation; either version 2.1
' of the License, or (at your option) any later version.
' 
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free
' Software Foundation, Inc., 59 Temple Place, Suite 330, Boston,
' MA 02111-1307, USA.
' 
' FormulaEngine - A library for parsing and managing formulas
' Copyright © 2007 Eugene Ciloci
'

' Classes for converting a parse tree into a postfix expression

Friend MustInherit Class ParseTreeElement

	Protected Sub New()

	End Sub

	Public MustOverride Sub AddAsRPN(ByVal dest As IList)
End Class

Friend Class UnaryOperatorElement
	Inherits ParseTreeElement

	Private MyOperator As UnaryOperator
	Private MyArgument As ParseTreeElement

	Public Sub New(ByVal [operator] As UnaryOperator, ByVal argument As ParseTreeElement)
		MyOperator = [operator]
		MyArgument = argument
	End Sub

	Public Overrides Sub AddAsRPN(ByVal dest As System.Collections.IList)
		MyArgument.AddAsRPN(dest)
		dest.Add(MyOperator)
	End Sub
End Class

Friend Class BinaryOperatorElement
	Inherits ParseTreeElement

	Private MyArguments As ParseTreeElement()
	Private MyOperations As BinaryOperator()

	Public Sub New()

	End Sub

	Public Sub AcceptValues(ByVal values As System.Collections.IList)
		Debug.Assert(values.Count >= 2, "must have at least 2 values")

		MyArguments = New ParseTreeElement(((values.Count + 1) / 2) - 1) {}
		MyOperations = New BinaryOperator(((values.Count - 1) / 2) - 1) {}

		Dim index As Integer = 0

		For i As Integer = 0 To values.Count - 1 Step 2
			MyArguments(index) = values.Item(i)
			index += 1
		Next

		index = 0

		For i As Integer = 1 To values.Count - 1 Step 2
			MyOperations(index) = values.Item(i)
			index += 1
		Next
	End Sub

	Public Overrides Sub AddAsRPN(ByVal dest As IList)
		Dim element As ParseTreeElement

		element = MyArguments(0)
		element.AddAsRPN(dest)
		element = MyArguments(1)
		element.AddAsRPN(dest)
		dest.Add(MyOperations(0))

		Dim opIndex As Integer = 1

		For i As Integer = 2 To MyArguments.Length - 1
			element = MyArguments(i)
			element.AddAsRPN(dest)
			dest.Add(MyOperations(opIndex))
			opIndex += 1
		Next
	End Sub
End Class

Friend Class ContainerElement
	Inherits ParseTreeElement

	Private MyValue As Object

	Public Sub New(ByVal value As Object)
		MyValue = value
	End Sub

	Public Overrides Sub AddAsRPN(ByVal dest As IList)
		dest.Add(MyValue)
	End Sub
End Class

Friend Class FunctionCallElement
	Inherits ParseTreeElement

	Private MyFunctionName As String
	Private MyArguments As ParseTreeElement()

	Public Sub New()

	End Sub

	Public Sub AcceptValues(ByVal values As System.Collections.IList)
		Dim functionName As String = DirectCast(values.Item(0), String)

		values.RemoveAt(0)

		Dim arr(values.Count - 1) As ParseTreeElement
		values.CopyTo(arr, 0)
		System.Array.Reverse(arr)

		MyArguments = arr
		MyFunctionName = functionName
	End Sub

	Public Overrides Sub AddAsRPN(ByVal dest As System.Collections.IList)
		For Each element As ParseTreeElement In MyArguments
			element.AddAsRPN(dest)
		Next

		Dim funcCall As New FunctionCallOperator(MyFunctionName, MyArguments.Length)
		dest.Add(funcCall)
	End Sub
End Class