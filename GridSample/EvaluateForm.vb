Imports ciloci.FormulaEngine

' Simple form to allow user to evaluate any expression
Public Class EvaluateForm

	Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
		MyBase.OnLoad(e)
		Me.SetResult(Nothing)
	End Sub

	Private Sub SetResult(ByVal value As Object)
		Dim text As String

		If value Is Nothing Then
			text = "(null)"
		Else
			text = value.ToString()
		End If

		Me.lblResult.Text = text
	End Sub

	Private Sub cmdEvaluate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEvaluate.Click
		' Get the expression
		Dim expression As String = Me.edExpression.Text
		' Try to create a formula from it
		Dim f As Formula = Me.MainForm.CreateFormula(expression)

		Dim result As Object = Nothing

		If Not f Is Nothing Then
			' If it's a valid formula then evaluate it
			result = f.Evaluate()
		End If

		' Show the result to the user
		Me.SetResult(result)
	End Sub

	Private ReadOnly Property MainForm() As MainForm
		Get
			Return Me.Site.GetService(GetType(MainForm))
		End Get
	End Property
End Class