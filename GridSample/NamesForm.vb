Imports ciloci.FormulaEngine

' Dialog to manage defined names in the formula engine.  Shows how you can use named references to setup aliases to formulas.
Friend Class NamesForm
	Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call

	End Sub

	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents cmdAdd As System.Windows.Forms.Button
	Friend WithEvents cmdDelete As System.Windows.Forms.Button
	Friend WithEvents edName As System.Windows.Forms.TextBox
	Friend WithEvents edFormula As System.Windows.Forms.TextBox
	Friend WithEvents lbNames As System.Windows.Forms.ListBox
	Friend WithEvents cmdOk As System.Windows.Forms.Button
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.edName = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.lbNames = New System.Windows.Forms.ListBox
		Me.edFormula = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.cmdAdd = New System.Windows.Forms.Button
		Me.cmdDelete = New System.Windows.Forms.Button
		Me.cmdOk = New System.Windows.Forms.Button
		Me.SuspendLayout()
		'
		'edName
		'
		Me.edName.Location = New System.Drawing.Point(8, 32)
		Me.edName.Name = "edName"
		Me.edName.Size = New System.Drawing.Size(272, 20)
		Me.edName.TabIndex = 0
		Me.edName.Text = ""
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(83, 16)
		Me.Label1.TabIndex = 1
		Me.Label1.Text = "Defined names:"
		'
		'lbNames
		'
		Me.lbNames.Location = New System.Drawing.Point(8, 56)
		Me.lbNames.Name = "lbNames"
		Me.lbNames.Size = New System.Drawing.Size(272, 108)
		Me.lbNames.TabIndex = 2
		'
		'edFormula
		'
		Me.edFormula.Location = New System.Drawing.Point(8, 200)
		Me.edFormula.Name = "edFormula"
		Me.edFormula.Size = New System.Drawing.Size(272, 20)
		Me.edFormula.TabIndex = 3
		Me.edFormula.Text = ""
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(8, 176)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(53, 16)
		Me.Label2.TabIndex = 4
		Me.Label2.Text = "Refers to:"
		'
		'cmdAdd
		'
		Me.cmdAdd.Location = New System.Drawing.Point(288, 64)
		Me.cmdAdd.Name = "cmdAdd"
		Me.cmdAdd.TabIndex = 7
		Me.cmdAdd.Text = "Add"
		'
		'cmdDelete
		'
		Me.cmdDelete.Location = New System.Drawing.Point(288, 96)
		Me.cmdDelete.Name = "cmdDelete"
		Me.cmdDelete.TabIndex = 8
		Me.cmdDelete.Text = "Delete"
		'
		'cmdOk
		'
		Me.cmdOk.Location = New System.Drawing.Point(288, 32)
		Me.cmdOk.Name = "cmdOk"
		Me.cmdOk.TabIndex = 9
		Me.cmdOk.Text = "Ok"
		'
		'NamesForm
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(368, 229)
		Me.Controls.Add(Me.cmdOk)
		Me.Controls.Add(Me.cmdDelete)
		Me.Controls.Add(Me.cmdAdd)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.edFormula)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.edName)
		Me.Controls.Add(Me.lbNames)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "NamesForm"
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Text = "Define Name"
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private MyEngine As FormulaEngine
	Private MyOwner As MainForm

	Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
		Me.LoadNames()
		Dim ms As MessageService = Me.Site.GetService(GetType(MessageService))
		ms.SetActiveDialog(Me)
	End Sub

	Protected Overrides Sub OnClosed(ByVal e As System.EventArgs)
		MyBase.OnClosed(e)
		Dim ms As MessageService = Me.Site.GetService(GetType(MessageService))
		ms.SetActiveDialog(MyOwner)
	End Sub

	Private Sub LoadNames()
		' Get the names of all named references with formulas in the engine
		Me.lbNames.Items.AddRange(MyEngine.GetNamedReferences())
		Me.cmdDelete.Enabled = Me.lbNames.Items.Count > 0
	End Sub

	' Undefine a name
	Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
		' Get the selected reference
		Dim ref As IReference = Me.SelectedReference
		' Get its formula
		Dim f As Formula = MyEngine.GetFormulaAt(ref)
		' Remove the formula
		MyEngine.RemoveFormula(f)
		' Recalculate any formulas that depended on the name
		MyEngine.Recalculate(ref)

		' Take care of UI stuff
		Me.lbNames.Items.Remove(ref)
		Me.edFormula.Clear()
		Me.edName.Clear()
	End Sub

	Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
		Me.DoAdd()
	End Sub

	' Define a name and associated formula
	Private Function DoAdd() As Boolean
		Dim name As String = Me.edName.Text
		' Try to create the named reference
		Dim ref As INamedReference = Me.CreateNamedReference(name)

		If ref Is Nothing Then
			Return False
		End If

		' Try to create the formula
		If Me.CreateFormula(Me.edFormula.Text, ref) = False Then
			Return False
		End If

		' Recalculate any formulas that depend on this name
		MyEngine.Recalculate(ref)

		' Do UI stuff
		Me.lbNames.Items.Add(ref)
		Me.lbNames.SelectedItem = ref
		Return True
	End Function

	' Try to create a formula
	Private Function CreateFormula(ByVal expression As String, ByVal ref As INamedReference) As Boolean
		' Get the main form to create a formula
		Dim f As Formula = MyOwner.CreateFormula(Me.edFormula.Text)

		If f Is Nothing Then
			Return False
		End If

		' Users will expect to be able to define named ranges.  For this to work, we have to change
		' the formula's result type to allow it to return references since, by default, it won't.
		f.ResultType = OperandType.Self

		Try
			' Try to add the formula
			MyEngine.AddFormula(f, ref)
		Catch ex As Exception
			Me.ShowMessage(ex.Message, MessageType.Error)
			Return False
		End Try

		Return True
	End Function

	Private Function CreateNamedReference(ByVal name As String) As INamedReference
		Try
			Dim ref As INamedReference = MyEngine.ReferenceFactory.Named(name)
			Return ref
		Catch ex As Exception
			Me.ShowMessage(ex.Message, MessageType.Error)
			Return Nothing
		End Try
	End Function

	' When the user presses this button, we want to do an Add if the user changed the name or expression.  If
	' nothing was changed, then we just close the form.
	Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
		If Me.edFormula.Text.Length = 0 Or Me.edName.Text.Length = 0 Then
			' User didn't change the form so just exit
			Me.Close()
			Return
		End If

		Dim ref As INamedReference = Me.CreateNamedReference(Me.edName.Text)

		If ref Is Nothing Then
			Return
		End If

		Dim f As Formula = MyEngine.GetFormulaAt(ref)

		If Not f Is Nothing Then
			' Formula already exists with this name
			If Me.edFormula.Text = f.ToString() Then
				' Formula hasn't changed 
				Me.Close()
				Return
			Else
				' Remove existing formula
				MyEngine.RemoveFormula(f)
			End If
		End If

		' Do an add, then close the form if it succeeded
		If Me.DoAdd() = True Then
			Me.Close()
		End If
	End Sub

	Private Sub lbNames_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbNames.SelectedIndexChanged
		Me.cmdDelete.Enabled = Me.lbNames.SelectedIndex <> -1

		Dim selected As IReference = Me.SelectedReference

		If Not selected Is Nothing Then
			Dim target As Formula = MyEngine.GetFormulaAt(selected)
			Me.edFormula.Text = target.ToString()
			Me.edName.Text = selected.ToString()
		End If
	End Sub

	Public Sub Initialize()
		MyEngine = Me.Site.GetService(GetType(FormulaEngine))
		MyOwner = Me.Site.GetService(GetType(MainForm))
	End Sub

	Private Sub ShowMessage(ByVal message As String, ByVal mt As MessageType)
		MyOwner.ShowMessage(message, mt)
	End Sub

	Private ReadOnly Property SelectedReference() As INamedReference
		Get
			Return Me.lbNames.SelectedItem
		End Get
	End Property
End Class