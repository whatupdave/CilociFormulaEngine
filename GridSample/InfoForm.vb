Imports ciloci.FormulaEngine

' Shows various statistics about the formula engine
Public Class InfoForm
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
	Friend WithEvents lblFormulaCount As System.Windows.Forms.Label
	Friend WithEvents lblReferenceCount As System.Windows.Forms.Label
	Friend WithEvents lbFunctions As System.Windows.Forms.ListBox
	Friend WithEvents Button1 As System.Windows.Forms.Button
	Friend WithEvents edDependencyGraph As System.Windows.Forms.TextBox

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim Label2 As System.Windows.Forms.Label
		Dim Label3 As System.Windows.Forms.Label
		Dim Label4 As System.Windows.Forms.Label
		Dim Label5 As System.Windows.Forms.Label
		Me.lblFormulaCount = New System.Windows.Forms.Label
		Me.lblReferenceCount = New System.Windows.Forms.Label
		Me.lbFunctions = New System.Windows.Forms.ListBox
		Me.edDependencyGraph = New System.Windows.Forms.TextBox
		Me.Button1 = New System.Windows.Forms.Button
		Label2 = New System.Windows.Forms.Label
		Label3 = New System.Windows.Forms.Label
		Label4 = New System.Windows.Forms.Label
		Label5 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		'
		'Label2
		'
		Label2.AutoSize = True
		Label2.Location = New System.Drawing.Point(0, 9)
		Label2.Name = "Label2"
		Label2.Size = New System.Drawing.Size(78, 13)
		Label2.TabIndex = 1
		Label2.Text = "Formula Count:"
		'
		'Label3
		'
		Label3.AutoSize = True
		Label3.Location = New System.Drawing.Point(147, 9)
		Label3.Name = "Label3"
		Label3.Size = New System.Drawing.Size(91, 13)
		Label3.TabIndex = 2
		Label3.Text = "Reference Count:"
		'
		'Label4
		'
		Label4.AutoSize = True
		Label4.Location = New System.Drawing.Point(0, 41)
		Label4.Name = "Label4"
		Label4.Size = New System.Drawing.Size(103, 13)
		Label4.TabIndex = 3
		Label4.Text = "Dependency Graph:"
		'
		'Label5
		'
		Label5.AutoSize = True
		Label5.Location = New System.Drawing.Point(0, 169)
		Label5.Name = "Label5"
		Label5.Size = New System.Drawing.Size(56, 13)
		Label5.TabIndex = 4
		Label5.Text = "Functions:"
		'
		'lblFormulaCount
		'
		Me.lblFormulaCount.AutoSize = True
		Me.lblFormulaCount.Location = New System.Drawing.Point(84, 9)
		Me.lblFormulaCount.Name = "lblFormulaCount"
		Me.lblFormulaCount.Size = New System.Drawing.Size(39, 13)
		Me.lblFormulaCount.TabIndex = 5
		Me.lblFormulaCount.Text = "Label6"
		'
		'lblReferenceCount
		'
		Me.lblReferenceCount.AutoSize = True
		Me.lblReferenceCount.Location = New System.Drawing.Point(244, 9)
		Me.lblReferenceCount.Name = "lblReferenceCount"
		Me.lblReferenceCount.Size = New System.Drawing.Size(39, 13)
		Me.lblReferenceCount.TabIndex = 6
		Me.lblReferenceCount.Text = "Label6"
		'
		'lbFunctions
		'
		Me.lbFunctions.ColumnWidth = 70
		Me.lbFunctions.FormattingEnabled = True
		Me.lbFunctions.Location = New System.Drawing.Point(61, 169)
		Me.lbFunctions.MultiColumn = True
		Me.lbFunctions.Name = "lbFunctions"
		Me.lbFunctions.Size = New System.Drawing.Size(210, 134)
		Me.lbFunctions.TabIndex = 7
		'
		'edDependencyGraph
		'
		Me.edDependencyGraph.Location = New System.Drawing.Point(108, 41)
		Me.edDependencyGraph.Multiline = True
		Me.edDependencyGraph.Name = "edDependencyGraph"
		Me.edDependencyGraph.ReadOnly = True
		Me.edDependencyGraph.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.edDependencyGraph.Size = New System.Drawing.Size(206, 120)
		Me.edDependencyGraph.TabIndex = 8
		Me.edDependencyGraph.WordWrap = False
		'
		'Button1
		'
		Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.Button1.Location = New System.Drawing.Point(123, 314)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(75, 23)
		Me.Button1.TabIndex = 9
		Me.Button1.Text = "Ok"
		Me.Button1.UseVisualStyleBackColor = True
		'
		'InfoForm
		'
		Me.AcceptButton = Me.Button1
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(321, 349)
		Me.Controls.Add(Me.Button1)
		Me.Controls.Add(Me.edDependencyGraph)
		Me.Controls.Add(Me.lbFunctions)
		Me.Controls.Add(Me.lblReferenceCount)
		Me.Controls.Add(Me.lblFormulaCount)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.MinimumSize = New System.Drawing.Size(300, 300)
		Me.Name = "InfoForm"
		Me.ShowInTaskbar = False
		Me.Text = "Formula engine information"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

#End Region

	Private MyEngine As FormulaEngine

	Public Sub SetEngine(ByVal engine As FormulaEngine)
		MyEngine = engine
	End Sub

	Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
		Me.lblFormulaCount.Text = MyEngine.FormulaCount
		Me.lblReferenceCount.Text = MyEngine.Info.ReferenceCount

		Dim functionNames As String() = MyEngine.FunctionLibrary.GetFunctionNames()
		System.Array.Sort(Of String)(functionNames)
		Me.lbFunctions.Items.AddRange(functionNames)
		Me.edDependencyGraph.Text = MyEngine.Info.DependencyDump
	End Sub
End Class