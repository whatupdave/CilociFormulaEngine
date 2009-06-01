<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EvaluateForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Dim Label1 As System.Windows.Forms.Label
		Dim Label2 As System.Windows.Forms.Label
		Me.edExpression = New System.Windows.Forms.TextBox
		Me.cmdEvaluate = New System.Windows.Forms.Button
		Me.lblResult = New System.Windows.Forms.Label
		Me.Panel1 = New System.Windows.Forms.Panel
		Me.Panel2 = New System.Windows.Forms.Panel
		Me.Splitter1 = New System.Windows.Forms.Splitter
		Me.Panel3 = New System.Windows.Forms.Panel
		Label1 = New System.Windows.Forms.Label
		Label2 = New System.Windows.Forms.Label
		Me.Panel1.SuspendLayout()
		Me.Panel2.SuspendLayout()
		Me.Panel3.SuspendLayout()
		Me.SuspendLayout()
		'
		'Label1
		'
		Label1.Dock = System.Windows.Forms.DockStyle.Top
		Label1.Location = New System.Drawing.Point(0, 0)
		Label1.Name = "Label1"
		Label1.Size = New System.Drawing.Size(400, 26)
		Label1.TabIndex = 0
		Label1.Text = "Type any expression into the textbox to have the formula engine evaluate it."
		'
		'Label2
		'
		Label2.AutoSize = True
		Label2.Dock = System.Windows.Forms.DockStyle.Left
		Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Label2.Location = New System.Drawing.Point(0, 0)
		Label2.Name = "Label2"
		Label2.Size = New System.Drawing.Size(58, 16)
		Label2.TabIndex = 3
		Label2.Text = "Result:"
		'
		'edExpression
		'
		Me.edExpression.Dock = System.Windows.Forms.DockStyle.Fill
		Me.edExpression.Location = New System.Drawing.Point(0, 0)
		Me.edExpression.Multiline = True
		Me.edExpression.Name = "edExpression"
		Me.edExpression.Size = New System.Drawing.Size(318, 31)
		Me.edExpression.TabIndex = 1
		'
		'cmdEvaluate
		'
		Me.cmdEvaluate.Location = New System.Drawing.Point(3, 3)
		Me.cmdEvaluate.Name = "cmdEvaluate"
		Me.cmdEvaluate.Size = New System.Drawing.Size(75, 23)
		Me.cmdEvaluate.TabIndex = 2
		Me.cmdEvaluate.Text = "Evaluate"
		Me.cmdEvaluate.UseVisualStyleBackColor = True
		'
		'lblResult
		'
		Me.lblResult.Dock = System.Windows.Forms.DockStyle.Fill
		Me.lblResult.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblResult.Location = New System.Drawing.Point(58, 0)
		Me.lblResult.Name = "lblResult"
		Me.lblResult.Size = New System.Drawing.Size(342, 25)
		Me.lblResult.TabIndex = 4
		Me.lblResult.Text = "Label3"
		'
		'Panel1
		'
		Me.Panel1.Controls.Add(Me.edExpression)
		Me.Panel1.Controls.Add(Me.Panel3)
		Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
		Me.Panel1.Location = New System.Drawing.Point(0, 26)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(400, 31)
		Me.Panel1.TabIndex = 5
		'
		'Panel2
		'
		Me.Panel2.Controls.Add(Me.lblResult)
		Me.Panel2.Controls.Add(Label2)
		Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.Panel2.Location = New System.Drawing.Point(0, 61)
		Me.Panel2.Name = "Panel2"
		Me.Panel2.Size = New System.Drawing.Size(400, 25)
		Me.Panel2.TabIndex = 6
		'
		'Splitter1
		'
		Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.Splitter1.Location = New System.Drawing.Point(0, 57)
		Me.Splitter1.Name = "Splitter1"
		Me.Splitter1.Size = New System.Drawing.Size(400, 4)
		Me.Splitter1.TabIndex = 7
		Me.Splitter1.TabStop = False
		'
		'Panel3
		'
		Me.Panel3.Controls.Add(Me.cmdEvaluate)
		Me.Panel3.Dock = System.Windows.Forms.DockStyle.Right
		Me.Panel3.Location = New System.Drawing.Point(318, 0)
		Me.Panel3.Name = "Panel3"
		Me.Panel3.Size = New System.Drawing.Size(82, 31)
		Me.Panel3.TabIndex = 3
		'
		'EvaluateForm
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(400, 86)
		Me.Controls.Add(Me.Panel1)
		Me.Controls.Add(Me.Splitter1)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Me.Panel2)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
		Me.MinimumSize = New System.Drawing.Size(408, 110)
		Me.Name = "EvaluateForm"
		Me.ShowInTaskbar = False
		Me.Text = "Evaluate an expression"
		Me.Panel1.ResumeLayout(False)
		Me.Panel1.PerformLayout()
		Me.Panel2.ResumeLayout(False)
		Me.Panel2.PerformLayout()
		Me.Panel3.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents edExpression As System.Windows.Forms.TextBox
	Friend WithEvents cmdEvaluate As System.Windows.Forms.Button
	Friend WithEvents lblResult As System.Windows.Forms.Label
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents Panel2 As System.Windows.Forms.Panel
	Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
	Friend WithEvents Panel3 As System.Windows.Forms.Panel
End Class
