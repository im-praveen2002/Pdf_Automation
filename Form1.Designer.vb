<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(517, 726)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(161, 26)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "EXCEL UPDATE"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(447, 81)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(293, 20)
        Me.TextBox2.TabIndex = 5
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(447, 16)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(293, 20)
        Me.TextBox1.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(255, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 19)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Customer Input"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(255, 82)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(116, 19)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Current Directory"
        '
        'TreeView1
        '
        Me.TreeView1.HideSelection = False
        Me.TreeView1.Location = New System.Drawing.Point(109, 149)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(988, 544)
        Me.TreeView1.TabIndex = 9
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(1124, 380)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 10
        Me.Button4.Text = "CLEAR"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(784, 14)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 20)
        Me.Button5.TabIndex = 11
        Me.Button5.Text = "OK"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1315, 863)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.TreeView1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Button2)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button2 As Button
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents TreeView1 As TreeView
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
End Class
