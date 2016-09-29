<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class YesNoDialog
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.Question = New System.Windows.Forms.Label()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.No = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(449, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'Question
        '
        Me.Question.AutoSize = True
        Me.Question.Location = New System.Drawing.Point(12, 36)
        Me.Question.Name = "Question"
        Me.Question.Size = New System.Drawing.Size(39, 13)
        Me.Question.TabIndex = 4
        Me.Question.Text = "Label1"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(15, 71)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(45, 30)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "Yes"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'No
        '
        Me.No.Location = New System.Drawing.Point(66, 71)
        Me.No.Name = "No"
        Me.No.Size = New System.Drawing.Size(45, 30)
        Me.No.TabIndex = 6
        Me.No.Text = "No"
        Me.No.UseVisualStyleBackColor = True
        '
        'YesNoDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(449, 113)
        Me.Controls.Add(Me.No)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Question)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "YesNoDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "YesNoDialog"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents Question As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents No As System.Windows.Forms.Button
End Class
