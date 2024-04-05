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
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblCommand = New System.Windows.Forms.Label()
        Me.txtRegion = New System.Windows.Forms.TextBox()
        Me.txtEdition = New System.Windows.Forms.TextBox()
        Me.txtIssue = New System.Windows.Forms.TextBox()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBoxRegion = New System.Windows.Forms.ComboBox()
        Me.ComboBoxEdition = New System.Windows.Forms.ComboBox()
        Me.ComboBoxIssue = New System.Windows.Forms.ComboBox()
        Me.ComboBoxTitle = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(13, 180)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(225, 24)
        Me.Button10.TabIndex = 6
        Me.Button10.Text = "Create Single Page PDFs"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button7.Location = New System.Drawing.Point(13, 210)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(225, 24)
        Me.Button7.TabIndex = 7
        Me.Button7.Text = "Create PDF Book"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblCommand)
        Me.GroupBox1.Controls.Add(Me.txtRegion)
        Me.GroupBox1.Controls.Add(Me.txtEdition)
        Me.GroupBox1.Controls.Add(Me.txtIssue)
        Me.GroupBox1.Controls.Add(Me.txtTitle)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(277, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(483, 139)
        Me.GroupBox1.TabIndex = 33
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Command Line.."
        '
        'lblCommand
        '
        Me.lblCommand.Location = New System.Drawing.Point(169, 26)
        Me.lblCommand.Name = "lblCommand"
        Me.lblCommand.Size = New System.Drawing.Size(298, 99)
        Me.lblCommand.TabIndex = 37
        '
        'txtRegion
        '
        Me.txtRegion.Enabled = False
        Me.txtRegion.Location = New System.Drawing.Point(36, 103)
        Me.txtRegion.Name = "txtRegion"
        Me.txtRegion.Size = New System.Drawing.Size(99, 20)
        Me.txtRegion.TabIndex = 36
        '
        'txtEdition
        '
        Me.txtEdition.Enabled = False
        Me.txtEdition.Location = New System.Drawing.Point(36, 76)
        Me.txtEdition.Name = "txtEdition"
        Me.txtEdition.Size = New System.Drawing.Size(99, 20)
        Me.txtEdition.TabIndex = 35
        '
        'txtIssue
        '
        Me.txtIssue.Enabled = False
        Me.txtIssue.Location = New System.Drawing.Point(36, 49)
        Me.txtIssue.Name = "txtIssue"
        Me.txtIssue.Size = New System.Drawing.Size(99, 20)
        Me.txtIssue.TabIndex = 34
        '
        'txtTitle
        '
        Me.txtTitle.Enabled = False
        Me.txtTitle.Location = New System.Drawing.Point(36, 22)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(99, 20)
        Me.txtTitle.TabIndex = 33
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.ComboBoxRegion)
        Me.GroupBox2.Controls.Add(Me.ComboBoxEdition)
        Me.GroupBox2.Controls.Add(Me.ComboBoxIssue)
        Me.GroupBox2.Controls.Add(Me.ComboBoxTitle)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(247, 136)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Tag = ""
        Me.GroupBox2.Text = "Parameters..."
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(24, 111)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 13)
        Me.Label4.TabIndex = 36
        Me.Label4.Text = "Region:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(25, 84)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 13)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "Edition:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(34, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 13)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Issue:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(39, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Title:"
        '
        'ComboBoxRegion
        '
        Me.ComboBoxRegion.FormattingEnabled = True
        Me.ComboBoxRegion.Location = New System.Drawing.Point(90, 103)
        Me.ComboBoxRegion.Name = "ComboBoxRegion"
        Me.ComboBoxRegion.Size = New System.Drawing.Size(136, 21)
        Me.ComboBoxRegion.TabIndex = 5
        '
        'ComboBoxEdition
        '
        Me.ComboBoxEdition.FormattingEnabled = True
        Me.ComboBoxEdition.Location = New System.Drawing.Point(90, 76)
        Me.ComboBoxEdition.Name = "ComboBoxEdition"
        Me.ComboBoxEdition.Size = New System.Drawing.Size(136, 21)
        Me.ComboBoxEdition.TabIndex = 4
        '
        'ComboBoxIssue
        '
        Me.ComboBoxIssue.FormattingEnabled = True
        Me.ComboBoxIssue.Location = New System.Drawing.Point(90, 49)
        Me.ComboBoxIssue.Name = "ComboBoxIssue"
        Me.ComboBoxIssue.Size = New System.Drawing.Size(136, 21)
        Me.ComboBoxIssue.TabIndex = 3
        '
        'ComboBoxTitle
        '
        Me.ComboBoxTitle.FormattingEnabled = True
        Me.ComboBoxTitle.Location = New System.Drawing.Point(90, 22)
        Me.ComboBoxTitle.Name = "ComboBoxTitle"
        Me.ComboBoxTitle.Size = New System.Drawing.Size(136, 21)
        Me.ComboBoxTitle.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(163, 240)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AcceptButton = Me.Button10
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(270, 279)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.Button7)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.Text = "PrepPDF"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button10 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lblCommand As Label
    Friend WithEvents txtRegion As TextBox
    Friend WithEvents txtEdition As TextBox
    Friend WithEvents txtIssue As TextBox
    Friend WithEvents txtTitle As TextBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents ComboBoxRegion As ComboBox
    Friend WithEvents ComboBoxEdition As ComboBox
    Friend WithEvents ComboBoxIssue As ComboBox
    Friend WithEvents ComboBoxTitle As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
End Class
