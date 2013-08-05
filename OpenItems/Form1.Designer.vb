<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.lblStatus = New System.Windows.Forms.ToolStripStatusLabel
        Me.lblBGW = New System.Windows.Forms.ToolStripStatusLabel
        Me.pgrWorking = New System.Windows.Forms.ProgressBar
        Me.lstStatus = New System.Windows.Forms.ListBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.BGW = New System.ComponentModel.BackgroundWorker
        Me.BS = New System.Windows.Forms.BindingSource(Me.components)
        Me.Restart = New System.Windows.Forms.Timer(Me.components)
        Me.StatusStrip1.SuspendLayout()
        CType(Me.BS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblStatus, Me.lblBGW})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 256)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(449, 22)
        Me.StatusStrip1.TabIndex = 0
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'lblStatus
        '
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(313, 17)
        Me.lblStatus.Spring = True
        Me.lblStatus.Text = "ToolStripStatusLabel1"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblBGW
        '
        Me.lblBGW.Name = "lblBGW"
        Me.lblBGW.Size = New System.Drawing.Size(121, 17)
        Me.lblBGW.Text = "ToolStripStatusLabel1"
        '
        'pgrWorking
        '
        Me.pgrWorking.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pgrWorking.Location = New System.Drawing.Point(0, 246)
        Me.pgrWorking.Name = "pgrWorking"
        Me.pgrWorking.Size = New System.Drawing.Size(449, 10)
        Me.pgrWorking.TabIndex = 1
        '
        'lstStatus
        '
        Me.lstStatus.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstStatus.FormattingEnabled = True
        Me.lstStatus.Location = New System.Drawing.Point(10, 26)
        Me.lstStatus.Name = "lstStatus"
        Me.lstStatus.Size = New System.Drawing.Size(427, 212)
        Me.lstStatus.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(0, 0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(449, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'BGW
        '
        '
        'Restart
        '
        Me.Restart.Enabled = True
        Me.Restart.Interval = 1000
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(449, 278)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lstStatus)
        Me.Controls.Add(Me.pgrWorking)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Name = "Form1"
        Me.Text = "Downloading Open Items"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.BS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents lblStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents pgrWorking As System.Windows.Forms.ProgressBar
    Friend WithEvents lstStatus As System.Windows.Forms.ListBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lblBGW As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents BGW As System.ComponentModel.BackgroundWorker
    Friend WithEvents BS As System.Windows.Forms.BindingSource
    Friend WithEvents Restart As System.Windows.Forms.Timer

End Class
