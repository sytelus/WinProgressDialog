'
'The DEMO for the .Net wrapper for the IProgressDialog which shows standard Windows progress dialogs
'Author: Shital Shah
'Date: Dec, 2003
'TLB source: http://www.msjogren.net/dotnet/eng/samples/vb6_progdlg.asp
'

Public Class Form1
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(40, 24)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(144, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Start Processing"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(240, 141)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Win Progress Dialog Demo"
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Just an array of strings
    Private m_Activities As String() = New String() {"first", "few entries", "don't show up", "because of", "default 2 sec delay", "Learning General Relativity...", "Watching movies...", "Learning Game Theory...", "Writing Code Project articles...", "Meeting people...", "Exploring Juneue...", "Reading news...", "Theorizing instruction-memory equivalece..."}


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim progressDialog As New WinProgressDialog.ProgressDialog
        Try
            progressDialog.ProgressBarVisible = False
            'Specify parameters for this progress dialog
            progressDialog.Show(Me.Handle.ToInt32, "Christmas Time Status", "Spending Christmas Time", m_Activities.Length)

            '****Dialog will appear after 2-3 sec of delay****

            For dayIndex As Integer = 0 To m_Activities.Length - 1
                'Checking if user pressed cancel is optional
                If progressDialog.UpdateProgress(dayIndex, m_Activities(dayIndex)) Then
                    MsgBox("You cancelled!!")
                    Exit For
                End If

                'Just make this For loop slow
                Threading.Thread.CurrentThread.Sleep(2000)
            Next
        Finally
            'Dialog must be disposed off
            progressDialog.Dispose()
        End Try
    End Sub
End Class
