Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim button As New CustomButton
        button.Parent = Me
        button.Name = "txt"
        button.Text = "Hola manol"
        button.Width = GroupBox1.Width - 6 * 2
        button.Height = (GroupBox1.Height - 6 * 2) \ 4
        button.Top = 6
        button.Left = +6

        GroupBox1.Controls.Add(button)
    End Sub
End Class