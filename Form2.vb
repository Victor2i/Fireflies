Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.StartNewGame(NumericUpDown1.Value, ComboBox1.SelectedIndex)
    End Sub
End Class