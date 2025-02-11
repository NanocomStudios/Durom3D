Public Class RenderDialog

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form1.renderVideo(TextBox1.Text, ComboBox1.SelectedIndex, ComboBox2.SelectedIndex)
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SaveFileDialog1.ShowDialog()
        TextBox1.Text = SaveFileDialog1.FileName
    End Sub
End Class