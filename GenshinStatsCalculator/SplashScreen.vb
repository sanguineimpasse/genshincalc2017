Public Class SplashScreen
    Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles pbLoading.Click

    End Sub

    Private Sub LoadingTimer_Tick(sender As Object, e As EventArgs) Handles LoadingTimer.Tick
        lblLoading.Text = pbLoading.Value & "%"

        pbLoading.Value += 1

        If pbLoading.Value = 100 Then
            LoadingTimer.Stop()
            Disclaimer()
        End If
    End Sub
    Private Sub Disclaimer()
        Dim msg As String
        msg = "Copyright Disclaimer under section 107 of the Copyright Act 1976, allowance is made for 'fair use' for purposes such as criticism, comment, news reporting, teaching, scholarship, education and research. Fair use is a use permitted by copyright statute that might otherwise be infringing."
        MessageBox.Show(msg, "Disclaimer", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.Hide()
        Form1.Show()
    End Sub
End Class