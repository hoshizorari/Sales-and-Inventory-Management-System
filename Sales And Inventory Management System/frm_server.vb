Public Class frm_server

    Private Sub frm_server_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txt_password.UseSystemPasswordChar = True
        txt_server.Text = My.Settings.host
        txt_database.Text = My.Settings.db
        txt_username.Text = My.Settings.user
        txt_password.Text = My.Settings.pass
    End Sub

    Private Sub btn_config_Click(sender As Object, e As EventArgs) Handles btn_config.Click
        If txt_server.Text = "" Then
            MessageBox.Show("Enter Data Source", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf txt_database.Text = "" Then
            MessageBox.Show("Enter Database", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf txt_username.Text = "" Then
            MessageBox.Show("Enter username", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf txt_password.Text = "" Then
            MessageBox.Show("Enter Password", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            My.Settings.connection = "Data Source=" & txt_server.Text.Trim() & ";Initial Catalog=" & txt_database.Text.Trim() & ";User ID=" & txt_username.Text.Trim() & ";Password=" & txt_password.Text.Trim()
            My.Settings.host = txt_server.Text
            My.Settings.db = txt_database.Text
            My.Settings.user = txt_username.Text
            My.Settings.pass = txt_password.Text
            My.Settings.Save()
            Me.Hide()
        End If
    End Sub

    Private Sub btn_close_Click(sender As Object, e As EventArgs) Handles btn_close.Click
        Me.Hide()
        frm_login.Show()
    End Sub
End Class