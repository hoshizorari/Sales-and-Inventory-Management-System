Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar

Public Class loginconfirm
    Dim attempts As Integer = 0
    Private Sub btn_login_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        Dim passcode As String = "ANIMEPC"
        If txt_passcode.Text = passcode Then
            Me.Hide()
            frm_login.Show()

        ElseIf txt_passcode.Text = Nothing Then
            MsgBox("Please Enter Your Username/Password !", vbExclamation)
            txt_passcode.Focus()
            attempts += 1
            If attempts = 3 Then
                MsgBox("Incorrect Password !", vbExclamation, "Password Incorrect !")
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MsgBox("Are You Sure You Want to Exit ?", vbInformation + vbYesNo) = vbYes Then
            Application.Exit()
        Else
            Return
        End If
    End Sub
End Class