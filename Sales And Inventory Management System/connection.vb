Imports System.Data.SqlClient

Module connection
    Dim connection As String
    Dim connstring As String
    Dim myconnection As SqlConnection = New SqlConnection
    Public cmd As New SqlCommand
    Public dr As SqlDataReader
    Public i As Integer
    Dim dtable As New DataTable
    Dim da As New SqlDataAdapter

    Public constring As New SqlConnection(My.Settings.connection)

    Public Function Getbillno() As String
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring
        Try
            myconnection.Open()
            Dim query As String
            query = "SELECT * FROM tbl_stock ORDER by id desc"
            cmd = New SqlCommand(query, myconnection)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                Getbillno = CLng(dr.Item("billno").ToString) + 1
            Else
                Getbillno = Date.Now.ToString("yyyy") & "1"

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try

    End Function

    Public Sub reload()
        Try
            myconnection.Open()

            Dim dtable As New DataTable
            Dim query As String
            query = "SELECT * FROM tbl_stock"
            Using cmd As New SqlCommand(query, myconnection)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dtable)
                End Using
            End Using

            frm_mainAdmin.dgv_reportList.DataSource = dtable
        Catch ex As Exception
            MessageBox.Show("An error occurred while loading data: " & ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

End Module
