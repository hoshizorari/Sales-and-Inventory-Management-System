Imports System.Data.SqlClient
Public Class frm_login
    Dim connection As String
    Dim connstring As String
    Dim myconnection As SqlConnection = New SqlConnection
    Dim names As String
    Dim username As String
    Dim password As String
    Dim role As String
    Dim attempts As Integer = 0


    '===================================================================================
    '===================================================================================
    'LOGIN

    Private Sub btn_login_Click(sender As Object, e As EventArgs) Handles btn_login.Click


        If txt_username.Text = Nothing And txt_password.Text = Nothing Then
            MsgBox("Please Enter Your Username/Password !", vbExclamation)
            txt_username.Focus()
            attempts += 1
            Count.Text = attempts
            If attempts = 3 Then
                MsgBox("Maximum Login Attempts Reached. Application will now Close !", vbExclamation)
                txt_password.Enabled = False
                txt_username.Enabled = False
                Dim adminUsername As String = "3435"
                Dim password As String = InputBox("Enter Your Password to Confirm Deletion:", "Password Required !")
                If password = adminUsername Then
                    txt_username.Enabled = True
                    txt_password.Enabled = True
                    attempts = 0
                    Count.Text = attempts
                ElseIf password <> "" Then
                    MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                    Application.Exit()
                ElseIf password = Nothing Then
                    MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                    Application.Exit()
                End If
            End If
        ElseIf txt_username.Text = Nothing Then
            MsgBox("Please Enter Your Username !", vbExclamation)
            txt_username.Focus()
            attempts += 1
            Count.Text = attempts
            If attempts = 3 Then
                MsgBox("Maximum Login Attempts Reached. Application will now Close !", vbExclamation)
                txt_password.Enabled = False
                txt_username.Enabled = False
                Dim adminUsername As String = "3435"
                Dim password As String = InputBox("Enter Your Password to Confirm Deletion:", "Password Required !")
                If password = adminUsername Then
                    txt_username.Enabled = True
                    txt_password.Enabled = True
                    attempts = 0
                    Count.Text = attempts
                ElseIf password <> "" Then
                    MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                    Application.Exit()
                ElseIf password = Nothing Then
                    MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                    Application.Exit()
                End If
            End If
        ElseIf txt_password.Text = Nothing Then
            MsgBox("Please Enter Your Password !", vbExclamation)
            txt_password.Focus()
            attempts += 1
            Count.Text = attempts
            If attempts = 3 Then
                MsgBox("Maximum Login Attempts Reached. Application will now Close !", vbExclamation)
                txt_password.Enabled = False
                txt_username.Enabled = False
                Dim adminUsername As String = "3435"

                Dim password As String = InputBox("Enter Your Password to Confirm Deletion:", "Password Required !")
                If password = adminUsername Then
                    txt_username.Enabled = True
                    txt_password.Enabled = True
                    attempts = 0
                    Count.Text = attempts
                ElseIf password <> "" Then
                    MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                    Application.Exit()
                ElseIf password = Nothing Then
                    MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                    Application.Exit()
                End If
            End If
        Else
            Try
                connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
                connstring = connection
                myconnection.ConnectionString = connstring
                myconnection.Open()
                'to select from the table which user will log in
                Dim query As String
                query = "Select * from tbl_user Where [Username] = '" & txt_username.Text & "' AND [Password]='" & txt_password.Text & "'"
                Dim command As SqlCommand = New SqlCommand(query, myconnection)
                Dim dr As SqlDataReader
                Dim blnfound As Boolean
                blnfound = False
                dr = command.ExecuteReader
                Dim found As New Boolean
                dr.Read()
                If dr.HasRows Then
                    found = True
                    names = dr.Item("Name").ToString
                    username = dr.Item("Username").ToString
                    password = dr.Item("Password").ToString
                    role = dr.Item("Role").ToString
                Else
                    found = False
                    names = ""
                    username = ""
                    password = ""
                    role = ""
                End If
                If found = True Then
                    If StrComp(txt_username.Text, username, CompareMethod.Binary) Or StrComp(txt_password.Text, password, CompareMethod.Binary) Then
                        MsgBox("Warning : Wrong Username or Password !", vbExclamation)
                        txt_username.Focus()
                        txt_username.Clear()
                        txt_password.Clear()
                        attempts += 1
                        Count.Text = attempts
                        If attempts = 3 Then
                            MsgBox("Maximum Login Attempts Reached. Application will now Close !", vbExclamation)
                            txt_password.Enabled = False
                            txt_username.Enabled = False
                            Dim adminUsername As String = "3435"
                            ' Requesting password from the user
                            Dim password As String = InputBox("Enter Your Password to Confirm Deletion:", "Password Required !")
                            If password = adminUsername Then
                                txt_username.Enabled = True
                                txt_password.Enabled = True
                                attempts = 0
                                Count.Text = attempts
                            ElseIf password <> "" Then
                                MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                                Application.Exit()
                            ElseIf password = Nothing Then
                                MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                                Application.Exit()
                            End If
                        End If
                        Return
                    Else
                        If UCase(role) = "ADMIN" Then
                            Me.Hide()
                            frm_mainAdmin.Show()
                            frm_mainAdmin.TabControl1.SelectedIndex = 0
                            viewproduct()
                            viewdataUSer()
                            viewSupplier()
                            viewdataReturn()
                            UpdateCounts()
                            viewdataReport()
                            viewdataReportInv()
                            viewdataReportRet()
                            PopulateSupplierComboBox()
                            PopulateReturnComboBox()
                            CheckBox1.CheckState = 0
                            myconnection.Close()
                            frm_mainAdmin.lbl_admin.Text = username
                            attempts = 0
                            Count.Text = attempts
                        ElseIf UCase(role) = "EMPLOYEE" Then
                            Me.Hide()
                            frm_mainCashier.Show()
                            frm_mainCashier.TabControl1.SelectedIndex = 0
                            viewSupplier1()
                            viewdataReturn1()
                            viewdataReport1()
                            viewdataReportInv1()
                            viewdataReportRet1()
                            PopulateReturnComboBoxemp()
                            CheckBox1.CheckState = 0
                            myconnection.Close()
                            frm_mainCashier.txt_searchPro.Focus()
                            frm_mainCashier.lbl_employee.Text = username
                            attempts = 0
                            Count.Text = attempts
                        End If
                    End If
                    ' Display logged-in username

                Else
                    MsgBox("Warning : Wrong Username or Password !", vbExclamation)
                    txt_username.Focus()
                    txt_username.Clear()
                    txt_password.Clear()
                    attempts += 1
                    Count.Text = attempts
                    If attempts = 3 Then
                        MsgBox("Maximum Login Attempts Reached. Application will now Close !", vbExclamation)
                        txt_password.Enabled = False
                        txt_username.Enabled = False
                        Dim adminUsername As String = "3435"
                        ' Requesting password from the user
                        Dim password As String = InputBox("Enter Your Password to Confirm Deletion:", "Password Required !")
                        If password = adminUsername Then
                            txt_username.Enabled = True
                            txt_password.Enabled = True
                            attempts = 0
                            Count.Text = attempts
                        ElseIf password <> "" Then
                            MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                            Application.Exit()
                        ElseIf password = Nothing Then
                            MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                            Application.Exit()
                        End If
                    End If
                End If
                txt_password.Clear()
                txt_username.Clear()
                myconnection.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End If
        'myconnection.Close()
    End Sub

    Private Sub btn_exit_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        If MsgBox("Are You Sure You Want to Exit ?", vbInformation + vbYesNo) = vbYes Then
            Application.Exit()
        Else
            Return
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            txt_password.UseSystemPasswordChar = False
        Else
            txt_password.UseSystemPasswordChar = True
        End If
    End Sub

    '===================================================================================
    '===================================================================================
    'VIEW DATA ADMIN


    Sub viewdataUSer()
        Try
            Dim con1 As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "SELECT Name, Username, Password, Role FROM tbl_user"
            Dim adapter As New SqlDataAdapter(sql, con1)
            Dim data As New DataTable("tbl_user")
            adapter.Fill(data)
            frm_mainAdmin.dgv_userList.DataSource = data
            Dim sql1 As String
            sql1 = "SELECT * FROM tbl_user"
            Dim adapter1 As New SqlDataAdapter(sql1, con1)
            Dim cmd As New SqlCommand(sql1, con1)
            con1.Open()
            Dim myreader As SqlDataReader = cmd.ExecuteReader
            myreader.Read()
            con1.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try

    End Sub

    Sub viewproduct()
        Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
        Dim sql As String
        sql = "SELECT supname, prodate, proname, progroup, stockno, price FROM tbl_product"
        Dim adapter As New SqlDataAdapter(sql, myconnection)
        Dim data As New DataTable("tbl_product")
        adapter.Fill(data)
        frm_mainAdmin.dgv_productList.DataSource = data
        Dim sql1 As String
        sql1 = "SELECT * FROM tbl_product"
        Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
        Dim cmd As New SqlCommand(sql1, myconnection)
        myconnection.Open()
        Dim myreader As SqlDataReader = cmd.ExecuteReader
        myreader.Read()
        myconnection.Close()
    End Sub

    Sub viewSupplier()
        Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
        Dim sql As String
        sql = "SELECT supname, supproname, address, contact, email FROM tbl_supplier"
        Dim adapter As New SqlDataAdapter(sql, myconnection)
        Dim data As New DataTable("tbl_supplier")
        adapter.Fill(data)
        frm_mainAdmin.dgv_supplierList.DataSource = data
        Dim sql1 As String
        sql1 = "SELECT * FROM tbl_supplier"
        Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
        Dim cmd As New SqlCommand(sql1, myconnection)
        myconnection.Open()
        Dim myreader As SqlDataReader = cmd.ExecuteReader
        myreader.Read()
        myconnection.Close()
    End Sub

    Sub viewdataReturn()
        Try
            Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "SELECT RDate, Cname, Contact, Pname, Qty, Issue, Acttaken FROM tbl_return"
            Dim adapter As New SqlDataAdapter(sql, myconnection)
            Dim data As New DataTable("tbl_return")
            adapter.Fill(data)
            frm_mainAdmin.dgv_returnList.DataSource = data
            Dim sql1 As String
            sql1 = "SELECT * FROM tbl_return"
            Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
            Dim cmd As New SqlCommand(sql1, myconnection)
            myconnection.Open()
            Dim myreader As SqlDataReader = cmd.ExecuteReader
            myreader.Read()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Sub viewdataReport()
        Try
            Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "SELECT billno, billdate, cname, contact, proname, progroup, price, qty, total, ototal, amount FROM tbl_stock"
            Dim adapter As New SqlDataAdapter(sql, myconnection)
            Dim data As New DataTable("tbl_stock")
            adapter.Fill(data)
            frm_mainAdmin.dgv_reportList.DataSource = data
            Dim sql1 As String
            sql1 = "SELECT * FROM tbl_stock"
            Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
            Dim cmd As New SqlCommand(sql1, myconnection)
            myconnection.Open()
            Dim myreader As SqlDataReader = cmd.ExecuteReader
            myreader.Read()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Sub viewdataReportInv()
        Try
            Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "Select supname, prodate, proname, progroup, stockno, price FROM tbl_product"
            Dim adapter As New SqlDataAdapter(sql, myconnection)
            Dim data As New DataTable("tbl_product")
            adapter.Fill(data)
            frm_mainAdmin.dgv_inventoryList.DataSource = data
            Dim sql1 As String
            sql1 = "Select * FROM tbl_product"
            Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
            Dim cmd As New SqlCommand(sql1, myconnection)
            myconnection.Open()
            Dim myreader As SqlDataReader = cmd.ExecuteReader
            myreader.Read()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Private Sub UpdateCounts()
        frm_mainAdmin.lbl_stocks.Text = "STOCKS : " & frm_mainAdmin.dgv_productList.RowCount
        frm_mainAdmin.lbl_user.Text = "USERS : " & frm_mainAdmin.dgv_userList.RowCount
        frm_mainAdmin.lbl_supplier.Text = "SUPPLIERS : " & frm_mainAdmin.dgv_supplierList.RowCount
        frm_mainAdmin.lbl_return.Text = "RETURNS : " & frm_mainAdmin.dgv_returnList.RowCount
    End Sub

    Private Sub PopulateSupplierComboBox()
        Dim connection As String = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        Dim connstring As String = connection

        Dim query As String = "SELECT DISTINCT supname FROM tbl_supplier"

        Using myconnection As New SqlConnection(connstring)
            Try
                Using command As New SqlCommand(query, myconnection)
                    myconnection.Open()

                    Using reader As SqlDataReader = command.ExecuteReader()
                        frm_mainAdmin.cbo_supplier.Items.Clear()

                        If reader.HasRows Then
                            While reader.Read()
                                frm_mainAdmin.cbo_supplier.Items.Add(reader("supname").ToString())
                            End While
                        End If
                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End Using
    End Sub

    Private Sub PopulateReturnComboBox()
        Dim connection As String = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        Dim connstring As String = connection

        Dim query As String = "SELECT DISTINCT proname FROM tbl_product"

        Using myconnection As New SqlConnection(connstring)
            Try
                Using command As New SqlCommand(query, myconnection)
                    myconnection.Open()

                    Using reader As SqlDataReader = command.ExecuteReader()
                        frm_mainAdmin.cbo_returnpro.Items.Clear()

                        If reader.HasRows Then
                            While reader.Read()
                                frm_mainAdmin.cbo_returnpro.Items.Add(reader("proname").ToString())
                            End While
                        End If
                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End Using
    End Sub

    Sub viewdataReportRet()
        Try
            Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "Select RDate, Cname, Contact, Pname, Qty, Issue, Acttaken FROM tbl_return"
            Dim adapter As New SqlDataAdapter(sql, myconnection)
            Dim data As New DataTable("tbl_return")
            adapter.Fill(data)
            frm_mainAdmin.dgv_returnListrep.DataSource = data
            Dim sql1 As String
            sql1 = "Select * FROM tbl_return"
            Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
            Dim cmd As New SqlCommand(sql1, myconnection)
            myconnection.Open()
            Dim myreader As SqlDataReader = cmd.ExecuteReader
            myreader.Read()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    '===================================================================================
    '===================================================================================
    'VIEW DATA EMPLOYEE

    Sub viewSupplier1()
        Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
        Dim sql As String
        sql = "SELECT supname, supproname, address, contact, email FROM tbl_supplier"
        Dim adapter As New SqlDataAdapter(sql, myconnection)
        Dim data As New DataTable("tbl_supplier")
        adapter.Fill(data)
        frm_mainCashier.dgv_supplierList.DataSource = data
        Dim sql1 As String
        sql1 = "SELECT * FROM tbl_supplier"
        Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
        Dim cmd As New SqlCommand(sql1, myconnection)
        myconnection.Open()
        Dim myreader As SqlDataReader = cmd.ExecuteReader
        myreader.Read()
        myconnection.Close()
    End Sub

    Sub viewdataReturn1()
        Try
            Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "SELECT RDate, Cname, Contact, Pname, Qty, Issue, Acttaken FROM tbl_return"
            Dim adapter As New SqlDataAdapter(sql, myconnection)
            Dim data As New DataTable("tbl_return")
            adapter.Fill(data)
            frm_mainCashier.dgv_returnList.DataSource = data
            Dim sql1 As String
            sql1 = "SELECT * FROM tbl_return"
            Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
            Dim cmd As New SqlCommand(sql1, myconnection)
            myconnection.Open()
            Dim myreader As SqlDataReader = cmd.ExecuteReader
            myreader.Read()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Sub viewdataReport1()
        Try
            Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "SELECT billno, billdate, cname, contact, proname, progroup, price, qty, total, ototal, amount FROM tbl_stock"
            Dim adapter As New SqlDataAdapter(sql, myconnection)
            Dim data As New DataTable("tbl_stock")
            adapter.Fill(data)
            frm_mainCashier.dgv_reportList.DataSource = data
            Dim sql1 As String
            sql1 = "SELECT * FROM tbl_stock"
            Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
            Dim cmd As New SqlCommand(sql1, myconnection)
            myconnection.Open()
            Dim myreader As SqlDataReader = cmd.ExecuteReader
            myreader.Read()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Sub viewdataReportInv1()
        Try
            Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "Select supname, prodate, proname, progroup, stockno, price FROM tbl_product"
            Dim adapter As New SqlDataAdapter(sql, myconnection)
            Dim data As New DataTable("tbl_product")
            adapter.Fill(data)
            frm_mainAdmin.dgv_inventoryList.DataSource = data
            Dim sql1 As String
            sql1 = "Select * FROM tbl_product"
            Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
            Dim cmd As New SqlCommand(sql1, myconnection)
            myconnection.Open()
            Dim myreader As SqlDataReader = cmd.ExecuteReader
            myreader.Read()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Private Sub PopulateReturnComboBoxemp()
        Dim connection As String = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        Dim connstring As String = connection

        Dim query As String = "SELECT DISTINCT proname FROM tbl_product"

        Using myconnection As New SqlConnection(connstring)
            Try
                Using command As New SqlCommand(query, myconnection)
                    myconnection.Open()

                    Using reader As SqlDataReader = command.ExecuteReader()
                        frm_mainCashier.cbo_returnpro.Items.Clear()

                        If reader.HasRows Then
                            While reader.Read()
                                frm_mainCashier.cbo_returnpro.Items.Add(reader("proname").ToString())
                            End While
                        End If
                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End Using
    End Sub

    Sub viewdataReportRet1()
        Try
            Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "Select RDate, Cname, Contact, Pname, Qty, Issue, Acttaken FROM tbl_return"
            Dim adapter As New SqlDataAdapter(sql, myconnection)
            Dim data As New DataTable("tbl_return")
            adapter.Fill(data)
            frm_mainCashier.dgv_returnListrep.DataSource = data
            Dim sql1 As String
            sql1 = "Select * FROM tbl_return"
            Dim adapter1 As New SqlDataAdapter(sql1, myconnection)
            Dim cmd As New SqlCommand(sql1, myconnection)
            myconnection.Open()
            Dim myreader As SqlDataReader = cmd.ExecuteReader
            myreader.Read()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

End Class