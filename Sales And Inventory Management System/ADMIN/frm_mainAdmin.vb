Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Diagnostics.Eventing.Reader
Imports System.Drawing.Printing
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class frm_mainAdmin
    Dim connection As String
    Dim connstring As String
    Dim myconnection As SqlConnection = New SqlConnection

    '===================================================================================
    '===================================================================================
    'MANAGE USER

    Private Sub btn_registerUser_Click(sender As Object, e As EventArgs) Handles btn_registerUser.Click
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring
        If txt_name.Text = String.Empty Then
            MsgBox("Please Fill Up Name !", vbExclamation)
            txt_name.Focus()
            Return
        ElseIf txt_username.Text = String.Empty Then
            MsgBox("Please Fill Up Username !", vbExclamation)
            txt_username.Focus()
            Return
        ElseIf txt_password.Text = String.Empty Then
            MsgBox("Please Fill Up Password !", vbExclamation)
            txt_password.Focus()
            Return
        ElseIf cbo_role.Text = String.Empty Then
            MsgBox("Please Choose Role !", vbExclamation)
            cbo_role.Focus()
            Return
        Else
            Try
                Dim dt As New DataTable
                Dim da As New SqlDataAdapter

                myconnection.Open()
                Dim query As String
                query = "Select * from tbl_user Where [Name] = '" & txt_name.Text & "' or [Username] = '" & txt_username.Text & "'"
                Dim command As SqlCommand = New SqlCommand(query, myconnection)
                dt = New DataTable
                da = New SqlDataAdapter(command)
                da.Fill(dt)
                myconnection.Close()

                If dt.Rows.Count > 0 Then
                    MsgBox("Duplicate User !", vbExclamation)
                    clearUser()
                Else
                    myconnection.Open()
                    Dim query1 As String
                    query1 = "Insert into tbl_user ([Name],[Username],[Password], [Role]) values (@Name, @Username, @Password, @Role)"
                    Dim cmd As SqlCommand = New SqlCommand(query1, myconnection)

                    cmd.Parameters.AddWithValue("@Name", txt_name.Text)
                    cmd.Parameters.AddWithValue("@Username", txt_username.Text)
                    cmd.Parameters.AddWithValue("@Password", txt_password.Text)
                    cmd.Parameters.AddWithValue("@role", cbo_role.Text)
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()
                    MsgBox("New User Register Success !", vbInformation, "Success")
                    txt_name.Focus()
                    clearUser()
                    viewdataUSer()
                    myconnection.Close()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End If
    End Sub

    Sub clearUser()
        txt_name.Clear()
        txt_username.Clear()
        txt_password.Clear()
        txt_searchUser.Clear()
        cbo_role.SelectedIndex = -1
    End Sub

    Private Sub btn_cancelUser_Click(sender As Object, e As EventArgs) Handles btn_cancelUser.Click
        clearUser()
        cbo_role.Enabled = True
        txt_name.Enabled = True
        txt_name.Focus()
    End Sub

    Sub viewdataUSer()
        Try
            Dim con1 As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "SELECT Name, Username, Password, Role FROM tbl_user"
            Dim adapter As New SqlDataAdapter(sql, con1)
            Dim data As New DataTable("tbl_user")
            adapter.Fill(data)
            dgv_userList.DataSource = data
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

    Private Sub btn_updateUser_Click(sender As Object, e As EventArgs) Handles btn_updateUser.Click
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring
        If txt_name.Text = String.Empty Then
            MsgBox("Please Choose a User to Edit !", vbExclamation)
            Return
        ElseIf txt_username.Text = String.Empty Then
            MsgBox("Username is Empty !", vbExclamation)
            Return
        ElseIf txt_password.Text = String.Empty Then
            MsgBox("Password is Empty !", vbExclamation)
            Return
        Else
            Try
                myconnection.Open()
                cbo_role.Enabled = True
                txt_name.Enabled = True
                Dim query As String
                query = "Update tbl_user Set Username = @Username, Password = @Password Where Name = @Name"
                Dim cmd As SqlCommand = New SqlCommand(query, myconnection)

                cmd.Parameters.AddWithValue("@Name", txt_name.Text)
                cmd.Parameters.AddWithValue("@Username", txt_username.Text)
                cmd.Parameters.AddWithValue("@Password", txt_password.Text)

                cmd.ExecuteNonQuery()
                cmd.Dispose()
                MsgBox("Success Update User !", vbInformation, "Success")
                txt_name.Enabled = True
                txt_name.Focus()
                clearUser()
                viewdataUSer()
                myconnection.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End If
    End Sub

    Private Sub txt_name_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_name.KeyPress
        'Letters Only
        If Not Char.IsLetter(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields is Letters Only!", vbExclamation)
        End If
    End Sub

    Private Sub dgv_userList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_userList.CellContentClick
        Dim colNamepro As String = dgv_productList.Columns(e.ColumnIndex).Name
        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()
            If colNamepro = "EditPro" Then
                Dim dt As DataGridViewRow = dgv_userList.SelectedRows(0)
                txt_name.Text = dt.Cells(2).Value.ToString()
                txt_username.Text = dt.Cells(3).Value.ToString()
                txt_password.Text = dt.Cells(4).Value.ToString()
                cbo_role.Text = dt.Cells(5).Value.ToString()
                txt_name.Enabled = False
            ElseIf colNamepro = "DeletePro" Then
                Dim adminUsername As String = "ADMIN"

                Dim password As String = InputBox("Enter Your Password to Confirm Deletion:", "Password Required !")
                If password = adminUsername Then
                    Dim dt As DataGridViewRow = dgv_userList.SelectedRows(0)
                    txt_name.Text = dt.Cells(2).Value.ToString()
                    txt_username.Text = dt.Cells(3).Value.ToString()
                    txt_password.Text = dt.Cells(4).Value.ToString()
                    cbo_role.Text = dt.Cells(5).Value.ToString()

                    Dim query As String = "DELETE FROM tbl_user WHERE [Name] = '" & txt_name.Text & "'"
                    Using cmd As New SqlCommand(query, myconnection)
                        cmd.ExecuteNonQuery()
                    End Using
                    MsgBox("Success Delete User !", vbInformation, "Success")
                    txt_name.Enabled = True
                    dgv_userList.Rows.RemoveAt(dgv_userList.SelectedRows(0).Index)
                    txt_name.Focus()
                    clearUser()
                    viewdataUSer()
                ElseIf password <> "" Then
                    MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                Else
                    MsgBox("Deletion Cancelled.", vbExclamation, "Operation Cancelled")
                End If
                myconnection.Close()
            End If
            myconnection.Close()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Private Sub txt_searchUser_TextChanged(sender As Object, e As EventArgs) Handles txt_searchUser.TextChanged
        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim query As String
            query = "Select Name, Username, Password, Role from tbl_user Where [Name] like '%" & txt_searchUser.Text & "%' or  [Username] like '%" & txt_searchUser.Text & "%' or  [Role] like '%" & txt_searchUser.Text & "%' "
            Dim dt As New DataTable
            Dim command As SqlCommand = New SqlCommand(query, myconnection)
            Dim da As New SqlDataAdapter
            da.SelectCommand = command
            da.Fill(dt)
            dgv_userList.DataSource = dt
            myconnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Private Sub txt_password_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txt_password.Validating
        Dim minContactLength As Integer = 8

        If txt_password.TextLength > 0 AndAlso (txt_password.TextLength < minContactLength) Then
            MessageBox.Show("Password should be " & minContactLength & " characters and above.", "Invalid Password !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            e.Cancel = True
            txt_password.Clear()
        End If
    End Sub
    '===================================================================================
    '===================================================================================
    'FORM LOAD 

    Private Sub frm_mainAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        viewdataUSer()
        viewproduct()
        viewSupplier()
        viewdataReturn()
        viewdataReport()
        viewdataReportInv()
        viewdataReportRet()
        UpdateCounts()
        PopulateSupplierComboBox()
        PopulateReturnComboBox()
        dtp_prodate.Value = DateTime.Now
        dtp_return.Value = DateTime.Now
        dtp_billdate.Value = DateTime.Now
        dtp_startInv.Value = DateTime.Now
        dtp_endInv.Value = DateTime.Now
        lbl_billno.Text = Getbillno()
        sumtotal()
        AddHandler dgv_productList.RowsAdded, AddressOf UpdateCounts
        AddHandler dgv_userList.RowsAdded, AddressOf UpdateCounts
        AddHandler dgv_supplierList.RowsAdded, AddressOf UpdateCounts
        AddHandler dgv_returnList.RowsAdded, AddressOf UpdateCounts
    End Sub

    Private Sub UpdateCounts()
        lbl_stocks.Text = "STOCKS : " & dgv_productList.RowCount
        lbl_user.Text = "USERS : " & dgv_userList.RowCount
        lbl_supplier.Text = "SUPPLIERS : " & dgv_supplierList.RowCount
        lbl_return.Text = "RETURNS : " & dgv_returnList.RowCount
    End Sub

    Private Sub TabControl1_Click(sender As Object, e As EventArgs) Handles TabControl1.Click
        clearProduct()
        clearSupplier()
        clearUser()
        clearReturn()
        clearStock()
        clearReport()
        clearReportInv()
        clearReportRet()
        UpdateCounts()
        cbo_supplier.Enabled = True
        cbo_product.Enabled = True
        lbl_billno.Text = Getbillno()
        txt_name.Focus()
        txt_supplierName.Focus()
        txt_searchPro.Focus()
        txt_creturn.Focus()
        txt_searchMonth.Focus()
        txt_searchProMonth.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring
        If MsgBox("Are You Sure You Want to Logout ?", vbInformation + vbYesNo) = vbYes Then
            Me.Hide()
            clearProduct()
            clearUser()
            clearSupplier()
            clearReturn()
            clearStock()
            clearReport()
            clearReportInv()
            clearReportRet()
            frm_login.Show()
            frm_login.txt_username.Focus()
            myconnection.Close()
        Else
            Return
        End If
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl2.SelectedIndexChanged
        txt_searchMonth.Focus()
        txt_searchProMonth.Focus()
        txt_searchret.Focus()
    End Sub

    '===================================================================================
    '===================================================================================
    'MANAGE PRODUCTS

    Private Sub btn_saveProduct_Click(sender As Object, e As EventArgs) Handles btn_saveProduct.Click
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring
        If cbo_supplier.Text = String.Empty Then
            MsgBox("Please Choose a Supplier !", vbExclamation)
            Return
        ElseIf cbo_product.Text = String.Empty Then
            MsgBox("Please Choose a Product !", vbExclamation)
            cbo_product.Focus()
            Return
        ElseIf cbo_progroup.Text = String.Empty Then
            MsgBox("Please Choose Product Group !", vbExclamation)
            cbo_progroup.Focus()
            Return
        ElseIf txt_stockno.Text = String.Empty Then
            MsgBox("Please Fill Up No. of Stock !", vbExclamation)
            txt_stockno.Focus()
            Return
        ElseIf txt_price.Text = String.Empty Then
            MsgBox("Please Fill Up Price !", vbExclamation)
            txt_price.Focus()
            Return
        Else

            Try
                Dim dt As New DataTable
                Dim da As New SqlDataAdapter
                Dim currentDate As String = DateTime.Now.Date.ToString("yyyy/MM/dd")

                myconnection.Open()
                Dim query1 As String
                query1 = "Select * from tbl_product Where [proname] = '" & cbo_product.Text & "'"
                Dim command As SqlCommand = New SqlCommand(query1, myconnection)
                dt = New DataTable
                da = New SqlDataAdapter(command)
                da.Fill(dt)
                myconnection.Close()

                If dt.Rows.Count > 0 Then
                    MsgBox("Duplicate Product !", vbExclamation)
                    clearProduct()
                Else
                    myconnection.Open()
                    Dim query As String
                    query = "Insert into tbl_product ([supname], [prodate], [proname], [progroup], [stockno], [price]) values (@supname, @prodate, @proname, @progroup, @stockno, @price)"
                    Dim cmd As SqlCommand = New SqlCommand(query, myconnection)

                    cmd.Parameters.AddWithValue("@supname", cbo_supplier.Text)
                    cmd.Parameters.AddWithValue("@prodate", currentDate)
                    cmd.Parameters.AddWithValue("@proname", cbo_product.Text)
                    cmd.Parameters.AddWithValue("@progroup", cbo_progroup.Text)
                    cmd.Parameters.AddWithValue("@stockno", txt_stockno.Text)
                    cmd.Parameters.AddWithValue("@price", txt_price.Text)
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()

                    MsgBox("New Product Successfully Added !", vbInformation, "Success")

                    txt_name.Focus()
                    clearProduct()
                    viewproduct()
                    viewdataReportInv()
                    PopulateReturnComboBox()
                    myconnection.Close()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End If
    End Sub

    Private Sub btn_updateProduct_Click(sender As Object, e As EventArgs) Handles btn_updateProduct.Click
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring
        If cbo_product.Text = String.Empty Then
            MsgBox("Please Choose a Product to Edit !", vbExclamation)
            Return
        ElseIf txt_stockno.Text = String.Empty Then
            MsgBox("Invalid Quantity !", vbExclamation)
            Return
        ElseIf txt_stockno.Text = String.Empty Then
            MsgBox("Invalid Price !", vbExclamation)
            Return
        Else
            Try
                myconnection.Open()
                Dim query As String
                query = "Update tbl_product Set progroup = @progroup, stockno = @stockno, price = @price Where proname = @proname"
                Dim cmd As SqlCommand = New SqlCommand(query, myconnection)

                cmd.Parameters.AddWithValue("@proname", cbo_product.Text)
                cmd.Parameters.AddWithValue("@progroup", cbo_progroup.Text)
                cmd.Parameters.AddWithValue("@stockno", txt_stockno.Text)
                cmd.Parameters.AddWithValue("@price", txt_price.Text)

                cmd.ExecuteNonQuery()
                cmd.Dispose()
                MsgBox("Success Update Product !", vbInformation, "Success")
                txt_name.Focus()
                cbo_supplier.Enabled = True
                cbo_product.Enabled = True
                clearProduct()
                viewproduct()
                viewdataReportInv()
                PopulateReturnComboBox()
                myconnection.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

    End Sub

    Sub clearProduct()
        cbo_product.Items.Clear()
        cbo_progroup.SelectedIndex = -1
        txt_stockno.Clear()
        txt_price.Clear()
        txt_searchProID.Clear()
        dtp_prodate.Value = DateTime.Now
        cbo_supplier.SelectedIndex = -1
    End Sub

    Sub viewproduct()
        Dim con1 As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
        Dim sql As String
        sql = "SELECT supname, prodate, proname, progroup, stockno, price FROM tbl_product"
        Dim adapter As New SqlDataAdapter(sql, con1)
        Dim data As New DataTable("tbl_product")
        adapter.Fill(data)
        dgv_productList.DataSource = data
        Dim sql1 As String
        sql1 = "SELECT * FROM tbl_product"
        Dim adapter1 As New SqlDataAdapter(sql1, con1)
        Dim cmd As New SqlCommand(sql1, con1)
        con1.Open()
        Dim myreader As SqlDataReader = cmd.ExecuteReader
        myreader.Read()
        con1.Close()
    End Sub

    Private Sub btn_clearProduct_Click(sender As Object, e As EventArgs) Handles btn_clearProduct.Click
        clearProduct()
        txt_name.Focus()
        cbo_supplier.Enabled = True
        cbo_product.Enabled = True
    End Sub

    Private Sub txt_stockno_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_stockno.KeyPress
        'Numbers Only
        If Not Char.IsNumber(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields Is Numbers Only!", vbExclamation)
        End If
    End Sub

    Private Sub txt_price_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_price.KeyPress
        'Numbers Only
        If Not Char.IsNumber(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields Is Numbers Only!", vbExclamation)
        End If
    End Sub

    Private Sub dgv_productList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_productList.CellContentClick

        Dim colNamepro As String = dgv_productList.Columns(e.ColumnIndex).Name
        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()
            If colNamepro = "EditPro" Then
                Dim dt As DataGridViewRow = dgv_productList.SelectedRows(0)
                cbo_supplier.Text = dt.Cells(2).Value.ToString()
                dtp_prodate.Text = dt.Cells(3).Value.ToString()
                cbo_product.Text = dt.Cells(4).Value.ToString()
                cbo_progroup.Text = dt.Cells(5).Value.ToString()
                txt_stockno.Text = dt.Cells(6).Value.ToString()
                txt_price.Text = dt.Cells(7).Value.ToString()
                cbo_product.Enabled = False
                cbo_supplier.Enabled = False
            ElseIf colNamepro = "DeletePro" Then
                Dim adminUsername As String = "ADMIN"

                Dim password As String = InputBox("Enter Your Password to Confirm Deletion:", "Password Required !")

                If password = adminUsername Then
                    Dim dt1 As DataGridViewRow = dgv_productList.SelectedRows(0)
                    cbo_supplier.Text = dt1.Cells(2).Value.ToString()
                    dtp_prodate.Text = dt1.Cells(3).Value.ToString()
                    cbo_product.Text = dt1.Cells(4).Value.ToString()
                    cbo_progroup.Text = dt1.Cells(5).Value.ToString()
                    txt_stockno.Text = dt1.Cells(6).Value.ToString()
                    txt_price.Text = dt1.Cells(7).Value.ToString()

                    Dim query As String = "DELETE FROM tbl_product WHERE [proname] = '" & cbo_product.Text & "'"
                    Using cmd As New SqlCommand(query, myconnection)
                        cmd.ExecuteNonQuery()
                    End Using
                    MsgBox("Success Delete Product !", vbInformation, "Success")
                    dgv_productList.Rows.RemoveAt(dgv_productList.SelectedRows(0).Index)
                    clearProduct()
                    viewproduct()
                    viewdataReportInv()
                    PopulateReturnComboBox()
                ElseIf password <> "" Then
                    MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                Else
                    MsgBox("Deletion Cancelled.", vbExclamation, "Operation Cancelled")
                End If
                myconnection.Close()
            End If
            myconnection.Close()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Private Sub txt_searchProID_TextChanged(sender As Object, e As EventArgs) Handles txt_searchProID.TextChanged
        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim query As String
            query = "Select supname, prodate, proname, progroup, stockno, price from tbl_product Where [supname] like '%" & txt_searchProID.Text & "%' or  [prodate] like '%" & txt_searchProID.Text & "%' or   [proname] like '%" & txt_searchProID.Text & "%' or  [progroup] like '%" & txt_searchProID.Text & "%' "
            Dim dt As New DataTable
            Dim command As SqlCommand = New SqlCommand(query, myconnection)
            Dim da As New SqlDataAdapter
            da.SelectCommand = command
            da.Fill(dt)
            dgv_productList.DataSource = dt
            myconnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Private Sub txt_progroup_KeyPress(sender As Object, e As KeyPressEventArgs)
        'Letters Only
        If Not Char.IsLetter(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields is Letters Only!", vbExclamation)
        End If
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
                        cbo_supplier.Items.Clear()

                        If reader.HasRows Then
                            While reader.Read()
                                cbo_supplier.Items.Add(reader("supname").ToString())
                            End While
                        End If
                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using
    End Sub

    Private Sub cbo_supplier_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cbo_supplier.SelectedIndexChanged
        If cbo_supplier.SelectedItem IsNot Nothing Then
            Dim selectedSupplier As String = cbo_supplier.SelectedItem.ToString()
            Dim connection As String = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            Dim query As String = "SELECT supproname FROM tbl_supplier WHERE supname = @supname"

            Using myconnection As New SqlConnection(connection)
                myconnection.ConnectionString = connection

                Try
                    Using command As New SqlCommand(query, myconnection)
                        command.Parameters.AddWithValue("@supname", selectedSupplier)

                        myconnection.Open()

                        Dim reader As SqlDataReader = command.ExecuteReader()

                        cbo_product.Items.Clear()

                        While reader.Read()
                            cbo_product.Items.Add(reader("supproname").ToString())
                        End While

                        reader.Close()

                        If cbo_product.Items.Count = 0 Then
                            cbo_product.Text = "No products found for this supplier."
                        End If
                    End Using

                Catch ex As Exception
                    MsgBox("Error: " & ex.Message)
                End Try
            End Using
        End If
    End Sub

    '===================================================================================
    '===================================================================================
    'MANAGE SUPPLIER

    Private Sub btn_addSupplier_Click(sender As Object, e As EventArgs) Handles btn_addSupplier.Click
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring
        If txt_supplierName.Text = String.Empty Then
            MsgBox("Please Fill Up Supplier Name !", vbExclamation)
            txt_supplierName.Focus()
            Return
        ElseIf txt_productname.Text = String.Empty Then
            MsgBox("Please Fill Up Product Name !", vbExclamation)
            txt_productname.Focus()
            Return
        ElseIf txt_address.Text = String.Empty Then
            MsgBox("Please Fill Up Address !", vbExclamation)
            txt_address.Focus()
            Return
        ElseIf txt_supplierNumber.Text = String.Empty Then
            MsgBox("Please Fill Up Contact!", vbExclamation)
            txt_supplierNumber.Focus()
            Return
        ElseIf txt_email.Text = String.Empty Then
            MsgBox("Please Fill Up Email !", vbExclamation)
            txt_email.Focus()
            Return
        Else

            Try
                myconnection.Open()
                Dim query As String
                query = "Insert into tbl_supplier ([supname], [supproname], [address], [contact], [email]) values (@supname, @supproname, @address, @contact, @email)"
                Dim cmd As SqlCommand = New SqlCommand(query, myconnection)

                cmd.Parameters.AddWithValue("@supname", txt_supplierName.Text)
                cmd.Parameters.AddWithValue("@supproname", txt_productname.Text)
                cmd.Parameters.AddWithValue("@address", txt_address.Text)
                cmd.Parameters.AddWithValue("@contact", txt_supplierNumber.Text)
                cmd.Parameters.AddWithValue("@email", txt_email.Text)
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                MsgBox("New Supplier Successfully Added !", vbInformation, "Success")
                PopulateSupplierComboBox()
                clearSupplier()
                viewSupplier()
                myconnection.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End If
    End Sub

    Sub clearSupplier()
        txt_supplierNumber.Clear()
        txt_supplierName.Clear()
        txt_productname.Clear()
        txt_address.Clear()
        txt_email.Clear()
        txt_searchSupplier.Clear()
    End Sub

    Sub viewSupplier()
        Dim con1 As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
        Dim sql As String
        sql = "SELECT supname, supproname, address, contact, email FROM tbl_supplier"
        Dim adapter As New SqlDataAdapter(sql, con1)
        Dim data As New DataTable("tbl_supplier")
        adapter.Fill(data)
        dgv_supplierList.DataSource = data
        Dim sql1 As String
        sql1 = "SELECT * FROM tbl_supplier"
        Dim adapter1 As New SqlDataAdapter(sql1, con1)
        Dim cmd As New SqlCommand(sql1, con1)
        con1.Open()
        Dim myreader As SqlDataReader = cmd.ExecuteReader
        myreader.Read()
        con1.Close()
    End Sub

    Private Sub btn_updateSupplier_Click(sender As Object, e As EventArgs) Handles btn_updateSupplier.Click
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring
        If txt_supplierName.Text = String.Empty Then
            MsgBox("Please Choose a Supplier to Edit !", vbExclamation)
            Return
        ElseIf txt_productname.Text = String.Empty Then
            MsgBox("Product is Empty !", vbExclamation)
            Return
        ElseIf txt_address.Text = String.Empty Then
            MsgBox("Address is Empty !", vbExclamation)
            Return
        ElseIf txt_supplierNumber.Text = String.Empty Then
            MsgBox("Supplier Number is Empty !", vbExclamation)
            Return
        ElseIf txt_email.Text = String.Empty Then
            MsgBox("Empty is Empty !", vbExclamation)
            Return
        Else
            Try
                myconnection.Open()
                Dim query As String
                query = "Update tbl_supplier Set  supproname = @supproname, address = @address, contact = @contact, email = @email Where supname = @supname"
                Dim cmd As SqlCommand = New SqlCommand(query, myconnection)

                cmd.Parameters.AddWithValue("@supname", txt_supplierName.Text)
                cmd.Parameters.AddWithValue("@supproname", txt_productname.Text)
                cmd.Parameters.AddWithValue("@address", txt_address.Text)
                cmd.Parameters.AddWithValue("@contact", txt_supplierNumber.Text)
                cmd.Parameters.AddWithValue("@email", txt_email.Text)

                cmd.ExecuteNonQuery()
                cmd.Dispose()
                MsgBox("Success Update Supplier !", vbInformation, "Success")
                PopulateSupplierComboBox()
                txt_supplierName.Enabled = True
                txt_supplierName.Focus()
                clearSupplier()
                viewSupplier()
                myconnection.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub btn_cancelSupplier_Click(sender As Object, e As EventArgs) Handles btn_cancelSupplier.Click
        clearSupplier()
        txt_supplierName.Focus()
        txt_supplierName.Enabled = True
    End Sub

    Private Sub txt_supplierNumber_KeyPress_1(sender As Object, e As KeyPressEventArgs) Handles txt_supplierNumber.KeyPress
        'Contact Number Only
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
            MsgBox("This Field is Contact Number Only !", vbExclamation)
        End If
    End Sub

    Private Sub txt_supplierNumber_Validating_1(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txt_supplierNumber.Validating
        Dim minContactLength As Integer = 11 ' Minimum length of the contact number
        Dim maxContactLength As Integer = 12 ' Maximum length of the contact number

        If txt_supplierNumber.TextLength > 0 AndAlso (txt_supplierNumber.TextLength < minContactLength OrElse txt_supplierNumber.TextLength > maxContactLength) Then
            MessageBox.Show("Contact number should be between " & minContactLength & " and " & maxContactLength & " digits.", "Invalid Contact", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            e.Cancel = True
            txt_supplierNumber.Clear()
        End If
    End Sub

    Private Sub txt_email_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txt_email.Validating
        If txt_email.Text <> String.Empty Then
            Dim isValidEmail As Boolean

            isValidEmail = Regex.IsMatch(txt_email.Text, "\A(?:[a-z0-9!#$%&'*+/=?^_'{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_'{|}~-]+)*@(?:[a-z0-9-](?:[a-z0-9]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase)
            If Not isValidEmail Then
                MsgBox("Not a valid email!", vbExclamation)
                e.Cancel = True ' Cancel the event to prevent focus change
                txt_email.Clear()
            End If
        End If
    End Sub

    Private Sub dgv_supplierList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_supplierList.CellContentClick
        Dim colNamepro As String = dgv_supplierList.Columns(e.ColumnIndex).Name
        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()
            If colNamepro = "EditSup" Then
                Dim dt As DataGridViewRow = dgv_supplierList.SelectedRows(0)
                txt_supplierName.Text = dt.Cells(2).Value.ToString()
                txt_productname.Text = dt.Cells(3).Value.ToString()
                txt_address.Text = dt.Cells(4).Value.ToString()
                txt_supplierNumber.Text = dt.Cells(5).Value.ToString()
                txt_email.Text = dt.Cells(6).Value.ToString()
                txt_supplierName.Enabled = False
            ElseIf colNamepro = "DeleteSup" Then
                Dim adminUsername As String = "ADMIN"

                Dim password As String = InputBox("Enter Your Password to Confirm Deletion:", "Password Required !")
                If password = adminUsername Then
                    Dim dt1 As DataGridViewRow = dgv_supplierList.SelectedRows(0)
                    txt_supplierName.Text = dt1.Cells(2).Value.ToString()
                    txt_productname.Text = dt1.Cells(3).Value.ToString()
                    txt_address.Text = dt1.Cells(4).Value.ToString()
                    txt_supplierNumber.Text = dt1.Cells(5).Value.ToString()
                    txt_email.Text = dt1.Cells(6).Value.ToString()

                    Dim query As String = "DELETE from tbl_supplier Where [supproname] = '" & txt_productname.Text & "'"
                    Using cmd As New SqlCommand(query, myconnection)
                        cmd.ExecuteNonQuery()
                    End Using
                    MsgBox("Success Delete Supplier !", vbInformation, "Success")
                    PopulateSupplierComboBox()
                    txt_supplierName.Enabled = True
                    dgv_supplierList.Rows.RemoveAt(dgv_supplierList.SelectedRows(0).Index)
                    txt_supplierName.Focus()
                    clearSupplier()
                    viewSupplier()
                    viewproduct()
                ElseIf password <> "" Then
                    MsgBox("Incorrect Password. Deletion Cancelled !", vbExclamation, "Password Incorrect !")
                Else
                    MsgBox("Deletion Cancelled.", vbExclamation, "Operation Cancelled")
                End If
                myconnection.Close()
            End If
            myconnection.Close()
            myconnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Private Sub txt_searchSupplier_TextChanged(sender As Object, e As EventArgs) Handles txt_searchSupplier.TextChanged
        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim query As String
            query = "Select supname, supproname, address, contact, email from tbl_supplier Where [supname] like '%" & txt_searchSupplier.Text & "%' or  [supproname] like '%" & txt_searchSupplier.Text & "%'"
            Dim dt As New DataTable
            Dim command As SqlCommand = New SqlCommand(query, myconnection)
            Dim da As New SqlDataAdapter
            da.SelectCommand = command
            da.Fill(dt)
            dgv_supplierList.DataSource = dt
            myconnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Private Sub txt_supplierName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_supplierName.KeyPress
        'Letters Only
        If Not Char.IsLetter(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields is Letters Only!", vbExclamation)
        End If
    End Sub

    '===================================================================================
    '===================================================================================
    'MANAGE SALES

    Public Sub ADDLIST()
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring

        Dim exist As Boolean = False, numrow As Integer = 0, numtext As Integer
        For Each itm As DataGridViewRow In dgv_stockList.Rows
            If itm.Cells(0).Value IsNot Nothing Then
                If itm.Cells(1).Value.ToString = txt_searchPro.Text Then
                    exist = True
                    numrow = itm.Index
                    numtext = CInt(itm.Cells(5).Value)
                    Exit For
                End If
            End If
        Next
        If exist = False Then
            Try
                myconnection.Open()
                Dim query As String
                query = "Select * from tbl_product Where proname like '%" & txt_searchPro.Text & "%'"
                cmd = New SqlCommand(query, myconnection)
                Dim dr As SqlDataReader
                Dim blnfound As Boolean
                blnfound = False
                dr = cmd.ExecuteReader
                While dr.Read
                    If txt_searchPro.Text = String.Empty Then
                        Return
                    Else
                        blnfound = True
                        'create new row
                        '`proname`, `progroup`, `price`, `total`
                        Dim proname As String = dr.Item("proname")
                        Dim progroup As String = dr.Item("progroup")
                        Dim price As Decimal = dr.Item("price")

                        Dim total As Decimal
                        total = price * 1

                        dgv_stockList.Rows.Add(dgv_stockList.Rows.Count + 1, proname, progroup, price, 1, total)

                        txt_searchPro.Clear()
                        txt_searchPro.Focus()
                        sumtotal()

                    End If
                End While
                If blnfound = False Then
                    MsgBox("Product Not Found !", vbExclamation)
                    txt_searchPro.Clear()
                    myconnection.Close()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            myconnection.Close()
        Else
            dgv_stockList.Rows(numrow).Cells(5).Value = CInt("1") + numtext
            dgv_stockList.Rows(numrow).Cells(6).Value = dgv_stockList.Rows(numrow).Cells(5).Value * dgv_stockList.Rows(numrow).Cells(4).Value
            sumtotal()
        End If

    End Sub

    Private Sub txt_searchPro_KeyDown(sender As Object, e As KeyEventArgs) Handles txt_searchPro.KeyDown
        If e.KeyCode = Keys.Enter AndAlso Not String.IsNullOrEmpty(txt_searchPro.Text.Trim()) Then
            Dim connectionString As String = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"

            Dim myconnection As New SqlConnection(connectionString)
            Try
                Dim searchCode As String = txt_searchPro.Text.Trim()

                myconnection.Open()

                Dim query As String = "SELECT proname, progroup, price, stockno FROM tbl_product WHERE proname = @proname"

                Using cmd As New SqlCommand(query, myconnection)
                    cmd.Parameters.AddWithValue("@proname", searchCode)

                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim proname As String = reader("proname").ToString()
                            Dim progroup As String = reader("progroup").ToString()
                            Dim price As Decimal = Convert.ToDecimal(reader("price"))
                            Dim stockNo As Integer = Convert.ToInt32(reader("stockno"))

                            If stockNo > 0 Then
                                Dim existingItem As DataGridViewRow = dgv_stockList.Rows.Cast(Of DataGridViewRow)().FirstOrDefault(Function(row) row.Cells(1).Value.ToString() = proname)

                                If existingItem IsNot Nothing Then
                                    MsgBox("The product is already existing!", vbExclamation)
                                Else
                                    Dim total As Decimal = price * 1
                                    dgv_stockList.Rows.Add(dgv_stockList.Rows.Count + 1, proname, progroup, price, 1, total)

                                    Dim currentTotal As Decimal = Convert.ToDecimal(lbl_total.Text)
                                    Dim newTotal As Decimal = currentTotal + price
                                    lbl_total.Text = newTotal.ToString()
                                End If
                            Else
                                MsgBox("The product is out of stock!", vbExclamation)
                            End If
                        Else
                            MsgBox("Product not found !", vbExclamation)
                        End If
                    End Using
                End Using

                myconnection.Close()

                txt_searchPro.Clear()
                txt_searchPro.Focus()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub btn_savePur_Click(sender As Object, e As EventArgs) Handles btn_savePur.Click
        Dim myconnection As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
        Try
            myconnection.Open()
            If txt_cname.Text = "" Then
                MsgBox("Please Insert Customer Name !", vbExclamation)
                txt_cname.Focus()
                Return
            ElseIf txt_contact.Text = "" Then
                MsgBox("Please Insert Contact !", vbExclamation)
                txt_contact.Focus()
                Return
            ElseIf txt_amount.Text = "" Then
                MsgBox("Please Insert Amount !", vbExclamation)
                txt_amount.Focus()
                Return
            End If

            Dim amount As Integer = CInt(txt_amount.Text)
            Dim total As Integer = CInt(lbl_total.Text)

            If amount < total Then
                MsgBox("Not Enough Balance !", vbExclamation)
                txt_amount.Clear()
                txt_amount.Focus()
                Return
            End If

            Dim currentDate As String = DateTime.Now.Date.ToString("yyyy/MM/dd")

            For j As Integer = 0 To dgv_stockList.Rows.Count - 1
                Dim purchasedQty As Integer = Convert.ToInt32(dgv_stockList.Rows(j).Cells(4).Value)
                Dim productname As String = dgv_stockList.Rows(j).Cells(1).Value.ToString()

                Dim checkStockQuery As String = "SELECT stockno FROM tbl_product WHERE proname = @proname"
                Dim checkStockCmd As New SqlCommand(checkStockQuery, myconnection)
                checkStockCmd.Parameters.AddWithValue("@proname", productname)
                Dim availableStock As Integer = Convert.ToInt32(checkStockCmd.ExecuteScalar())

                If availableStock < purchasedQty Then
                    MsgBox("Not enough quantity available for Product Name : " & productname, vbExclamation)
                    Return
                End If
            Next

            Dim cmd As SqlCommand
            For i As Integer = 0 To dgv_stockList.Rows.Count - 1
                Dim query As String = "Insert into tbl_stock ([billno],[billdate],[cname], [contact], [proname], [progroup], [price], [qty], [total], [ototal], [amount]) values (@billno, @billdate, @cname, @contact, @proname, @progroup, @price, @qty, @total, @ototal, @amount)"
                cmd = New SqlCommand(query, myconnection)
                cmd.Parameters.AddWithValue("@billno", lbl_billno.Text)
                cmd.Parameters.AddWithValue("@billdate", currentDate)
                cmd.Parameters.AddWithValue("@cname", txt_cname.Text)
                cmd.Parameters.AddWithValue("@contact", txt_contact.Text)
                cmd.Parameters.AddWithValue("@proname", dgv_stockList.Rows(i).Cells(1).Value)
                cmd.Parameters.AddWithValue("@progroup", dgv_stockList.Rows(i).Cells(2).Value)
                cmd.Parameters.AddWithValue("@price", dgv_stockList.Rows(i).Cells(3).Value)
                cmd.Parameters.AddWithValue("@qty", dgv_stockList.Rows(i).Cells(4).Value)
                cmd.Parameters.AddWithValue("@total", dgv_stockList.Rows(i).Cells(5).Value)
                cmd.Parameters.AddWithValue("@ototal", lbl_total.Text)
                cmd.Parameters.AddWithValue("@amount", txt_amount.Text)
                cmd.ExecuteNonQuery()

                Dim updateQuery As String = "UPDATE tbl_product SET stockno = stockno - @qty WHERE proname = @proname"
                Dim updateCmd As New SqlCommand(updateQuery, myconnection)
                updateCmd.Parameters.AddWithValue("@qty", dgv_stockList.Rows(i).Cells(4).Value)
                updateCmd.Parameters.AddWithValue("@proname", dgv_stockList.Rows(i).Cells(1).Value)
                updateCmd.ExecuteNonQuery()
            Next

            MsgBox("New Transaction Save Success !" & vbNewLine & "Bill No. : " & lbl_billno.Text, vbInformation)
            clearStock()
            viewdataReport()
            viewproduct()
            txt_searchPro.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
            lbl_billno.Text = Getbillno()
        End Try
    End Sub

    Private Sub dgv_stockList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_stockList.CellContentClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim colNamesup As String = dgv_stockList.Columns(e.ColumnIndex).Name
            If colNamesup = "Addqty" Then
                Dim currentQuantity As Integer = CInt(dgv_stockList.Rows(e.RowIndex).Cells(4).Value)
                Dim pricePerItem As Decimal = CDec(dgv_stockList.Rows(e.RowIndex).Cells(3).Value)

                dgv_stockList.Rows(e.RowIndex).Cells(4).Value = currentQuantity + 1
                dgv_stockList.Rows(e.RowIndex).Cells(5).Value = (currentQuantity + 1) * pricePerItem
            ElseIf colNamesup = "Minusqty" Then
                Dim currentQuantity As Integer = CInt(dgv_stockList.Rows(e.RowIndex).Cells(4).Value)
                Dim pricePerItem As Decimal = CDec(dgv_stockList.Rows(e.RowIndex).Cells(3).Value)

                If currentQuantity > 1 Then
                    dgv_stockList.Rows(e.RowIndex).Cells(4).Value = currentQuantity - 1
                    dgv_stockList.Rows(e.RowIndex).Cells(5).Value = (currentQuantity - 1) * pricePerItem
                End If
            End If

            Dim totalAmount As Decimal = 0
            For Each row As DataGridViewRow In dgv_stockList.Rows
                totalAmount += CDec(row.Cells(5).Value)
            Next
            lbl_total.Text = totalAmount.ToString()

            Dim amountReceived As Decimal
            If Decimal.TryParse(txt_amount.Text, amountReceived) AndAlso amountReceived >= totalAmount Then
                Dim change As Decimal = amountReceived - totalAmount
                lbl_change.Text = change.ToString()
            Else
                lbl_change.Text = "00.00"
            End If
        End If
    End Sub

    Sub clearStock()
        txt_searchPro.Clear()
        dtp_billdate.Text = Now
        dgv_stockList.Rows.Clear()
        txt_cname.Clear()
        txt_contact.Clear()
        txt_amount.Clear()
        lbl_total.Text = "00.00"
        lbl_change.Text = "00.00"
    End Sub

    Private Sub txt_contact_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_contact.KeyPress
        'Contact Number Only
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txt_contact_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txt_contact.Validating

        Dim minContactLength As Integer = 11 ' Minimum length of the contact number
        Dim maxContactLength As Integer = 12 ' Maximum length of the contact number

        If txt_contact.TextLength > 0 AndAlso (txt_contact.TextLength < minContactLength OrElse txt_contact.TextLength > maxContactLength) Then
            MessageBox.Show("Contact number should be between " & minContactLength & " and " & maxContactLength & " digits.", "Invalid Contact", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            e.Cancel = True
            txt_contact.Clear()
        End If
    End Sub

    Private Sub txt_amount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_amount.KeyPress
        'Numbers Only
        If Not Char.IsNumber(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields Is Numbers Only!", vbExclamation)
        End If
    End Sub

    Private Sub txt_cname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_cname.KeyPress
        'Letters Only
        If Not Char.IsLetter(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields Is Letters Only!", vbExclamation)
        End If
    End Sub

    Public Sub sumtotal()
        Dim colsum As Decimal = 0
        For Each row As DataGridViewRow In dgv_stockList.Rows
            If Not row.IsNewRow AndAlso IsNumeric(row.Cells(6).Value) Then
                colsum += Convert.ToDecimal(row.Cells(6).Value)
            End If
        Next
        Try
            lbl_total.Text = colsum.ToString("#,##0.00")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btn_remove_Click(sender As Object, e As EventArgs) Handles btn_remove.Click
        Try
            If dgv_stockList.SelectedCells.Count > 0 Then
                Dim selectedCellIndex As Integer = dgv_stockList.SelectedCells(0).RowIndex
                Dim confirmationMessage As String = "Are you sure you want to remove this item?"
                If MessageBox.Show(confirmationMessage, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    dgv_stockList.Rows.RemoveAt(selectedCellIndex)

                    ' Recalculate the total amount
                    Dim totalAmount As Double = 0
                    For Each row As DataGridViewRow In dgv_stockList.Rows
                        Dim quantity As Integer = CInt(row.Cells(4).Value)
                        Dim itemPrice As Double = CDbl(row.Cells(3).Value)
                        totalAmount += quantity * itemPrice
                    Next
                    lbl_total.Text = totalAmount.ToString()
                    txt_searchPro.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox("An error occurred: " & ex.Message)
        End Try
    End Sub

    Private Sub txt_amount_TextChanged(sender As Object, e As EventArgs) Handles txt_amount.TextChanged
        If txt_amount.Text = Nothing Then
            lbl_change.Text = "00.00"
        Else
            lbl_change.Text = txt_amount.Text - lbl_total.Text
        End If
    End Sub

    '===================================================================================
    '===================================================================================
    'MANAGE RETURN

    Private Sub PopulateReturnComboBox()
        Dim connection As String = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        Dim connstring As String = connection

        Dim query As String = "SELECT DISTINCT proname FROM tbl_product"

        Using myconnection As New SqlConnection(connstring)
            Try
                Using command As New SqlCommand(query, myconnection)
                    myconnection.Open()

                    Using reader As SqlDataReader = command.ExecuteReader()
                        cbo_returnpro.Items.Clear()

                        If reader.HasRows Then
                            While reader.Read()
                                cbo_returnpro.Items.Add(reader("proname").ToString())
                            End While
                        End If
                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using
    End Sub

    Private Sub btn_saveReturn_Click(sender As Object, e As EventArgs) Handles btn_saveReturn.Click
        connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        connstring = connection
        myconnection.ConnectionString = connstring
        If txt_creturn.Text = String.Empty Then
            MsgBox("Please Fill Up Name !", vbExclamation)
            txt_creturn.Focus()
            Return
        ElseIf txt_conreturn.Text = String.Empty Then
            MsgBox("Please Fill Up Contact !", vbExclamation)
            txt_conreturn.Focus()
            Return
        ElseIf cbo_returnpro.Text = String.Empty Then
            MsgBox("Please Fill Up Product Name !", vbExclamation)
            cbo_returnpro.Focus()
            Return
        ElseIf txt_qtyreturn.Text = String.Empty Then
            MsgBox("Please Fill Up Quantity !", vbExclamation)
            txt_qtyreturn.Focus()
            Return
        ElseIf txt_issue.Text = String.Empty Then
            MsgBox("Please Fill Up Issue !", vbExclamation)
            txt_issue.Focus()
            Return
        ElseIf cbo_action.Text = String.Empty Then
            MsgBox("Please Choose What Action Will be Taken !", vbExclamation)
            cbo_action.Focus()
            Return
        Else
            Try
                Dim dt As New DataTable
                Dim da As New SqlDataAdapter
                Dim currentDate As String = DateTime.Now.Date.ToString("yyyy/MM/dd")

                myconnection.Open()
                Dim query1 As String
                query1 = "Insert into tbl_return ([RDate], [Cname], [Contact], [Pname], [Qty], [Issue], [Acttaken]) values (@RDate, @Cname, @Contact, @Pname, @Qty, @Issue, @Acttaken)"
                Dim cmd As SqlCommand = New SqlCommand(query1, myconnection)

                cmd.Parameters.AddWithValue("@RDate", currentDate)
                cmd.Parameters.AddWithValue("@Cname", txt_creturn.Text)
                cmd.Parameters.AddWithValue("@Contact", txt_conreturn.Text)
                cmd.Parameters.AddWithValue("@Pname", cbo_returnpro.Text)
                cmd.Parameters.AddWithValue("@Qty", txt_qtyreturn.Text)
                cmd.Parameters.AddWithValue("@Issue", txt_issue.Text)
                cmd.Parameters.AddWithValue("@Acttaken", cbo_action.Text)
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                MsgBox("Product Return Success !", vbInformation, "Success")
                txt_creturn.Focus()
                clearReturn()
                viewdataReturn()
                viewdataReportRet()
                myconnection.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End If
    End Sub

    Sub clearReturn()
        txt_creturn.Clear()
        txt_conreturn.Clear()
        txt_qtyreturn.Clear()
        txt_issue.Clear()
        dtp_return.Value = DateTime.Now
        cbo_action.SelectedIndex = -1
        cbo_returnpro.SelectedIndex = -1
    End Sub

    Private Sub btn_clearre_Click(sender As Object, e As EventArgs) Handles btn_clearre.Click
        clearReturn()
    End Sub

    Sub viewdataReturn()
        Try
            Dim con1 As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "Select RDate, Cname, Contact, Pname, Qty, Issue, Acttaken FROM tbl_return"
            Dim adapter As New SqlDataAdapter(sql, con1)
            Dim data As New DataTable("tbl_return")
            adapter.Fill(data)
            dgv_returnList.DataSource = data
            Dim sql1 As String
            sql1 = "Select * FROM tbl_return"
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

    Private Sub txt_searchproname_TextChanged(sender As Object, e As EventArgs) Handles txt_searchproname.TextChanged
        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim query As String
            query = "Select RDate, Cname, Contact, Pname, Qty, Issue, Acttaken from tbl_return Where [RDate] like '%" & txt_searchproname.Text & "%' or [Cname] like '%" & txt_searchproname.Text & "%'  or [Pname] like '%" & txt_searchproname.Text & "%' or [Acttaken] like '%" & txt_searchproname.Text & "%'"
            Dim dt As New DataTable
            Dim command As SqlCommand = New SqlCommand(query, myconnection)
            Dim da As New SqlDataAdapter
            da.SelectCommand = command
            da.Fill(dt)
            dgv_returnList.DataSource = dt
            myconnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    Private Sub txt_conreturn_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_conreturn.KeyPress
        'Numbers Only
        If Not Char.IsNumber(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields Is Numbers Only!", vbExclamation)
        End If
    End Sub

    Private Sub txt_creturn_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_creturn.KeyPress
        'Letters Only
        If Not Char.IsLetter(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields is Letters Only!", vbExclamation)
        End If
    End Sub

    Private Sub txt_qtyreturn_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_qtyreturn.KeyPress
        'Numbers Only
        If Not Char.IsNumber(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields Is Numbers Only!", vbExclamation)
        End If
    End Sub

    Private Sub txt_issue_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_issue.KeyPress
        'Letters Only
        If Not Char.IsLetter(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) Then

            e.Handled = True
            MsgBox("This Fields is Letters Only!", vbExclamation)
        End If
    End Sub

    '===================================================================================
    '===================================================================================
    'SALES REPORT

    Sub viewdataReport()
        Try
            Dim con1 As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "Select billno, billdate, cname, contact, proname, progroup, price, qty, total, ototal, amount FROM tbl_stock"
            Dim adapter As New SqlDataAdapter(sql, con1)
            Dim data As New DataTable("tbl_stock")
            adapter.Fill(data)
            dgv_reportList.DataSource = data
            Dim sql1 As String
            sql1 = "Select * FROM tbl_stock"
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

    Sub clearReport()
        txt_searchMonth.Clear()
    End Sub

    Private Sub txt_searchMonth_TextChanged_1(sender As Object, e As EventArgs) Handles txt_searchMonth.TextChanged

        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim query As String
            query = "Select billno, billdate, cname, contact, proname, progroup, price, qty, total, ototal, amount from tbl_stock Where [billno] like '%" & txt_searchMonth.Text & "%' or [proname] like '%" & txt_searchMonth.Text & "%' or [cname] like '%" & txt_searchMonth.Text & "%'  or [progroup] like '%" & txt_searchMonth.Text & "%' "
            Dim dt As New DataTable
            Dim command As SqlCommand = New SqlCommand(query, myconnection)
            Dim da As New SqlDataAdapter
            da.SelectCommand = command
            da.Fill(dt)
            dgv_reportList.DataSource = dt
            myconnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    '===================================================================================
    '===================================================================================
    'INVENTORY REPORT

    Sub viewdataReportInv()
        Try
            Dim con1 As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "Select supname, prodate, proname, progroup, stockno, price FROM tbl_product"
            Dim adapter As New SqlDataAdapter(sql, con1)
            Dim data As New DataTable("tbl_product")
            adapter.Fill(data)
            dgv_inventoryList.DataSource = data
            Dim sql1 As String
            sql1 = "Select * FROM tbl_product"
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

    Sub clearReportInv()
        txt_searchProMonth.Clear()
    End Sub

    Private Sub txt_searchProMonth_TextChanged(sender As Object, e As EventArgs) Handles txt_searchProMonth.TextChanged
        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim query As String
            query = "Select supname, prodate, proname, progroup, stockno, price from tbl_product Where [supname] like '%" & txt_searchProMonth.Text & "%' or [progroup] like '%" & txt_searchProMonth.Text & "%' or [proname] like '%" & txt_searchProMonth.Text & "%'"
            Dim dt As New DataTable
            Dim command As SqlCommand = New SqlCommand(query, myconnection)
            Dim da As New SqlDataAdapter
            da.SelectCommand = command
            da.Fill(dt)
            dgv_inventoryList.DataSource = dt
            myconnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    '===================================================================================
    '===================================================================================
    'RETURN REPORT

    Sub viewdataReportRet()
        Try
            Dim con1 As New SqlConnection("Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;")
            Dim sql As String
            sql = "Select RDate, Cname, Contact, Pname, Qty, Issue, Acttaken FROM tbl_return"
            Dim adapter As New SqlDataAdapter(sql, con1)
            Dim data As New DataTable("tbl_return")
            adapter.Fill(data)
            dgv_returnListrep.DataSource = data
            Dim sql1 As String
            sql1 = "Select * FROM tbl_return"
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

    Sub clearReportRet()
        txt_searchProMonth.Clear()
    End Sub

    Private Sub txt_searchret_TextChanged(sender As Object, e As EventArgs) Handles txt_searchret.TextChanged
        Try
            connection = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
            connstring = connection
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim query As String
            query = "Select RDate, Cname, Contact, Pname, Qty, Issue, Acttaken from tbl_return Where [RDate] like '%" & txt_searchret.Text & "%' or [Cname] like '%" & txt_searchret.Text & "%' or [Pname] like '%" & txt_searchret.Text & "%' or [Acttaken] like '%" & txt_searchret.Text & "%'"
            Dim dt As New DataTable
            Dim command As SqlCommand = New SqlCommand(query, myconnection)
            Dim da As New SqlDataAdapter
            da.SelectCommand = command
            da.Fill(dt)
            dgv_returnListrep.DataSource = dt
            myconnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub

    '===================================================================================
    '===================================================================================
    'SALES FILTERS

    Sub Sales()
        Dim connection As String = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        Dim query As String = "SELECT billno, billdate, cname, contact, proname, progroup, price, qty, total, ototal, amount FROM tbl_stock WHERE billdate >= @startdate AND billdate <= @enddate ORDER BY billdate ASC"

        Using myconnection As New SqlConnection(connection)
            Try
                Dim startdate As String = dtp_start.Value.ToString("yyyy/MM/dd")
                Dim enddate As String = dtp_end.Value.ToString("yyyy/MM/dd")

                Dim adapter As New SqlDataAdapter(query, myconnection)
                adapter.SelectCommand.Parameters.AddWithValue("@startdate", startdate)
                adapter.SelectCommand.Parameters.AddWithValue("@enddate", enddate)

                Dim data As New DataTable()
                adapter.Fill(data)
                dgv_reportList.DataSource = data
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using
    End Sub

    Private Sub btn_searchDate_Click(sender As Object, e As EventArgs) Handles btn_searchDate.Click
        Sales()
    End Sub

    Private Sub btn_viewall_Click(sender As Object, e As EventArgs) Handles btn_viewall.Click
        Try
            dtp_start.Value = DateTime.Now
            dtp_end.Value = DateTime.Now
            viewdataReport()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '===================================================================================
    '===================================================================================
    'INVENTORY FILTERS

    Sub Inventory()
        Dim connection As String = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        Dim query As String = "SELECT supname, prodate, proname, progroup, stockno, price FROM tbl_product WHERE prodate >= @startdate AND prodate <= @enddate ORDER BY prodate ASC"

        Using myconnection As New SqlConnection(connection)
            Try
                Dim startdate As String = dtp_startInv.Value.ToString("yyyy/MM/dd")
                Dim enddate As String = dtp_endInv.Value.ToString("yyyy/MM/dd")

                Dim adapter As New SqlDataAdapter(query, myconnection)
                adapter.SelectCommand.Parameters.AddWithValue("@startdate", startdate)
                adapter.SelectCommand.Parameters.AddWithValue("@enddate", enddate)

                Dim data As New DataTable()
                adapter.Fill(data)
                dgv_inventoryList.DataSource = data
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using
    End Sub

    Private Sub btn_searchInv_Click(sender As Object, e As EventArgs) Handles btn_searchInv.Click
        Inventory()
    End Sub

    Private Sub btn_all_Click(sender As Object, e As EventArgs) Handles btn_all.Click
        Try
            dtp_startInv.Value = DateTime.Now
            dtp_endInv.Value = DateTime.Now
            viewdataReportInv()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '===================================================================================
    '===================================================================================
    'RETURN FILTERS

    Sub ReturnRep()
        Dim connection As String = "Server=DESKTOP-B55CCOO\SQLEXPRESS01;Database=SalesAndInventory;Trusted_Connection=True;"
        Dim query As String = "SELECT RDate, Cname, Contact, Pname, Qty, Issue, Acttaken FROM tbl_return WHERE RDate >= @startdate AND RDate <= @enddate ORDER BY RDate ASC"

        Using myconnection As New SqlConnection(connection)
            Try
                Dim startdate As String = dtp_startRet.Value.ToString("yyyy/MM/dd")
                Dim enddate As String = dtp_endRet.Value.ToString("yyyy/MM/dd")

                Dim adapter As New SqlDataAdapter(query, myconnection)
                adapter.SelectCommand.Parameters.AddWithValue("@startdate", startdate)
                adapter.SelectCommand.Parameters.AddWithValue("@enddate", enddate)

                Dim data As New DataTable()
                adapter.Fill(data)
                dgv_returnListrep.DataSource = data
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using
    End Sub

    Private Sub btn_searchRet_Click(sender As Object, e As EventArgs) Handles btn_searchRet.Click
        ReturnRep()
    End Sub

    Private Sub btn_allRet_Click(sender As Object, e As EventArgs) Handles btn_allRet.Click
        Try
            dtp_startRet.Value = DateTime.Now
            dtp_endRet.Value = DateTime.Now
            viewdataReportRet()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '===================================================================================
    '===================================================================================
    'SALES PRINT

    Private printDocument As New Printing.PrintDocument()

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim font As New Font("Times New Roman", 12, FontStyle.Regular)
        Dim titleFont As New Font("Times New Roman", 16, FontStyle.Bold)
        Dim startX As Integer = 30
        Dim startY As Integer = 100
        Dim rowHeight As Integer = 20
        Dim columnWidth As Integer = 100
        Dim columnSpacing As Integer = 15

        Dim title As String = "Sales Report"

        Dim currentDate As String = "Date Printed: " & DateTime.Now.ToString("MMMM dd, yyyy HH:mm")

        e.Graphics.DrawString(title, titleFont, Brushes.Black, New PointF(startX, startY))
        startY += titleFont.Height + 10

        e.Graphics.DrawString(currentDate, font, Brushes.Black, New PointF(startX, startY))
        startY += font.Height + 20

        For col As Integer = 0 To dgv_reportList.Columns.Count - 1
            Dim xPosition As Integer = startX + (col * (columnWidth + columnSpacing))
            e.Graphics.DrawString(dgv_reportList.Columns(col).HeaderText, font, Brushes.Black, xPosition, startY)
        Next

        startY += rowHeight

        Dim rowsAbove As Integer = Math.Min(6, dgv_reportList.Rows.Count)
        For row As Integer = 0 To rowsAbove - 1
            For col As Integer = 0 To dgv_reportList.Columns.Count - 1
                Dim xPosition As Integer = startX + (col * (columnWidth + columnSpacing))
                Dim cellValue As String = dgv_reportList.Rows(row).Cells(col).Value.ToString()
                e.Graphics.DrawString(cellValue, font, Brushes.Black, xPosition, startY + (row * rowHeight))
            Next
        Next

        startY += (rowsAbove * rowHeight)

    End Sub

    Private Sub btn_salesprint_Click(sender As Object, e As EventArgs) Handles btn_salesprint.Click
        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub btn_salespreview_Click(sender As Object, e As EventArgs) Handles btn_salespreview.Click
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.StartPosition = FormStartPosition.CenterScreen
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub btn_salesSetup_Click(sender As Object, e As EventArgs) Handles btn_salesSetup.Click
        PageSetupDialog1.Document = PrintDocument1
        If PageSetupDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.DefaultPageSettings = PageSetupDialog1.PageSettings
        End If
    End Sub

    '===================================================================================
    '===================================================================================
    'INVENTORY PRINT

    Private Sub btn_inprint_Click(sender As Object, e As EventArgs) Handles btn_inprint.Click
        PrintDialog1.Document = PrintDocument2
        If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument2.Print()
        End If
    End Sub

    Private Sub btn_invview_Click(sender As Object, e As EventArgs) Handles btn_invview.Click
        PrintPreviewDialog1.Document = PrintDocument2
        PrintPreviewDialog1.StartPosition = FormStartPosition.CenterScreen
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub btn_invset_Click(sender As Object, e As EventArgs) Handles btn_invset.Click
        PageSetupDialog1.Document = PrintDocument2
        If PageSetupDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.DefaultPageSettings = PageSetupDialog1.PageSettings
        End If
    End Sub

    Private Sub PrintDocument2_PrintPage_1(sender As Object, e As PrintPageEventArgs) Handles PrintDocument2.PrintPage

        Dim font As New Font("Times New Roman", 12, FontStyle.Regular)
        Dim titleFont As New Font("Times New Roman", 16, FontStyle.Bold)
        Dim startX As Integer = 30
        Dim startY As Integer = 100
        Dim rowHeight As Integer = 20
        Dim columnWidth As Integer = 100
        Dim columnSpacing As Integer = 15

        Dim title As String = "Inventory Report"

        Dim currentDate As String = "Date Printed: " & DateTime.Now.ToString("MMMM dd, yyyy HH:mm")

        e.Graphics.DrawString(title, titleFont, Brushes.Black, New PointF(startX, startY))
        startY += titleFont.Height + 10

        e.Graphics.DrawString(currentDate, font, Brushes.Black, New PointF(startX, startY))
        startY += font.Height + 20

        For col As Integer = 0 To dgv_inventoryList.Columns.Count - 1
            Dim xPosition As Integer = startX + (col * (columnWidth + columnSpacing))
            e.Graphics.DrawString(dgv_inventoryList.Columns(col).HeaderText, font, Brushes.Black, xPosition, startY)
        Next

        startY += rowHeight

        Dim rowsAbove As Integer = Math.Min(6, dgv_inventoryList.Rows.Count)
        For row As Integer = 0 To rowsAbove - 1
            For col As Integer = 0 To dgv_inventoryList.Columns.Count - 1
                Dim xPosition As Integer = startX + (col * (columnWidth + columnSpacing))
                Dim cellValue As String = dgv_inventoryList.Rows(row).Cells(col).Value.ToString()
                e.Graphics.DrawString(cellValue, font, Brushes.Black, xPosition, startY + (row * rowHeight))
            Next
        Next

        startY += (rowsAbove * rowHeight)
    End Sub

    '===================================================================================
    '===================================================================================
    'RETURN PRINT

    Private Sub btn_viewRet_Click(sender As Object, e As EventArgs) Handles btn_viewRet.Click
        PrintPreviewDialog1.Document = PrintDocument3
        PrintPreviewDialog1.StartPosition = FormStartPosition.CenterScreen
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub btn_Retset_Click(sender As Object, e As EventArgs) Handles btn_Retset.Click
        PageSetupDialog1.Document = PrintDocument3
        If PageSetupDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.DefaultPageSettings = PageSetupDialog1.PageSettings
        End If
    End Sub

    Private Sub btn_printRet_Click(sender As Object, e As EventArgs) Handles btn_printRet.Click
        PrintDialog1.Document = PrintDocument3
        If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument3.Print()
        End If
    End Sub

    Private Sub PrintDocument3_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument3.PrintPage

        Dim font As New Font("Times New Roman", 12, FontStyle.Regular)
        Dim titleFont As New Font("Times New Roman", 16, FontStyle.Bold)
        Dim startX As Integer = 30
        Dim startY As Integer = 100
        Dim rowHeight As Integer = 20
        Dim columnWidth As Integer = 100
        Dim columnSpacing As Integer = 15

        Dim title As String = "Return Report"

        Dim currentDate As String = "Date Printed: " & DateTime.Now.ToString("MMMM dd, yyyy HH:mm")

        e.Graphics.DrawString(title, titleFont, Brushes.Black, New PointF(startX, startY))
        startY += titleFont.Height + 10

        e.Graphics.DrawString(currentDate, font, Brushes.Black, New PointF(startX, startY))
        startY += font.Height + 20

        For col As Integer = 0 To dgv_returnListrep.Columns.Count - 1
            Dim xPosition As Integer = startX + (col * (columnWidth + columnSpacing))
            e.Graphics.DrawString(dgv_returnListrep.Columns(col).HeaderText, font, Brushes.Black, xPosition, startY)
        Next

        startY += rowHeight

        Dim rowsAbove As Integer = Math.Min(6, dgv_returnListrep.Rows.Count)
        For row As Integer = 0 To rowsAbove - 1
            For col As Integer = 0 To dgv_returnListrep.Columns.Count - 1
                Dim xPosition As Integer = startX + (col * (columnWidth + columnSpacing))
                Dim cellValue As String = dgv_returnListrep.Rows(row).Cells(col).Value.ToString()
                e.Graphics.DrawString(cellValue, font, Brushes.Black, xPosition, startY + (row * rowHeight))
            Next
        Next

        startY += (rowsAbove * rowHeight)
    End Sub

    Private Sub txt_supplierNumber_TextChanged(sender As Object, e As EventArgs) Handles txt_supplierNumber.TextChanged

    End Sub
End Class