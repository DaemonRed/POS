Imports MySql.Data.MySqlClient
Imports System.IO

Module Methods

    Public conStr As String = Nothing
    Public con As MySqlConnection = Nothing
    Dim rdr As MySqlDataReader = Nothing
    Dim cmd As MySqlCommand = New MySqlCommand()

    Dim query As String = Nothing


    'Adds a variable number of items to a specified Database
    Function addItems(ByVal ParamArray items As String()) As Boolean
        For i As Integer = 0 To items.Length
            If CheckDuplicateByName("ID", "itemsTable", items(i)) Then
                OpenDB(con)
                query = "INSERT INTO itemsTable VALUES (@itemName)"
                cmd.CommandText = query
                cmd.Connection = con

                With cmd
                    .Parameters.AddWithValue("itemName", items(i))
                End With
                CloseDB(con)
            End If
        Next
        Return True
    End Function

    Public Function updateNameByID(ByVal ID As Integer, ByVal newName As String) As Boolean
        query = "UPDATE items SET name = @newName WHERE itemid = @itemId"
        Try
            OpenDB(con)
            With cmd
                .Connection = con
                .CommandText = query
                .Parameters.Clear()
                .Parameters.AddWithValue("@newName", newName)
                .Parameters.AddWithValue("@itemId", ID)
            End With
            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            MsgBox("Error updating item name" + vbNewLine + ex.Message())
            Return False
        End Try
    End Function

    Public Function updatePriceByID(ByVal ID As Integer, ByVal newPrice As Decimal) As Boolean
        query = "UPDATE items SET selling_price = @newPrice WHERE itemid = @itemId"
        Try
            OpenDB(con)

            With cmd
                .Connection = con
                .CommandText = query
                .Parameters.Clear()
                .Parameters.AddWithValue("@newPrice", newPrice)
                .Parameters.AddWithValue("@itemId", ID)
            End With

            cmd.ExecuteNonQuery()
            Return True

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        CloseDB(con)

    End Function

    Public Function updateQuantityByID(ByVal ID As Integer, ByVal newQuantity As Integer) As Boolean
        query = "UPDATE items SET quantity = @newQuantity WHERE itemid = @itemId"
        Try
            OpenDB(con)

            With cmd
                .Connection = con
                .CommandText = query
                .Parameters.Clear()
                .Parameters.AddWithValue("@newPrice", newQuantity)
                .Parameters.AddWithValue("@itemId", ID)
            End With

            cmd.ExecuteNonQuery()
            Return True

        Catch ex As Exception
            MsgBox("Error updating item quantity " + vbNewLine + ex.Message)
            Return False
        End Try

        CloseDB(con)
    End Function

    Public Function getPriceByID(ByVal ID As Integer) As Decimal
        Dim price As Decimal
        query = "SELECT price FROM items WHERE itemid = @itemId"

        Try
            OpenDB(con)
            With cmd
                .CommandText = query
                .Parameters.Clear()
                .Parameters.AddWithValue("@itemId", ID)
            End With

            price = cmd.ExecuteScalar()

            CloseDB(con)
        Catch ex As Exception
            CloseDB(con)
        Finally
        End Try

        Return price
    End Function

    Public Function getNameByID(ByVal ID As Integer) As String
        Dim name As String = Nothing
        query = "SELECT name FROM items WHERE itemid = @itemId"

        Try
            OpenDB(con)
            With cmd
                .CommandText = query
                .Parameters.Clear()
                .Parameters.AddWithValue("@itemId", ID)
            End With

            name = cmd.ExecuteScalar()

            CloseDB(con)
        Catch ex As Exception
            CloseDB(con)
        Finally
        End Try

        Return name
    End Function

    Public Function getQuantityByID(ByVal ID As Integer) As Integer
        Dim Qty As Integer = Nothing

        query = "SELECT quantity FROM items WHERE itemid = @itemId"

        Try
            OpenDB(con)
            With cmd
                .CommandText = query
                .Parameters.Clear()
                .Parameters.AddWithValue("@itemId", ID)
            End With

            Qty = cmd.ExecuteScalar()

            CloseDB(con)
        Catch ex As Exception
            CloseDB(con)
        Finally
        End Try

        Return Qty
    End Function

    Function getRecords(commandText As MySqlCommand) As DataTable
        Dim D As New MySqlDataAdapter
        D.SelectCommand = commandText

        Dim T As New DataTable

        Try
            D.Fill(T)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return T

    End Function

    Public Sub ValidateKeyDouble(e As KeyPressEventArgs)
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." Then
            e.KeyChar = Nothing
        End If
    End Sub

    Public Function OpenDB(ByRef con As MySqlConnection) As Boolean
        Try
            If con.State <> ConnectionState.Open Then
                con.Open()
                Return True
            End If
        Catch ex As MySqlException
            MsgBox("Error connecting to server")
            Return False
        End Try

        Return False
    End Function

    Public Function CloseDB(ByRef con As MySqlConnection) As Boolean
        Try
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        Return True
    End Function

    Sub FitTableData(ByRef grid As DataGridView)
        Dim columncount As Integer = grid.Columns.Count()
        Dim i As Integer = 0

        While i < columncount - 1
            grid.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            i = i + 1
        End While
        grid.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
    End Sub

    Function CheckDuplicateByName(ByVal fieldName As String, ByVal tableName As String, ByVal arg As String) As Boolean
        Try
            CloseDB(con)
            query = "SELECT COUNT(" + fieldName + " FROM " + tableName + " WHERE name = '" + arg.ToUpper + "')"

            cmd.CommandText = query
            cmd.Connection = con

            Dim x As Short = cmd.ExecuteScalar()

            If x > 0 Then
                CheckDuplicateByName = False
            Else
                CheckDuplicateByName = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        CloseDB(con)

        Return CheckDuplicateByName
    End Function

End Module


