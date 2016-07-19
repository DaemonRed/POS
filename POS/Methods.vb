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


