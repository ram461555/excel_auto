Imports System.Data.OleDb

Module Module1

    Sub Main()
        Dim dta As OleDbDataAdapter

        Dim dts As DataSet
        Dim excel As String


        Console.WriteLine("Enter file path")
        excel = Console.ReadLine()

        Dim conn As OleDbConnection
        conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excel + ";Extended Properties='Excel 12.0 Xml;HDR=YES'")
        conn.Open()
        dta = New OleDbDataAdapter("Select * From [Sheet1$]", conn)
        dts = New DataSet
        dta.Fill(dts, "[Sheet1$]")

        Dim xlrow As String = "", xleachrow As DataColumnCollection, flag As Boolean, rowIndex As Int32
        flag = True


        xleachrow = dts.Tables(0).Columns
        Dim list As New List(Of String)
        'C:\Users\Rameshwar\Desktop\project.xlsx
        For Each row As DataRow In dts.Tables(0).Rows
            If IsLineEmpty(row) Then
                If xlrow.Length > 0 Then
                    list.Add(xlrow)
                End If
                xlrow = ""
                Continue For
            End If
            'rowIndex = dts.Tables(0).Rows.IndexOf(row)
            'MsgBox(rowIndex)
            If row(2).ToString() = "" And IsNextLineEmpty(row, dts) Then
                xlrow += "R" + row(0).ToString() + " " + ChrW(34) + row(1).ToString() + ChrW(34) + "};" + vbCrLf
            ElseIf row(2).ToString() = "" Then
                xlrow += "R" + row(0).ToString() + " " + ChrW(34) + row(1).ToString() + ChrW(34) + "," + vbCrLf
            Else
                xlrow += row(0).ToString() + " " + ChrW(34) + row(1).ToString() + ChrW(34) + " " + row(2).ToString() + " [" + row(3).ToString() + "] " + vbCrLf
            End If

        Next
        list.Add(xlrow)

        For Each line As String In list
            Dim query As New OleDbCommand("Insert into [Sheet2$] values ('" & line & "')", conn)
            query.ExecuteNonQuery()

        Next
        conn.Close()
        Console.ReadKey()
    End Sub

    Function IsNextLineEmpty(row As DataRow, dts As DataSet) As Boolean
        Dim rowIndex As Int32, nextrow As String, table1
        rowIndex = dts.Tables(0).Rows.IndexOf(row) + 1
        Dim table = dts.Tables(0)
        If table.Rows.Count > 0 Then
            table1 = table.Rows(rowIndex)
            If table1.IsNull(0) Then
                Return False
            Else
                Return True
            End If
        End If
        'If table.Rows.Count > 0 Then
        '    Return False
        'End If
        'MsgBox(nextrow)

    End Function

    Function IsLineEmpty(row As DataRow) As Boolean
        For Each value In row.ItemArray
            If Not IsDBNull(value) Then
                Return False
            End If
        Next
        Return True
    End Function

End Module
