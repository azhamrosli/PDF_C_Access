Imports System.IO
Imports System.Data.OleDb

Module LoadCBOValue

    Public Sub LoadCBO(ByVal cbo As ComboBox, ByVal strSQL As String, ByVal strTextField As String, ByVal strValueField As String)

        Dim li As MyListItem
        Dim dr As OleDbDataReader

        Try
            dr = DataHandler.GetDataReader(strSQL, Conn)
            cbo.Items.Clear()
            Do While dr.Read
                li = New MyListItem
                li.Text = dr(strTextField)
                li.Value = dr(strValueField)
                cbo.Items.Add(li)
            Loop
        Catch exp As Exception
            MsgBox(exp.Message, MsgBoxStyle.Critical)
            Exit Sub
        Finally

            If Not dr Is Nothing Then
                dr.Close()
            End If
        End Try
    End Sub

End Module