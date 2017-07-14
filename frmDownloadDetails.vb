Imports System.Data.OleDb
Public Class frmDownloadDetails

    Private Sub frmDownloadDetails_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub



    Private Sub frmDownloadDetails_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        RefreshList()
        RefreshControl()
        AuditorListRefresh()
        'PANYW 2009.1
        TaxAgentListRefresh()
        'PANYW 2009.1 END
        DirectorListRefresh()
        optMenuntut.Enabled = False
        optMenyerah.Enabled = False
        S60FDownload()

    End Sub

    Private Sub RefreshControl()

        txtName.Text = ""
        txtPosition.Text = ""
        txtIC.Text = ""
        'chkKeepRecord.Checked = False
        ' HS : C2008.7 : Set as default to true
        chkKeepRecord.Checked = True
        chkRKST.Checked = False
        optMenuntut.Checked = True

    End Sub
    Private Sub RefreshList()
        lstViewAuditor.Items.Clear()
        'PANYW 2009.1
        lstViewTaxAgent.Items.Clear()
        'PANYW 2009.1 END
        lstViewDirector.Items.Clear()
    End Sub
    Private Sub AuditorListRefresh()

        Dim li As ListViewItem
        Dim dr As OleDbDataReader
        Dim cSQL As String

        cSQL = "SELECT * FROM [AUDITOR_PROFILE]"
        dr = DataHandler.GetDataReader(cSQL, Conn)
        Do While dr.Read()
            li = lstViewAuditor.Items.Add(dr("AD_KEY"))
            With li
                .SubItems.Add(dr("AD_CO_NAME"))
                .SubItems.Add(dr("AD_ADD"))
            End With
        Loop
        dr.Close()
    End Sub

    'PANYW 2009.1
    Private Sub TaxAgentListRefresh()

        Dim li As New ListViewItem
        Dim dr As OleDbDataReader
        Dim intIndexCount As Integer = 0
        Dim cSQL As String

        cSQL = "SELECT [TA_KEY], [TA_CO_NAME], [TA_ADD_LINE1], [TA_ADD_LINE2], [TA_ADD_LINE3], [TA_DEFAULT] FROM [TAXA_PROFILE]"
        dr = DataHandler.GetDataReader(cSQL, Conn)
        Do While dr.Read()
            li = lstViewTaxAgent.Items.Add(dr("TA_KEY"))
            With li
                .SubItems.Add(dr("TA_CO_NAME"))
                .SubItems.Add(dr("TA_ADD_LINE1") + dr("TA_ADD_LINE2") + dr("TA_ADD_LINE3"))
            End With
            If dr("TA_DEFAULT") = "1" Then
                lstViewTaxAgent.Items(intIndexCount).Checked = True

            End If
            intIndexCount += 1
        Loop
        dr.Close()
    End Sub
    'PANYW 2009.1 END

    Private Sub DirectorListRefresh()

        Dim li As ListViewItem
        Dim dr As OleDbDataReader
        Dim cSQL As String

        cSQL = "SELECT * FROM [DIRECTORS_PROFILE] WHERE [DIR_REF_NO] ='" & frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value & "' " & _
         "AND [DIR_YA] = '" & frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value & "' ORDER BY [DIR_REF_NO],[DIR_YA],[DIR_NAME]"

        dr = DataHandler.GetDataReader(cSQL, Conn)
        Do While dr.Read()
            li = lstViewDirector.Items.Add(dr("DIR_KEY"))
            With li
                .SubItems.Add(dr("DIR_NAME"))
                .SubItems.Add(dr("DIR_IC"))
            End With
        Loop
        dr.Close()
    End Sub

    Private Sub lstViewDirector_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lstViewDirector.ItemCheck

        If lstViewDirector.CheckedItems.Count > 0 Then
            If e.CurrentValue = False Then
                lstViewDirector.CheckedItems(0).Checked = False
                lstViewDirector.Items(e.Index).Checked = True
            End If
        Else
            lstViewDirector.Items(e.Index).Selected = True
        End If

    End Sub

    Private Sub chkRKST_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkRKST.Click

        If chkRKST.Checked = True Then
            optMenuntut.Enabled = True
            optMenyerah.Enabled = True
        Else
            optMenuntut.Enabled = False
            optMenyerah.Enabled = False
        End If

    End Sub

    Private Sub lstViewAuditor_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lstViewAuditor.ItemCheck

        If lstViewAuditor.CheckedItems.Count > 0 Then
            If e.CurrentValue = False Then
                lstViewAuditor.CheckedItems(0).Checked = False
                lstViewAuditor.Items(e.Index).Checked = True
            End If
        Else
            lstViewAuditor.Items(e.Index).Selected = True
        End If
    End Sub

    'PANYW 2009.1
    Private Sub lstViewTaxAgent_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lstViewTaxAgent.ItemCheck
        If lstViewTaxAgent.CheckedItems.Count > 0 Then
            If e.CurrentValue = False Then
                lstViewTaxAgent.CheckedItems(0).Checked = False
                lstViewTaxAgent.Items(e.Index).Checked = True
            End If
        Else
            lstViewTaxAgent.Items(e.Index).Selected = True
        End If
    End Sub
    'PANYW 2009.1 END

    Private Sub lstViewDirector_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles lstViewDirector.ItemSelectionChanged
        If e.IsSelected = True Then
            lstViewDirector.Items(e.ItemIndex).Checked = True

            txtName.Text = lstViewDirector.SelectedItems(0).SubItems(1).Text
            txtIC.Text = lstViewDirector.SelectedItems(0).SubItems(2).Text
            txtPosition.Text = "Director"
        End If

    End Sub

    Private Sub lstViewAuditor_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles lstViewAuditor.ItemSelectionChanged
        If e.IsSelected = True Then
            lstViewAuditor.Items(e.ItemIndex).Checked = True
        End If

    End Sub

    Private Sub tsbtnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbtnPost.Click
        Dim dr As OleDbDataReader
        Dim cSQL As String
        'amy2008august
        cSQL = "SELECT [PL_S60F] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' AND [PL_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
        dr = DataHandler.GetDataReader(cSQL, Conn)
        If dr.Read() Then
            If IsDBNull(dr("PL_S60F")) = False Then
                If dr("PL_S60F") = "Y" Then
                    cSQL = "SELECT IH_INSTALLMENTS " _
                      & " FROM [INVESTMENT_HOLDING]" _
                      & " where IH_REF_NO = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' And IH_YA='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value) & "'"

                Else
                    cSQL = "SELECT TC_TP_INSTALL , TC_BUSINESS FROM Tax_Computation WHERE TC_REF_NO = '" & frmDownloadMainMenu.dgdDownload.CurrentRow.Cells(1).Value & "'  AND TC_YA = '" & frmDownloadMainMenu.dgdDownload.CurrentRow.Cells(2).Value & "'" _
                            & " Order By TC_BUSINESS"
                    '     MsgBox(cSQL)
                End If
                dr = DataHandler.GetDataReader(cSQL, Conn)
                If dr.Read() Then
                    If CDbl(dr(0)) = 0 Then
                        'If dr(0).ToString = "0" Or dr(0).ToString = "0.00" Then
                        If MsgBox("Is your Installment made figure equal to RM 0.00?", vbYesNo + vbInformation, "Taxcom") = vbYes Then
                            frmDownloadPost.ShowDialog()
                        Else
                            Exit Sub
                        End If
                    Else
                        frmDownloadPost.ShowDialog()
                    End If
                End If
            End If
        End If

    End Sub

    ' === csNgoh Julai 2008 === '
    Private Sub S60FDownload()
        Dim strSQL As String
        Dim dr As OleDbDataReader

        strSQL = "SELECT [PL_S60F] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' AND [PL_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
        dr = DataHandler.GetDataReader(strSQL, Conn)
        If dr.Read() Then
            If IsDBNull(dr("PL_S60F")) = False Then
                If dr("PL_S60F") = "Y" Then
                    optMenyerah.Visible = False
                Else
                    optMenyerah.Visible = True
                End If
            End If
        End If
        dr.Close()
    End Sub
    ' === end csNgoh Julai 2008 === '


End Class