Imports System.IO
Imports System.Data.OleDb

Public Class frmDownloadMainMenu

    Public Sub GridRefresh(ByVal OptionMode As String)
        Dim cSQL As String
        Dim prmOledb(1) As OleDbParameter
        prmOledb(0) = New OleDbParameter("@TC_VAL", Trim(String.Concat(txtSearchContent.Text, "%")))
        prmOledb(1) = New OleDbParameter("@IH_VAL", Trim(String.Concat(txtSearchContent.Text, "%")))
        cSQL = GetTCQuery() + " UNION " + GetIHQuery()
        DataHandler.RunSQLtoDGD(cSQL, dgdDownload, prmOledb)

    End Sub

    Private Sub ManualSearchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsddlManual.Click
        tsbtnSearch.Image = My.Resources.pngSearch
        tsbtnSearch.Text = "Manual Search"
        txtSearchContent.Text = ""
    End Sub

    Private Sub AutoSearchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsddlAuto.Click
        tsbtnSearch.Image = My.Resources.pngAuto
        tsbtnSearch.Text = "Auto Search"
        txtSearchContent.Text = ""
        If cboSearchCriteria.Text = "-All Record-" And tsbtnSearch.Text = "Auto Search" Then
            GridRefresh("[ALL]")
        End If
    End Sub

    Private Function GetTCQuery() As String
        Dim strSQL As String = ""
        Dim SearchByVal As String = ""

        strSQL = "Select TAX_COMPUTATION.TC_KEY as RECORDCOUNT, TAX_COMPUTATION.TC_REF_NO as RefNO,  TAX_COMPUTATION.TC_YA as YA, TAX_COMPUTATION.TC_CO_NAME as comName" & _
                " from TAX_COMPUTATION where TC_BUSINESS = 1 and TC_YA >= '2008'"

        If (GetOptionVal("TAX_COMPUTATION") <> "ALL") Then
            SearchByVal = "@TC_VAL"
        Else
            SearchByVal = "[ALL]"
        End If
        GetTCQuery = ChangeSQL(SearchByVal, GetOptionVal("TAX_COMPUTATION"), strSQL)
    End Function

    Private Function GetIHQuery() As String

        Dim strSQL As String = ""
        Dim SearchByVal As String = ""

        strSQL = "Select INVESTMENT_HOLDING.IH_KEY, INVESTMENT_HOLDING.IH_REF_NO, INVESTMENT_HOLDING.IH_YA, TAXP_PROFILE.TP_COM_NAME" & _
                " from [INVESTMENT_HOLDING] INNER JOIN [TAXP_PROFILE] ON INVESTMENT_HOLDING.IH_REF_NO=TAXP_PROFILE.TP_REF_NO where IH_YA >= '2008'"

        If (GetOptionVal("INVESTMENT_HOLDING") <> "ALL") Then
            SearchByVal = "@IH_VAL"
        Else
            SearchByVal = "[ALL]"
        End If
        GetIHQuery = ChangeSQL(SearchByVal, GetOptionVal("INVESTMENT_HOLDING"), strSQL)

    End Function

    Private Function ChangeSQL(ByVal SearchByVal As String, ByVal OptionVal As String, ByVal cSQL As String)

        If Trim(SearchByVal) <> "[ALL]" And Trim(OptionVal) <> "[ALL]" Then
            cSQL = cSQL + " and " & OptionVal & " like " & SearchByVal
        End If
        cSQL = cSQL + " Order By 1"
        ChangeSQL = cSQL
    End Function

    Private Function GetOptionVal(ByVal strTable As String)

        GetOptionVal = ""

        If strTable = "TAX_COMPUTATION" Then
            If cboSearchCriteria.Text = "-All Record-" Then
                GetOptionVal = "ALL"
            ElseIf cboSearchCriteria.Text = "C Reference No." Then
                GetOptionVal = "TC_REF_NO"
            ElseIf cboSearchCriteria.Text = "Year of Assessment" Then
                GetOptionVal = "TC_YA"
            ElseIf cboSearchCriteria.Text = "" Then
                GetOptionVal = "ALL"
            End If

        ElseIf strTable = "INVESTMENT_HOLDING" Then
            If cboSearchCriteria.Text = "-All Record-" Then
                GetOptionVal = "ALL"
            ElseIf cboSearchCriteria.Text = "C Reference No." Then
                GetOptionVal = "IH_REF_NO"
            ElseIf cboSearchCriteria.Text = "Year of Assessment" Then
                GetOptionVal = "IH_YA"
            ElseIf cboSearchCriteria.Text = "" Then
                GetOptionVal = "ALL"

            End If
        End If

    End Function

    Private Sub frmDownloadMainMenu_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim frm As New Form
        Init.PublicFunc.InitVar()
        'frm = New frmDownloadMainMenu
        ' === csNgoh C2008.5 === '
        cboSearchCriteria.Text = "-All Record-" 'GridRefresh("[ALL]")
        If _strRefNum <> "" Then
            cboSearchCriteria.Text = "C Reference No."
            txtSearchContent.Text = _strRefNum
        End If
        ' === end csNgoh C2008.5 === '
    End Sub

    Private Sub tsbtnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbtnPost.Click
        If Not dgdDownload.SelectedRows.Count > 0 Then
            MsgBox("No record is selected!", MsgBoxStyle.Information, "Caution")
            Exit Sub
        End If
        ' HS : 2009: C2008.7 : Add Borang selection
        If Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2008" Then
            BorangSelector.Borang = BorangSelector.BorangEnum.BorangC2008
            BorangSelector.Year = Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value())
            BorangSelector.RefNo = Trim(Me.dgdDownload.SelectedRows(0).Cells(1).Value)
        End If

        If Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2009" Then
            BorangSelector.Borang = BorangSelector.BorangEnum.BorangC2009
            BorangSelector.Year = Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value())
            BorangSelector.RefNo = Trim(Me.dgdDownload.SelectedRows(0).Cells(1).Value)
        End If
        '===NgKL C2010.1==='
        If Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2010" Then
            BorangSelector.Borang = BorangSelector.BorangEnum.BorangC2010
            BorangSelector.Year = Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value())
            BorangSelector.RefNo = Trim(Me.dgdDownload.SelectedRows(0).Cells(1).Value)
        End If
        '===NgKL C2010.1 End==='
        'weihong
        If Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2011" Then
            BorangSelector.Borang = BorangSelector.BorangEnum.BorangC2011
            BorangSelector.Year = Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value())
            BorangSelector.RefNo = Trim(Me.dgdDownload.SelectedRows(0).Cells(1).Value)
        End If
        'LEESH 04-APR-2012
        If Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2012" Then
            BorangSelector.Borang = BorangSelector.BorangEnum.BorangC2012
            BorangSelector.Year = Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value())
            BorangSelector.RefNo = Trim(Me.dgdDownload.SelectedRows(0).Cells(1).Value)
        End If
        'LEESH END
        If Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2013" Then
            BorangSelector.Borang = BorangSelector.BorangEnum.BorangC2013
            BorangSelector.Year = Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value())
            BorangSelector.RefNo = Trim(Me.dgdDownload.SelectedRows(0).Cells(1).Value)
        End If
        'simkh 2014
        If Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2014" Then
            BorangSelector.Borang = BorangSelector.BorangEnum.BorangC2014
            BorangSelector.Year = Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value())
            BorangSelector.RefNo = Trim(Me.dgdDownload.SelectedRows(0).Cells(1).Value)
        End If
        'simkh end

        'simkh 2015 su8.1
        If Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2015" Then
            BorangSelector.Borang = BorangSelector.BorangEnum.BorangC2015
            BorangSelector.Year = Trim(Me.dgdDownload.SelectedRows(0).Cells(2).Value())
            BorangSelector.RefNo = Trim(Me.dgdDownload.SelectedRows(0).Cells(1).Value)
        End If
        'simkh end
        frmDownloadDetails.ShowDialog()
    End Sub

    Private Sub tsbtnSearch_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbtnSearch.ButtonClick

        Dim strMode As String = ""
        If cboSearchCriteria.SelectedIndex <> 0 And txtSearchContent.Text = Nothing Then
            MsgBox("Please specify the search words!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
        ElseIf cboSearchCriteria.SelectedItem = Nothing Then
            MsgBox("Please select a search criteria!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
        Else

            If cboSearchCriteria.Text = "-All Record-" Then
                strMode = "[ALL]"
            ElseIf cboSearchCriteria.Text = "C Reference No." Then
                strMode = "[REFNO]"
            ElseIf cboSearchCriteria.Text = "Year of Assessment" Then
                strMode = "[YEAR]"
            End If
            GridRefresh(strMode)
        End If

    End Sub

    Private Sub txtSearchContent_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearchContent.TextChanged
        Dim strMode As String = ""
        If tsbtnSearch.Text = "Auto Search" Then

            If cboSearchCriteria.Text = "-All Record-" Then
                strMode = "[ALL]"
            ElseIf cboSearchCriteria.Text = "C Reference No." Then
                strMode = "[REFNO]"
            ElseIf cboSearchCriteria.Text = "Year of Assessment" Then
                strMode = "[YEAR]"
            End If
            If txtSearchContent.Text <> "" Then
                GridRefresh(strMode)
            End If
        ElseIf tsbtnSearch.Text = "Manual Search " Then
            'tsbtnSearch.PerformClick()
        End If
    End Sub

    Private Sub cboSearchCriteria_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSearchCriteria.SelectedIndexChanged
        txtSearchContent.Text = ""
        'If tsbtnSearch.Text = "Auto Search" Then
        'tsbtnSearch.Image = My.Resources.pngAuto
        If cboSearchCriteria.Text = "-All Record-" Then
            AutoSearchToolStripMenuItem_Click(sender, e)
            'GridRefresh("[ALL]")
        End If
    End Sub

    Private Sub dgdDownload_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgdDownload.CellDoubleClick
        'If Not dgdDownload.SelectedRows.Count > 0 Then
        '    MsgBox("No record is selected!", MsgBoxStyle.Information, "Caution")
        '    Exit Sub
        'End If
        'frmDownloadDetails.ShowDialog()
        ' HS : 2009 : C2008.7 : Simplified the function'
        tsbtnPost.PerformClick()
    End Sub

End Class
