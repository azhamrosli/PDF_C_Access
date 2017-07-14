Public Class frmDownloadPost

    Private Sub frmDownload_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.txtSaveFile.Text = ""
        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2008" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2008.pdf"
            'NGOHCS C2009.1 (SU11)
            lblRemark.Visible = False
        End If

        'HS : 2009 : C2008.7 : Add for 2009 Borang C path selection
        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2009" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2009.pdf"
            'NGOHCS C2009.1 (SU11)
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2009 and onwards, Form C PDF format is not for submission to LHDNM by Tax Agent"
        End If

        '===NgKL C2010.1==='
        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2010" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2010.pdf"
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2010 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If

        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) >= "2010" Then
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2010 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If
        '===NgKL C2010.1 End==='

        'weihong C2011.4'
        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2011" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2011.pdf"
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2010 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If

        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) >= "2011" Then
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2011 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If
        'endweihong C2011.4'

        'LEESH 04-APR-2012
        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2012" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2012.pdf"
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2012 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If

        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) >= "2012" Then
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2012 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If
        'LEESH END

        'LEESH 10-JULY-2013
        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2013" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2013.pdf"
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2013 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If

        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) >= "2013" Then
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2013 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If
        'LEESH END

        'simkh 2014
        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2014" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2014.pdf"
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2014 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If

        'If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) >= "2014" Then
        '    lblRemark.Visible = True
        '    lblRemark.Text = "Remark: YA 2014 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        'End If
        'simkh end

        'simkh 2015
        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2015" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2015.pdf"
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2015 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If
        'simkh end

        'azham 2016
        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2016" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2016.pdf"
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2016 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If
        'azham 2016


        'azham 2017

        If Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value()) = "2017" Then
            txtOpenFile.Text = Application.StartupPath & "\Template\BorangC_2017.pdf"
            lblRemark.Visible = True
            lblRemark.Text = "Remark: YA 2017 and onwards, Form C PDF format is not for submission to LHDNM by all Tax Payer(Including Tax Agent)."
        End If
        'azham 2017
    End Sub

    Private Sub tsbtnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbtnPost.Click
        If txtOpenFile.Text = "" Then
            MsgBox("Please select template!", MsgBoxStyle.Critical)
        ElseIf txtSaveFile.Text = "" Then
            MsgBox("Please select a location to save!", MsgBoxStyle.Critical)
        Else
            ConExport()
        End If
    End Sub

    Private Sub ConExport()
        Select Case BorangSelector.Borang
            Case BorangSelector.BorangEnum.BorangC2008
                PostBorangC2008()
            Case BorangSelector.BorangEnum.BorangC2009
                PostBorangC2009()
                '===NgKL C2010.1==='
            Case BorangSelector.BorangEnum.BorangC2010
                PostBorangC2010()
                '===NgKL C2010.1 End==='
                'weihong
            Case BorangSelector.BorangEnum.BorangC2011
                PostBorangC2011()
                'endweihong
                'LEESH 04-APR-2012
            Case BorangSelector.BorangEnum.BorangC2012
                PostBorangC2012()
                'LEESH END
            Case BorangSelector.BorangEnum.BorangC2013
                PostBorangC2013()
                'LEESH END
                'simkh 2014
            Case BorangSelector.BorangEnum.BorangC2014
                PostBorangC2014()
                'simkh end
                'simkh 2015 su8.1
            Case BorangSelector.BorangEnum.BorangC2015
                PostBorangC2015()
                'simkh end
            Case BorangSelector.BorangEnum.BorangC2016
                PostBorangC2016()
                'azham end
            Case BorangSelector.BorangEnum.BorangC2017
                PostBorangC2017()
                'azham end
        End Select
    End Sub

    Private Sub PostBorangC2008()
        ' Data Export to PDF Borang C 2008
        Dim BrgC2008 = New BorangC2008
        Try
            BrgC2008.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2008.CheckFieldEmpty()
            BrgC2008.Name()
            BrgC2008.Page1()
            BrgC2008.Page2()
            BrgC2008.Page3()
            BrgC2008.Page4()
            BrgC2008.Page5()
            BrgC2008.Page6()
            BrgC2008.Page7()
            BrgC2008.Page8()
            BrgC2008.Page9()
            BrgC2008.Page10()
            BrgC2008.Page11()
            BrgC2008.Page12()
            BrgC2008.Page13()
            BrgC2008.Slip() ' HS : C2008.7 , Slip need to fill in information

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub

    Private Sub PostBorangC2009()
        'Data Export to PDF Borang C 2009
        Dim BrgC2009 = New BorangC2009
        Try
            BrgC2009.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2009.CheckFieldEmpty()
            BrgC2009.Name()
            BrgC2009.Page1()
            BrgC2009.Page2()
            BrgC2009.Page3() ' Combination of Page3 and Page4
            BrgC2009.Page4()
            BrgC2009.Page5()
            BrgC2009.Page6()
            BrgC2009.Page7()
            BrgC2009.Page8() ' combination of Page8 and Page9
            BrgC2009.Page9()
            BrgC2009.Page10()
            BrgC2009.Page11()
            BrgC2009.Page12()
            BrgC2009.Page13()
            BrgC2009.Page14()
            BrgC2009.Page15()
            BrgC2009.Slip()

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub

    Private Sub PostBorangC2010()
        'Data Export to PDF Borang C 2010
        Dim BrgC2010 = New BorangC2010
        Try
            BrgC2010.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2010.CheckFieldEmpty()
            BrgC2010.Name()
            BrgC2010.Page1()
            BrgC2010.Page2()
            BrgC2010.Page3() ' Combination of Page3 and Page4
            BrgC2010.Page4()
            BrgC2010.Page5()
            BrgC2010.Page6()
            BrgC2010.Page7()
            BrgC2010.Page8() ' combination of Page8 and Page9
            BrgC2010.Page9()
            BrgC2010.Page10()
            BrgC2010.Page11()
            BrgC2010.Page12()
            BrgC2010.Page13()
            BrgC2010.Page14()
            BrgC2010.Page15()
            BrgC2010.Slip()

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub

    Private Sub PostBorangC2011()
        'Data Export to PDF Borang C 2011
        Dim BrgC2011 = New BorangC2011
        Try
            BrgC2011.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2011.CheckFieldEmpty()
            BrgC2011.Name()
            BrgC2011.Page1()
            BrgC2011.Page2()
            BrgC2011.Page3() ' Combination of Page3 and Page4
            BrgC2011.Page4()
            BrgC2011.Page5()
            BrgC2011.Page6()
            BrgC2011.Page7()
            BrgC2011.Page8() ' combination of Page8 and Page9
            BrgC2011.Page9()
            BrgC2011.Page10()
            BrgC2011.Page11()
            BrgC2011.Page12()
            BrgC2011.Page13()
            BrgC2011.Page14()
            BrgC2011.Page15()
            BrgC2011.Slip()

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub

    Private Sub PostBorangC2012()
        'Data Export to PDF Borang C 2011
        Dim BrgC2012 = New BorangC2012

        Try
            BrgC2012.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2012.CheckFieldEmpty()
            BrgC2012.Name()
            BrgC2012.Page1()
            BrgC2012.Page2()
            BrgC2012.Page3() ' Combination of Page3 and Page4
            BrgC2012.Page4()
            BrgC2012.Page5()
            BrgC2012.Page6()
            BrgC2012.Page7()
            BrgC2012.Page8() ' combination of Page8 and Page9
            BrgC2012.Page9()
            BrgC2012.Page10()
            BrgC2012.Page11()
            BrgC2012.Page12()
            BrgC2012.Page13()
            BrgC2012.Page14()
            BrgC2012.Page15()
            BrgC2012.Slip()

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub

    Private Sub PostBorangC2013()
        'Data Export to PDF Borang C 2011
        Dim BrgC2013 = New BorangC2013

        Try
            BrgC2013.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2013.CheckFieldEmpty()
            BrgC2013.Name()
            BrgC2013.Page1()
            BrgC2013.Page2()
            BrgC2013.Page3() ' Combination of Page3 and Page4
            BrgC2013.Page4()
            BrgC2013.Page5()
            BrgC2013.Page6()
            BrgC2013.Page7()
            BrgC2013.Page8() ' combination of Page8 and Page9
            BrgC2013.Page9()
            BrgC2013.Page10()
            BrgC2013.Page11()
            BrgC2013.Page12()
            BrgC2013.Page13()
            BrgC2013.Page14()
            BrgC2013.Page15()
            BrgC2013.Slip()

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub

    'simkh 2014
    Private Sub PostBorangC2014()
        'Data Export to PDF Borang C 2011
        Dim BrgC2014 = New BorangC2014

        Try
            BrgC2014.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2014.CheckFieldEmpty()
            BrgC2014.Name()
            BrgC2014.Page1()
            BrgC2014.Page2()
            BrgC2014.Page3() ' Combination of Page3 and Page4
            BrgC2014.Page4()
            BrgC2014.Page5()
            BrgC2014.Page6()
            BrgC2014.Page7()
            BrgC2014.Page8() ' combination of Page8 and Page9
            BrgC2014.Page9()
            BrgC2014.Page10()
            BrgC2014.Page11()
            BrgC2014.Page12()
            BrgC2014.Page13()
            BrgC2014.Page14()
            BrgC2014.Page15()
            BrgC2014.Slip()

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub
    'simkh end

    'simkh 2015 su8.1
    Private Sub PostBorangC2015()
        'Data Export to PDF Borang C 2011
        Dim BrgC2015 = New BorangC2015

        Try
            BrgC2015.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2015.CheckFieldEmpty()
            BrgC2015.Name()
            BrgC2015.Page1()
            BrgC2015.Page2()
            BrgC2015.Page3() ' Combination of Page3 and Page4
            BrgC2015.Page4()
            BrgC2015.Page5()
            BrgC2015.Page6()
            BrgC2015.Page7()
            BrgC2015.Page8() ' combination of Page8 and Page9
            BrgC2015.Page9()
            BrgC2015.Page10()
            BrgC2015.Page11()
            BrgC2015.Page12()
            BrgC2015.Page13()
            BrgC2015.Page14()
            BrgC2015.Page15()
            BrgC2015.Slip()

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub
    'simkh end
    'simkh 2015 su8.1
    Private Sub PostBorangC2016()
        'Data Export to PDF Borang C 2011
        Dim BrgC2016 = New BorangC2016

        Try
            BrgC2016.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2016.CheckFieldEmpty()
            BrgC2016.Name()
            BrgC2016.Page1()
            BrgC2016.Page2()
            BrgC2016.Page3() ' Combination of Page3 and Page4
            BrgC2016.Page4()
            BrgC2016.Page5()
            BrgC2016.Page6()
            BrgC2016.Page7()
            BrgC2016.Page8() ' combination of Page8 and Page9
            BrgC2016.Page9()
            BrgC2016.Page10()
            BrgC2016.Page11()
            BrgC2016.Page12()
            BrgC2016.Page13()
            BrgC2016.Page14()
            BrgC2016.Page15()
            BrgC2016.Slip()

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub
    'simkh end

    Private Sub PostBorangC2017()
        'Data Export to PDF Borang C 2011
        Dim BrgC2017 = New BorangC2017

        Try
            BrgC2017.InitBorang()
        Catch
            Exit Sub
        End Try
        Try
            lblProgress.Text = "Progress - Exporting"
            BrgC2017.CheckFieldEmpty()
            BrgC2017.Name()
            BrgC2017.Page1()
            BrgC2017.Page2()
            BrgC2017.Page3() ' Combination of Page3 and Page4
            BrgC2017.Page4()
            BrgC2017.Page5()
            BrgC2017.Page6()
            BrgC2017.Page7()
            BrgC2017.Page8() ' combination of Page8 and Page9
            BrgC2017.Page9()
            BrgC2017.Page10()
            BrgC2017.Page11()
            BrgC2017.Page12()
            BrgC2017.Page13()
            BrgC2017.Page14()
            BrgC2017.Page15()
            BrgC2017.Slip()

            lblProgress.Text = "Progress - Finish Export"
            openFile()
        Catch
            'MsgBox("Template is not correct!", MsgBoxStyle.Critical, "Caution!")
            MsgBox("Fail to export!", MsgBoxStyle.Critical, "Caution!")
            Exit Sub
        Finally
            Me.Dispose()
        End Try
    End Sub

    Private Sub btnSearchFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchFile.Click

        ' HS : C2008.7
        OpenFileDialog.InitialDirectory() = Application.StartupPath & "\Template"
        OpenFileDialog.Filter = "Adobe PDF File (*.pdf)|*.pdf|All File(*.*)|*.*"
        OpenFileDialog.FileName = "BorangC_" + BorangSelector.Year

        If Not OpenFileDialog.ShowDialog = System.Windows.Forms.DialogResult.Cancel Then
            txtOpenFile.Text = OpenFileDialog.FileName
        End If

    End Sub

    Private Sub btnSaveFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveFile.Click

        ' HS : C2008.7
        SaveFileDialog.InitialDirectory() = "C:\"
        SaveFileDialog.Filter = "Adobe PDF File (*.pdf)|*.pdf|All File(*.*)|*.*"
        SaveFileDialog.FileName = "BorangC_" + BorangSelector.RefNo + "_" + BorangSelector.Year

        If Not SaveFileDialog.ShowDialog = System.Windows.Forms.DialogResult.Cancel Then
            txtSaveFile.Text = SaveFileDialog.FileName
        End If

    End Sub

    Private Sub openFile()
        Dim proc As New Process()

        With proc.StartInfo
            .FileName = txtSaveFile.Text
            .UseShellExecute = True
            .WindowStyle = ProcessWindowStyle.Maximized
        End With

        proc.Start()
        proc.Close()
        proc.Dispose()
    End Sub

End Class