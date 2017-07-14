Imports iTextSharp.text.pdf
Imports System.Data.OleDb
Imports System.IO

Public Class BorangC2012

    'Dim Total As Double
    Private dr As OleDbDataReader
    Private dr2 As OleDbDataReader
    Private dr3 As OleDbDataReader
    Private dr4 As OleDbDataReader
    Private pdfTemplate As String '= frmDownloadPost.txtOpenFile.Text
    Private newFile As String '= frmDownloadPost.txtSaveFile.Text
    Private pdfReader As PdfReader
    Private pdfStamper As PdfStamper '(pdfReader, New FileStream( _
    '  newFile, FileMode.Create))
    Private pdfFormFields As AcroFields '= pdfStamper.AcroFields
    Private pdfFieldPath As String = "topmostSubform[0]."
    ' variable use for L1 - L8 calculation
    Private dSales As Double = 0
    Private dOS As Double = 0
    Private dDep As Double = 0
    Private dPur As Double = 0
    Private dA As Double = 0
    Private dNA As Double = 0
    Private dCS As Double = 0
    Private dGP As Double = 0
    Private BSCode As String = ""
    Private countSpace As Integer = 0
    Private i As Integer = 0

    'Dim BSCode As String = ""
    ' variable use to chooped the String

    Public strCropped As String = ""
    Public strRemainder As String = ""
    Dim strCropped1 As String, strCropped2 As String, strCropped3 As String, chkLength As Integer
    Private totalA25 As Double = 0
    Private totalA26 As Double = 0
    Private totalA27 As Double = 0
    Private totalA28 As Double = 0
    Private totalA29 As Double = 0
    Private totalA30 As Double = 0
    Private totalA31 As Double = 0
    Private totalA32 As Double = 0
    Private totalA34 As Double = 0
    Private totalA35 As Double = 0

    Public Sub initBorang()
        Try
            pdfTemplate = frmDownloadPost.txtOpenFile.Text.ToString()
            newFile = frmDownloadPost.txtSaveFile.Text.ToString()
            pdfReader = New PdfReader(pdfTemplate)
            pdfStamper = New PdfStamper(pdfReader, New FileStream( _
                newFile, FileMode.Create))
            'Dim pdfReader2 = New PdfReader(pdfTemplate)
            'Dim pdfStamper2 = New PdfStamper(pdfReader2, New FileStream( _
            '    newFile, FileMode.Create))
            'pdfFormFields = pdfStamper.AcroFields
            'pdfReader = pdfReader2
            'pdfStamper = pdfStamper2
            'pdfFormFields = pdfStamper2.AcroFields
            pdfFormFields = pdfStamper.AcroFields
        Catch ex As Exception
            MsgBox("Please select a correct template!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try

    End Sub

    Public Sub Name()
        Dim intIndex As Integer = 0
        Dim pdfFieldFullPath As String = ""
        ' HS : C2008.7 : Fit name in Nama Field
        Dim lenname1 As Integer

        Dim lngIC As Long
        lngIC = frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value
        ' end testing
        Try
            Do While intIndex <= 20
                ' testing

                intIndex = intIndex + 1
                pdfFieldFullPath = pdfFieldPath + "Page" + (intIndex + 3).ToString + "[0]."
                If intIndex < 19 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "Ruj" + intIndex.ToString, lngIC.ToString("0000000000"))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "Ruj" + intIndex.ToString + "_2", lngIC.ToString("0000000000"))
                End If
                ' HS : C2008.7 : Fit name in Nama field
                'pdfFormFields.SetField(pdfFieldFullPath + "Nama" + intIndex.ToString + "_1", Mid(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value.ToString.ToUpper(), 1, 45))
                'If Len(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value) > 30 Then
                'pdfFormFields.SetField(pdfFieldFullPath + "Nama" + intIndex.ToString + "_2", Mid(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value.ToString.ToUpper(), 46))
                'Else
                'pdfFormFields.SetField(pdfFieldFullPath + "Nama" + intIndex.ToString + "_2", "")
                'End If
                If Right(Mid(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value.ToString.ToUpper(), 1, 45), 44) <> "" Then
                    If InStr(Right(Mid(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value.ToString.ToUpper(), 1, 45), 44), "") Then
                        lenname1 = Len(InStr(Right(Mid(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value.ToString.ToUpper(), 1, 45), 44), ""))
                        pdfFormFields.SetField(pdfFieldFullPath + "Nama" + intIndex.ToString + "_1", Mid(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value.ToString.ToUpper(), 1, 44 - (lenname1 + 1)))
                        pdfFormFields.SetField(pdfFieldFullPath + "Nama" + intIndex.ToString + "_2", Mid(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value.ToString.ToUpper(), 45 - (lenname1 + 1)))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "Nama" + intIndex.ToString + "_1", Mid(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value.ToString.ToUpper(), 1, 45))
                    If Len(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value) > 30 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "Nama" + intIndex.ToString + "_2", Mid(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(3).Value.ToString.ToUpper(), 46))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "Nama" + intIndex.ToString + "_2", "")
                    End If
                End If
                ' HS : C2008.7 : Fit name in Nama field : end
            Loop

        Catch ex As Exception
            MsgBox("Name is not fill!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try

    End Sub

    Public Sub CheckFieldEmpty()

        Dim de As DictionaryEntry
        'Dim pdfFieldFullPath As String = pdfFieldPath + "Page3[0]."
        'Dim intIndex As Integer = 0

        For Each de In pdfReader.AcroFields.Fields
            pdfFormFields.SetField(de.Key.ToString, RTrim("---"))
        Next
        'pdfFormFields.SetField(pdfFieldPath + "Page5[0].E3a" ,"sadas")
        'Loop

    End Sub

#Region "Pages Function"

    Public Sub Page1()

        'pdfFieldPath = Application.StartupPath & "\2008c.pdf"
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page3[0]."
        Dim Total As Long = 0
        Dim I As Integer = 0 'for Multi Business Source
        'Dim rows As Integer = 0
        ' HS : C2008.7 : Count Total for A5
        Dim TotalPerniagaan As Long = 0

        Try
            '================== From part 1 , part A , part B - B4   ===========================' 
            ' part 1

            cSQL = "Select TP_COM_NAME, TP_REF_NO, TP_EMPLOYER_NO, TP_RESIDENCE, TP_COUNTRY, TP_ROC_NO, TP_ACC_PERIOD_FR, TP_ACC_PERIOD_TO, TP_EFILING, TP_PUBLIC_ORDER " _
                  & " from TAXP_PROFILE" _
                  & " where TP_REF_NO = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then

                If IsDBNull(dr("TP_COM_NAME")) = False Then
                    If Len(dr("TP_COM_NAME")) > 28 Then
                        CutLine(dr("TP_COM_NAME").ToString.ToUpper(), 28)
                        pdfFormFields.SetField(pdfFieldFullPath + "I_1", strCropped)   'Mid(dr("TP_COM_NAME").ToString.ToUpper(), 1, 28))
                        'pdfFormFields.SetField(pdfFieldFullPath + "I_1", Mid(dr("TP_COM_NAME").ToString.ToUpper(), 1, 28))
                        pdfFormFields.SetField(pdfFieldFullPath + "I_2", Mid(LTrim(strRemainder), 1, 28))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "I_1", dr("TP_COM_NAME")).ToString.ToUpper()
                        pdfFormFields.SetField(pdfFieldFullPath + "I_2", space(28))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "I_1", space(28))
                    pdfFormFields.SetField(pdfFieldFullPath + "I_2", space(28))
                End If

                If IsDBNull(dr("TP_ROC_NO")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "II", dr("TP_ROC_NO"))
                End If

                If IsDBNull(dr("TP_EMPLOYER_NO")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "III", Mid(dr("TP_EMPLOYER_NO"), 1, 10))
                End If

                If IsDBNull(dr("TP_RESIDENCE")) = False Then
                    If dr("TP_RESIDENCE") = "1" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "IV_1", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "IV_2", "")
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "IV_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "IV_2", "X")
                    End If
                End If

                If IsDBNull(dr("TP_COUNTRY")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "V", dr("TP_COUNTRY"))
                End If

                If IsDBNull(dr("TP_REF_NO")) = False Then
                    Dim lngRefNo As Long
                    lngRefNo = dr("TP_REF_NO")
                    pdfFormFields.SetField(pdfFieldFullPath + "VI", Mid(lngRefNo.ToString("0000000000"), 1, 10))
                End If

                If IsDBNull(dr("TP_ACC_PERIOD_FR")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "VII", Mid(Format(dr("TP_ACC_PERIOD_FR"), "ddMMyyyy"), 1, 8))
                End If

                If IsDBNull(dr("TP_ACC_PERIOD_TO")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "VIII", Mid(Format(dr("TP_ACC_PERIOD_TO"), "ddMMyyyy"), 1, 8))
                End If
                If dr("TP_PUBLIC_ORDER") = "0" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "IX_1", "X")
                    pdfFormFields.SetField(pdfFieldFullPath + "IX_2", "")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "IX_1", "")
                    pdfFormFields.SetField(pdfFieldFullPath + "IX_2", "X")
                End If

                If frmDownloadDetails.chkKeepRecord.Checked = True Then

                    pdfFormFields.SetField(pdfFieldFullPath + "X_1", "X")
                    pdfFormFields.SetField(pdfFieldFullPath + "X_2", "")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "X_1", "")
                    pdfFormFields.SetField(pdfFieldFullPath + "X_2", "X")
                End If

                If frmDownloadDetails.chkRKST.Checked = True Then
                    If frmDownloadDetails.optMenuntut.Checked = True Then
                        pdfFormFields.SetField(pdfFieldFullPath + "XI", "1")
                    ElseIf frmDownloadDetails.optMenyerah.Checked = True Then
                        pdfFormFields.SetField(pdfFieldFullPath + "XI", "2")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "XI", "3")
                End If
            Else
                'If Tax Payer Table has no record "
            End If
            dr.Close()

            cSQL = "Select TC_CB_CHECK" _
              & " from TAX_COMPUTATION" _
              & " where TC_REF_NO = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
              & "' And TC_YA='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                If IsDBNull(dr("TC_CB_CHECK")) = False Then
                    If dr("TC_CB_CHECK") = 0 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "XII", "2")
                    ElseIf dr("TC_CB_CHECK") = 1 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "XII", "1")
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "XII", "3")
                    End If
                End If
            Else
                ' If Tax Computation has no record
            End If
            dr.Close()

            ' Kedudukan cukai : Set default to Blank
            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", " ")

            cSQL = "SELECT [PL_S60F] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO]='" _
                    & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' AND [PL_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                If IsDBNull(dr("PL_S60F")) = False Then
                    If dr("PL_S60F") = "Y" Then
                        cSQL = "SELECT * " _
                          & " FROM [INVESTMENT_HOLDING]" _
                          & " where IH_REF_NO = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                          & "' And IH_YA='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows(0).Cells(2).Value) & "'"
                        dr = DataHandler.GetDataReader(cSQL, Conn)
                        If dr.Read() Then
                            If IsDBNull(dr("IH_CHECK")) = False Then
                                If dr("IH_CHECK") = 0 Then
                                    pdfFormFields.SetField(pdfFieldFullPath + "XII", "2")
                                ElseIf dr("IH_CHECK") = 1 Then
                                    pdfFormFields.SetField(pdfFieldFullPath + "XII", "1")
                                Else
                                    pdfFormFields.SetField(pdfFieldFullPath + "XII", "3")
                                End If
                            End If

                            ' Kedudukan cukai
                            Dim FigureH1, FigureH2, FigureH3 As Double
                            FigureH1 = CDbl(dr("IH_ITP"))
                            FigureH2 = CDbl(dr("IH_ITP"))
                            FigureH3 = CDbl(dr("IH_TP_BAL"))

                            If FigureH1 < 0 Then
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "X")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                            End If

                            If (FigureH1 < 0 And FigureH3 < 0) Then
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "X")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                            End If

                            If (FigureH1 > 0 And FigureH3 < 0) Then
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "X")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                            End If

                            If (FigureH1 = 0 And FigureH3 < 0) Then
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "X")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                            End If

                            If (FigureH1 > 0 And FigureH3 > 0) Then
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "X")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                            End If

                            If (FigureH1 = 0 And FigureH2 = 0 And FigureH3 = 0) Then
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "X")
                            End If
                        End If
                        dr.Close()

                        ' New field TAX_COMPUTATION.TC_NB_CHKCB for XIII '
                        pdfFormFields.SetField(pdfFieldFullPath + "XIII", "2")
                    Else
                        ' New field TAX_COMPUTATION.TC_NB_CHKCB for XIII '
                        cSQL = "SELECT [TC_NB_CHKCB] FROM [TAX_COMPUTATION] WHERE [TC_REF_NO]='" _
                                & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                                & "' AND [TC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                                & " ORDER by TC_BUSINESS"
                        dr = DataHandler.GetDataReader(cSQL, Conn)
                        dr.Read()

                        'If Trim(dr("TC_NB_CHKCB")) = "1" Then
                        'pdfFormFields.SetField(pdfFieldFullPath + "XIII", "1")
                        'Else
                        'pdfFormFields.SetField(pdfFieldFullPath + "XIII", "2")
                        'End If
                        dr.Close()

                        ' Kedudukan cukai
                        Dim Figure1, Figure2, Figure3 As Double
                        cSQL = "SELECT [TC_TP_PAYABLE] , [TC_TP_PAYABLE_BAL] FROM [TAX_COMPUTATION] WHERE [TC_REF_NO]='" _
                        & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                        & "' AND [TC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                        & " ORDER by TC_BUSINESS"
                        dr = DataHandler.GetDataReader(cSQL, Conn)
                        dr.Read()
                        Figure1 = CDbl(dr("TC_TP_PAYABLE"))
                        Figure2 = CDbl(dr("TC_TP_PAYABLE"))
                        Figure3 = CDbl(dr("TC_TP_PAYABLE_BAL"))

                        If Figure1 < 0 Then
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "X")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                        End If

                        If (Figure1 < 0 And Figure3 < 0) Then
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "X")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                        End If

                        If (Figure1 > 0 And Figure3 < 0) Then
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "X")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                        End If

                        If (Figure1 = 0 And Figure3 < 0) Then
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "X")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                        End If

                        If (Figure1 > 0 And Figure3 > 0) Then
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "X")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "")
                        End If

                        If (Figure1 = 0 And Figure2 = 0 And Figure3 = 0) Then
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_1", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_2", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_3", "")
                            pdfFormFields.SetField(pdfFieldFullPath + "KCukai_4", "X")
                        End If
                        dr.Close()
                    End If
                End If
            End If
            dr.Close()

            'weihong XIII Syarikat Kecil dan Sederhana
            cSQL = "SELECT SME" _
                   & " FROM [BALANCE_SHEET]" _
                   & " WHERE [BS_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                   & "' And BS_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                If Trim(dr("SME")) = "1" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "XIII", "1")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "XIII", "2")
                End If
            End If
            dr.Close()
            'endweihong

            ' Bahagian A
            cSQL = "SELECT [PL_S60F] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' And PL_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            dr.Read()
            If dr("PL_S60F") = "Y" Then

                cSQL = "Select BC_CODE" _
                        & " from BUSINESS_SOURCE" _
                        & " where BC_KEY = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' and BC_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "' and BC_SOURCENO = 1" _
                        & " order by BC_SOURCENO"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                If dr.Read() Then
                    ' 1 Business Source only
                    pdfFormFields.SetField(pdfFieldFullPath + "A1_1", dr("BC_CODE").ToString())
                    pdfFormFields.SetField(pdfFieldFullPath + "A1_2", CDbl("0").ToString.Replace(",", ""))

                    'Other Business Source
                    pdfFormFields.SetField("A2_1", space(5))
                    pdfFormFields.SetField("A2_2", "0")
                    pdfFormFields.SetField("A3_1", space(5))
                    pdfFormFields.SetField("A3_2", "0")
                    pdfFormFields.SetField("A4_1", space(5))
                    pdfFormFields.SetField("A4_2", "0")
                    pdfFormFields.SetField("A5_1", space(5))
                    pdfFormFields.SetField("A5_2", "0")

                    ''Perkongsian
                    'pdfFormFields.SetField("C6_1", "---")
                    pdfFormFields.SetField("A6_2", "0")
                    'pdfFormFields.SetField("C7_1", "---")
                    pdfFormFields.SetField("A7_2", "0")
                    'pdfFormFields.SetField("C8_1", "---")
                    pdfFormFields.SetField("A8_2", "0")
                    'pdfFormFields.SetField("C9_1", "---")
                    pdfFormFields.SetField("A9_2", "0")
                    'pdfFormFields.SetField("C10_1", "---")
                    pdfFormFields.SetField("A10_2", "0")

                    'Calculate Total
                    pdfFormFields.SetField(pdfFieldFullPath + "A11", "0")    'A1 + A2 + A3 + ... + A1
                End If
                dr.Close()
            Else
                ' ==== If S60F IS NOT Checked , Then TAX COMPUTATION is Load ==== '
                ' Allow Multi Business Source
                ' !!!!!!!!!!!!not sure for Business Source  > 6 !!!!!!!
                'NGOHCS 2009 
                Dim strArray As String()
                ReDim strArray(10)

                cSQL = "Select BC_CODE" _
                        & " from BUSINESS_SOURCE" _
                        & " where BC_KEY = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' and BC_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                        & " order by BC_SOURCENO"

                dr = DataHandler.GetDataReader(cSQL, Conn)

                Do While dr.Read()
                    strArray(I) = dr("BC_CODE")
                    I = I + 1
                    If I < 6 Then '6
                        pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_1", dr("BC_CODE")) ' Business Source
                    End If
                Loop

                While I < 5
                    I = I + 1
                    pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_1", space(5)) ' Business Source
                End While
                dr.Close()

                Dim intBCCode As Integer = 0
                Dim intBCSource As Integer = 0
                ' Amount according Business Source
                cSQL = "Select TC_BUSINESS, TC_SI_NET_STAT_IN" _
                        & " from TAX_COMPUTATION" _
                        & " where TC_REF_NO = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' and TC_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                        & " order by TC_BUSINESS"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                I = 0
                Do While dr.Read()
                    If I <= 10 Then
                        I = I + 1
                        If I < 5 Then ' 6
                            pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_2", CDbl(dr("TC_SI_NET_STAT_IN")).ToString.Replace(",", "")) ' Amount
                        Else
                            If I = 5 Then
                                intBCSource = 4
                            End If
                            If CDbl(dr("TC_SI_NET_STAT_IN")) > 0 Then
                                intBCCode = intBCCode + 1
                                intBCSource = I - 1
                            End If
                            Total = Total + CDbl(dr("TC_SI_NET_STAT_IN"))
                            pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_2", Total.ToString.Replace(",", "")) ' Amount
                        End If
                    End If
                    If I >= 5 Then
                        TotalPerniagaan = TotalPerniagaan + CDbl(dr("TC_SI_NET_STAT_IN"))
                    End If
                Loop

                If intBCCode > 1 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "A5_1", space(5)) ' Description
                ElseIf intBCCode = 1 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "A5_1", strArray(intBCSource))
                End If
                pdfFormFields.SetField(pdfFieldFullPath + "A5_2", TotalPerniagaan.ToString.Replace(",", "")) ' Amount

                While I < 5        ' Still less than 5 source
                    I = I + 1
                    pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_2", "0")
                End While

                dr.Close()

                ''Perkongsian
                'pdfFormFields.SetField("C6_1", "---")
                pdfFormFields.SetField("A6_2", "0")
                'pdfFormFields.SetField("C7_1", "---")
                pdfFormFields.SetField("A7_2", "0")
                'pdfFormFields.SetField("C8_1", "---")
                pdfFormFields.SetField("A8_2", "0")
                'pdfFormFields.SetField("C9_1", "---")
                pdfFormFields.SetField("A9_2", "0")
                'pdfFormFields.SetField("C10_1", "---")
                pdfFormFields.SetField("A10_2", "0")

                'weihong
                ''Perkongsian
                ' === CYS 2010 ===
                Dim cSQL2 As String
                Dim strKey As String
                Dim sourceNo As String
                Dim StatPartnership As Double = 0.0 'weihong

                cSQL = "Select [TCP_KEY], [TCP_PARTNERSHIP]" _
                        & " from [TAX_COMPUTATION_PARTNER]" _
                        & " where [TCP_REF_NO] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' and [TCP_YA] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                If dr.Read() Then
                    strKey = dr("TCP_KEY")
                    StatPartnership = CDbl(dr("TCP_PARTNERSHIP")) 'weihong
                Else
                    strKey = "0"
                End If
                dr.Close()

                cSQL = "Select [PS_FILE_NO2],[PS_FILE_NO3]" _
                        & " from [TAXP_PARTNERSHIP]" _
                        & " where [PS_KEY] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' and [PS_YA] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                        & " order by [PS_SOURCENO]"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                I = 5




                Do While dr.Read()
                    If I < 10 Then  'HoGie
                        strArray(I) = dr("PS_FILE_NO2") + dr("PS_FILE_NO3")
                    End If

                    I = I + 1

                    If I < 11 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_1", dr("PS_FILE_NO2") + dr("PS_FILE_NO3"))
                    End If
                Loop

                While I < 11
                    I = I + 1
                    pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_1", space(10))

                End While
                dr.Close()



                Dim rowcount As Integer = 0

                Dim lngTotalPartnerAmount As Long = 0

                cSQL = "Select [PS_SOURCENO]" _
                        & " from [TAXP_PARTNERSHIP]" _
                        & " where [PS_KEY] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' and [PS_YA] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                        & " order by [PS_SOURCENO]"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                I = 5

                Do While dr.Read()

                    'Ho Gie

                    sourceNo = dr("PS_SOURCENO")
                    'Ho Gie - End

                    cSQL2 = "Select [PN_TOTAL_STAT_INCOME]" _
                        & " from [INCOME_PARTNERSHIP]" _
                        & " where [TCP_KEY] = " & strKey & " AND [PN_SOURCENO] = " & sourceNo _
                        & " order by [PN_SOURCENO]"
                    dr2 = DataHandler.GetDataReader(cSQL2, Conn)
                    If dr2.Read() Then

                        'strArray(I) = dr2("PN_TOTAL_STAT_INCOME")

                        rowcount += 1


                        I = I + 1
                        'If I < 10 Then
                        If rowcount < 5 Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_2", CDbl(dr2("PN_TOTAL_STAT_INCOME")).ToString.Replace(",", ""))

                            'ElseIf I >= 10 Then
                        ElseIf rowcount > 4 Then
                            ' Total Perkongsian for source more than 5 
                            lngTotalPartnerAmount = lngTotalPartnerAmount + CDbl(dr2("PN_TOTAL_STAT_INCOME"))
                            pdfFormFields.SetField(pdfFieldFullPath + "A10_2", lngTotalPartnerAmount.ToString.Replace(",", ""))
                        End If


                    Else
                        I = I + 1
                        If I < 11 Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_2", "0")
                        End If
                    End If

                    dr2.Close()
                Loop

                While I < 10        ' Still less than 5 source
                    I = I + 1
                    pdfFormFields.SetField(pdfFieldFullPath + "A" + I.ToString + "_2", "0")
                End While

                dr.Close()
                'endweihong

                'Calculate Total
                cSQL = "Select TC_NB_INT_NET, TC_NB_RENT_NET, TC_NB_SUNDRY, TC_NB_ADDITION, TC_TP_CURR_LOSS," _
                    & " TC_TP_PROSPECTING, TC_TP_PREOP_BS," _
                    & " TC_TP_APRV_DONATION, TC_TP_ZAKAT," _
                    & " TC_AI_ROYALTY, TC_NB_OTH_AGGR_STAT, TC_TP_AGGR_IN, TC_TP_AGGR_IN_LOSS, TC_TP_TOTAL_OTH_EXP, TC_TP_CLAIM, TC_TP_TOTAL_INCOME," _
                    & " TC_SI_TOT, TC_SI_BS_LOSS_BF, TC_SI_AGGREGATE, TC_TP_EXP_ALLOWED," _
                    & " TC_BUSINESS" _
                    & " from TAX_COMPUTATION" _
                    & " WHERE [TC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' and TC_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                    & " Order By TC_BUSINESS"

                dr = DataHandler.GetDataReader(cSQL, Conn)
                If dr.Read() Then
                    pdfFormFields.SetField(pdfFieldFullPath + "A11", CDbl(dr("TC_SI_TOT")).ToString.Replace(",", "") + StatPartnership)    'A1 + A2 + A3 + ... + A10 'weihong+ StatPartnership
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "A11", "0")
                End If
                dr.Close()
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 3!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
        ' ==== End Page 1 Borang C 2010 ==== '
    End Sub

    Public Sub Page2()
        ' Combination of some part from Page2 to Page3
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page4[0]."
        Dim pdfFieldFullPath2 As String = pdfFieldPath + "Page5[0]."
        Dim Total As Long = 0
        Dim I As Integer = 0 'for Multi Business Source
        Dim PLRefNo, amount, DetailType, KEY, KeySourceno, Refno1, YearAccess, business As String
        Dim TotalAmount1, TotalAmount2, TotalAmount3, TotalAmount4, TotalAmount5, TotalAmount6, TotalAmount7, _
        TotalA1, LessA1, totalSum As Double
        PLRefNo = ""
        Refno1 = Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value)
        amount = ""
        TotalAmount1 = 0
        TotalAmount2 = 0
        TotalAmount3 = 0
        TotalAmount4 = 0
        TotalAmount5 = 0
        TotalAmount6 = 0
        TotalAmount7 = 0
        DetailType = ""
        YearAccess = Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value)
        KEY = ""
        KeySourceno = ""

        Try
            ' Bahagian A
            cSQL = "SELECT [PL_S60F] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO]='" _
                & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                & "' And PL_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            dr.Read()
            If dr("PL_S60F") = "Y" Then ' Investment Holding '

                cSQL = "Select BC_CODE" _
                        & " from BUSINESS_SOURCE" _
                        & " where BC_KEY = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                        & "' and BC_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) _
                        & "' and BC_SOURCENO = 1" _
                        & " order by BC_SOURCENO"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                If dr.Read() Then
                    pdfFormFields.SetField(pdfFieldFullPath + "A12", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A13", "0")    ' A11 - A12
                End If
                dr.Close()

                cSQL = "SELECT IH_DIVIDEND_NET, IH_INTEREST_NET, IH_RENTAL_NET, IH_EXP_ALLOWED," _
                    & " IH_APPR_DONATION,IH_ZAKAT, IH_PIONEER_CHARGE, IH_OP_HQ_CHARGE, IH_FOREIGN_CHARGE," _
                    & " IH_NET_EXDIV_TOTAL, IH_EXP_ALLOWED, IH_TOTAL_OTH_EXP, IH_APPR_DONATION, IH_ZAKAT, IH_CLAIM,IH_TOTAL_INCOME , IH_TOTAL_INCOME," _
                    & " IH_ROYALTY, IH_OTHER_INCOME, IH_ADDITION, IH_TP_AGGR_IN, IH_OTHER_EXPENSES" _
                    & " FROM [INVESTMENT_HOLDING]" _
                    & " WHERE [IH_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' and IH_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

                dr = DataHandler.GetDataReader(cSQL, Conn)
                If dr.Read() Then
                    ' IH_TP_AGGR_IN for new field - donation
                    TotalA1 = Trim(CDbl(dr("IH_TP_AGGR_IN")))

                    pdfFormFields.SetField(pdfFieldFullPath + "A14", CDbl(dr("IH_INTEREST_NET")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A15", CDbl(dr("IH_RENTAL_NET")) + CDbl(dr("IH_ROYALTY")).ToString.Replace(",", ""))
                    If (CDbl(dr("IH_OTHER_INCOME")) - CDbl(dr("IH_OTHER_EXPENSES"))) >= 0 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "A16", (CDbl(dr("IH_OTHER_INCOME")) - CDbl(dr("IH_OTHER_EXPENSES"))).ToString.Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "A16", "0")
                    End If
                    pdfFormFields.SetField(pdfFieldFullPath + "A17", CDbl(dr("IH_ADDITION")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A18", CDbl(dr("IH_TP_AGGR_IN")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A19", CDbl(dr("IH_TP_AGGR_IN")).ToString.Replace(",", ""))

                    pdfFormFields.SetField(pdfFieldFullPath + "A20", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A21", CDbl(dr("IH_TP_AGGR_IN")).ToString.Replace(",", ""))

                    pdfFormFields.SetField(pdfFieldFullPath + "A22", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A23", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A24", CDbl(dr("IH_EXP_ALLOWED")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A25", CDbl(dr("IH_TOTAL_OTH_EXP")).ToString.Replace(",", ""))

                    ' variable used to store amount to stamp in Page3()
                    totalA25 = CDbl(dr("IH_TOTAL_OTH_EXP"))
                    totalA31 = CDbl(dr("IH_ZAKAT"))
                    totalA32 = CDbl(dr("IH_CLAIM"))

                End If
                dr.Close()

                ' New field for donation
                cSQL = "SELECT [PL_KEY] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" & Refno1 & "' AND [PL_YA] = '" & YearAccess & "'"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                If dr.Read() Then
                    PLRefNo = Trim(dr("PL_KEY"))
                End If
                dr.Close()

                cSQL = "SELECT * FROM [OTHER_EXAPPRDONATION] WHERE [EXOAD_KEY] = " & PLRefNo & ""
                dr2 = DataHandler.GetDataReader(cSQL, Conn)
                Do While dr2.Read()
                    business = Trim(dr2("EXOAD_SOURCENO"))
                    If dr2("EXOAD_DETAIL") = "Yes" Then
                        cSQL = "SELECT * FROM [OTHER_EXAPPRDONATION_DETAIL] WHERE [EXOADD_KEY] = " & PLRefNo & " AND [EXOADD_SOURCENO] = " & business & " AND EXOADD_EXOADKEY =" & dr2("EXOAD_EXOADKEY")
                        dr = DataHandler.GetDataReader(cSQL, Conn)
                        Do While dr.Read()
                            Select Case Trim(dr("EXOADD_TYPE"))
                                Case "Gifts of Money to Government"
                                    TotalAmount1 = TotalAmount1 + CDbl(dr("EXOADD_AMOUNT"))
                                Case "Gifts of Money to Approved institutions"
                                    TotalAmount2 = TotalAmount2 + CDbl(dr("EXOADD_AMOUNT"))
                                Case "Gifts of Money or Contribution for Approved Sports Activities"
                                    TotalAmount3 = TotalAmount3 + CDbl(dr("EXOADD_AMOUNT"))
                                Case "Gifts of Money or Contribution for Approved Project of National Interest"
                                    TotalAmount4 = TotalAmount4 + CDbl(dr("EXOADD_AMOUNT"))
                                Case "Gifts of Artifacts, Manuscripts or paintings"
                                    TotalAmount5 = TotalAmount5 + CDbl(dr("EXOADD_AMOUNT"))
                                Case "Gifts of Money to Library"
                                    TotalAmount6 = TotalAmount6 + CDbl(dr("EXOADD_AMOUNT"))
                                Case "Gifts of Paintings to National Art Gallery"
                                    TotalAmount7 = TotalAmount7 + CDbl(dr("EXOADD_AMOUNT"))
                            End Select
                        Loop
                        dr.Close()
                    Else
                        Select Case Trim(dr2("EXOAD_TYPE"))
                            Case "Gifts of Money to Government"
                                TotalAmount1 = TotalAmount1 + CDbl(dr2("EXOAD_AMOUNT"))
                            Case "Gifts of Money to Approved institutions"
                                TotalAmount2 = TotalAmount2 + CDbl(dr2("EXOAD_AMOUNT"))
                            Case "Gifts of Money or Contribution for Approved Sports Activities"
                                TotalAmount3 = TotalAmount3 + CDbl(dr2("EXOAD_AMOUNT"))
                            Case "Gifts of Money or Contribution for Approved Project of National Interest"
                                TotalAmount4 = TotalAmount4 + CDbl(dr2("EXOAD_AMOUNT"))
                            Case "Gifts of Artifacts, Manuscripts or paintings"
                                TotalAmount5 = TotalAmount5 + CDbl(dr2("EXOAD_AMOUNT"))
                            Case "Gifts of Money to Library"
                                TotalAmount6 = TotalAmount6 + CDbl(dr2("EXOAD_AMOUNT"))
                            Case "Gifts of Paintings to National Art Gallery"
                                TotalAmount7 = TotalAmount7 + CDbl(dr2("EXOAD_AMOUNT"))
                        End Select
                    End If
                Loop
                dr2.Close()

                LessA1 = 0

                pdfFormFields.SetField(pdfFieldFullPath + "A26", TotalAmount1.ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "A27A", TotalAmount2.ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "A27B", TotalAmount3.ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "A27C", TotalAmount4.ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "A28", TotalAmount5.ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "A29", TotalAmount6.ToString.Replace(",", ""))

                totalSum = TotalAmount2 + TotalAmount3 + TotalAmount4
                pdfFormFields.SetField(pdfFieldFullPath + "A27", totalSum.ToString.Replace(",", ""))

                Select Case Val(YearAccess)
                    Case Is >= 2009
                        ' "Restricted to 10% of (AI)"
                        LessA1 = TotalA1 * 10 / 100
                        LessA1 = FormatNumber(CDbl(LessA1), 0)
                        If totalSum >= LessA1 Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A27", LessA1.ToString.Replace(",", ""))
                        End If
                End Select

                ' variable used to store amount to stamp in Page3()
                totalA26 = TotalAmount1
                If totalSum >= LessA1 Then
                    totalA27 = LessA1
                Else
                    totalA27 = totalSum
                End If
                totalA28 = TotalAmount5
                totalA29 = TotalAmount6
                totalA30 = TotalAmount7
            Else
                'Calculate Total
                cSQL = "Select TC_NB_INT_NET, TC_NB_RENT_NET, TC_NB_SUNDRY, TC_NB_ADDITION, TC_TP_CURR_LOSS," _
                    & " TC_TP_PROSPECTING, TC_TP_PREOP_BS," _
                    & " TC_TP_APRV_DONATION, TC_TP_ZAKAT," _
                    & " TC_AI_ROYALTY, TC_NB_OTH_AGGR_STAT, TC_TP_AGGR_IN, TC_TP_AGGR_IN_LOSS, TC_TP_TOTAL_OTH_EXP, TC_TP_CLAIM, TC_TP_TOTAL_INCOME," _
                    & " TC_SI_TOT, TC_SI_BS_LOSS_BF, TC_SI_AGGREGATE, TC_TP_EXP_ALLOWED," _
                    & " TC_BUSINESS, TC_SUNDRY_EXP, TC_NB_ROYALTY" _
                    & " from TAX_COMPUTATION" _
                    & " WHERE [TC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' and TC_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                    & " Order By TC_BUSINESS"

                dr = DataHandler.GetDataReader(cSQL, Conn)
                If dr.Read() Then
                    pdfFormFields.SetField(pdfFieldFullPath + "A12", CDbl(dr("TC_SI_BS_LOSS_BF")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A13", CDbl(dr("TC_SI_AGGREGATE")).ToString.Replace(",", ""))    ' C11 - C12

                    pdfFormFields.SetField(pdfFieldFullPath + "A14", CDbl(dr("TC_NB_INT_NET")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A15", (CDbl(dr("TC_NB_RENT_NET")) + CDbl(dr("TC_NB_ROYALTY"))).ToString.Replace(",", ""))
                    If (CDbl(dr("TC_NB_SUNDRY")) - CDbl(dr("TC_SUNDRY_EXP"))) >= 0 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "A16", (CDbl(dr("TC_NB_SUNDRY")) - CDbl(dr("TC_SUNDRY_EXP"))).ToString.Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "A16", "0")
                    End If
                    pdfFormFields.SetField(pdfFieldFullPath + "A17", CDbl(dr("TC_NB_ADDITION")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A18", CDbl(dr("TC_NB_OTH_AGGR_STAT")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A19", CDbl(dr("TC_TP_AGGR_IN")).ToString.Replace(",", ""))

                    pdfFormFields.SetField(pdfFieldFullPath + "A20", CDbl(dr("TC_TP_CURR_LOSS")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A21", CDbl(dr("TC_TP_AGGR_IN_LOSS")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A22", CDbl(dr("TC_TP_PROSPECTING")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A23", CDbl(dr("TC_TP_PREOP_BS")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A24", CDbl(dr("TC_TP_EXP_ALLOWED")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A25", CDbl(dr("TC_TP_TOTAL_OTH_EXP")).ToString.Replace(",", ""))

                    ' variable used to store amount to stamp in Page3()
                    totalA25 = CDbl(dr("TC_TP_TOTAL_OTH_EXP"))
                    totalA31 = CDbl(dr("TC_TP_ZAKAT"))
                    totalA32 = CDbl(dr("TC_TP_CLAIM"))

                    ' New field for donation
                    ' TC_BUSINESS,TC_TP_AGGR_IN  for the next section - Donation
                    TotalA1 = Trim(CDbl(dr("TC_TP_AGGR_IN")))

                    cSQL = "SELECT [PL_KEY] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" & Refno1 & "' AND [PL_YA] = '" & YearAccess & "'"
                    dr = DataHandler.GetDataReader(cSQL, Conn)
                    If dr.Read() Then
                        PLRefNo = Trim(dr("PL_KEY"))
                    End If
                    dr.Close()

                    cSQL = "SELECT  [BC_SOURCENO] FROM [BUSINESS_SOURCE] " _
                    & " WHERE [BC_KEY]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' and BC_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
                    dr = DataHandler.GetDataReader(cSQL, Conn)
                    Do While dr.Read()
                        business = Trim(dr("BC_SOURCENO"))
                        cSQL = "SELECT * FROM [OTHER_EXAPPRDONATION] WHERE [EXOAD_KEY] = " & PLRefNo & " AND [EXOAD_SOURCENO] = " & business & ""

                        dr3 = DataHandler.GetDataReader(cSQL, Conn)
                        Do While dr3.Read()
                            If Trim(dr3("EXOAD_DETAIL")) = "Yes" Then
                                cSQL = "SELECT * FROM [OTHER_EXAPPRDONATION_DETAIL] WHERE [EXOADD_KEY] = " & PLRefNo & " AND [EXOADD_SOURCENO] = " & business & " AND EXOADD_EXOADKEY =" & dr3("EXOAD_EXOADKEY")

                                dr4 = DataHandler.GetDataReader(cSQL, Conn)
                                Do While dr4.Read()
                                    Select Case dr4("EXOADD_TYPE")
                                        Case "Gifts of Money to Government"
                                            TotalAmount1 = TotalAmount1 + CDbl(dr4("EXOADD_AMOUNT"))
                                        Case "Gifts of Money to Approved institutions"
                                            TotalAmount2 = TotalAmount2 + CDbl(dr4("EXOADD_AMOUNT"))
                                        Case "Gifts of Money or Contribution for Approved Sports Activities"
                                            TotalAmount3 = TotalAmount3 + CDbl(dr4("EXOADD_AMOUNT"))
                                        Case "Gifts of Money or Contribution for Approved Project of National Interest"
                                            TotalAmount4 = TotalAmount4 + CDbl(dr4("EXOADD_AMOUNT"))
                                        Case "Gifts of Artifacts, Manuscripts or paintings"
                                            TotalAmount5 = TotalAmount5 + CDbl(dr4("EXOADD_AMOUNT"))
                                        Case "Gifts of Money to Library"
                                            TotalAmount6 = TotalAmount6 + CDbl(dr4("EXOADD_AMOUNT"))
                                        Case "Gifts of Paintings to National Art Gallery"
                                            TotalAmount7 = TotalAmount7 + CDbl(dr4("EXOADD_AMOUNT"))
                                    End Select
                                Loop
                                dr4.Close()
                            Else
                                Select Case dr3("EXOAD_TYPE")
                                    Case "Gifts of Money to Government"
                                        TotalAmount1 = TotalAmount1 + CDbl(dr3("EXOAD_AMOUNT"))
                                    Case "Gifts of Money to Approved institutions"
                                        TotalAmount2 = TotalAmount2 + CDbl(dr3("EXOAD_AMOUNT"))
                                    Case "Gifts of Money or Contribution for Approved Sports Activities"
                                        TotalAmount3 = TotalAmount3 + CDbl(dr3("EXOAD_AMOUNT"))
                                    Case "Gifts of Money or Contribution for Approved Project of National Interest"
                                        TotalAmount4 = TotalAmount4 + CDbl(dr3("EXOAD_AMOUNT"))
                                    Case "Gifts of Artifacts, Manuscripts or paintings"
                                        TotalAmount5 = TotalAmount5 + CDbl(dr3("EXOAD_AMOUNT"))
                                    Case "Gifts of Money to Library"
                                        TotalAmount6 = TotalAmount6 + CDbl(dr3("EXOAD_AMOUNT"))
                                    Case "Gifts of Paintings to National Art Gallery"
                                        TotalAmount7 = TotalAmount7 + CDbl(dr3("EXOAD_AMOUNT"))
                                End Select
                            End If
                        Loop
                        dr3.Close()
                    Loop
                    dr.Close()

                    LessA1 = 0

                    pdfFormFields.SetField(pdfFieldFullPath + "A26", TotalAmount1.ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A27A", TotalAmount2.ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A27B", TotalAmount3.ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A27C", TotalAmount4.ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A28", TotalAmount5.ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "A29", TotalAmount6.ToString.Replace(",", ""))

                    totalSum = TotalAmount2 + TotalAmount3 + TotalAmount4
                    pdfFormFields.SetField(pdfFieldFullPath + "A27", totalSum.ToString.Replace(",", ""))

                    Select Case Val(YearAccess)
                        Case Is >= 2009
                            ' "Restricted to 10% of (AI)"
                            LessA1 = TotalA1 * 10 / 100
                            LessA1 = FormatNumber(CDbl(LessA1), 0)
                            If totalSum >= LessA1 Then
                                pdfFormFields.SetField(pdfFieldFullPath + "A27", LessA1.ToString.Replace(",", ""))
                            End If
                    End Select

                    ' variable used to store amount to stamp in Page3()
                    totalA26 = TotalAmount1
                    If totalSum >= LessA1 Then
                        totalA27 = LessA1
                    Else
                        totalA27 = totalSum
                    End If
                    totalA28 = TotalAmount5
                    totalA29 = TotalAmount6
                    totalA30 = TotalAmount7
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "A12", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A13", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A14", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A15", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A16", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A17", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A18", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A19", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A20", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A21", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A22", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A23", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A24", "0")
                    pdfFormFields.SetField(pdfFieldFullPath + "A25", "0")

                    ' variable used to store amount to stamp in Page3()
                    totalA25 = 0
                End If
                dr.Close()

            End If
            dr.Close()

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 2!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page3()
        ' Combination of Page3 and Page4
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page5[0]."
        Dim pdfFieldFullPath2 As String = pdfFieldPath + "Page6[0]."
        Dim totalsumA33 As Double
        Dim Total As Long = 0
        Dim I As Integer = 0 'for Multi Business Source

        Try
            pdfFormFields.SetField(pdfFieldFullPath + "A30", totalA30.ToString.Replace(",", ""))
            pdfFormFields.SetField(pdfFieldFullPath + "A31", totalA31.ToString.Replace(",", ""))
            pdfFormFields.SetField(pdfFieldFullPath + "A32", totalA32.ToString.Replace(",", ""))

            totalsumA33 = totalA25 - (totalA26 + totalA27 + totalA28 + totalA29 + totalA30 + totalA31 + totalA32)

            If totalsumA33 <= 0 Then
                pdfFormFields.SetField(pdfFieldFullPath + "A33", "0")
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "A33", totalsumA33.ToString.Replace(",", ""))
            End If

            cSQL = "SELECT [PL_S60F] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' And PL_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                If dr("PL_S60F") = "Y" Then
                    ' ==== If S60F IS Checked , Then INVESTMENT HOLDING is Load ==== '
                    ' INVESTMENT HOLDING Part C
                    cSQL = "SELECT IH_DIVIDEND_NET, IH_INTEREST_NET, IH_RENTAL_NET, IH_EXP_ALLOWED," _
                       & " IH_APPR_DONATION,IH_ZAKAT,IH_CLAIM, IH_PIONEER_CHARGE, IH_OP_HQ_CHARGE, IH_FOREIGN_CHARGE," _
                       & " IH_STAT_DIVIDEND, IH_CHARGEABLE_IN" _
                       & " FROM [INVESTMENT_HOLDING]" _
                       & " WHERE [IH_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                       & "' AND [IH_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
                    dr = DataHandler.GetDataReader(cSQL, Conn)
                    If dr.Read() Then
                        If IsDBNull(dr("IH_STAT_DIVIDEND")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A34", CDbl(dr("IH_STAT_DIVIDEND")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "A34", "0")
                        End If
                        If IsDBNull(dr("IH_PIONEER_CHARGE")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A35", CDbl(dr("IH_PIONEER_CHARGE")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "A35", "0")
                        End If
                        If IsDBNull(dr("IH_CHARGEABLE_IN")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A36", CDbl(dr("IH_CHARGEABLE_IN")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "A36", "0")
                        End If
                        If IsDBNull(dr("IH_FOREIGN_CHARGE")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A37", CDbl(dr("IH_FOREIGN_CHARGE")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "A37", "0")
                        End If
                        If IsDBNull(dr("IH_CHARGEABLE_IN")) = False Then
                            pdfFormFields.SetField("B1", CDbl(dr("IH_CHARGEABLE_IN")) + CDbl(dr("IH_FOREIGN_CHARGE")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "B1", "0")
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "A34", "0")
                        pdfFormFields.SetField(pdfFieldFullPath + "A35", "0")
                        pdfFormFields.SetField(pdfFieldFullPath + "A36", "0")
                        pdfFormFields.SetField(pdfFieldFullPath + "A37", "0")
                        pdfFormFields.SetField(pdfFieldFullPath + "B1", "0")
                    End If
                Else

                    ' ==== If S60F IS NOT Checked , Then TAX COMPUTATION is Load ==== '
                    ' TAX COMPUTATION Part C

                    cSQL = "SELECT TC_STAT_DIVIDEND, TC_TP_PIONEER_CHARGE, TC_TP_CHARGEABLE, TC_TP_FOREIGN_CHARGE" _
                            & " FROM TAX_COMPUTATION" _
                            & " WHERE [TC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                            & "' AND [TC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                            & " ORDER BY [TC_BUSINESS]"

                    dr = DataHandler.GetDataReader(cSQL, Conn)
                    If dr.Read() Then
                        If IsDBNull(dr("TC_STAT_DIVIDEND")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A34", CDbl(dr("TC_STAT_DIVIDEND")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "A34", "0")
                        End If
                        If IsDBNull(dr("TC_TP_PIONEER_CHARGE")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A35", CDbl(dr("TC_TP_PIONEER_CHARGE")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "A35", "0")
                        End If
                        If IsDBNull(dr("TC_TP_CHARGEABLE")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A36", CDbl(dr("TC_TP_CHARGEABLE")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "A36", "0")
                        End If
                        If IsDBNull(dr("TC_TP_FOREIGN_CHARGE")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "A37", CDbl(dr("TC_TP_FOREIGN_CHARGE")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "A37", "0")
                        End If
                        If IsDBNull(dr("TC_TP_CHARGEABLE")) = False Then
                            pdfFormFields.SetField("B1", CDbl(dr("TC_TP_CHARGEABLE")) + CDbl(dr("TC_TP_FOREIGN_CHARGE")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "B1", "0")
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "A34", "0")
                        pdfFormFields.SetField(pdfFieldFullPath + "A35", "0")
                        pdfFormFields.SetField(pdfFieldFullPath + "A36", "0")
                        pdfFormFields.SetField(pdfFieldFullPath + "A37", "0")
                        pdfFormFields.SetField(pdfFieldFullPath + "B1", "0")
                    End If
                End If
            End If

            ' ==== If S60F is Checked , Run INVESTMENT HOLDING ===== '
            ' Investment Holding part A

            cSQL = "SELECT [PL_S60F] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' AND [PL_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            dr.Read()
            If dr("PL_S60F") = "Y" Then
                cSQL = "SELECT IH_APP_CHARGE_IN1, IH_APP_CHARGE_IN2, IH_APP_CHARGE_IN3, IH_APP_CHARGE_IN4, IH_APP_CHARGE_IN4A, IH_APP_CHARGE_IN5, IH_APP_CHARGE_IN6, IH_RATE1, IH_RATE2, IH_RATE3, IH_RATE4, IH_RATE4A, IH_RATE5, IH_RATE6," _
                      & " IH_SEC6B_REBATE, IH_ITP_SETOFF, IH_ITP_SETOFF_OTH, IH_SEC132, IH_SEC133, IH_TOTAL_TAX_CHARGED, IH_TOTAL_SETOFF, IH_ITP_SETOFF_110B, IH_ITP ,  IH_INSTALLMENTS , IH_TP_BAL" _
                      & " FROM [INVESTMENT_HOLDING]" _
                      & " WHERE [IH_REF_NO] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' And IH_YA='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
                dr = DataHandler.GetDataReader(cSQL, Conn)

                If dr.Read() Then
                    If IsDBNull(dr("IH_APP_CHARGE_IN1")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B2_1", Replace(CDbl(dr("IH_APP_CHARGE_IN1")), ",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B2_1", "0")
                    End If
                    If IsDBNull(dr("IH_RATE1")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B2_2", Replace((Convert.ToDouble((dr("IH_APP_CHARGE_IN1")) * CDbl(dr("IH_RATE1")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B2_2", "000")
                    End If
                    If IsDBNull(dr("IH_APP_CHARGE_IN2")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B3_1", Replace(CDbl(dr("IH_APP_CHARGE_IN2")), ",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B3_1", "0")
                    End If
                    If IsDBNull(dr("IH_RATE2")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B3_2", Replace((Convert.ToDouble((dr("IH_APP_CHARGE_IN2")) * CDbl(dr("IH_RATE2")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B3_2", "000")
                    End If
                    If IsDBNull(dr("IH_APP_CHARGE_IN3")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B4_1", Replace(CDbl(dr("IH_APP_CHARGE_IN3")), ",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B4_1", "0")
                    End If
                    If IsDBNull(dr("IH_RATE3")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B4_2", Replace((Convert.ToDouble((dr("IH_APP_CHARGE_IN3")) * CDbl(dr("IH_RATE3")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B4_2", "000")
                    End If
                    If IsDBNull(dr("IH_APP_CHARGE_IN4")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B5_1", Replace(CDbl(dr("IH_APP_CHARGE_IN4")), ",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B5_1", "0")
                    End If
                    If IsDBNull(dr("IH_RATE4")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B5_2", Replace((Convert.ToDouble((dr("IH_APP_CHARGE_IN4")) * CDbl(dr("IH_RATE4")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B5_2", "000")
                    End If
                    If IsDBNull(dr("IH_APP_CHARGE_IN4A")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B5A_1", Replace(CDbl(dr("IH_APP_CHARGE_IN4A")), ",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B5A_1", "0")
                    End If
                    If IsDBNull(dr("IH_RATE4A")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B5A_2", Replace((Convert.ToDouble((dr("IH_APP_CHARGE_IN4A")) * CDbl(dr("IH_RATE4A")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B5A_2", "000")
                    End If
                    If IsDBNull(dr("IH_APP_CHARGE_IN5")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B6_1", Replace(CDbl(dr("IH_APP_CHARGE_IN5")), ",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B6_1", "0")
                    End If
                    If IsDBNull(dr("IH_RATE5")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B6_2", Replace((Convert.ToDouble((dr("IH_APP_CHARGE_IN5")) * CDbl(dr("IH_RATE5")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B6_2", "000")
                    End If
                    If IsDBNull(dr("IH_APP_CHARGE_IN6")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B7_1", Replace(CDbl(dr("IH_APP_CHARGE_IN6")), ",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B7_1", "0")
                    End If

                    If IsDBNull(dr("IH_RATE6")) = False Then
                        Dim LastRate As String
                        LastRate = CDbl(dr("IH_RATE6"))
                        If Val(LastRate) = 0 Then
                            LastRate = "000"
                        Else
                            If Len(LastRate) = 1 Then
                                LastRate = " " + LastRate + "0"
                            Else
                                If Len(LastRate) = 2 Then
                                    LastRate = LastRate + "0"
                                End If
                            End If
                        End If
                        pdfFormFields.SetField(pdfFieldFullPath + "B7_2", LastRate)
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B7_2", "000")
                    End If

                    If IsDBNull(dr("IH_RATE6")) = False And IsDBNull(dr("IH_APP_CHARGE_IN6")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B7_3", Replace((Convert.ToDouble((dr("IH_APP_CHARGE_IN6")) * CDbl(dr("IH_RATE6")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B7_3", "000")
                    End If

                    ''Total = (CDbl(dr("IH_APP_CHARGE_IN1")) * 0.05 * 100) + (CDbl(dr("IH_APP_CHARGE_IN2")) * 0.08 * 100) + _
                    ''        (CDbl(dr("IH_APP_CHARGE_IN3")) * 0.1 * 100) + (CDbl(dr("IH_APP_CHARGE_IN4")) * 0.15 * 100) + _
                    ''        (CDbl(dr("IH_APP_CHARGE_IN4A")) * 0.2 * 100) + (CDbl(dr("IH_APP_CHARGE_IN5")) * 0.26 * 100) + _
                    ''        (CDbl(dr("IH_APP_CHARGE_IN6")) * CDbl(dr("IH_RATE")) * 100)

                    If IsDBNull(dr("IH_TOTAL_TAX_CHARGED")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "B8", Replace((Convert.ToDouble((dr("IH_TOTAL_TAX_CHARGED"))).ToString("0.00")), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "B8", "000")
                    End If
                    If IsDBNull(dr("IH_SEC6B_REBATE")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B9", Replace((Convert.ToDouble((dr("IH_SEC6B_REBATE"))).ToString("0.00")), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B9", "000")
                    End If
                    If IsDBNull(dr("IH_ITP_SETOFF_110B")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B10", Replace((Convert.ToDouble((dr("IH_ITP_SETOFF_110B"))).ToString("0.00")), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B10", "000")
                    End If
                    If IsDBNull(dr("IH_ITP_SETOFF")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B11", Replace((Convert.ToDouble((dr("IH_ITP_SETOFF"))).ToString("0.00")), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B11", "000")
                    End If
                    If IsDBNull(dr("IH_ITP_SETOFF_OTH")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B12", Replace((Convert.ToDouble((dr("IH_ITP_SETOFF_OTH"))).ToString("0.00")), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B12", "000")
                    End If
                    If IsDBNull(dr("IH_SEC132")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B13", Replace((Convert.ToDouble((dr("IH_SEC132"))).ToString("0.00")), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B13", "000")
                    End If
                    If IsDBNull(dr("IH_SEC133")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B14", Replace((Convert.ToDouble((dr("IH_SEC133"))).ToString("0.00")), ".", "").Replace(",", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B14", "000")
                    End If
                    If IsDBNull(dr("IH_TOTAL_SETOFF")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B15", Replace((Convert.ToDouble((dr("IH_TOTAL_SETOFF"))) + Convert.ToDouble(dr("IH_SEC6B_REBATE"))).ToString("0.00"), ".", "").Replace(",", "").Replace("-", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B15", "000")
                    End If
                    If IsDBNull(dr("IH_ITP")) = False Then
                        If CDbl(dr("IH_ITP")) >= 0 Then
                            pdfFormFields.SetField(pdfFieldFullPath2 + "B16", Replace((Convert.ToDouble((dr("IH_ITP"))).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                            pdfFormFields.SetField(pdfFieldFullPath2 + "B17", "000")
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath2 + "B16", "000")
                            pdfFormFields.SetField(pdfFieldFullPath2 + "B17", Replace((Convert.ToDouble((dr("IH_ITP"))).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B16", "000")
                        pdfFormFields.SetField(pdfFieldFullPath2 + "B17", "000")
                    End If

                    ' Investment Holding part B
                    If IsDBNull(dr("IH_ITP")) = False Then
                        If (dr("IH_ITP")) >= 0 Then
                            pdfFormFields.SetField(pdfFieldFullPath2 + "C1", Replace((Convert.ToDouble((dr("IH_ITP"))).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath2 + "C1", "000")
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "C1", "000")
                    End If
                    If IsDBNull(dr("IH_INSTALLMENTS")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath2 + "C2", Replace((Convert.ToDouble((dr("IH_INSTALLMENTS"))).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "C2", "000")
                    End If
                    If IsDBNull(dr("IH_TP_BAL")) = False Then
                        If CDbl(dr("IH_TP_BAL")) >= 0 Then
                            pdfFormFields.SetField(pdfFieldFullPath2 + "C3", Replace((Convert.ToDouble((dr("IH_TP_BAL"))).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                            pdfFormFields.SetField(pdfFieldFullPath2 + "C4", "000")
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath2 + "C3", "000")
                            pdfFormFields.SetField(pdfFieldFullPath2 + "C4", Replace((Convert.ToDouble((dr("IH_TP_BAL"))).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath2 + "C3", "000")
                        pdfFormFields.SetField(pdfFieldFullPath2 + "C4", "000")
                    End If
                End If
            Else

                ' ===== If S60F NOT Checked , Run TAX COMPUTATION ====== '
                ' Tax Computation part A

                cSQL = "Select TC_TP_APP_CHARGEABLE1, TC_TP_APP_CHARGEABLE2, TC_TP_APP_CHARGEABLE3, TC_TP_APP_CHARGEABLE4, TC_TP_APP_CHARGEABLE4A, TC_TP_APP_CHARGEABLE5, TC_TP_APP_CHARGEABLE6, TC_TP_RATE1, TC_TP_RATE2, TC_TP_RATE3, TC_TP_RATE4, TC_TP_RATE4A, TC_TP_RATE5, TC_TP_RATE6," _
                        & " TC_TP_SEC6B_REBATE, TC_TP_SEC110, TC_TP_SEC110_OTHERS, TC_TP_SEC132, TC_TP_SEC133, TC_TP_INSTALL, TC_TP_RATE5_CHARGEABLE, TC_TP_PAYABLE, TC_TP_TOT_SETOFF , TC_TP_TOT_TAX_CHARGED, TC_TP_SEC110B," _
                        & " TC_BUSINESS" _
                        & " FROM TAX_COMPUTATION" _
                        & " WHERE [TC_REF_NO] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' And TC_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                        & " Order by TC_BUSINESS"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                dr.Read()
                'Add Checking statement here ( continue )
                'pdfFormFields.SetField(pdfFieldFullPath + "A1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "B2_1", CDbl(dr("TC_TP_APP_CHARGEABLE1")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B2_2", Replace((FormatNumber(Convert.ToDouble(CDbl(dr("TC_TP_APP_CHARGEABLE1")) * CDbl(dr("TC_TP_RATE1")) / 100))), ".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B3_1", CDbl(dr("TC_TP_APP_CHARGEABLE2")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B3_2", Replace((Convert.ToDouble(CDbl(dr("TC_TP_APP_CHARGEABLE2")) * CDbl(dr("TC_TP_RATE2")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B4_1", CDbl(dr("TC_TP_APP_CHARGEABLE3")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B4_2", Replace((Convert.ToDouble(CDbl(dr("TC_TP_APP_CHARGEABLE3")) * CDbl(dr("TC_TP_RATE3")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B5_1", CDbl(dr("TC_TP_APP_CHARGEABLE4")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B5_2", Replace((Convert.ToDouble(CDbl(dr("TC_TP_APP_CHARGEABLE4")) * CDbl(dr("TC_TP_RATE4")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B5A_1", CDbl(dr("TC_TP_APP_CHARGEABLE4A")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B5A_2", Replace((Convert.ToDouble(CDbl(dr("TC_TP_APP_CHARGEABLE4A")) * CDbl(dr("TC_TP_RATE4A")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B6_1", CDbl(dr("TC_TP_APP_CHARGEABLE5")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B6_2", Replace((Convert.ToDouble(CDbl(dr("TC_TP_APP_CHARGEABLE5")) * CDbl(dr("TC_TP_RATE5")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B7_1", CDbl(dr("TC_TP_APP_CHARGEABLE6")).ToString.Replace(",", ""))

                If IsDBNull(dr("TC_TP_RATE6")) = False Then
                    Dim LastRate As String
                    LastRate = CDbl(dr("TC_TP_RATE6"))
                    If Val(LastRate) = 0 Then
                        LastRate = "000"
                    Else
                        If Len(LastRate) = 1 Then
                            LastRate = " " + LastRate + "0"
                        Else
                            If Len(LastRate) = 2 Then
                                LastRate = LastRate + "0"
                            End If
                        End If
                    End If
                    pdfFormFields.SetField(pdfFieldFullPath + "B7_2", LastRate)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "B7_2", "000")
                End If

                'pdfFormFields.SetField(pdfFieldFullPath + "B7_2", Replace(CDbl(dr("TC_TP_RATE6")).ToString("00"), ".", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "B7_3", Replace((Convert.ToDouble(CDbl(dr("TC_TP_APP_CHARGEABLE6")) * CDbl(dr("TC_TP_RATE6")) / 100)).ToString("0.00"), ".", "").Replace(",", ""))

                pdfFormFields.SetField(pdfFieldFullPath + "B8", Replace((Convert.ToDouble(dr("TC_TP_TOT_TAX_CHARGED")).ToString("0.00")), ".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "B9", Replace((Convert.ToDouble(dr("TC_TP_SEC6B_REBATE")).ToString("0.00")), ".", "").Replace(",", ""))

                pdfFormFields.SetField(pdfFieldFullPath2 + "B10", Replace((Convert.ToDouble(dr("TC_TP_SEC110B")).ToString("0.00")), ".", "").Replace(",", ""))

                pdfFormFields.SetField(pdfFieldFullPath2 + "B11", Replace((Convert.ToDouble(dr("TC_TP_SEC110")).ToString("0.00")), ".", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "B12", Replace((Convert.ToDouble(dr("TC_TP_SEC110_OTHERS")).ToString("0.00")), ".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "B13", Replace((Convert.ToDouble(dr("TC_TP_SEC132")).ToString("0.00")), ".", "").Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "B14", Replace((Convert.ToDouble(dr("TC_TP_SEC133")).ToString("0.00")), ".", "").Replace(",", ""))

                pdfFormFields.SetField(pdfFieldFullPath2 + "B15", Replace((Convert.ToDouble(dr("TC_TP_TOT_SETOFF")) + Convert.ToDouble(dr("TC_TP_SEC6B_REBATE"))).ToString("0.00"), ".", "").Replace(",", "").Replace("-", ""))

                If CDbl(dr("TC_TP_PAYABLE")) >= 0 Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "B16", Replace((Convert.ToDouble(dr("TC_TP_PAYABLE")).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                    pdfFormFields.SetField(pdfFieldFullPath2 + "B17", "000")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "B16", "000")
                    pdfFormFields.SetField(pdfFieldFullPath2 + "B17", Replace((Convert.ToDouble(dr("TC_TP_PAYABLE")).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                End If

                ' Tax Computation part B

                If CDbl(dr("TC_TP_PAYABLE")) >= 0 Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "C1", Replace((Convert.ToDouble(dr("TC_TP_PAYABLE")).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "C1", "000")
                End If
                pdfFormFields.SetField(pdfFieldFullPath2 + "C2", Replace((Convert.ToDouble(dr("TC_TP_INSTALL")).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))

                Dim TotalB4 As Double
                If CDbl(dr("TC_TP_PAYABLE")) >= 0 Then
                    TotalB4 = Convert.ToDouble(dr("TC_TP_PAYABLE")) - Convert.ToDouble(dr("TC_TP_INSTALL"))
                Else
                    TotalB4 = 0 - Convert.ToDouble(dr("TC_TP_INSTALL"))
                End If

                If TotalB4 >= 0 Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "C3", Replace((TotalB4.ToString("0.00")), ".", "").Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath2 + "C4", "000")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "C3", "000")
                    pdfFormFields.SetField(pdfFieldFullPath2 + "C4", Replace((TotalB4.ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                End If

            End If

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 3!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page4()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page6[0]."
        Dim Total As Long = 0
        Dim I As Integer = 0 'for Multi Business Source

        Try

            ' Part D
            cSQL = "Select OE_CLAIMCODE, OE_AMOUNT" _
                    & " from OTHER_EXPENDITURE" _
                    & " WHERE [OE_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' AND [OE_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

            dr = DataHandler.GetDataReader(cSQL, Conn)

            Do While dr.Read()
                I = I + 1
                If IsDBNull(dr("OE_CLAIMCODE")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "D" + I.ToString + "_1", dr("OE_CLAIMCODE"))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "D" + I.ToString + "_1", "0")
                End If
                If IsDBNull(dr("OE_AMOUNT")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "D" + I.ToString + "_2", CDbl(dr("OE_AMOUNT")).ToString.Replace(",", ""))
                    Total = Total + CDbl(dr("OE_AMOUNT"))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "D" + I.ToString + "_2", "0")
                End If
            Loop
            While I < 10
                I = I + 1
                pdfFormFields.SetField(pdfFieldFullPath + "D" + I.ToString + "_2", "0")
            End While
            pdfFormFields.SetField(pdfFieldFullPath + "D11", Total)

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 4!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page5()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page7[0]."
        Dim TotalE5 As Double = 0.0
        Dim TotalE5b As Double = 0.0
        Dim TotalE11 As Double = 0.0
        Dim Total As Double = 0
        Dim I As Integer = 0
        Dim dblE11QC As Double = 0.0
        Dim dblE11RATE As Double = 0.0
        Dim dblE11IARATE As Double = 0.0
        Dim TotalF1b As Double = 0.0
        'NGOHCS 2009
        Dim dr2 As OleDbDataReader
        Dim TotalAA As Double = 0.0
        Dim TotalIA As Double = 0.0
        Dim boolAccelerated As String = False 'weihong
        Dim dblAAAccelerated As Double = 0.0 'weihong
        Dim dblIAAccelerated As Double = 0.0 'weihong
        Dim TotalE11a As Double = 0.0 'weihong
        Dim TotalE11b As Double = 0.0 'weihong
        Dim TotalUtilCA As Double = 0.0 'weihong
        Dim TotalCAForYr As Double = 0.0 'weihong
        Dim TotalAcceleratedCA As Double = 0.0 'weihong
        Dim TotalBalBF As Double = 0.0 'weihong


        'Ho Gie - e6a to e10a and e6b to e10b
        Dim rowCount As Integer = 0

        Dim TotalPerkongsian As Long = 0
        Dim TotalBakiHantarHadapan As Long = 0
        'Ho Gie - e6a to e10a and e6b to e10b End 

        Try
            ' === continue here === '
            ' Part E
            'weihong add TC_CB_CA_CURR, TC_CB_CA_ACA, TC_CB_CA_BAL_BF
            cSQL = "Select TC_CB_CA_UTIL, TC_CB_CA_ABAL_CF, TC_CB_CA_CURR, TC_CB_CA_ACA, TC_CB_CA_BAL_BF" _
                    & " from TAX_COMPUTATION" _
                    & " WHERE [TC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' AND [TC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                    & " Order By TC_BUSINESS"

            dr = DataHandler.GetDataReader(cSQL, Conn)
            I = 0

            Do While dr.Read()
                I = I + 1
                If I <= 10 Then
                    If I < 6 Then
                        If IsDBNull(dr("TC_CB_CA_UTIL")) = False Then
                            If I <> 3 Then
                                pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "a", CDbl(dr("TC_CB_CA_UTIL")).ToString.Replace(",", ""))
                                ''weihong
                                'TotalUtilCA = TotalUtilCA + CDbl(dr("TC_CB_CA_UTIL"))
                                ''endweihong
                            Else
                                pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "a", CDbl(dr("TC_CB_CA_UTIL")).ToString.Replace(",", ""))
                                ''weihong
                                'TotalUtilCA = TotalUtilCA + CDbl(dr("TC_CB_CA_UTIL"))
                                ''endweihong
                            End If

                            If I = 5 Then
                                TotalE5 = TotalE5 + CDbl(dr("TC_CB_CA_UTIL"))
                                ''weihong
                                'TotalUtilCA = TotalUtilCA + CDbl(dr("TC_CB_CA_UTIL"))
                                ''endweihong
                            End If
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "a", "0")
                        End If
                    Else
                        TotalE5 = TotalE5 + CDbl(dr("TC_CB_CA_UTIL"))
                        ''weihong
                        'TotalUtilCA = TotalUtilCA + CDbl(dr("TC_CB_CA_UTIL"))
                        ''endweihong
                        pdfFormFields.SetField(pdfFieldFullPath + "E5a", TotalE5.ToString.Replace(",", ""))
                    End If
                    ''weihong
                    'TotalCAForYr = TotalCAForYr + CDbl(dr("TC_CB_CA_CURR"))
                    TotalAcceleratedCA = TotalAcceleratedCA + CDbl(dr("TC_CB_CA_ACA"))
                    'TotalBalBF = TotalBalBF + CDbl(dr("TC_CB_CA_BAL_BF"))
                    ''endweihong
                    If I < 6 Then
                        If IsDBNull(dr("TC_CB_CA_ABAL_CF")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "b", CDbl(dr("TC_CB_CA_ABAL_CF")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "b", "0")
                        End If
                        If I = 5 Then
                            TotalE5b = TotalE5b + CDbl(dr("TC_CB_CA_ABAL_CF"))
                        End If
                    Else
                        TotalE5b = TotalE5b + CDbl(dr("TC_CB_CA_ABAL_CF"))
                        pdfFormFields.SetField(pdfFieldFullPath + "E5b", TotalE5b.ToString.Replace(",", ""))
                    End If
                    'Total = Total + CDbl(dr("OE_AMOUNT"))
                End If
            Loop
            While I < 5
                I = I + 1
                If I <> 3 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "a", "0")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "a", "0")
                End If
                pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "b", "0")

            End While


            'Ho gie - TUNTUTAN ELAUN JADUAL 3 - e6a to e10a and e6b to e10b
            cSQL = "SELECT PN_SOURCENO , PN_CA_BA , PN_CF " & _
                    "FROM INCOME_PARTNERSHIP IP , TAX_COMPUTATION_PARTNER TCP " & _
                    "WHERE TCP.TCP_KEY = IP.TCP_KEY " & _
                    "AND [TCP_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "'" & _
                    " AND [TCP_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" & _
                    " ORDER BY  PN_SOURCENO"



            dr = DataHandler.GetDataReader(cSQL, Conn)
            I = 5

            Do While dr.Read()

                rowCount += 1

                I = I + 1
                If rowCount < 5 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "a", CDbl(dr("PN_CA_BA")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "E" + I.ToString + "b", CDbl(dr("PN_CF")).ToString.Replace(",", ""))
                    'ElseIf I >= 10 Then
                ElseIf rowCount > 4 Then
                    ' Total Perniagaan for source more than 5 
                    TotalPerkongsian = TotalPerkongsian + CDbl(dr("PN_CA_BA"))
                    TotalBakiHantarHadapan = TotalBakiHantarHadapan + CDbl(dr("PN_CF"))
                    pdfFormFields.SetField(pdfFieldFullPath + "E10a", TotalPerkongsian.ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "E10b", TotalBakiHantarHadapan.ToString.Replace(",", ""))
                End If
            Loop

            ''Ho gie - TUNTUTAN ELAUN JADUAL 3 - e6a to e10a and e6b to e10b - end

            'pdfFormFields.SetField(pdfFieldFullPath + "E6a", "0")
            'pdfFormFields.SetField(pdfFieldFullPath + "E6b", "0")
            'pdfFormFields.SetField(pdfFieldFullPath + "E7a", "0")
            'pdfFormFields.SetField(pdfFieldFullPath + "E7b", "0")
            'pdfFormFields.SetField(pdfFieldFullPath + "E8a", "0")
            'pdfFormFields.SetField(pdfFieldFullPath + "E8b", "0")
            'pdfFormFields.SetField(pdfFieldFullPath + "E9a", "0")
            'pdfFormFields.SetField(pdfFieldFullPath + "E9b", "0")
            'pdfFormFields.SetField(pdfFieldFullPath + "E10a", "0")
            'pdfFormFields.SetField(pdfFieldFullPath + "E10b", "0")

            pdfFormFields.SetField(pdfFieldFullPath + "E11", "0")
            ' E11
            ' HS : C2008.7 : CA_ADD_CURR_AMT Added

            'NGOHCS CA2008  
            'weihong Add CA_ACCELERATED
            cSQL = "Select CA_RATE_AA, CA_QUALIFYING_COST, CA_RATE_IA, CA_REMAIN_QC, CA_TWDV" _
                & " from CA" _
                & " WHERE [CA_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' AND [CA_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "' AND" _
                & " [CA_MODE] = 'ADD' AND CA_KEY NOT IN (SELECT DISTINCT CA_KEY FROM CA_DISPOSAL WHERE CA_DISP_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "')"
            dr = DataHandler.GetDataReader(cSQL, ConnCA)
            Do While dr.Read()
                If IsDBNull(dr("CA_RATE_AA")) Then
                    dblE11RATE = 0
                Else
                    dblE11RATE = CDbl(dr("CA_RATE_AA"))
                End If
                If IsDBNull(dr("CA_QUALIFYING_COST")) Then
                    dblE11QC = 0
                Else
                    dblE11QC = CDbl(dr("CA_QUALIFYING_COST"))
                End If
                If IsDBNull(dr("CA_RATE_IA")) Then
                    dblE11IARATE = 0
                Else
                    dblE11IARATE = CDbl(dr("CA_RATE_IA"))
                End If

                If IsDBNull(dr("CA_RATE_AA")) = False Or IsDBNull(dr("CA_QUALIFYING_COST")) = False Or IsDBNull(dr("CA_RATE_IA")) = False Then
                    'LeeCC 2011.5 ctrl transfer
                    Dim dblAA As Double
                    Dim dblIA As Double

                    'TotalAA = TotalAA + Math.Round((CDbl(dr("CA_RATE_AA")) / 100) * CDbl(dr("CA_QUALIFYING_COST")), 2)
                    'TotalIA = TotalIA + Math.Round((CDbl(dr("CA_RATE_IA")) / 100) * CDbl(dr("CA_QUALIFYING_COST")), 2)

                    'LeeCC 2011.5 ctrl transfer 
                    dblAA = Math.Round((CDbl(dr("CA_RATE_AA")) / 100) * CDbl(dr("CA_QUALIFYING_COST")), 2)
                    dblIA = Math.Round((CDbl(dr("CA_RATE_IA")) / 100) * CDbl(dr("CA_QUALIFYING_COST")), 2)

                    If (dblAA + dblIA) > CDbl(dr("CA_TWDV")) Then
                        dblAA = CDbl(dr("CA_TWDV"))
                        dblIA = 0
                    End If
                    TotalAA += dblAA
                    TotalIA += dblIA
                End If
            Loop

            'weihong Add CA_ACCELERATED
            cSQL = "Select CA_KEY, CA_RATE_AA, CA_QUALIFYING_COST, CA_RATE_IA, CA_REMAIN_QC, CA_TWDV" _
                           & " from CA" _
                           & " WHERE [CA_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' AND [CA_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "' AND" _
                           & " [CA_MODE] = 'ADD' AND CA_KEY IN (SELECT DISTINCT CA_KEY FROM CA_DISPOSAL WHERE CA_DISP_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "')"

            dr = DataHandler.GetDataReader(cSQL, ConnCA)
            Do While dr.Read()
                If IsDBNull(dr("CA_RATE_AA")) Then
                    dblE11RATE = 0
                Else
                    dblE11RATE = CDbl(dr("CA_RATE_AA"))
                End If
                If IsDBNull(dr("CA_REMAIN_QC")) Then
                    dblE11QC = 0
                Else
                    dblE11QC = CDbl(dr("CA_REMAIN_QC"))
                End If
                If IsDBNull(dr("CA_RATE_IA")) Then
                    dblE11IARATE = 0
                Else
                    dblE11IARATE = CDbl(dr("CA_RATE_IA"))
                End If


                If IsDBNull(dr("CA_RATE_AA")) = False Or IsDBNull(dr("CA_REMAIN_QC")) = False Or IsDBNull(dr("CA_RATE_IA")) = False Then
                    cSQL = "SELECT count([CA_KEY]) as [NumRecord], sum(CA_DISP_QC) as [CA_IA_TOTAL], sum(CA_DISP_TWDV) from CA_DISPOSAL where CA_DISP_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "' and ca_key= " & Val(dr("CA_KEY"))
                    dr2 = DataHandler.GetDataReader(cSQL, ConnCA)

                    'LeeCC 2011.5 ctrl transfer
                    Dim dblAA As Double
                    Dim dblIA As Double

                    If dr2.Read() Then
                        If Val(dr2("NumRecord")) > 0 Then
                            'TotalAA = TotalAA + Math.Round((CDbl(dr("CA_RATE_AA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST")) - CDbl(dr2("CA_IA_TOTAL"))), 2)
                            'TotalIA = TotalIA + Math.Round((CDbl(dr("CA_RATE_IA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST")) - CDbl(dr2("CA_IA_TOTAL"))), 2)

                            'LeeCC 2011.5 ctrl transfer 
                            dblAA = Math.Round((CDbl(dr("CA_RATE_AA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST")) - CDbl(dr2("CA_IA_TOTAL"))), 2)
                            dblIA = Math.Round((CDbl(dr("CA_RATE_IA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST")) - CDbl(dr2("CA_IA_TOTAL"))), 2)

                            If (dblAA + dblIA) > (CDbl(dr("CA_TWDV")) - CDbl(dr(2))) Then
                                dblAA = CDbl(dr("CA_TWDV")) - CDbl(dr2(2))
                                dblIA = 0
                            End If
                        Else
                            'TotalAA = TotalAA + Math.Round((CDbl(dr("CA_RATE_AA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST"))), 2)
                            'TotalIA = TotalIA + Math.Round((CDbl(dr("CA_RATE_IA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST"))), 2)

                            'LeeCC 2011.5 ctrl transfer 
                            dblAA = Math.Round((CDbl(dr("CA_RATE_AA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST"))), 2)
                            dblIA = Math.Round((CDbl(dr("CA_RATE_IA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST"))), 2)

                            If (dblAA + dblIA) > CDbl(dr("CA_TWDV")) Then
                                dblAA = CDbl(dr("CA_TWDV"))
                                dblIA = 0
                            End If
                        End If
                    Else
                        'TotalAA = TotalAA + Math.Round((CDbl(dr("CA_RATE_AA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST"))), 2)
                        'TotalIA = TotalIA + Math.Round((CDbl(dr("CA_RATE_IA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST"))), 2)

                        'LeeCC 2011.5 ctrl transfer 
                        dblAA = Math.Round((CDbl(dr("CA_RATE_AA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST"))), 2)
                        dblIA = Math.Round((CDbl(dr("CA_RATE_IA")) / 100) * (CDbl(dr("CA_QUALIFYING_COST"))), 2)

                        If (dblAA + dblIA) > CDbl(dr("CA_TWDV")) Then
                            dblAA = CDbl(dr("CA_TWDV"))
                            dblIA = 0
                        End If
                    End If
                    dr2.Close()

                    TotalAA += dblAA
                    TotalIA += dblIA
                End If
            Loop
            'NGOHCS CA2008 END

            'weihong Total accelerated is dblAAAccelerated + dblIAAccelerated
            If TotalAcceleratedCA > 0 Then
                'LeeCC Modify Accelerated CA SU3.4 start
                cSQL = "SELECT TC_BUSINESS FROM TAX_COMPUTATION WHERE TC_REF_NO = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "'" & _
                " AND TC_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

                dr2 = DataHandler.GetDataReader(cSQL, Conn)

                Dim dblUtilizedCACurr As Double = 0
                Dim dblUtilizedACA As Double = 0

                While dr2.Read
                    cSQL = "SELECT TC_CB_CA_CURR, TC_CB_CA_UTIL, TC_CB_CA_BAL_BF, TC_CB_CA_ACA FROM TAX_COMPUTATION" & _
                    " WHERE TC_REF_NO = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "'" & _
                    " AND TC_YA = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" & _
                    " AND TC_BUSINESS = " & dr2(0).ToString

                    dr3 = DataHandler.GetDataReader(cSQL, Conn)

                    While dr3.Read
                        If CDbl(dr3(3)) > 0 Then
                            dblUtilizedCACurr = CDbl(dr3(1)) - (CDbl(dr3(0)) + CDbl(dr3(2)) - CDbl(dr3(3)))
                            If dblUtilizedCACurr <= 0 Then
                                TotalE11b += CDbl(dr3(3))
                            Else
                                dblUtilizedACA = (dblUtilizedCACurr - CDbl(dr3(3)))
                                If dblUtilizedACA = 0 Then
                                    TotalE11a += (CDbl(dr3(3)))
                                Else
                                    TotalE11a += dblUtilizedCACurr
                                    TotalE11b += (-(dblUtilizedACA))
                                End If
                            End If
                        End If
                    End While
                    dr3.Close()
                End While
                dr2.Close()
                'TotalE11a = TotalUtilCA - (TotalCAForYr + TotalBalBF - TotalAcceleratedCA)
                'If TotalE11a > 0 Then
                '    pdfFormFields.SetField(pdfFieldFullPath + "E11a", Format(TotalE11a, 0).ToString.Replace(",", ""))
                '    TotalE11b = TotalAcceleratedCA - TotalE11a
                '    If TotalE11b > 0 Then
                '        pdfFormFields.SetField(pdfFieldFullPath + "E11b", Format(TotalE11b, 0).ToString.Replace(",", ""))
                '    Else
                '        pdfFormFields.SetField(pdfFieldFullPath + "E11b", 0)
                '    End If
                'Else
                '    pdfFormFields.SetField(pdfFieldFullPath + "E11a", 0)
                '    pdfFormFields.SetField(pdfFieldFullPath + "E11b", Format(TotalAcceleratedCA, 0).ToString.Replace(",", ""))
                'End If
            Else
                TotalE11a = 0
                TotalE11b = 0
            End If
            pdfFormFields.SetField(pdfFieldFullPath + "E11a", TotalE11a)
            pdfFormFields.SetField(pdfFieldFullPath + "E11b", TotalE11b)
            'LeeCC Modify Accelerated CA SU3.4 end
            'endweihong

            TotalE11 = Math.Round(TotalAA, 0) + Math.Round(TotalIA, 0)
            pdfFormFields.SetField(pdfFieldFullPath + "E11", Format(TotalE11, 0).ToString.Replace(",", ""))
            ' E12
            cSQL = "Select Sum(TC_CB_CA_DISALLOW)" _
                   & " from TAX_COMPUTATION" _
                   & " WHERE [TC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' AND [TC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                If IsDBNull(dr(0)) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "E12", CDbl(dr(0)).ToString.Replace(",", "")) 'E12 Get from Tax Computation
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "E12", "0")
                End If
            End If

            ' ===== Part F =====
            cSQL = "Select TC_CB_LS_D_BL, TC_CB_LS_BALS_CF, TC_CB_LS_SAMOUNT, TC_CB_LS_BLNCF,TC_NB_AMTNOTCARRYBCK" _
                 & " from TAX_COMPUTATION" _
                 & " WHERE [TC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                 & "' AND [TC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                 & " Order By TC_BUSINESS"
            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then
                TotalF1b = CDbl(dr("TC_CB_LS_BALS_CF"))
                If TotalF1b < 0 Then
                    TotalF1b = 0.0
                End If
                pdfFormFields.SetField(pdfFieldFullPath + "F1a", CDbl(dr("TC_CB_LS_D_BL")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "F1b", TotalF1b.ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "F1Aa", CDbl(dr("TC_CB_LS_SAMOUNT")).ToString.Replace(",", ""))
                If CDbl(dr("TC_CB_LS_SAMOUNT")) = 0 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "F1Ab", "0")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "F1Ab", CDbl(dr("TC_NB_AMTNOTCARRYBCK")).ToString.Replace(",", ""))
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "F1a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F1b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F1Aa", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F1Ab", "0")
            End If

            ' ===== Part F : New Field  =====
            cSQL = "Select  TC_NB_CARRYBCKLOSS, TC_NB_AMTNOTCARRYBCK, TC_NB_CHKCB" _
            & " from TAX_COMPUTATION" _
            & " WHERE [TC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
            & "' AND [TC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
            & " Order By TC_BUSINESS"
            dr = DataHandler.GetDataReader(cSQL, Conn)

            '==Oscar2009.1=='
            ' Do not show amount in Part F1B (b) of Borang C 2009 when the check box for Carry-back losses is not selected. 
            If dr.Read() Then
                pdfFormFields.SetField(pdfFieldFullPath + "F1Ba", CDbl(dr("TC_NB_CARRYBCKLOSS")).ToString.Replace(",", ""))
                If (dr("TC_NB_CHKCB")) = 0 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "F1Bb", "0")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "F1Bb", CDbl(dr("TC_NB_AMTNOTCARRYBCK")).ToString.Replace(",", ""))
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "F1Ba", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F1Bb", "0")
            End If
            '==End Oscar2009.1=='

            'Original Coding
            'If dr.Read() Then
            '    pdfFormFields.SetField(pdfFieldFullPath + "F1Ba", CDbl(dr("TC_NB_CARRYBCKLOSS")).ToString.Replace(",", ""))
            '    pdfFormFields.SetField(pdfFieldFullPath + "F1Bb", CDbl(dr("TC_NB_AMTNOTCARRYBCK")).ToString.Replace(",", ""))
            'Else
            '    pdfFormFields.SetField(pdfFieldFullPath + "F1Ba", "0")
            '    pdfFormFields.SetField(pdfFieldFullPath + "F1Bb", "0")
            'End If

            ' Get Data from Loss Claim
            cSQL = "Select LC_PIONEER_AMT, LC_PIONEER_BF, LC_SERVICE_AMT, LC_SERVICE_BF," _
                    & " LC_HQ_AMT, LC_HQ_BF, LC_SHIP_AMT, LC_SHIP_BF" _
                    & " from LOSS_CLAIM" _
                    & " WHERE [LC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' AND [LC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then
                pdfFormFields.SetField(pdfFieldFullPath + "F2a", CDbl(dr("LC_PIONEER_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "F2b", CDbl(dr("LC_PIONEER_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "F3a", CDbl(dr("LC_SERVICE_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "F3b", CDbl(dr("LC_SERVICE_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "F4a", CDbl(dr("LC_HQ_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "F4b", CDbl(dr("LC_HQ_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "F5a", CDbl(dr("LC_SHIP_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "F5b", CDbl(dr("LC_SHIP_BF")).ToString.Replace(",", ""))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "F2a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F2b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F3a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F3b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F4a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F4b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F5a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "F5b", "0")
            End If

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 5!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page6()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page8[0]."
        Dim Total As Double = 0
        Dim TotalClaim As Double = 0
        Dim Total1 As Double = 0
        Dim Total2 As Double = 0
        Dim Total3 As Double = 0

        Try
            ' Get Data from Incentive Claim
            cSQL = "Select IC_INVEST_AMT, IC_INVEST_BF, IC_INDUST_AMT, IC_INDUST_BF," _
                    & " IC_INFRA_AMT, IC_INFRA_BF, IC_SECT7A_AMT, IC_SECT7A_BF," _
                    & " IC_SECT7B_AMT, IC_SECT7B_BF," _
                    & " IC_EXPORT_AMT, IC_EXPORT_BF, IC_AGRI_AMT, IC_AGRI_BF," _
                    & " IC_INTCO_AMT, IC_INTCO_BF, IC_SERVICE_AMT, IC_SERVICE_BF," _
                    & " IC_SPECIAL_AMT, IC_SPECIAL_BF, IC_BIO_AMT, IC_BIO_CF," _
                    & " IC_SCH4, IC_SCH4B, IC_TOTAL, IC_TRANSFER" _
                    & " from INCENTIVE_CLAIM" _
                    & " WHERE [IC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' AND [IC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then

                pdfFormFields.SetField(pdfFieldFullPath + "G1a", CDbl(dr("IC_INVEST_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G1b", CDbl(dr("IC_INVEST_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G2a", CDbl(dr("IC_INDUST_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G2b", CDbl(dr("IC_INDUST_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G3a", CDbl(dr("IC_INFRA_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G3b", CDbl(dr("IC_INFRA_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G4a", CDbl(dr("IC_SECT7A_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G4b", CDbl(dr("IC_SECT7A_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G5a", CDbl(dr("IC_SECT7B_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G5b", CDbl(dr("IC_SECT7B_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6a", CDbl(dr("IC_EXPORT_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6b", CDbl(dr("IC_EXPORT_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6Aa", CDbl(dr("IC_AGRI_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6Ab", CDbl(dr("IC_AGRI_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6Ba", CDbl(dr("IC_INTCO_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6Bb", CDbl(dr("IC_INTCO_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6Ca", CDbl(dr("IC_SERVICE_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6Cb", CDbl(dr("IC_SERVICE_BF")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6Da", CDbl(dr("IC_SPECIAL_AMT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G6Db", CDbl(dr("IC_SPECIAL_BF")).ToString.Replace(",", ""))
                If IsDBNull(dr("IC_BIO_AMT")) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "G6Ea", "0")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "G6Ea", CDbl(dr("IC_BIO_AMT")).ToString.Replace(",", ""))
                End If
                If IsDBNull(dr("IC_BIO_CF")) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "G6Eb", "0")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "G6Eb", CDbl(dr("IC_BIO_CF")).ToString.Replace(",", ""))
                End If

                pdfFormFields.SetField(pdfFieldFullPath + "G7", CDbl(dr("IC_SCH4")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G8", CDbl(dr("IC_SCH4B")).ToString.Replace(",", ""))

                ' [G1(a)- G6A(a)]
                If IsDBNull(dr("IC_BIO_AMT")) Then
                    Total = CDbl(dr("IC_INVEST_AMT")) + CDbl(dr("IC_INDUST_AMT")) + CDbl(dr("IC_INFRA_AMT")) + _
                            CDbl(dr("IC_SECT7A_AMT")) + CDbl(dr("IC_SECT7B_AMT")) + CDbl(dr("IC_EXPORT_AMT")) + _
                            CDbl(dr("IC_AGRI_AMT")) + CDbl(dr("IC_INTCO_AMT")) + CDbl(dr("IC_SERVICE_AMT")) + _
                            CDbl(dr("IC_SPECIAL_AMT")) + 0
                Else
                    Total = CDbl(dr("IC_INVEST_AMT")) + CDbl(dr("IC_INDUST_AMT")) + CDbl(dr("IC_INFRA_AMT")) + _
                            CDbl(dr("IC_SECT7A_AMT")) + CDbl(dr("IC_SECT7B_AMT")) + CDbl(dr("IC_EXPORT_AMT")) + _
                            CDbl(dr("IC_AGRI_AMT")) + CDbl(dr("IC_INTCO_AMT")) + CDbl(dr("IC_SERVICE_AMT")) + _
                            CDbl(dr("IC_SPECIAL_AMT")) + CDbl(dr("IC_BIO_AMT"))
                End If

                cSQL = "SELECT TC_TP_PROSPECTING , TC_TP_PREOP_BS, TC_TP_QUALIFY_AGRICULTURE " _
                        & "FROM TAX_COMPUTATION " _
                        & "WHERE TC_REF_NO = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                        & "' AND [TC_YA]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "' AND [TC_BUSINESS]=1"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                Do While dr.Read()
                    Total1 = Total1 + CDbl(dr("TC_TP_PROSPECTING"))
                    Total2 = Total2 + CDbl(dr("TC_TP_PREOP_BS"))
                    Total3 = Total3 + CDbl(dr("TC_TP_QUALIFY_AGRICULTURE"))

                    'pdfFormFields.SetField(pdfFieldFullPath + "G9", (Total + CDbl(dr("TC_TP_PROSPECTING")) + CDbl(dr("TC_TP_PREOP_BS"))).ToString.Replace(",", ""))  '[G1(a)- G6A(a)] + C22 + C23
                    'pdfFormFields.SetField(pdfFieldFullPath + "G9", Total.ToString.Replace(",", ""))  '[G1(a)- G6A(a)] + C22 + C23
                Loop
                TotalClaim = Total + Total1 + Total2 + Total3

                pdfFormFields.SetField(pdfFieldFullPath + "G9", TotalClaim.ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "G10", Total.ToString.Replace(",", ""))
                'pdfFormFields.SetField(pdfFieldFullPath + "G9", CDbl(dr("IC_TOTAL")).ToString.Replace(",", ""))
                'pdfFormFields.SetField(pdfFieldFullPath + "G10", CDbl(dr("IC_TRANSFER")).ToString.Replace(",", ""))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "G1a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G1b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G2a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G2b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G3a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G3b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G4a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G4b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G5a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G5b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6a", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6b", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Aa", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Ab", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Ba", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Bb", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Ca", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Cb", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Da", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Db", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Ea", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G6Eb", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G7", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G8", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G9", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "G10", "0")
            End If

            ' Bahagian H : fill the code acccording to the top 10 description
            ' Get Data from Income Transfer
            cSQL = "SELECT * FROM [INCOME_TRANSFER]" _
                        & " WHERE [IT_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                        & "' And IT_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then
                Dim dic As New Dictionary(Of String, String)
                Dim counter As Integer = 0
                Dim totalDic As Double

                dic.Add("501", dr("IT_1"))
                dic.Add("502", dr("IT_2"))
                dic.Add("503", dr("IT_3"))
                dic.Add("504", dr("IT_4"))
                dic.Add("505", dr("IT_5"))
                dic.Add("506", dr("IT_6"))
                dic.Add("507", dr("IT_6A"))
                dic.Add("508", dr("IT_7"))
                dic.Add("509", dr("IT_7A"))
                dic.Add("510", dr("IT_7B"))
                dic.Add("511", dr("IT_8"))
                dic.Add("512", dr("IT_9"))
                dic.Add("513", dr("IT_10"))
                dic.Add("514", dr("IT_11"))
                dic.Add("515", dr("IT_12"))
                dic.Add("516", dr("IT_13"))
                dic.Add("517", dr("IT_14"))
                dic.Add("518", dr("IT_15"))
                dic.Add("519", dr("IT_16"))
                dic.Add("520", dr("IT_18"))
                dic.Add("521", dr("IT_19"))
                dic.Add("522", dr("IT_20"))
                dic.Add("523", dr("IT_21"))
                dic.Add("524", dr("IT_23"))
                dic.Add("525", dr("IT_24"))
                dic.Add("526", dr("IT_25"))
                dic.Add("527", dr("IT_527"))
                dic.Add("528", dr("IT_528"))
                dic.Add("529", dr("IT_529"))
                dic.Add("530", dr("IT_530"))

                For Each key As String In dic.Keys
                    If counter <= 9 Then
                        If dic(key) <> "0" Then
                            'MsgBox(key.ToString + "," + dic(key).ToString)
                            counter = counter + 1
                            totalDic = totalDic + CDbl(dic(key).ToString)
                            pdfFormFields.SetField(pdfFieldFullPath + "H" + counter.ToString + "_1", key.ToString)
                            pdfFormFields.SetField(pdfFieldFullPath + "H" + counter.ToString + "_2", _
                                CDbl(dic(key)).ToString.Replace(",", ""))
                        Else
                            countSpace = countSpace + 1
                        End If
                    End If
                Next
                '===NgKL C2010.1==='
                If countSpace > 20 Then
                    counter = counter + 1
                    For i = counter To 10
                        pdfFormFields.SetField(pdfFieldFullPath + "H" + i.ToString + "_2", "0")
                    Next
                End If
                '===NgKL C20101.1 End==='
                pdfFormFields.SetField(pdfFieldFullPath + "H11", totalDic.ToString.Replace(",", ""))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "H1_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H2_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H3_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H4_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H5_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H6_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H7_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H8_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H9_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H10_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "H11", "0")
            End If

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 6!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page7()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page9[0]."
        Dim pdfFieldFullPath2 As String = pdfFieldPath + "Page10[0]."
        Dim Total As Double = 0
        Dim Total4A As Double
        Dim Total6 As Double
        Dim UntungAsing As Double = 0 'weihong
        Dim strPLKey As String
        Dim dr2 As OleDbDataReader

        Try
            ' =========Part I==============
            cSQL = "SELECT EA_CREDIT, EA_EXEMPT" _
         & " FROM [EXEMPT_ACCOUNT]" _
         & " WHERE [EA_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
         & "' And EA_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then
                If IsDBNull(dr("EA_CREDIT")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "I1", CDbl(dr("EA_CREDIT")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "I1", "0")
                End If
                If IsDBNull(dr("EA_EXEMPT")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "I2", CDbl(dr("EA_EXEMPT")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "I2", "0")
                End If
                If IsDBNull(dr("EA_CREDIT")) = False Or IsDBNull(dr("EA_EXEMPT")) = False Then
                    If (CDbl(dr("EA_CREDIT")) - CDbl(dr("EA_EXEMPT"))) < 0 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "I3_1", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "I3_1", "")
                    End If
                    pdfFormFields.SetField(pdfFieldFullPath + "I3_2", Replace(CDbl(dr("EA_CREDIT")) - CDbl(dr("EA_EXEMPT")), "-", "").ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "I3_1", "")
                    pdfFormFields.SetField(pdfFieldFullPath + "I3_2", "0")
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "I1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "I2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "I3_1", "")
                pdfFormFields.SetField(pdfFieldFullPath + "I3_2", "0")
            End If

            ' =======part J=========
            cSQL = "SELECT PY_INCOME" _
                   & " FROM [PRECEDING_YEAR]" _
                   & " WHERE [PY_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                   & "' And PY_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                If IsDBNull(dr("PY_INCOME")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "J1", CDbl(dr("PY_INCOME")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "J1", "0")
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "J1", "0")
            End If

            ' ======== Part K===========
            cSQL = "SELECT DP_DISPOSAL, DP_DECLARE" _
            & " FROM [DISPOSAL]" _
            & " WHERE [DP_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
            & "' And DP_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                If Trim(dr("DP_DISPOSAL").ToString) = "Yes" Then
                    If Trim(dr("DP_DECLARE").ToString) = "Yes" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "K1_1", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "K2_1", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "K1_2", " ")
                        pdfFormFields.SetField(pdfFieldFullPath + "K2_2", " ")
                    ElseIf Trim(dr("DP_DECLARE").ToString) = "No" Or Trim(dr("DP_DECLARE").ToString) = "" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "K1_1", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "K2_1", " ")
                        pdfFormFields.SetField(pdfFieldFullPath + "K1_2", " ")
                        pdfFormFields.SetField(pdfFieldFullPath + "K2_2", "X")
                    End If
                ElseIf Trim(dr("DP_DISPOSAL").ToString) = "No" Or Trim(dr("DP_DISPOSAL").ToString) = "" Then
                    If dr("DP_DECLARE").ToString = "Yes" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "K1_1", " ")
                        'LEESH 06-APR-2012
                        pdfFormFields.SetField(pdfFieldFullPath + "K2_1", " ")
                        'LEESH END
                        pdfFormFields.SetField(pdfFieldFullPath + "K1_2", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "K2_2", " ")
                    ElseIf Trim(dr("DP_DECLARE").ToString) = "No" Or Trim(dr("DP_DECLARE").ToString) = "" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "K1_1", " ")
                        pdfFormFields.SetField(pdfFieldFullPath + "K2_1", " ")
                        pdfFormFields.SetField(pdfFieldFullPath + "K1_2", "X")
                        'LEESH 06-APR-2012
                        pdfFormFields.SetField(pdfFieldFullPath + "K2_2", " ")
                        'LEESH END
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "K1_1", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "K2_1", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "K1_2", "X")
                    'LEESH 06-APR-2012
                    pdfFormFields.SetField(pdfFieldFullPath + "K2_2", " ")
                    'LEESH END
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "K1_1", " ")
                pdfFormFields.SetField(pdfFieldFullPath + "K2_1", " ")
                pdfFormFields.SetField(pdfFieldFullPath + "K1_2", "X")
                'LEESH 06-APR-2012
                pdfFormFields.SetField(pdfFieldFullPath + "K2_2", " ")
                'LEESH END
            End If

            ' ======== Part L===========
            L8GetBSCode()

            Total = L8Calculation()

            pdfFormFields.SetField(pdfFieldFullPath + "L7_1", "")
            'L2  - L8 , L11 - L18 from Profit And Loss
            cSQL = "SELECT PL_SALES, PL_OP_STK, PL_PURCHASES, PL_PRO_COST, PL_CLS_STK," _
                    & " PL_EXP_INT, PL_EXP_INTRESTRICT, PL_LAWYER_COST, PL_TECH_FEE, PL_CONTRACT_EXP," _
                    & " PL_EXP_SALARY, PL_EMPL_STOCK, PL_ROYALTY, PL_EXP_RENT, PL_EXP_MAINTENANCE," _
                    & " PL_OTH_BSIN_NONSOURCE, PL_RND, PL_KEY, PL_COGS, PL_OTH_BSIN , PL_NONTAX_IN" _
                    & " FROM [PROFIT_LOSS_ACCOUNT]" _
                    & " WHERE [PL_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' And PL_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

            dr = DataHandler.GetDataReader(cSQL, Conn)



            strPLKey = 0
            If dr.Read() Then
                pdfFormFields.SetField(pdfFieldFullPath + "L1", Mid(BSCode, 1, 5))
                pdfFormFields.SetField(pdfFieldFullPath + "L2", dSales.ToString) 'CDbl(dr("PL_SALES")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "L3", dOS.ToString) 'CDbl(dr("PL_OP_STK")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "L4", dPur.ToString) 'CDbl(dr("PL_PURCHASES")).ToString.Replace(",", ""))
                Total4A = dDep + dA + dNA
                pdfFormFields.SetField(pdfFieldFullPath + "L4A", CDbl(Total4A).ToString) 'CDbl(dr("PL_PRO_COST")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "L5", CDbl(dCS).ToString)
                'Total = CDbl(dr("PL_OP_STK")) + CDbl(dr("PL_PURCHASES")) + CDbl(dr("PL_PURCHASES")) - CDbl(dr("PL_CLS_STK"))
                Total6 = dOS + dPur + Total4A - dCS
                pdfFormFields.SetField(pdfFieldFullPath + "L6", CDbl(Total6).ToString)

                If (dSales - Total6) < 0 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "L7_1", "X")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L7_1", "")
                End If

                pdfFormFields.SetField(pdfFieldFullPath + "L7_2", CDbl(Replace((dSales - Total6), "-", "")).ToString.Replace(",", ""))

                If IsDBNull(dr("PL_OTH_BSIN_NONSOURCE")) = False Then
                    ' HS : Set value to 0 if -ve
                    If Total < 0 Then
                        Total = 0
                    End If
                    pdfFormFields.SetField(pdfFieldFullPath + "L8", CDbl(Total + dr("PL_OTH_BSIN_NONSOURCE")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L8", "0")
                End If

                If IsDBNull(dr("PL_EXP_INTRESTRICT")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "L11", CDbl(dr("PL_EXP_INT")) + CDbl(dr("PL_EXP_INTRESTRICT")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L11", "0")
                End If
                If IsDBNull(dr("PL_LAWYER_COST")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "L12", CDbl(dr("PL_LAWYER_COST")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L12", "0")
                End If
                'If dr.IsDBNull(0) = True Then
                If IsDBNull(dr("PL_TECH_FEE")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "L12A", CDbl(dr("PL_TECH_FEE")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L12A", "0")
                End If
                'End If
                If IsDBNull(dr("PL_CONTRACT_EXP")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L13", CDbl(dr("PL_CONTRACT_EXP")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L13", "0")
                End If
                If IsDBNull(dr("PL_EXP_SALARY")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L14", CDbl(dr("PL_EXP_SALARY")).ToString.Replace(",", ""))
                    'If dr.IsDBNull(0) = True Then
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L14", "0")
                End If
                If IsDBNull(dr("PL_EMPL_STOCK")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L14A", CDbl(dr("PL_EMPL_STOCK")).ToString.Replace(",", ""))
                    'End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L14A", "0")
                End If
                If IsDBNull(dr("PL_ROYALTY")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L15", CDbl(dr("PL_ROYALTY")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L15", "0")
                End If
                If IsDBNull(dr("PL_EXP_RENT")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L16", CDbl(dr("PL_EXP_RENT")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L16", "0")
                End If
                If IsDBNull(dr("PL_EXP_MAINTENANCE")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L17", CDbl(dr("PL_EXP_MAINTENANCE")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L17", "0")
                End If
                If IsDBNull(dr("PL_RND")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L18", CDbl(dr("PL_RND")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L18", "0")
                End If
                If IsDBNull(dr("PL_KEY")) = False Then
                    strPLKey = dr("PL_KEY")
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "L1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L3", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L4", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L4A", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L5", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L6", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L7_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L8", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L11", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L12", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L12A", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L13", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L14", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L14A", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L15", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L16", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L17", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L18", "0")
            End If

            'L9  - L10 From Profit And Loss + Exempt Dividend
            'weihong in cSQL add PL_OTH_BSIN_UNREALGT, PL_NONTAX_IN_REALG, PL_NONTAX_IN_UNREALG
            cSQL = "SELECT PROFIT_LOSS_ACCOUNT.PL_OTH_IN, PROFIT_LOSS_ACCOUNT.PL_NONTAX_IN, PL_OTH_BSIN_UNREALGT, PL_NONTAX_IN_REALG, PL_NONTAX_IN_UNREALG, PL_OTH_BSIN_REALGT" _
                    & " FROM [PROFIT_LOSS_ACCOUNT]" _
                    & " WHERE PROFIT_LOSS_ACCOUNT.PL_REF_NO ='" _
                    & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' AND PROFIT_LOSS_ACCOUNT.PL_YA ='" _
                    & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                'weihong
                UntungAsing = CDbl(dr("PL_OTH_BSIN_UNREALGT")) + CDbl(dr("PL_NONTAX_IN_REALG")) + CDbl(dr("PL_NONTAX_IN_UNREALG")) + CDbl(dr("PL_OTH_BSIN_REALGT"))  'weihong
                pdfFormFields.SetField(pdfFieldFullPath + "L8_1", UntungAsing)
                'endweihong
                cSQL = "SELECT SUM(ED_AMOUNT) FROM EXEMPT_DIVIDEND WHERE ED_KEY = " & strPLKey
                dr2 = DataHandler.GetDataReader(cSQL, Conn)

                If dr2.Read() Then
                    If IsDBNull(dr2(0)) = False Then
                        If IsDBNull(dr("PL_OTH_IN")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "L9", CDbl(dr("PL_OTH_IN")) + CDbl(dr2(0)).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "L9", "0")
                        End If
                        If IsDBNull(dr("PL_NONTAX_IN")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "L10", CDbl(dr("PL_NONTAX_IN")) - CDbl(dr2(0)).ToString.Replace(",", ""))
                            'weihong
                            'If (CDbl(dr("PL_NONTAX_IN")) - CDbl(dr2(0)).ToString.Replace(",", "") - UntungAsing + CDbl(dr("PL_OTH_BSIN_REALGT"))) > 0 Then
                            '    pdfFormFields.SetField(pdfFieldFullPath + "L10", CDbl(dr("PL_NONTAX_IN")) - CDbl(dr2(0)).ToString.Replace(",", "") - UntungAsing + CDbl(dr("PL_OTH_BSIN_REALGT")))
                            'Else
                            '    pdfFormFields.SetField(pdfFieldFullPath + "L10", "0")
                            'End If
                            'endweihong
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "L10", "0")
                        End If
                    Else
                        If IsDBNull(dr("PL_OTH_IN")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "L9", CDbl(dr("PL_OTH_IN")).ToString.Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "L9", "0")
                        End If
                        If IsDBNull(dr("PL_NONTAX_IN")) = False Then
                            pdfFormFields.SetField(pdfFieldFullPath + "L10", CDbl(dr("PL_NONTAX_IN")).ToString.Replace(",", ""))
                            'weihong
                            'If (CDbl(dr("PL_NONTAX_IN")) - UntungAsing + CDbl(dr("PL_OTH_BSIN_REALGT"))) > 0 Then
                            '    pdfFormFields.SetField(pdfFieldFullPath + "L10", CDbl(dr("PL_NONTAX_IN")) - UntungAsing + CDbl(dr("PL_OTH_BSIN_REALGT")))
                            'Else
                            '    pdfFormFields.SetField(pdfFieldFullPath + "L10", "0")
                            'End If
                            'endweihong
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "L10", "0")
                        End If
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L9", CDbl(dr("PL_OTH_IN")).ToString.Replace(",", ""))
                    pdfFormFields.SetField(pdfFieldFullPath + "L10", CDbl(dr("PL_NONTAX_IN")).ToString.Replace(",", ""))
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "L8_1", "0") 'weihong
                pdfFormFields.SetField(pdfFieldFullPath + "L9", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L10", "0")
            End If
        Catch ex As Exception
            MsgBox("Some important data is not fill in page 7!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page8()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page10[0]."
        Dim pdfFieldFullPath2 As String = pdfFieldPath + "Page11[0]."
        Dim Total As Double
        Dim RugiAsing As Double = 0

        Try
            'weihong add in cSQL PL_OTHER_EXP_UNREALOSS, PL_OTHER_EXP_REALOSS
            cSQL = "SELECT PL_SALES, PL_OP_STK, PL_PURCHASES, PL_PRO_COST, PL_CLS_STK," _
                              & " PL_EXP_INT, PL_EXP_INTRESTRICT, PL_LAWYER_COST, PL_TECH_FEE, PL_CONTRACT_EXP," _
                              & " PL_EXP_SALARY, PL_EMPL_STOCK, PL_ROYALTY, PL_EXP_RENT, PL_EXP_MAINTENANCE," _
                              & " PL_OTH_BSIN_NONSOURCE, PL_OTHER_EXP_UNREALOSS, PL_OTHER_EXP_REALOSS,PL_OTHER_EXRLOSSFOREIGNT" _
                              & " FROM [PROFIT_LOSS_ACCOUNT]" _
                              & " WHERE [PL_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                              & "' And PL_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"

            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then
                Total = CDbl(dr("PL_EXP_INT")) + CDbl(dr("PL_EXP_INTRESTRICT")) + CDbl(dr("PL_LAWYER_COST")) + CDbl(dr("PL_TECH_FEE")) + CDbl(dr("PL_CONTRACT_EXP")) _
                     + CDbl(dr("PL_EXP_SALARY")) + CDbl(dr("PL_EMPL_STOCK")) + CDbl(dr("PL_ROYALTY")) + CDbl(dr("PL_EXP_RENT")) + CDbl(dr("PL_EXP_MAINTENANCE")) _
                     + CDbl(dr("PL_OTH_BSIN_NONSOURCE"))
                RugiAsing = CDbl(dr("PL_OTHER_EXP_UNREALOSS")) + CDbl(dr("PL_OTHER_EXP_REALOSS")) + CDbl(dr("PL_OTHER_EXRLOSSFOREIGNT")) 'hogie - mod on 28 July 2011 , added PL_OTHER_EXRLOSSFOREIGNT '
            Else
                Total = 0
                RugiAsing = 0
            End If

            ' L19 - L24 From Profit And Loss
            cSQL = "SELECT PL_ADVERT, PL_TRAVEL, PL_OTHER_EXP, PL_NET_PROFIT_LOSS, PL_DISALLOWED_EXP, PL_TOT_EXP" _
                    & " FROM [PROFIT_LOSS_ACCOUNT]" _
                    & " WHERE [PL_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' And PL_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                pdfFormFields.SetField(pdfFieldFullPath + "L19", CDbl(dr("PL_ADVERT")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "L20", CDbl(dr("PL_TRAVEL")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "L22_1", RugiAsing) 'weihong
                'pdfFormFields.SetField(pdfFieldFullPath + "L21", CDbl(dr("PL_OTHER_EXP")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "L21", CDbl(dr("PL_OTHER_EXP")).ToString.Replace(",", "") - RugiAsing) 'pdfFormFields.SetField(pdfFieldFullPath + "L21", CDbl(dr("PL_OTHER_EXP")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "L22", CDbl(dr("PL_TOT_EXP")).ToString.Replace(",", "")) '(Total + CDbl(dr("PL_ADVERT")) + CDbl(dr("PL_TRAVEL")) + CDbl(dr("PL_OTHER_EXP"))).ToString.Replace(",", "")) ' L11 + L12 + ... + L21
                If dr("PL_NET_PROFIT_LOSS") < 0 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "L23_1", "X")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L23_1", "")
                End If
                pdfFormFields.SetField(pdfFieldFullPath + "L23_2", (CDbl(Replace(dr("PL_NET_PROFIT_LOSS"), "-", ""))).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath + "L24", (CDbl(dr("PL_DISALLOWED_EXP"))).ToString.Replace(",", ""))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "L19", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L20", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L21", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L22_1", "0") 'weihong
                pdfFormFields.SetField(pdfFieldFullPath + "L22", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L23_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L24", "0")

            End If

            pdfFormFields.SetField(pdfFieldFullPath + "L34_1", "")
            ' L25 - L45 From Balance Sheet
            cSQL = "SELECT BS_TRANSPORT, BS_MACHINERY, BS_LAND, BS_OTH_FA, BS_CURYEARFA, BS_INVESTMENT, BS_TRADE_DEBTORS," _
                    & " BS_OTH_DEBTORS, BS_STOCK, BS_LOAN_DIRECTOR, BS_CASH, BS_OTH_CA, BS_INVESTMENT, BS_LOAN, BS_TRADE_CR," _
                    & " BS_OTHER_CR, BS_LOAN_FR_DIR, BS_OTH_LIAB, BS_LT_LIAB, BS_TOT_CA, BS_TOT_ASSETS, BS_CASH, BS_TOT_CUR_LIAB, BS_TOT_LIAB" _
                    & " FROM [BALANCE_SHEET]" _
                    & " WHERE [BS_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' And BS_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                pdfFormFields.SetField(pdfFieldFullPath + "L25", CDbl(dr("BS_TRANSPORT")))
                pdfFormFields.SetField(pdfFieldFullPath + "L26", CDbl(dr("BS_MACHINERY")))
                pdfFormFields.SetField(pdfFieldFullPath + "L27", CDbl(dr("BS_LAND")))
                pdfFormFields.SetField(pdfFieldFullPath + "L28", CDbl(dr("BS_OTH_FA")))
                pdfFormFields.SetField(pdfFieldFullPath + "L29", CDbl(dr("BS_TRANSPORT")) + CDbl(dr("BS_MACHINERY")) + CDbl(dr("BS_LAND")) + CDbl(dr("BS_OTH_FA")))  'L25 + L26 + L27 + L28
                pdfFormFields.SetField(pdfFieldFullPath + "L29A", CDbl(dr("BS_CURYEARFA")))
                pdfFormFields.SetField(pdfFieldFullPath + "L30", CDbl(dr("BS_INVESTMENT")))

                pdfFormFields.SetField(pdfFieldFullPath + "L31", CDbl(dr("BS_TRADE_DEBTORS")))
                pdfFormFields.SetField(pdfFieldFullPath + "L32", CDbl(dr("BS_OTH_DEBTORS")))
                pdfFormFields.SetField(pdfFieldFullPath + "L32A", CDbl(dr("BS_STOCK")))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L33", CDbl(dr("BS_LOAN_DIRECTOR")))

                If CDbl(dr("BS_CASH")) < 0 Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L34_1", "X")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "L34_1", "")
                End If
                pdfFormFields.SetField(pdfFieldFullPath2 + "L34_2", CDbl(Replace(dr("BS_CASH"), "-", "")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L35", CDbl(Replace(dr("BS_OTH_CA"), "-", "")).ToString.Replace(",", ""))

                Total = CDbl(dr("BS_TRADE_DEBTORS")) + CDbl(dr("BS_OTH_DEBTORS")) + CDbl(dr("BS_STOCK")) + CDbl(dr("BS_LOAN_DIRECTOR")) _
                        + CDbl(dr("BS_OTH_CA")) + CDbl(dr("BS_INVESTMENT"))

                pdfFormFields.SetField(pdfFieldFullPath2 + "L36", CDbl(dr("BS_TOT_CA")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L37", CDbl(dr("BS_TOT_ASSETS")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L38", CDbl(dr("BS_LOAN")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L39", CDbl(dr("BS_TRADE_CR")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L40", CDbl(dr("BS_OTHER_CR")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L41", CDbl(dr("BS_LOAN_FR_DIR")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L42", CDbl(dr("BS_OTH_LIAB")).ToString.Replace(",", ""))

                Total = CDbl(dr("BS_TRANSPORT")) + CDbl(dr("BS_TRADE_CR")) + CDbl(dr("BS_LOAN")) + CDbl(dr("BS_LOAN_DIRECTOR")) _
                            + CDbl(dr("BS_OTH_CA")) + CDbl(dr("BS_INVESTMENT"))

                'pdfFormFields.SetField(pdfFieldFullPath + "L43", Total.ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L43", CDbl(dr("BS_TOT_CUR_LIAB")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L44", CDbl(dr("BS_LT_LIAB")).ToString.Replace(",", ""))
                pdfFormFields.SetField(pdfFieldFullPath2 + "L45", CDbl(dr("BS_TOT_LIAB")).ToString.Replace(",", ""))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "L25", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L26", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L27", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L28", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L29", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L29A", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L30", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L31", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L32", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L32A", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L33", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L34_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L35", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L36", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L37", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L38", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L39", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L40", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L41", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L42", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L43", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L44", "0")
                pdfFormFields.SetField(pdfFieldFullPath2 + "L45", "0")
            End If

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 8!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page9()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page11[0]."
        Dim Total As Double

        Try
            pdfFormFields.SetField(pdfFieldFullPath + "L47_1", "")
            pdfFormFields.SetField(pdfFieldFullPath + "L49_1", "")
            pdfFormFields.SetField(pdfFieldFullPath + "L50_1", "")
            ' L46 - L50
            cSQL = "SELECT BS_CAPITAL, BS_PNL_APPR_ACC, BS_PNL_APPR_ACC, BS_RESERVE_ACC, BS_TOT_EQUITY, BS_TOT_LIAB_EQ" _
                   & " FROM [BALANCE_SHEET]" _
                   & " WHERE [BS_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                   & "' And BS_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                If IsDBNull(dr("BS_PNL_APPR_ACC")) = False Then
                    Total = CDbl(dr("BS_PNL_APPR_ACC")) + CDbl(dr("BS_PNL_APPR_ACC")) + CDbl(dr("BS_RESERVE_ACC"))
                Else
                    Total = 0
                End If

                If IsDBNull(dr("BS_CAPITAL")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "L46", CDbl(dr("BS_CAPITAL")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L46", "0")
                End If

                If IsDBNull(dr("BS_PNL_APPR_ACC")) = False Then
                    If CDbl(dr("BS_PNL_APPR_ACC")) < 0 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "L47_1", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "L47_1", "")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L47_1", "")
                End If

                If IsDBNull(dr("BS_PNL_APPR_ACC")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "L47_2", Replace(CDbl(dr("BS_PNL_APPR_ACC")), "-", "").ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L47_2", "0")
                End If

                If IsDBNull(dr("BS_RESERVE_ACC")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "L48", Replace(CDbl(dr("BS_RESERVE_ACC")), "-", "").ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L48", "0")
                End If

                If IsDBNull(dr("BS_TOT_EQUITY")) = False Then
                    If CDbl(dr("BS_TOT_EQUITY")) < 0 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "L49_1", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "L49_1", "")
                    End If
                    pdfFormFields.SetField(pdfFieldFullPath + "L49_2", CDbl(Replace(dr("BS_TOT_EQUITY"), "-", "")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L49_2", "0")
                End If

                If CDbl(dr("BS_TOT_LIAB_EQ")) < 0 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "L50_1", "X")
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "L50_1", "")
                End If

                ' remark
                pdfFormFields.SetField(pdfFieldFullPath + "L50_2", CDbl(Replace(dr("BS_TOT_LIAB_EQ"), "-", "").ToString.Replace(",", "")))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "L46", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L47_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L48", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L49_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "L50_2", "0")
            End If

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 9!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page10()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page12[0]."

        Try
            ' ========= Part M ============
            cSQL = "SELECT WT_107A_GROSS, WT_107A_TAX, WT_109_GROSS, WT_109_TAX, WT_109A_GROSS, WT_109A_TAX," _
                   & " WT_109B_GROSS, WT_109B_TAX, WT_109E_GROSS, WT_109E_TAX, WT_109F_GROSS, WT_109F_TAX" _
                   & " FROM [WITHHOLD_TAX]" _
                   & " WHERE [WT_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                   & "' And WT_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then
                If IsDBNull(dr("WT_107A_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M1_1", FormatNumber((CDbl(dr("WT_107A_GROSS"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M1_1", "0")
                End If
                If IsDBNull(dr("WT_107A_TAX")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M1_2", FormatNumber((CDbl(dr("WT_107A_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M1_2", "0")
                End If
                If IsDBNull(dr("WT_107A_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M1_3", FormatNumber((CDbl(dr("WT_107A_GROSS"))) - (CDbl(dr("WT_107A_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M1_3", "0")
                End If
                If IsDBNull(dr("WT_109_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M2_1", FormatNumber((CDbl(dr("WT_109_GROSS"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M2_1", "0")
                End If
                If IsDBNull(dr("WT_109_TAX")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M2_2", FormatNumber((CDbl(dr("WT_109_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M2_2", "0")
                End If
                If IsDBNull(dr("WT_109_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M2_3", FormatNumber((CDbl(dr("WT_109_GROSS")) - CDbl(dr("WT_109_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M2_3", "0")
                End If
                If IsDBNull(dr("WT_109A_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M3_1", FormatNumber((CDbl(dr("WT_109A_GROSS"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M3_1", "0")
                End If
                If IsDBNull(dr("WT_109A_TAX")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M3_2", FormatNumber((CDbl(dr("WT_109A_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M3_2", "0")
                End If
                If IsDBNull(dr("WT_109A_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M3_3", FormatNumber((CDbl(dr("WT_109A_GROSS")) - CDbl(dr("WT_109A_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M3_3", "0")
                End If
                If IsDBNull(dr("WT_109B_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M4_1", FormatNumber((CDbl(dr("WT_109B_GROSS"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M4_1", "0")
                End If
                If IsDBNull(dr("WT_109B_TAX")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M4_2", FormatNumber((CDbl(dr("WT_109B_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M4_2", "0")
                End If
                If IsDBNull(dr("WT_109B_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M4_3", FormatNumber((CDbl(dr("WT_109B_GROSS")) - CDbl(dr("WT_109B_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M4_3", "0")
                End If
                If IsDBNull(dr("WT_109E_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M5_1", FormatNumber((CDbl(dr("WT_109E_GROSS"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M5_1", "0")
                End If
                If IsDBNull(dr("WT_109E_TAX")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M5_2", FormatNumber((CDbl(dr("WT_109E_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M5_2", "0")
                End If
                If IsDBNull(dr("WT_109E_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M5_3", FormatNumber((CDbl(dr("WT_109E_GROSS")) - CDbl(dr("WT_109E_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M5_3", "0")
                End If
                ' HS :  New Field
                If IsDBNull(dr("WT_109F_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M6_1", FormatNumber((CDbl(dr("WT_109F_GROSS"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M6_1", "0")
                End If
                If IsDBNull(dr("WT_109F_TAX")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M6_2", FormatNumber((CDbl(dr("WT_109F_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M6_2", "0")
                End If
                If IsDBNull(dr("WT_109F_GROSS")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "M6_3", FormatNumber((CDbl(dr("WT_109F_GROSS")) - CDbl(dr("WT_109F_TAX"))), 0))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "M6_3", "0")
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "M1_1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M1_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M1_3", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M2_1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M2_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M2_3", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M3_1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M3_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M3_3", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M4_1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M4_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M4_3", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M5_1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M5_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M5_3", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M6_1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M6_2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "M6_3", "0")
            End If

            ' ========= Part N ============
            cSQL = "SELECT RC_1, RC_2, RC_3, RC_4, RC_5, RC_6," _
                    & " RC_7, RC_8, RC_9, RC_10, RC_11, RC_12" _
                    & " FROM [RELATED_COMPANY]" _
                    & " WHERE [RC_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' And RC_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then
                If IsDBNull(dr("RC_1")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N1", CDbl(dr("RC_1")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N1", "0")
                End If
                If IsDBNull(dr("RC_2")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N2", CDbl(dr("RC_2")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N2", "0")
                End If
                If IsDBNull(dr("RC_3")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N3", CDbl(dr("RC_3")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N3", "0")
                End If
                If IsDBNull(dr("RC_4")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N4", CDbl(dr("RC_4")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N4", "0")
                End If
                If IsDBNull(dr("RC_5")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N5", CDbl(dr("RC_5")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N5", "0")
                End If
                If IsDBNull(dr("RC_6")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N6", CDbl(dr("RC_6")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N6", "0")
                End If
                If IsDBNull(dr("RC_7")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N7", CDbl(dr("RC_7")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N7", "0")
                End If
                If IsDBNull(dr("RC_8")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N8", CDbl(dr("RC_8")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N8", "0")
                End If
                If IsDBNull(dr("RC_9")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N9", CDbl(dr("RC_9")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N9", "0")
                End If
                If IsDBNull(dr("RC_10")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N10", CDbl(dr("RC_10")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N10", "0")
                End If
                If IsDBNull(dr("RC_11")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N11", CDbl(dr("RC_11")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N11", "0")
                End If
                If IsDBNull(dr("RC_12")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "N12", CDbl(dr("RC_12")).ToString.Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "N12", "0")
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "N1", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N2", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N3", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N4", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N5", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N6", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N7", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N8", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N9", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N10", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N11", "0")
                pdfFormFields.SetField(pdfFieldFullPath + "N12", "0")
            End If

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 10!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page11()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page13[0]."
        Dim pdfFieldFullPath2 As String = pdfFieldPath + "Page14[0]."
        Dim strCoStatus As Array
        Dim strAddress As String = ""

        pdfFormFields.SetField(pdfFieldFullPath + "P1_1", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_2", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_3", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_4", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_5", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_6", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_7", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_8", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_9", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_10", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_11", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_12", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_13", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_14", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_15", "")
        pdfFormFields.SetField(pdfFieldFullPath + "P1_16", "")

        Try
            ' =========== Part P1 =============
            cSQL = "SELECT TP_CO_STATUS, TP_REG_ADD_LINE1, TP_REG_ADD_LINE2, TP_REG_ADD_LINE3, TP_REG_POSTCODE," _
        & " TP_REG_CITY, TP_REG_STATE, TP_TEL_NO, TP_CURR_ADD_LINE1, TP_CURR_ADD_LINE2, TP_CURR_ADD_LINE3," _
        & " TP_CURR_POSTCODE, TP_CURR_CITY, TP_CURR_STATE, TP_COM_ADD_LINE1, TP_COM_ADD_LINE2, TP_COM_ADD_LINE3, TP_COM_POSTCODE," _
        & " TP_COM_CITY, TP_COM_STATE, TP_ALT_ADD_LINE1, TP_ALT_ADD_LINE2, TP_ALT_ADD_LINE3," _
        & " TP_ALT_POSTCODE, TP_ALT_CITY, TP_ALT_STATE, TP_BANK, TP_BANK_ACC, TP_RECORD_KEPT, TP_BLOG" _
        & " FROM [TAXP_PROFILE]" _
        & " WHERE [TP_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then

                If Not String.IsNullOrEmpty(dr("TP_CO_STATUS")) And Not IsDBNull(dr("TP_CO_STATUS")) Then

                    strCoStatus = Split(dr("TP_CO_STATUS"), ",")

                    For intO1 As Integer = 0 To UBound(strCoStatus)
                        Select Case strCoStatus(intO1)
                            Case 0

                            Case 1
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_2", "X")
                            Case 2
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_3", "X")
                            Case 3
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_11", "X")
                            Case 4
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_4", "X")
                            Case 5
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_6", "X")
                            Case 6
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_9", "X")
                            Case 7
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_7", "X")
                            Case 8
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_13", "X")
                            Case 9
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_8", "X")
                            Case 10
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_10", "X")
                            Case 11
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_16", "X")
                            Case 12
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_12", "X")
                            Case 13
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_14", "X")
                            Case 14
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_15", "X")
                            Case 15
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_5", "X")
                            Case 16
                                pdfFormFields.SetField(pdfFieldFullPath + "P1_1", "X")
                        End Select
                    Next
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_2", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_3", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_11", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_4", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_9", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_7", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_13", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_8", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_10", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_16", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_12", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_14", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_15", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_5", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_1", " ")
                    pdfFormFields.SetField(pdfFieldFullPath + "P1_6", " ")
                End If
            End If

            ' =========== Part P2 =============
            ' P2 
            Dim strReg_ADD(2) As String
            If (IsDBNull(dr("TP_REG_ADD_LINE1")) = False) Or (IsDBNull(dr("TP_REG_ADD_LINE2")) = False) Or (IsDBNull(dr("TP_REG_ADD_LINE3")) = False) Then
                'strReg_ADD = (dr("TP_REG_ADD_LINE1")).ToString & " " & (dr("TP_REG_ADD_LINE2")).ToString & " " & (dr("TP_REG_ADD_LINE3")).ToString
                '=== NGKL 2010.2 ==='
                strReg_ADD(0) = dr("TP_REG_ADD_LINE1").ToString
                strReg_ADD(1) = dr("TP_REG_ADD_LINE2").ToString
                strReg_ADD(2) = dr("TP_REG_ADD_LINE3").ToString
                strReg_ADD = TextAddressSpliter(strReg_ADD, 24)
                If Not (strReg_ADD(0) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_1", strReg_ADD(0).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_1", "")
                End If
                If Not (strReg_ADD(1) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_2", strReg_ADD(1).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_2", "")
                End If
                If Not (strReg_ADD(2) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_3", strReg_ADD(2).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_3", "")
                End If
                ' === NGKL 2010.2 END ==='
                'If Len(strReg_ADD) < 24 Then
                '    pdfFormFields.SetField(pdfFieldFullPath + "P2_1", strReg_ADD.ToString.ToUpper)
                'Else
                '    CutLine(strReg_ADD.ToString.ToUpper(), 24)
                '    pdfFormFields.SetField(pdfFieldFullPath + "P2_1", strCropped)
                '    If Len(strReg_ADD) > 24 And (Len(strReg_ADD) <= 48) Then
                '        pdfFormFields.SetField(pdfFieldFullPath + "P2_2", strRemainder)
                '        pdfFormFields.SetField(pdfFieldFullPath + "P2_3", space(24))
                '    Else
                '        CutLine(strRemainder, 24)
                '        pdfFormFields.SetField(pdfFieldFullPath + "P2_2", strCropped)
                '        If Len(strRemainder) > 24 Then
                '            pdfFormFields.SetField(pdfFieldFullPath + "P2_3", strRemainder.Substring(0, 24))
                '        Else
                '            pdfFormFields.SetField(pdfFieldFullPath + "P2_3", strRemainder.ToString())
                '        End If
                '    End If
                'End If

                If ((IsDBNull(dr("TP_REG_POSTCODE"))) = True) Or dr("TP_REG_POSTCODE") = "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_4", space(5))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_4", dr("TP_REG_POSTCODE").ToString.ToUpper())
                End If

                If ((IsDBNull(dr("TP_REG_CITY"))) = True) Or dr("TP_REG_CITY") = "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_5", space(16))
                Else
                    If Len(dr("TP_REG_CITY")) > 16 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "P2_5", dr("TP_REG_CITY").ToString.ToUpper().Substring(0, 16))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "P2_5", dr("TP_REG_CITY").ToString.ToUpper())
                    End If
                End If

                If ((IsDBNull(dr("TP_REG_STATE"))) = True) Or dr("TP_REG_STATE") = "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_6", space(24))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P2_6", dr("TP_REG_STATE").ToString.ToUpper())
                End If
            End If

            ' P3
            If (IsDBNull(dr("TP_TEL_NO"))) = False Then
                If (InStr(dr("TP_TEL_NO"), "-") = 3) = True Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P3", " " + Mid(dr("TP_TEL_NO"), 1, 12).Replace("-", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P3", Mid(dr("TP_TEL_NO"), 1, 12).Replace("-", ""))
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "P3", space(12))
            End If

            ' P4
            Dim strCurr_ADD(2) As String
            If (IsDBNull(dr("TP_CURR_ADD_LINE1"))) = False Or (IsDBNull(dr("TP_CURR_ADD_LINE2"))) = False Or (IsDBNull(dr("TP_CURR_ADD_LINE3"))) = False Then
                'strCurr_ADD = (dr("TP_CURR_ADD_LINE1")).ToString & " " & (dr("TP_CURR_ADD_LINE2")).ToString & " " & (dr("TP_CURR_ADD_LINE3")).ToString
                'If Len(strCurr_ADD) < 24 Then
                '    pdfFormFields.SetField(pdfFieldFullPath + "P4_1", strCurr_ADD.ToString.ToUpper)
                'Else
                '    CutLine(strCurr_ADD.ToString.ToUpper(), 24)
                '    pdfFormFields.SetField(pdfFieldFullPath + "P4_1", strCropped)
                '    If Len(strCurr_ADD) > 24 And (Len(strCurr_ADD) <= 48) Then
                '        pdfFormFields.SetField(pdfFieldFullPath + "P4_2", strRemainder)
                '        pdfFormFields.SetField(pdfFieldFullPath + "P4_3", space(24))
                '    Else
                '        CutLine(strRemainder, 24)
                '        pdfFormFields.SetField(pdfFieldFullPath + "P4_2", strCropped)
                '        If Len(strRemainder) > 24 Then
                '            pdfFormFields.SetField(pdfFieldFullPath + "P4_3", strRemainder.Substring(0, 24))
                '        Else
                '            pdfFormFields.SetField(pdfFieldFullPath + "P4_3", strRemainder.ToString())
                '        End If
                '    End If
                'End If
                ' === NGKL 2010.2 ==='
                strCurr_ADD(0) = dr("TP_CURR_ADD_LINE1").ToString
                strCurr_ADD(1) = dr("TP_CURR_ADD_LINE2").ToString
                strCurr_ADD(2) = dr("TP_CURR_ADD_LINE3").ToString
                strCurr_ADD = TextAddressSpliter(strCurr_ADD, 24)
                If Not (strCurr_ADD(0) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_1", strCurr_ADD(0).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_1", "")
                End If
                If Not (strCurr_ADD(1) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_2", strCurr_ADD(1).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_2", "")
                End If
                If Not (strCurr_ADD(2) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_3", strCurr_ADD(2).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_3", "")
                End If
                ' === NGKL 2010.2 END ==='

                If IsDBNull(dr("TP_CURR_POSTCODE")) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_4", space(5))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_4", dr("TP_CURR_POSTCODE"))
                End If
                If IsDBNull(dr("TP_CURR_CITY")) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_5", space(16))
                Else
                    If Len(dr("TP_CURR_CITY")) > 16 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "P4_5", dr("TP_CURR_CITY").ToString.ToUpper().Substring(0, 16))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "P4_5", dr("TP_CURR_CITY").ToString.ToUpper())
                    End If
                End If
                If IsDBNull(dr("TP_CURR_STATE")) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_6", space(24))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P4_6", dr("TP_CURR_STATE").ToString.ToUpper())
                End If
            End If

            ' P5
            Dim strCom_ADD(2) As String
            If (IsDBNull(dr("TP_COM_ADD_LINE1"))) = False Or (IsDBNull(dr("TP_COM_ADD_LINE2"))) Or (IsDBNull(dr("TP_COM_ADD_LINE3"))) Then
                'strCom_ADD = (dr("TP_COM_ADD_LINE1")).ToString & " " & (dr("TP_COM_ADD_LINE2")).ToString & " " & (dr("TP_COM_ADD_LINE3")).ToString
                'If Len(strCom_ADD) < 24 Then
                '    pdfFormFields.SetField(pdfFieldFullPath + "P5_1", strCom_ADD.ToString.ToUpper)
                'Else
                '    Call CutLine(strCom_ADD.ToString.ToUpper(), 24)
                '    pdfFormFields.SetField(pdfFieldFullPath + "P5_1", strCropped)
                '    If Len(strCom_ADD) > 24 And (Len(strCom_ADD) <= 48) Then
                '        pdfFormFields.SetField(pdfFieldFullPath + "P5_2", strRemainder)
                '        pdfFormFields.SetField(pdfFieldFullPath + "P5_3", space(24))
                '    Else
                '        Call CutLine(strRemainder, 24)
                '        pdfFormFields.SetField(pdfFieldFullPath + "P5_2", strCropped)
                '        If Len(strRemainder) > 24 Then
                '            pdfFormFields.SetField(pdfFieldFullPath + "P5_3", strRemainder.Substring(0, 24))
                '        Else
                '            pdfFormFields.SetField(pdfFieldFullPath + "P5_3", strRemainder.ToString())
                '        End If
                '    End If
                'End If
                ' === NGKL 2010.2 ==='
                strCom_ADD(0) = dr("TP_COM_ADD_LINE1").ToString
                strCom_ADD(1) = dr("TP_COM_ADD_LINE2").ToString
                strCom_ADD(2) = dr("TP_COM_ADD_LINE3").ToString
                strCom_ADD = TextAddressSpliter(strCom_ADD, 24)
                If Not (strCom_ADD(0) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_1", strCom_ADD(0).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_1", "")
                End If
                If Not (strCom_ADD(1) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_2", strCom_ADD(1).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_2", "")
                End If
                If Not (strCom_ADD(2) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_3", strCom_ADD(2).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_3", "")
                End If
                ' === NGKL 2010.2 END ==='

                If IsDBNull(dr("TP_COM_POSTCODE")) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_4", space(5))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_4", dr("TP_COM_POSTCODE"))
                End If
                If IsDBNull(dr("TP_COM_CITY")) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_5", space(16))
                Else
                    If Len(dr("TP_COM_CITY")) > 16 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "P5_5", dr("TP_COM_CITY").ToString.ToUpper().Substring(0, 16))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "P5_5", dr("TP_COM_CITY").ToString.ToUpper())
                    End If
                End If
                If IsDBNull(dr("TP_COM_STATE")) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_6", space(24))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P5_6", dr("TP_COM_STATE").ToString.ToUpper())
                End If
            End If

            'P6
            If IsDBNull(dr("TP_BLOG")) Then
                pdfFormFields.SetField(pdfFieldFullPath + "P6", space(24))
            Else
                If Len(dr("TP_BLOG")) > 24 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P6", dr("TP_BLOG").ToString.ToUpper().Substring(0, 24))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P6", dr("TP_BLOG").ToString.ToUpper())
                End If
            End If
            ' P7
            If IsDBNull(dr("TP_BANK")) Then
                pdfFormFields.SetField(pdfFieldFullPath + "P7", space(24))
            Else
                If Len(dr("TP_BANK")) > 24 Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P7", dr("TP_BANK").ToString.ToUpper().Substring(0, 24))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P7", dr("TP_BANK").ToString.ToUpper())
                End If
            End If
            ' P8
            If IsDBNull(dr("TP_BANK_ACC")) Then
                pdfFormFields.SetField(pdfFieldFullPath + "P8", space(24))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "P8", dr("TP_BANK_ACC"))
            End If

            ' P9
            pdfFormFields.SetField(pdfFieldFullPath + "P9_1", "")
            pdfFormFields.SetField(pdfFieldFullPath + "P9_2", "")
            pdfFormFields.SetField(pdfFieldFullPath + "P9_3", "")
            If IsDBNull(dr("TP_RECORD_KEPT")) = False Then
                If dr("TP_RECORD_KEPT") = "3" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P9_1", "X")
                    pdfFormFields.SetField(pdfFieldFullPath + "P9_2", "")
                    pdfFormFields.SetField(pdfFieldFullPath + "P9_3", "")
                ElseIf dr("TP_RECORD_KEPT") = "1" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P9_1", "")
                    pdfFormFields.SetField(pdfFieldFullPath + "P9_2", "X")
                    pdfFormFields.SetField(pdfFieldFullPath + "P9_3", "")
                ElseIf dr("TP_RECORD_KEPT") = "2" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P9_1", "")
                    pdfFormFields.SetField(pdfFieldFullPath + "P9_2", "")
                    pdfFormFields.SetField(pdfFieldFullPath + "P9_3", "X")
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "P9_1", "")
                pdfFormFields.SetField(pdfFieldFullPath + "P9_2", "")
                pdfFormFields.SetField(pdfFieldFullPath + "P9_3", "")
            End If

            ' P10
            Dim strAlt_ADD(2) As String
            If IsDBNull(dr("TP_ALT_ADD_LINE1")) = False Or IsDBNull(dr("TP_ALT_ADD_LINE2")) = False Or IsDBNull(dr("TP_ALT_ADD_LINE3")) = False Then
                'strAlt_ADD = (dr("TP_ALT_ADD_LINE1")).ToString & " " & (dr("TP_ALT_ADD_LINE2")).ToString & " " & (dr("TP_ALT_ADD_LINE3")).ToString
                'If Len(strAlt_ADD) < 24 Then
                '    pdfFormFields.SetField(pdfFieldFullPath + "P10_1", strAlt_ADD.ToString.ToUpper)
                'Else
                '    CutLine(strAlt_ADD.ToString.ToUpper(), 24)
                '    pdfFormFields.SetField(pdfFieldFullPath + "P10_1", strCropped)
                '    If Len(strAlt_ADD) > 24 And (Len(strAlt_ADD) <= 48) Then
                '        pdfFormFields.SetField(pdfFieldFullPath + "P10_2", strRemainder)
                '        pdfFormFields.SetField(pdfFieldFullPath + "P10_3", space(24))
                '    Else
                '        CutLine(strRemainder, 24)
                '        pdfFormFields.SetField(pdfFieldFullPath + "P10_2", strCropped)
                '        If Len(strRemainder) > 24 Then
                '            pdfFormFields.SetField(pdfFieldFullPath + "P10_3", strRemainder.ToString.Substring(0, 24))
                '        Else
                '            pdfFormFields.SetField(pdfFieldFullPath + "P10_3", strRemainder.ToString())
                '        End If
                '    End If
                'End If
                ' === NGKL 2010.2 ==='
                strAlt_ADD(0) = dr("TP_ALT_ADD_LINE1").ToString
                strAlt_ADD(1) = dr("TP_ALT_ADD_LINE2").ToString
                strAlt_ADD(2) = dr("TP_ALT_ADD_LINE3").ToString
                strAlt_ADD = TextAddressSpliter(strAlt_ADD, 24)
                If Not (strAlt_ADD(0) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P10_1", strAlt_ADD(0).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P10_1", "")
                End If
                If Not (strAlt_ADD(1) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P10_2", strAlt_ADD(1).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P10_2", "")
                End If
                If Not (strAlt_ADD(2) Is Nothing) Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P10_3", strAlt_ADD(2).ToString.ToUpper)
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P10_3", "")
                End If
                ' === NGKL 2010.2 END ==='
            End If

            If IsDBNull(dr("TP_ALT_POSTCODE")) Or dr("TP_ALT_POSTCODE") = "" Then
                pdfFormFields.SetField(pdfFieldFullPath2 + "P10_4", space(5))
            Else
                pdfFormFields.SetField(pdfFieldFullPath2 + "P10_4", dr("TP_ALT_POSTCODE"))
            End If
            If IsDBNull(dr("TP_ALT_CITY")) Or dr("TP_ALT_CITY") = "" Then
                pdfFormFields.SetField(pdfFieldFullPath2 + "P10_5", space(16))
            Else
                If Len(dr("TP_ALT_CITY")) > 16 Then
                    pdfFormFields.SetField(pdfFieldFullPath2 + "P10_5", dr("TP_ALT_CITY").ToString.ToUpper().Substring(0, 16))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath2 + "P10_5", dr("TP_ALT_CITY").ToString.ToUpper())
                End If
            End If
            If IsDBNull(dr("TP_ALT_STATE")) Or dr("TP_ALT_STATE") = "" Then
                pdfFormFields.SetField(pdfFieldFullPath2 + "P10_6", space(24))
            Else
                pdfFormFields.SetField(pdfFieldFullPath2 + "P10_6", dr("TP_ALT_STATE").ToString.ToUpper())
            End If

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 11!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page12()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page14[0]."
        Dim strAddress As String = ""
        Dim strIndex As String = ""

        Try

            ' P10 - P16
            cSQL = "SELECT DIR_NAME, DIR_IC, DIR_TEL_NO, DIR_REFTYPE, DIR_REFNUM, DIR_REFNUM2, DIR_REFNUM3," _
            & " DIR_EQUITY, DIR_SALARY, DIR_ALLOW" _
            & " FROM [DIRECTORS_PROFILE]" _
            & " WHERE [DIR_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
            & "' And DIR_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
            & " ORDER BY [DIR_ORDER]"
            dr = DataHandler.GetDataReader(cSQL, Conn)

            Do While dr.Read()
                strIndex = strIndex + "I"
                If IsDBNull(dr("DIR_NAME")) = False And (dr("DIR_NAME")) <> "" Then
                    If Len(dr("DIR_NAME")) > 28 Then
                        CutLine(dr("DIR_NAME").ToString.ToUpper(), 28)
                        pdfFormFields.SetField(pdfFieldFullPath + "P11" + strIndex + "_1", strCropped)
                        pdfFormFields.SetField(pdfFieldFullPath + "P11" + strIndex + "_2", Mid(LTrim(strRemainder), 1, 28))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "P11" + strIndex + "_1", dr("DIR_NAME").ToString.ToUpper())
                        pdfFormFields.SetField(pdfFieldFullPath + "P11" + strIndex + "_2", space(28))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P11" + strIndex + "_1", space(28))
                    pdfFormFields.SetField(pdfFieldFullPath + "P11" + strIndex + "_2", space(28))
                End If

                If IsDBNull(dr("DIR_IC")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P12" + strIndex, Mid(Replace(dr("DIR_IC"), "-", ""), 1, 12))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P12" + strIndex, space(12))
                End If

                If IsDBNull(dr("DIR_TEL_NO")) = False And dr("DIR_TEL_NO") <> "" Then
                    If InStr(dr("DIR_TEL_NO"), "-") = 3 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "P13" + strIndex, " " + Mid(dr("DIR_TEL_NO"), 1, 12).Replace("-", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "P13" + strIndex, Mid(dr("DIR_TEL_NO"), 1, 12).Replace("-", ""))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P13" + strIndex, space(12))
                End If

                If IsDBNull(dr("DIR_REFTYPE")) = False And dr("DIR_REFTYPE") <> "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P14" + strIndex + "_1", dr("DIR_REFTYPE"))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P14" + strIndex + "_1", "--")
                End If

                If (IsDBNull(dr("DIR_REFNUM")) = False Or IsDBNull(dr("DIR_REFNUM2")) = False Or IsDBNull(dr("DIR_REFNUM3")) = False) And _
                    (dr("DIR_REFNUM") + dr("DIR_REFNUM2") + dr("DIR_REFNUM3")) <> "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P14" + strIndex + "_2", dr("DIR_REFNUM") + dr("DIR_REFNUM2") + dr("DIR_REFNUM3"))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P14" + strIndex + "_2", space(11))
                End If

                If IsDBNull(dr("DIR_EQUITY")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P15" + strIndex, Replace(FormatNumber((dr("DIR_EQUITY")), 2).ToString, ".", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P15" + strIndex, "000")
                End If

                If IsDBNull(dr("DIR_SALARY")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P16" + strIndex, dr("DIR_SALARY").Replace(",", ""))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P16" + strIndex, "0")
                End If

                If IsDBNull(dr("DIR_ALLOW")) = False Then
                    pdfFormFields.SetField(pdfFieldFullPath + "P17" + strIndex, dr("DIR_ALLOW").Replace(",", ""))
                    If strIndex.Length >= 3 Then
                        Exit Sub
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "P17" + strIndex, "0")
                    If strIndex.Length >= 3 Then
                        Exit Sub
                    End If
                End If
            Loop

            Dim intNoLoop As Integer = 0

            If strIndex <> "" Then
                intNoLoop = strIndex.Length()
            Else
                intNoLoop = 0
            End If
            While intNoLoop < 3
                strIndex = strIndex + "I"
                pdfFormFields.SetField(pdfFieldFullPath + "P11" + strIndex + "_1", space(28))
                pdfFormFields.SetField(pdfFieldFullPath + "P11" + strIndex + "_2", space(28))
                pdfFormFields.SetField(pdfFieldFullPath + "P12" + strIndex, space(12))
                pdfFormFields.SetField(pdfFieldFullPath + "P13" + strIndex, space(12))
                pdfFormFields.SetField(pdfFieldFullPath + "P14" + strIndex + "_1", "--")
                pdfFormFields.SetField(pdfFieldFullPath + "P14" + strIndex + "_2", space(11))
                pdfFormFields.SetField(pdfFieldFullPath + "P15" + strIndex, "000")
                pdfFormFields.SetField(pdfFieldFullPath + "P16" + strIndex, "0")
                pdfFormFields.SetField(pdfFieldFullPath + "P17" + strIndex, "0")
                intNoLoop = intNoLoop + 1
            End While

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 12!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page13()
        Dim intIndex As Integer = 0
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page15[0]."

        Try
            ' === Q1 - Q6 ===
            cSQL = "SELECT SHAREHOLDERS_PROFILE.SH_NAME, SHAREHOLDERS_PROFILE.SH_IC, SHAREHOLDERS_PROFILE.SH_COUNTRY, SHAREHOLDERS_PROFILE.SH_SHARE, SHAREHOLDERS_PROFILE.SH_PAR_VALUE, BALANCE_SHEET.BS_CAPITAL, SHAREHOLDERS_PROFILE.SH_CHECK, SHAREHOLDERS_PROFILE.SH_SHAREP" _
                   & " FROM [SHAREHOLDERS_PROFILE] INNER JOIN [BALANCE_SHEET] ON BALANCE_SHEET.BS_REF_NO = SHAREHOLDERS_PROFILE.SH_REF_NO And" _
                   & " SHAREHOLDERS_PROFILE.SH_YA = BALANCE_SHEET.BS_YA" _
                   & " WHERE [SH_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "' And SH_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                   & " ORDER BY SH_ORDER"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            Do While dr.Read()
                If intIndex <= 5 Then
                    intIndex = intIndex + 1

                    If IsDBNull(dr("SH_IC")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_1", Replace(dr("SH_IC"), "-", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_1", space(12))
                    End If

                    If IsDBNull(dr("SH_NAME")) = False Then
                        If Len(dr("SH_NAME")) > 28 Then
                            CutLine(dr("SH_NAME").ToString.ToUpper(), 28)
                            pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_2", strCropped)
                            pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_3", Mid(LTrim(strRemainder), 1, 28))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_2", dr("SH_NAME").ToString.ToUpper())
                            pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_3", space(28))
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_2", space(28))
                        pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_3", space(28))
                    End If

                    Dim Total2 As Double = 0.0
                    If dr("SH_SHAREP") = 0 Then 'Add
                        If IsDBNull(dr("BS_CAPITAL")) = False Then
                            If dr("BS_CAPITAL") = 0 Then
                                pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_4", "000")
                            Else
                                If IsDBNull(dr("SH_SHARE")) = False And IsDBNull(dr("SH_SHARE")) = False Then
                                    Total2 = ((dr("SH_SHARE")) * (dr("SH_PAR_VALUE")) / dr("BS_CAPITAL")) * 100
                                    pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_4", Replace(Total2.ToString("0.00"), ".", "")) 'Replace((CDbl(dr("SH_SHARE")) * CDbl(dr("SH_PAR_VALUE")) / (CDbl(dr("BS_CAPITAL")) * 100)).ToString("0.00"), ".", ""))
                                Else
                                    pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_4", "")
                                End If
                            End If
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_4", "000")
                        End If
                    Else
                        Total2 = dr("SH_SHAREP")
                        pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_4", Replace(Total2.ToString("0.00"), ".", ""))
                    End If

                    If IsDBNull(dr("SH_COUNTRY")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_5", dr("SH_COUNTRY"))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_5", "--")
                    End If
                End If
            Loop

            Do While intIndex <= 5
                intIndex = intIndex + 1
                pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_1", space(12))
                pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_2", space(28))
                pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_3", space(28))
                pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_4", "000")
                pdfFormFields.SetField(pdfFieldFullPath + "Q" + intIndex.ToString + "_5", "--")
            Loop

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 13!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page14()
        Dim intIndex As Integer = 0
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page16[0]."

        Try
            ' ==== Part R ====
            pdfFormFields.SetField(pdfFieldFullPath + "R1_1", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R1_2", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R1_3", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R1_4", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R1_5", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R2a_1", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R2a_2", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R2b_1", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R2b_2", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R2c_1", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R2c_2", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R3a_1", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R3a_2", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R3b_1", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R3b_2", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R3c_1", " ")
            pdfFormFields.SetField(pdfFieldFullPath + "R3c_2", " ")

            'Q1
            cSQL = "SELECT FE_TYPE, FE_AER, FE_CWER, FE_MCER, FE_APA, FE_CAPA, FE_MCAPA,FE_AERNOT ,FE_APANOT  " _
                        & " FROM [FOREIGNEQUITY]" _
                        & " WHERE [FE_REF_NO]='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                        & "' And FE_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                If IsDBNull(dr("FE_TYPE")) = False Then
                    If dr("FE_TYPE") = "70 % - 100 %" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_1", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_2", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_3", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_4", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_5", "")
                    End If

                    If dr("FE_TYPE") = "51 % - 69 %" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_2", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_3", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_4", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_5", "")
                    End If

                    If dr("FE_TYPE") = "20 % - 50 %" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_3", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_2", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_4", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_5", "")
                    End If

                    If dr("FE_TYPE") = "<= 19 %" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_4", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_2", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_3", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_5", "")
                    End If

                    If dr("FE_TYPE") = "NIL" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_5", "X")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_2", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_3", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R1_4", "")
                    End If
                End If

                ' If Not Appliable is checked
                If IsDBNull(dr("FE_AERNOT")) = False Then
                    If dr("FE_AERNOT").ToString = "Y" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "R2a_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R2a_2", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R2b_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R2b_2", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R2c_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R2c_2", "")
                    Else
                        'Q2a
                        If IsDBNull(dr("FE_AER")) = False Then
                            If dr("FE_AER").ToString = "N" Then
                                pdfFormFields.SetField(pdfFieldFullPath + "R2a_2", "X")
                                pdfFormFields.SetField(pdfFieldFullPath + "R2a_1", "")

                                ' set Q2b, Q2c to blank
                                pdfFormFields.SetField(pdfFieldFullPath + "R2b_1", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "R2b_2", "")

                                pdfFormFields.SetField(pdfFieldFullPath + "R2c_1", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "R2c_2", "")
                            Else
                                pdfFormFields.SetField(pdfFieldFullPath + "R2a_1", "X")
                                pdfFormFields.SetField(pdfFieldFullPath + "R2a_2", "")
                            End If
                        End If

                        'Q2b
                        If IsDBNull(dr("FE_MCER")) = False Then
                            If dr("FE_AER").ToString = "Y" Then
                                If dr("FE_CWER").ToString = "N" Then
                                    pdfFormFields.SetField(pdfFieldFullPath + "R2b_2", "X")
                                    pdfFormFields.SetField(pdfFieldFullPath + "R2b_1", "")
                                Else
                                    pdfFormFields.SetField(pdfFieldFullPath + "R2b_1", "X")
                                    pdfFormFields.SetField(pdfFieldFullPath + "R2b_2", "")
                                End If
                            End If
                        End If

                        'Q2c
                        If IsDBNull(dr("FE_CWER")) = False Then
                            If dr("FE_AER").ToString = "Y" Then
                                If dr("FE_MCER").ToString = "N" Then
                                    pdfFormFields.SetField(pdfFieldFullPath + "R2c_1", "")
                                    pdfFormFields.SetField(pdfFieldFullPath + "R2c_2", "X")
                                Else
                                    pdfFormFields.SetField(pdfFieldFullPath + "R2c_1", "X")
                                    pdfFormFields.SetField(pdfFieldFullPath + "R2c_2", "")
                                End If
                            End If
                        End If
                    End If
                End If

                ' If Not Appliable is checked
                If IsDBNull(dr("FE_APANOT")) = False Then
                    'Q3a, Q3b, Q3c
                    If dr("FE_APANOT").ToString = "Y" Then
                        pdfFormFields.SetField(pdfFieldFullPath + "R3a_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R3a_2", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R3b_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R3b_2", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R3c_1", "")
                        pdfFormFields.SetField(pdfFieldFullPath + "R3c_2", "")
                    Else
                        If IsDBNull(dr("FE_APA")) = False Then
                            If dr("FE_APA").ToString = "Y" Then
                                pdfFormFields.SetField(pdfFieldFullPath + "R3a_2", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "R3a_1", "X")
                                If dr("FE_MCAPA").ToString = "Y" Then
                                    pdfFormFields.SetField(pdfFieldFullPath + "R3b_1", "X")
                                    pdfFormFields.SetField(pdfFieldFullPath + "R3b_2", "")
                                Else
                                    pdfFormFields.SetField(pdfFieldFullPath + "R3b_2", "X")
                                    pdfFormFields.SetField(pdfFieldFullPath + "R3b_1", "")
                                End If

                                If dr("FE_CAPA").ToString = "Y" Then
                                    pdfFormFields.SetField(pdfFieldFullPath + "R3c_1", "X")
                                    pdfFormFields.SetField(pdfFieldFullPath + "R3c_2", "")
                                Else
                                    pdfFormFields.SetField(pdfFieldFullPath + "R3c_2", "X")
                                    pdfFormFields.SetField(pdfFieldFullPath + "R3c_1", "")
                                End If
                            Else
                                pdfFormFields.SetField(pdfFieldFullPath + "R3a_1", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "R3a_2", "X")

                                pdfFormFields.SetField(pdfFieldFullPath + "R3b_1", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "R3b_2", "")

                                pdfFormFields.SetField(pdfFieldFullPath + "R3c_1", "")
                                pdfFormFields.SetField(pdfFieldFullPath + "R3c_2", "")
                            End If
                        End If
                    End If
                End If
            End If

            ' ==== Part S ====
            If frmDownloadDetails.lstViewAuditor.CheckedItems.Count > 0 Then
                cSQL = "SELECT AD_CO_NAME, AD_ADD, AD_ADD_POSTCODE, AD_ADD_CITY, AD_ADD_STATE, AD_TEL_NO" _
                & " FROM [AUDITOR_PROFILE]" _
                & " WHERE [AD_KEY]= " & frmDownloadDetails.lstViewAuditor.CheckedItems(0).SubItems(0).Text

                dr = DataHandler.GetDataReader(cSQL, Conn)

                If dr.Read() Then
                    If IsDBNull(dr("AD_CO_NAME")) = False Then
                        If Len(dr("AD_CO_NAME")) > 24 Then
                            CutLine(dr("AD_CO_NAME").ToString.ToUpper(), 24)
                            pdfFormFields.SetField(pdfFieldFullPath + "S1_1", strCropped)
                            pdfFormFields.SetField(pdfFieldFullPath + "S1_2", Mid(LTrim(strRemainder), 1, 28))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "S1_1", dr("AD_CO_NAME").ToString.ToUpper())
                            pdfFormFields.SetField(pdfFieldFullPath + "S1_2", space(24))
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "S1_1", space(24))
                        pdfFormFields.SetField(pdfFieldFullPath + "S1_2", space(24))
                    End If
                    If IsDBNull(dr("AD_ADD")) = False Then
                        If Len(dr("AD_ADD").ToString()) < 24 Then
                            pdfFormFields.SetField(pdfFieldFullPath + "S2_1", dr("AD_ADD").ToString.ToUpper.Replace(vbCrLf, " "))
                        Else
                            Call CutLine(dr("AD_ADD").ToString.ToUpper(), 24)
                            pdfFormFields.SetField(pdfFieldFullPath + "S2_1", strCropped)
                            If Len(strRemainder) < 24 Then
                                pdfFormFields.SetField(pdfFieldFullPath + "S2_2", strRemainder.Replace(vbCrLf, " "))
                                pdfFormFields.SetField(pdfFieldFullPath + "S2_3", space(24))
                            Else
                                Call CutLine(strRemainder, 24)
                                pdfFormFields.SetField(pdfFieldFullPath + "S2_2", strCropped)
                                If Len(strRemainder) > 24 Then
                                    pdfFormFields.SetField(pdfFieldFullPath + "S2_3", strRemainder.Substring(0, 24).Replace(vbCrLf, " "))
                                Else
                                    pdfFormFields.SetField(pdfFieldFullPath + "S2_3", strRemainder)
                                End If
                            End If
                        End If
                    End If

                    If IsDBNull(dr("AD_ADD_POSTCODE")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "S2_4", Mid(dr("AD_ADD_POSTCODE"), 1, 5))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "S2_4", space(5))
                    End If

                    If IsDBNull(dr("AD_ADD_CITY")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "S2_5", Mid(dr("AD_ADD_CITY").ToString.ToUpper(), 1, 16))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "S2_5", space(16))
                    End If

                    If IsDBNull(dr("AD_ADD_STATE")) = False Then
                        pdfFormFields.SetField(pdfFieldFullPath + "S2_6", Mid(dr("AD_ADD_STATE").ToString.ToUpper(), 1, 24))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "S2_6", space(24))
                    End If

                    If IsDBNull(dr("AD_TEL_NO")) = False Then
                        If InStr(dr("AD_TEL_NO"), "-") = 3 Then
                            pdfFormFields.SetField(pdfFieldFullPath + "S3", " " + Replace(dr("AD_TEL_NO"), "-", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "S3", Replace(dr("AD_TEL_NO"), "-", ""))
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "S3", space(12))
                    End If

                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "S1_1", space(24))
                    pdfFormFields.SetField(pdfFieldFullPath + "S1_2", space(24))
                    pdfFormFields.SetField(pdfFieldFullPath + "S2_1", space(24))
                    pdfFormFields.SetField(pdfFieldFullPath + "S2_2", space(24))
                    pdfFormFields.SetField(pdfFieldFullPath + "S2_3", space(24))
                    pdfFormFields.SetField(pdfFieldFullPath + "S2_4", space(5))
                    pdfFormFields.SetField(pdfFieldFullPath + "S2_5", space(16))
                    pdfFormFields.SetField(pdfFieldFullPath + "S2_6", space(24))
                    pdfFormFields.SetField(pdfFieldFullPath + "S3", space(12))
                End If

            Else
                pdfFormFields.SetField(pdfFieldFullPath + "S1_1", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "S1_2", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "S2_1", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "S2_2", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "S2_3", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "S2_4", space(5))
                pdfFormFields.SetField(pdfFieldFullPath + "S2_5", space(16))
                pdfFormFields.SetField(pdfFieldFullPath + "S2_6", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "S3", space(12))
            End If

        Catch ex As Exception
            MsgBox("Some important data is not fill in page 14!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Page15()
        Dim intIndex As Integer = 0
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page17[0]."
        Dim strTA_ADD As String
        Try
            ' ===== Part T =====
            cSQL = "SELECT TA_CO_NAME, TA_ADD_LINE1, TA_ADD_LINE2, TA_ADD_LINE3, TA_ADD_POSTCODE, TA_ADD_CITY, TA_ADD_STATE, TA_TEL_NO," _
                   & " TA_LICENSE, TA_ROC_NO, TA_EMAIL" _
                   & " FROM [TAXA_PROFILE]"

            cSQL += " WHERE [TA_KEY] =" & frmDownloadDetails.lstViewTaxAgent.Items(frmDownloadDetails.lstViewTaxAgent.CheckedItems(0).Index).Text

            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then
                If IsDBNull(dr("TA_CO_NAME")) = False And dr("TA_CO_NAME") <> "" Then
                    If Len(dr("TA_CO_NAME")) > 24 Then
                        CutLine(dr("TA_CO_NAME").ToString.ToUpper(), 24)
                        pdfFormFields.SetField(pdfFieldFullPath + "T1_1", strCropped)
                        pdfFormFields.SetField(pdfFieldFullPath + "T1_2", Mid(LTrim(strRemainder), 1, 28))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "T1_1", dr("TA_CO_NAME").ToString.ToUpper())
                        pdfFormFields.SetField(pdfFieldFullPath + "T1_2", space(24))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "T1_1", space(24))
                    pdfFormFields.SetField(pdfFieldFullPath + "T1_2", space(24))
                End If

                If IsDBNull(dr("TA_ADD_LINE1")) = False And dr("TA_ADD_LINE1") <> "" Then

                    strTA_ADD = ""
                    strCropped = ""
                    strRemainder = ""


                    If Trim(dr("TA_ADD_LINE2")) <> "" Then
                        If Not Right(Trim(dr("TA_ADD_LINE1")), 1) = "," Then
                            strTA_ADD = dr("TA_ADD_LINE1") + ", "
                        Else
                            strTA_ADD = dr("TA_ADD_LINE1") + " "
                        End If
                    End If

                    If Trim(dr("TA_ADD_LINE3")) <> "" Then
                        If Not Right(Trim(dr("TA_ADD_LINE2")), 1) = "," Then
                            strTA_ADD = strTA_ADD + dr("TA_ADD_LINE2") + ", " + dr("TA_ADD_LINE3")
                        Else
                            strTA_ADD = strTA_ADD + dr("TA_ADD_LINE2") + " " + dr("TA_ADD_LINE3")
                        End If
                    Else
                        If Not Right(Trim(dr("TA_ADD_LINE2")), 1) = "," Then
                            strTA_ADD = strTA_ADD + dr("TA_ADD_LINE2")
                        Else
                            strTA_ADD = strTA_ADD + dr("TA_ADD_LINE2")
                        End If
                    End If

                    If Trim(dr("TA_ADD_LINE2")) = "" And Trim(dr("TA_ADD_LINE3")) = "" Then
                        strTA_ADD = dr("TA_ADD_LINE1")
                    End If

                    If Len(strTA_ADD) > 24 Then
                        Call CutLine(strTA_ADD, 24)

                        pdfFormFields.SetField(pdfFieldFullPath + "T2_1", strCropped.ToUpper)

                        If Len(strRemainder) > 24 Then
                            Call CutLine(strRemainder, 24)
                            pdfFormFields.SetField(pdfFieldFullPath + "T2_2", strCropped.ToUpper)
                            pdfFormFields.SetField(pdfFieldFullPath + "T2_3", strRemainder.ToUpper)
                        ElseIf Len(strRemainder) > 0 Then
                            pdfFormFields.SetField(pdfFieldFullPath + "T2_2", strRemainder.ToUpper)
                            pdfFormFields.SetField(pdfFieldFullPath + "T2_3", space(24))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "T2_2", space(24))
                            pdfFormFields.SetField(pdfFieldFullPath + "T2_3", space(24))
                        End If

                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "T2_1", strTA_ADD.ToUpper)
                        pdfFormFields.SetField(pdfFieldFullPath + "T2_2", space(24))
                        pdfFormFields.SetField(pdfFieldFullPath + "T2_3", space(24))
                    End If

                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "T2_1", space(24))
                    pdfFormFields.SetField(pdfFieldFullPath + "T2_2", space(24))
                    pdfFormFields.SetField(pdfFieldFullPath + "T2_3", space(24))
                End If

                If IsDBNull(dr("TA_ADD_POSTCODE")) = False And dr("TA_ADD_POSTCODE") <> "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "T2_4", Mid(dr("TA_ADD_POSTCODE"), 1, 5).ToString.ToUpper())
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "T2_4", space(5))
                End If

                If IsDBNull(dr("TA_ADD_CITY")) = False And dr("TA_ADD_CITY") <> "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "T2_5", Mid(dr("TA_ADD_CITY"), 1, 16).ToString.ToUpper())
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "T2_5", space(16))
                End If

                If IsDBNull(dr("TA_ADD_STATE")) = False And dr("TA_ADD_STATE") <> "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "T2_6", Mid(dr("TA_ADD_STATE"), 1, 24).ToString.ToUpper())
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "T2_6", space(24))
                End If

                If IsDBNull(dr("TA_TEL_NO")) = False And dr("TA_TEL_NO") <> "" Then
                    If InStr(dr("TA_TEL_NO"), "-") = 3 Then
                        pdfFormFields.SetField(pdfFieldFullPath + "T3", " " + Replace(dr("TA_TEL_NO"), "-", ""))
                    Else
                        pdfFormFields.SetField(pdfFieldFullPath + "T3", Replace(dr("TA_TEL_NO"), "-", ""))
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "T3", space(12))
                End If

                If IsDBNull(dr("TA_LICENSE")) = False And dr("TA_LICENSE") <> "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "T4", Mid(dr("TA_LICENSE"), 1, 15))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "T4", space(15))
                End If

                If IsDBNull(dr("TA_ROC_NO")) = False And dr("TA_ROC_NO") <> "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "T5", Mid(dr("TA_ROC_NO"), 1, 12))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "T5", space(12))
                End If

                If IsDBNull(dr("TA_EMAIL")) = False And dr("TA_EMAIL") <> "" Then
                    pdfFormFields.SetField(pdfFieldFullPath + "T6", Mid(dr("TA_EMAIL"), 1, 24))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "T6", space(24))
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "T1_1", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "T1_2", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "T2_1", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "T2_2", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "T2_3", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "T2_4", space(5))
                pdfFormFields.SetField(pdfFieldFullPath + "T2_5", space(16))
                pdfFormFields.SetField(pdfFieldFullPath + "T2_6", space(24))
                pdfFormFields.SetField(pdfFieldFullPath + "T3", space(12))
                pdfFormFields.SetField(pdfFieldFullPath + "T4", space(12))
                pdfFormFields.SetField(pdfFieldFullPath + "T5", space(12))
                pdfFormFields.SetField(pdfFieldFullPath + "T6", space(24))
            End If

            If frmDownloadDetails.dtpPrintDateA.Checked = True Then
                pdfFormFields.SetField(pdfFieldFullPath + "TTarikh", Mid(Format(frmDownloadDetails.dtpPrintDateA.Value, "ddMMyyyy"), 1, 8))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "TTarikh", "")
            End If

            '===== Part Borang Akuan =====
            If frmDownloadDetails.txtName.Text <> "" Then
                If Len(frmDownloadDetails.txtName.Text) > 28 Then
                    CutLine(frmDownloadDetails.txtName.Text.ToString.ToUpper(), 28)
                    pdfFormFields.SetField(pdfFieldFullPath + "Akuan_1", strCropped)
                    pdfFormFields.SetField(pdfFieldFullPath + "Akuan_2", Mid(LTrim(strRemainder), 1, 28))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "Akuan_1", frmDownloadDetails.txtName.Text.ToUpper)
                    pdfFormFields.SetField(pdfFieldFullPath + "Akuan_2", space(28))
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "Akuan_1", space(28))
                pdfFormFields.SetField(pdfFieldFullPath + "Akuan_2", space(28))
            End If

            If frmDownloadDetails.txtIC.Text <> "" Then
                pdfFormFields.SetField(pdfFieldFullPath + "Akuan_3", Mid(Replace(frmDownloadDetails.txtIC.Text, "-", ""), 1, 12).ToString.ToUpper())
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "Akuan_3", space(12))
            End If
            If frmDownloadDetails.txtPosition.Text <> "" Then
                If Len(frmDownloadDetails.txtPosition.Text) > 28 Then
                    CutLine(frmDownloadDetails.txtPosition.Text.ToString.ToUpper(), 28)
                    pdfFormFields.SetField(pdfFieldFullPath + "Akuan_5", strCropped)
                    pdfFormFields.SetField(pdfFieldFullPath + "Akuan_6", Mid(LTrim(strRemainder), 1, 28))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "Akuan_5", frmDownloadDetails.txtPosition.Text.ToUpper)
                    pdfFormFields.SetField(pdfFieldFullPath + "Akuan_6", space(28))
                End If
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "Akuan_5", space(28))
                pdfFormFields.SetField(pdfFieldFullPath + "Akuan_6", space(28))
            End If


            '==Liko 2010.2=='
            If frmDownloadDetails.dtpPrintDate.Checked = True Then
                pdfFormFields.SetField(pdfFieldFullPath + "Akuan_4", Mid(Format(frmDownloadDetails.dtpPrintDate.Value, "ddMMyyyy"), 1, 8))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "Akuan_4", "")
            End If
            '==End Liko2010.2=='



        Catch ex As Exception
            MsgBox("Some important data is not fill in page 15!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

    Public Sub Slip()
        Dim cSQL As String
        Dim pdfFieldFullPath As String = pdfFieldPath + "Page19[0]."
        Dim strTempString As String = ""
        Dim add1 As String = ""
        Dim lenadd1 As Integer
        Dim lenadd2 As Integer
        Dim add2 As String = ""
        Dim Total As Double

        Try
            ' ==== Part Borang C ====
            cSQL = "Select *  from TAXP_PROFILE where TP_REF_NO = '" _
                    & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)

            If dr.Read() Then
                strTempString = "C" + dr("TP_REF_NO")
                pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_1", strTempString) ' Borang C
                pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_1", strTempString) ' Borang R
                add1 = dr("TP_COM_NAME")
                lenadd1 = Len(add1)
                If lenadd1 > 0 Then
                    lenadd2 = 138 - lenadd1 / 2
                Else
                    lenadd2 = 0
                End If
                strTempString = add1 + spaceSlip(lenadd2) + spaceSlip(138) _
                    + " " + dr("TP_CURR_ADD_LINE1") + IIf(Trim(dr("TP_CURR_ADD_LINE1")) <> "", IIf(InStr(Trim(dr("TP_CURR_ADD_LINE1")), ","), " ", " , "), " ") _
                    + dr("TP_CURR_ADD_LINE2") + IIf(Trim(dr("TP_CURR_ADD_LINE2")) <> "", IIf(InStr(Trim(dr("TP_CURR_ADD_LINE2")), ","), " ", " , "), " ") _
                    + dr("TP_CURR_ADD_LINE3") + IIf(Trim(dr("TP_CURR_ADD_LINE3")) <> "", IIf(InStr(Trim(dr("TP_CURR_ADD_LINE3")), ","), " ", " , "), " ") _
                    + dr("TP_CURR_POSTCODE") + " " _
                    + dr("TP_CURR_CITY") + IIf(Trim(dr("TP_CURR_CITY")) <> "", IIf(InStr(Trim(dr("TP_CURR_CITY")), ","), " ", " , "), " ") _
                    + dr("TP_CURR_STATE")
                pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_3", strTempString.ToString.ToUpper) ' Borang C
                pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_3", strTempString.ToString.ToUpper) ' Borang R
                pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_4", dr("TP_ROC_NO").ToString.ToUpper) ' Borang C
                pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_4", dr("TP_ROC_NO").ToString.ToUpper) ' Borang R

                If IsDBNull(dr("TP_TEL_NO")) = False Then
                    ' Borang C
                    pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_7", dr("TP_TEL_NO"))
                    ' Borang R
                    pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_7", dr("TP_TEL_NO"))
                Else
                    pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_7", space(12)) ' Borang C
                    pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_7", space(12)) ' Borang R
                End If

                cSQL = "SELECT [PL_S60F] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO]='" _
                    & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' And PL_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
                dr = DataHandler.GetDataReader(cSQL, Conn)
                dr.Read()
                If dr("PL_S60F") = "Y" Then ' Investment Holding '
                    cSQL = "SELECT IH_TP_BAL" _
                        & " FROM [INVESTMENT_HOLDING]" _
                        & " WHERE [IH_REF_NO] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                        & "' And IH_YA='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
                    dr = DataHandler.GetDataReader(cSQL, Conn)
                    If dr.Read() Then
                        pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_2", Replace((Convert.ToDouble((dr("IH_TP_BAL"))).ToString("0.00")), ".", "").Replace(",", "").Replace("-", ""))
                    End If
                    dr.Close()
                Else
                    cSQL = "Select * FROM TAX_COMPUTATION" _
                            & " WHERE [TC_REF_NO] = '" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                            & "' And TC_YA ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'" _
                            & " Order by TC_BUSINESS"
                    dr = DataHandler.GetDataReader(cSQL, Conn)
                    If dr.Read() Then
                        Dim TotalB4 As Double
                        If CDbl(dr("TC_TP_PAYABLE")) >= 0 Then
                            TotalB4 = Convert.ToDouble(dr("TC_TP_PAYABLE")) - Convert.ToDouble(dr("TC_TP_INSTALL"))
                        Else
                            TotalB4 = 0 - Convert.ToDouble(dr("TC_TP_INSTALL"))
                        End If

                        If TotalB4 >= 0 Then
                            pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_2", Replace((TotalB4.ToString("0.00")), ".", "").Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_2", "000")
                        End If
                    End If
                    dr.Close()
                End If
            End If
            dr.Close()

            ' === Bahagian R  : Amount ===
            cSQL = "SELECT * from BORANGR2008 " _
                    & " WHERE BR08_REF_NO ='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value) _
                    & "' AND BR08_YA='" & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
            dr = DataHandler.GetDataReader(cSQL, Conn)
            If dr.Read() Then
                Total = CDbl(dr("BR08_BA_HUTANGKERAJAAN2")) + CDbl(dr("BR08_BA_CUKAITAKLAYAK"))
                pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_2", Replace((Total.ToString("0.00")), ".", "").Replace(",", ""))
            Else
                pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_2", "000")
            End If
            dr.Close()

            ' ==== Set Cek, Bank , Tarikh to blank
            pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_5", "   ")
            pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_6", "   ")
            pdfFormFields.SetField(pdfFieldFullPath + "Slip_C_8", "   ")

            pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_5", "   ")
            pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_6", "   ")
            pdfFormFields.SetField(pdfFieldFullPath + "Slip_R_8", "   ")

            pdfStamper.FormFlattening = False
            pdfStamper.Close()

        Catch ex As Exception
            MsgBox("Some important data is not fill in Slip Pengakuan!", MsgBoxStyle.Critical, "Caution")
            pdfStamper.Close()
        End Try
    End Sub

#End Region

    Private Function spaceSlip(ByVal intSpace As Integer) As String
        Dim intA As String = ""
        For i As Integer = 0 To (intSpace - 4)
            intA = intA + " "
        Next
        spaceSlip = intA
    End Function

    Public Function L8Calculation() As Double

        Dim cSQL As String = ""
        Dim dr4 As OleDbDataReader
        Dim dr3 As OleDbDataReader
        'Dim dr2 As OleDbDataReader
        Dim odSales As Double = 0.0
        Dim odOs As Double = 0.0
        Dim odPur As Double = 0.0
        Dim odDep As Double = 0.0
        Dim odA As Double = 0.0
        Dim odNA As Double = 0.0
        Dim odCS As Double = 0.0
        Dim TotalL8 As Double = 0.0

        cSQL = "SELECT [PL_KEY],[PL_MAINBUZ] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" _
        & frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value & "' AND [PL_YA] = '" _
        & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
        dr3 = DataHandler.GetDataReader(cSQL, Conn)

        If dr3.Read() Then
            cSQL = "SELECT SUM(cdbl(PLFS_AMOUNT)) FROM [PLFST_SALES] WHERE [PLFS_KEY] = " & dr3("PL_KEY") _
            & " AND NOT [PLFS_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                If IsDBNull(dr4(0)) = False Then
                    odSales = Convert.ToDouble(dr4(0))
                Else
                    odSales = 0.0
                End If
            Else
                odSales = 0.0
            End If
            cSQL = "SELECT SUM(cdbl(PLFOS_AMOUNT)) FROM [PLFST_OPENSTOCK] WHERE [PLFOS_KEY] = " _
            & dr3("PL_KEY") & " AND NOT [PLFOS_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                If IsDBNull(dr4(0)) = False Then
                    odOs = Convert.ToDouble(dr4(0))
                Else
                    odOs = 0.0
                End If
            Else
                odOs = 0.0
            End If
            cSQL = "SELECT SUM(cdbl(PLFPUR_AMOUNT)) FROM [PLFST_PURCHASE] WHERE [PLFPUR_KEY] = " _
            & dr3("PL_KEY") & " AND NOT [PLFPUR_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                If IsDBNull(dr4(0)) = False Then
                    odPur = Convert.ToDouble(dr4(0))
                Else
                    odPur = 0.0
                End If
            Else
                odPur = 0.0
            End If
            cSQL = "SELECT SUM(cdbl(EXDEP_AMOUNT)) FROM [EXPENSES_DEPRECIATION] WHERE [EXDEP_KEY] = " _
            & dr3("PL_KEY") & " AND NOT [EXDEP_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                If IsDBNull(dr4(0)) = False Then
                    odDep = Convert.ToDouble(dr4(0))
                Else
                    odDep = 0.0
                End If
            Else
                odDep = 0.0
            End If
            cSQL = "SELECT SUM(cdbl(EXA_AMOUNT)) FROM [EXPENSES_ALLOW] WHERE [EXA_KEY] = " _
            & dr3("PL_KEY") & " AND NOT [EXA_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                If IsDBNull(dr4(0)) = False Then
                    odA = Convert.ToDouble(dr4(0))
                Else
                    odA = 0.0
                End If
            Else
                odA = 0.0
            End If
            cSQL = "SELECT SUM(cdbl(EXNA_AMOUNT)) FROM [EXPENSES_NONALLOW] WHERE [EXNA_KEY] = " _
            & dr3("PL_KEY") & " AND NOT [EXNA_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                If IsDBNull(dr4(0)) = False Then
                    odNA = Convert.ToDouble(dr4(0))
                Else
                    odNA = 0.0
                End If
            Else
                odNA = 0.0
            End If
            cSQL = "SELECT SUM(cdbl(PLFCS_AMOUNT)) FROM [PLFST_CLOSESTOCK] WHERE [PLFCS_KEY] = " _
            & dr3("PL_KEY") & " AND NOT [PLFCS_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                If IsDBNull(dr4(0)) = False Then
                    odCS = Convert.ToDouble(dr4(0))
                Else
                    odCS = 0.0
                End If
            Else
                odCS = 0.0
            End If
        End If

        TotalL8 = odSales - (odOs + odPur + odDep + odA + odNA - odCS)

        L8Calculation = TotalL8

    End Function

    Public Sub L8GetBSCode()
        Dim cSQL As String = ""
        Dim dr2 As OleDbDataReader
        Dim dr3 As OleDbDataReader
        Dim dr4 As OleDbDataReader

        cSQL = "SELECT [PL_KEY],[PL_MAINBUZ] FROM [PROFIT_LOSS_ACCOUNT] WHERE [PL_REF_NO] = '" _
        & frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value & "' AND [PL_YA] = '" _
        & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "'"
        dr3 = DataHandler.GetDataReader(cSQL, Conn)
        If dr3.Read() Then
            cSQL = "Select BC_CODE from BUSINESS_SOURCE where BC_KEY = '" _
            & frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(1).Value & "' and BC_YA = '" _
            & Trim(frmDownloadMainMenu.dgdDownload.SelectedRows.Item(0).Cells(2).Value) & "' and BC_SOURCENO = " _
            & dr3("PL_MAINBUZ") & ""
            dr2 = DataHandler.GetDataReader(cSQL, Conn)
            If dr2.Read() Then
                If Trim(dr2("BC_CODE")) <> "" Then
                    BSCode = Trim(dr2("BC_CODE"))
                    'add code here
                Else
                    BSCode = "-"
                    'add code here
                End If
            End If
        End If

        dSales = 0
        dOS = 0
        dDep = 0
        dPur = 0
        dA = 0
        dNA = 0
        dCS = 0
        dGP = 0


        cSQL = "SELECT [PLFS_AMOUNT] FROM [PLFST_SALES] WHERE [PLFS_KEY] = " & dr3("PL_KEY") _
        & " AND [PLFS_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
        dr2 = DataHandler.GetDataReader(cSQL, Conn)
        If dr2.Read Then
            cSQL = "SELECT SUM(cdbl(PLFS_AMOUNT)) FROM [PLFST_SALES] WHERE [PLFS_KEY] = " & dr3("PL_KEY") _
            & " AND [PLFS_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                dSales = Convert.ToDouble(dr4(0))
            End If
        Else
            dSales = 0.0
        End If

        cSQL = "SELECT [PLFOS_AMOUNT] FROM [PLFST_OPENSTOCK] WHERE [PLFOS_KEY] = " & dr3("PL_KEY") _
        & " AND [PLFOS_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
        dr2 = DataHandler.GetDataReader(cSQL, Conn)
        If dr2.Read Then
            cSQL = "SELECT SUM(cdbl(PLFOS_AMOUNT)) FROM [PLFST_OPENSTOCK] WHERE [PLFOS_KEY] = " & dr3("PL_KEY") _
            & " AND [PLFOS_SOURCENO] = " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                dOS = Convert.ToDouble(dr4(0))
            End If
        Else
            dOS = 0
        End If

        cSQL = "SELECT [PLFPUR_AMOUNT] FROM [PLFST_PURCHASE] WHERE [PLFPUR_KEY] = " & dr3("PL_KEY") _
        & " AND [PLFPUR_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
        dr2 = DataHandler.GetDataReader(cSQL, Conn)
        If dr2.Read Then
            cSQL = "SELECT SUM(cdbl(PLFPUR_AMOUNT)) FROM [PLFST_PURCHASE] WHERE [PLFPUR_KEY] = " & dr3("PL_KEY") _
            & " AND [PLFPUR_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                dPur = Convert.ToDouble(dr4(0))
            End If
        Else
            dPur = 0
        End If

        cSQL = "SELECT [EXDEP_AMOUNT] FROM [EXPENSES_DEPRECIATION] WHERE [EXDEP_KEY] = " & dr3("PL_KEY") _
        & " AND [EXDEP_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
        dr2 = DataHandler.GetDataReader(cSQL, Conn)
        If dr2.Read Then
            cSQL = "SELECT SUM(cdbl(EXDEP_AMOUNT)) FROM [EXPENSES_DEPRECIATION] WHERE [EXDEP_KEY] = " & dr3("PL_KEY") _
            & " AND [EXDEP_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                dDep = Convert.ToDouble(dr4(0))
            End If
        Else
            dDep = 0
        End If

        cSQL = "SELECT [EXA_AMOUNT] FROM [EXPENSES_ALLOW] WHERE [EXA_KEY] = " & dr3("PL_KEY") _
        & " AND [EXA_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
        dr2 = DataHandler.GetDataReader(cSQL, Conn)
        If dr2.Read Then
            cSQL = "SELECT SUM(cdbl(EXA_AMOUNT)) FROM [EXPENSES_ALLOW] WHERE [EXA_KEY] = " & dr3("PL_KEY") _
            & " AND [EXA_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                dA = Convert.ToDouble(dr4(0))
            End If
        Else
            dA = 0
        End If

        cSQL = "SELECT [EXNA_AMOUNT] FROM [EXPENSES_NONALLOW] WHERE [EXNA_KEY] = " & dr3("PL_KEY") _
        & " AND [EXNA_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
        dr2 = DataHandler.GetDataReader(cSQL, Conn)
        If dr2.Read Then
            cSQL = "SELECT SUM(cdbl(EXNA_AMOUNT)) FROM [EXPENSES_NONALLOW] WHERE [EXNA_KEY] = " _
            & dr3("PL_KEY") & " AND [EXNA_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                dNA = Convert.ToDouble(dr4(0))
            End If
        Else
            dNA = 0
        End If

        cSQL = "SELECT [PLFCS_AMOUNT] FROM [PLFST_CLOSESTOCK] WHERE [PLFCS_KEY] = " & dr3("PL_KEY") _
        & " AND [PLFCS_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
        dr2 = DataHandler.GetDataReader(cSQL, Conn)
        If dr2.Read Then
            cSQL = "SELECT SUM(cdbl(PLFCS_AMOUNT)) FROM [PLFST_CLOSESTOCK] WHERE [PLFCS_KEY] = " _
            & dr3("PL_KEY") & " AND [PLFCS_SOURCENO] =  " & dr3("PL_MAINBUZ") & ""
            dr4 = DataHandler.GetDataReader(cSQL, Conn)
            If dr4.Read Then
                dCS = Convert.ToDouble(dr4(0))
            End If
        Else
            dCS = 0
        End If

        dGP = (dSales - (dOS + dPur + dDep + dA + dNA - dCS))

    End Sub

    Public Sub CutLine(ByVal strCut As String, ByVal cLength As Integer)
        Dim I As Integer
        Dim cPoint As Integer
        Dim strLineRef As String
        'Dim NOM As Integer
        Dim chkspace As Integer

        strLineRef = Trim(strCut)

        'LEESH 13-APR-2012
        strLineRef = strLineRef.Replace(vbCrLf, " ")
        strLineRef = Trim(strLineRef)
        'LEESH END

        If Mid(Trim(strLineRef), cLength, 1) <> " " And Mid(Trim(strLineRef), cLength + 1, 1) <> " " And Mid(Trim(strLineRef), cLength, 1) <> "," And Mid(Trim(strLineRef), cLength + 1, 1) <> "," Then
            I = 0
            chkspace = InStr(1, strLineRef, " ")
            If chkspace <> 0 Then 'pin

                If chkspace < cLength Then
                    While Mid(strLineRef, cLength - I, 1) <> " "  ' And i < cLength
                        I = I + 1
                    End While
                    cPoint = I
                    strCropped = LTrim(Mid(strLineRef, 1, cLength - cPoint))
                    strRemainder = LTrim(Mid(strLineRef, cLength - cPoint, Len(strLineRef) - (cLength - cPoint) + 1))
                Else
                    strCropped = LTrim(Mid(strLineRef, 1, cLength))
                    strRemainder = LTrim(Mid(strLineRef, cLength + 1, cLength))

                End If
            Else

                If Len(strLineRef) > cLength Then
                    strCropped = LTrim(Mid(strLineRef, 1, cLength))
                    strRemainder = LTrim(Mid(strLineRef, cLength + 1, Len(strLineRef) - cLength))
                Else
                    strCropped = LTrim(Mid(strLineRef, 1, cLength))
                End If

            End If
        Else
            strCropped = LTrim(Mid(strLineRef, 1, cLength))
            strRemainder = LTrim(Mid(strLineRef, cLength + 1, Len(strLineRef) - cLength))
        End If
    End Sub

    Private Function space(ByVal intSpace As Integer) As String
        Dim intA As String = ""
        For i As Integer = 0 To (intSpace - 4)
            intA = intA + " "
        Next
        intA = intA & "---"
        space = intA
    End Function

    Private Sub closePDF()
        MsgBox("Please select a correct template!", MsgBoxStyle.Critical, "Caution")
        pdfStamper.Close()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Function TextAddressSpliter(ByVal strAddress As String(), ByVal intLength As Integer) As String()
        Dim arrText(2) As String
        Dim i As Integer

        For i = 0 To strAddress.Length - 1
            If (strAddress(i).Length > intLength) Then
                If ((((strAddress(i)).EndsWith(",")) = False) And ((i + 1 < (strAddress.Length)) = True)) And ((i = strAddress.Length - 1) = False) Then
                    If (((strAddress(i).Replace(" ", "")).EndsWith(",")) = False) Then
                        strAddress(i) = strAddress(i) & ","
                    End If
                    If strAddress(i).Substring(0, intLength).LastIndexOf(" ") > 0 Or strAddress(i).Substring(0, intLength).LastIndexOf(",") > 0 Then
                        If strAddress(i).Substring(0, intLength).LastIndexOf(" ") > strAddress(i).Substring(0, intLength).LastIndexOf(",") Then
                            If strAddress(i).Substring(0, intLength + 1).EndsWith(" ") Then
                                strAddress(i + 1) = strAddress(i).Substring(intLength + 1) & strAddress(i + 1)
                                arrText(i) = strAddress(i).Substring(0, intLength)
                            Else
                                strAddress(i + 1) = strAddress(i).Substring(strAddress(i).Substring(0, intLength).LastIndexOf(" ") + 1) & strAddress(i + 1)
                                arrText(i) = strAddress(i).Substring(0, strAddress(i).Substring(0, intLength).LastIndexOf(" ") + 1)
                            End If
                        Else
                            If strAddress(i).Substring(0, intLength + 1).EndsWith(" ") Then
                                strAddress(i + 1) = strAddress(i).Substring(intLength + 1) & strAddress(i + 1)
                                arrText(i) = strAddress(i).Substring(0, intLength)
                            Else
                                strAddress(i + 1) = strAddress(i).Substring(strAddress(i).Substring(0, intLength).LastIndexOf(",") + 1) & strAddress(i + 1)
                                arrText(i) = strAddress(i).Substring(0, strAddress(i).Substring(0, intLength).LastIndexOf(",")) & ","
                            End If
                        End If
                    Else
                        strAddress(i + 1) = strAddress(i).Substring(intLength + 1) & strAddress(i + 1).ToString
                        arrText(i) = strAddress(i).Substring(0, intLength)
                    End If
                Else
                    If strAddress(i).Substring(0, intLength).LastIndexOf(" ") > strAddress(i).Substring(0, intLength).LastIndexOf(",") Then
                        If strAddress(i).Substring(0, intLength).LastIndexOf(" ") < 0 Then
                            arrText(i) = strAddress(i).Substring(0, intLength).ToString()
                        Else
                            arrText(i) = strAddress(i).Substring(0, strAddress(i).Substring(0, intLength).LastIndexOf(" ")).ToString()
                        End If
                    Else
                        If strAddress(i).Substring(0, intLength).LastIndexOf(",") < 0 Then
                            arrText(i) = strAddress(i).Substring(0, intLength).ToString()
                        Else
                            arrText(i) = strAddress(i).Substring(0, strAddress(i).Substring(0, intLength).LastIndexOf(",")).ToString()
                        End If
                    End If

                End If
            Else

                If i + 1 > strAddress.Length - 1 Then
                    If (((strAddress(i).Trim().EndsWith(",")) = False) And ((i = strAddress.Length - 1) = False)) Then
                        arrText(i) = strAddress(i).ToString() + ","
                    Else
                        arrText(i) = strAddress(i).ToString()
                    End If
                Else
                    If (((strAddress(i).Trim().EndsWith(",")) = False) And ((i = strAddress.Length - 1) = False)) And (String.IsNullOrEmpty(strAddress(i + 1)) = False) Then
                        arrText(i) = strAddress(i).ToString() + ","
                    Else
                        arrText(i) = strAddress(i).ToString()
                    End If
                End If
            End If
        Next
        Return arrText

    End Function

End Class
