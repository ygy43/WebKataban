Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports WebKataban.CdCst
Imports System.Linq.Enumerable

Public Class KHKatabanSeparator

    ''' <summary>
    ''' 形番分解情報の取得
    ''' </summary>
    ''' <param name="strKataban"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strKeyKata"></param>
    ''' <param name="strKataName"></param>
    ''' <param name="strSpecNo"></param>
    ''' <param name="strPriceNo"></param>
    ''' <param name="strItem1"></param>
    ''' <param name="strItemName1"></param>
    ''' <param name="strHyphen1"></param>
    ''' <param name="strStructure_div"></param>
    ''' <param name="strElement_div1"></param>
    ''' <param name="ds_table"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSeparatorData(ByVal strKataban As String, _
                              ByRef strSeries As String, ByRef strKeyKata As String, ByRef strKataName As String, _
                              ByRef strSpecNo As String, ByRef strPriceNo As String, ByRef strItem1() As String, _
                              ByRef strItemName1() As String, ByRef strHyphen1() As String, ByRef strStructure_div() As String, _
                              ByRef strElement_div1() As String, Optional ds_table As DataSet = Nothing) As Boolean

        GetSeparatorData = False
        ' 形番分解
        Dim strFirstHyphen As String = String.Empty
        Try
            If Separator(strKataban, strSeries, strItem1, strItemName1, strHyphen1, strElement_div1, strStructure_div, _
                         strFirstHyphen, strKataName, strSpecNo, strPriceNo, strKeyKata, ds_table) Then
                GetSeparatorData = True

                Dim strElement_div(strElement_div1.Length) As String

                strElement_div(0) = String.Empty
                For inti As Integer = 1 To strItem1.Length
                    strElement_div(inti) = strElement_div1(inti - 1)
                Next
                strElement_div1 = strElement_div
            End If
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 新形番分解処理
    ''' </summary>
    ''' <param name="strKata"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strItem"></param>
    ''' <param name="strItemName"></param>
    ''' <param name="strHyphen"></param>
    ''' <param name="strElement_div"></param>
    ''' <param name="strStructure_div"></param>
    ''' <param name="strFirseHyphen"></param>
    ''' <param name="strKataName"></param>
    ''' <param name="strSpecNo"></param>
    ''' <param name="strPriceNo"></param>
    ''' <param name="strKeyKata"></param>
    ''' <param name="ds_table"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function Separator(ByVal strKata As String, ByRef strSeries As String, ByRef strItem() As String, _
                          ByRef strItemName() As String, ByRef strHyphen() As String, ByRef strElement_div() As String, _
                          ByRef strStructure_div() As String, ByRef strFirseHyphen As String, ByRef strKataName As String, _
                          ByRef strSpecNo As String, ByRef strPriceNo As String, _
                          ByRef strKeyKata As String, Optional ds_table As DataSet = Nothing) As Boolean
        Separator = False

        Dim strNowKata As String = String.Empty
        Dim strItemNo() As Integer = Nothing
        Dim dt_Series As New DS_KatSep.kh_series_katabanDataTable
        Dim dt_ItemName As New DS_KatSep.kh_ktbn_strc_nm_mstDataTable
        Dim dt_Hyphen As New DS_KatSep.kh_kataban_strcDataTable
        Dim Mflag As Boolean = False
        Dim dr_Hyphen() As DataRow
        Dim dr_Series As DataRow
        Dim strNewHyphen() As String
        Dim Divlist As New ArrayList
        Dim dr_ItemName() As DataRow
        Dim dt_Option As New DS_KatOut.DT_OptionDataTable
        Dim dt_ElePattern As New DS_KatOut.kh_ele_patternDataTable
        Dim dt_VolStd As New DS_KatOut.kh_std_voltage_mstDataTable
        Dim dt_Stroke As New DS_KatOut.kh_strokeDataTable
        'オプション候補
        Dim dtStrcEle As New DS_KatSep.kh_kataban_strc_eleDataTable

        Try
            If ds_table Is Nothing Then
                '新形番分解のシリーズデータの取得
                Using da As New DS_KatSepTableAdapters.kh_series_katabanTableAdapter
                    dt_Series = da.GetData(strKata)
                End Using
                '新形番分解のItemNameデータの取得
                Using da As New DS_KatSepTableAdapters.kh_ktbn_strc_nm_mstTableAdapter
                    dt_ItemName = da.GetData(strKata)
                End Using

                '新形番分解のHyphenデータの取得
                Using da As New DS_KatSepTableAdapters.kh_kataban_strcTableAdapter
                    dt_Hyphen = da.GetData(strKata)
                End Using

                '新形番分解のOptionデータの取得
                Using da As New DS_KatSepTableAdapters.kh_kataban_strc_eleTableAdapter
                    dtStrcEle = da.GetDataByKata(strKata)
                End Using
            Else
                Dim dr() As DataRow = Nothing
                Dim strWhere As String = "'" & strKata & "' Like series_kataban + '%'"
                dr = ds_table.Tables("dt_Series").Select(strWhere)
                For inti As Integer = 0 To dr.Length - 1
                    dt_Series.ImportRow(dr(inti))
                Next
                dr = ds_table.Tables("dt_ItemName").Select(strWhere)
                For inti As Integer = 0 To dr.Length - 1
                    dt_ItemName.ImportRow(dr(inti))
                Next
                dr = ds_table.Tables("dt_Hyphen").Select(strWhere)
                For inti As Integer = 0 To dr.Length - 1
                    dt_Hyphen.ImportRow(dr(inti))
                Next
                dr = Nothing
            End If
            If dt_Series.Rows.Count <= 0 Then Exit Try
            If dt_ItemName.Rows.Count <= 0 Then Exit Try
            If dt_Hyphen.Rows.Count <= 0 Then Exit Try

            Dim DS_Tab As New DataSet

            'シリーズ毎に探します（一番長いシリーズから）
            For inti As Integer = 0 To dt_Series.Rows.Count - 1
                dr_Series = dt_Series.Rows(inti)
                strKeyKata = dr_Series("key_Kataban").ToString

                If dr_Series("hyphen_div").ToString = "1" And strKata.StartsWith(dr_Series("series_kataban").ToString & "-") Then
                    strNowKata = Mid(strKata, Len(dr_Series("series_kataban").ToString) + 2)
                Else
                    strNowKata = Mid(strKata, Len(dr_Series("series_kataban").ToString) + 1)
                End If

                '該当シリーズとキー形番のHyphenなど情報の取得
                dr_Hyphen = dt_Hyphen.Select("series_kataban='" & dr_Series("series_kataban").ToString _
                                            & "' AND key_Kataban='" & dr_Series("key_Kataban").ToString & "'")
                If dr_Hyphen.Length <= 0 Then Continue For 'データなければ、次のシリーズへ
                ReDim strHyphen(dr_Hyphen.Length - 1)
                ReDim strElement_div(dr_Hyphen.Length - 1)
                ReDim strStructure_div(dr_Hyphen.Length - 1)
                ReDim strItem(dr_Hyphen.Length - 1)
                ReDim strItemNo(dr_Hyphen.Length - 1)
                ReDim strNewHyphen(dr_Hyphen.Length - 1)
                For intj As Integer = 0 To strItemNo.Length - 1
                    strItemNo(intj) = 0
                Next
                For intj As Integer = 0 To dr_Hyphen.Length - 1
                    strHyphen(intj) = dr_Hyphen(intj)("hyphen_div").ToString
                    strElement_div(intj) = dr_Hyphen(intj)("element_div").ToString
                    strStructure_div(intj) = dr_Hyphen(intj)("structure_div").ToString
                Next
                strSeries = dr_Series("series_kataban").ToString

                Divlist = New ArrayList
                Divlist.Add(strHyphen)
                Divlist.Add(strStructure_div)
                Divlist.Add(strElement_div)
                Divlist.Add(strNewHyphen)

                DS_Tab = New DataSet
                dt_Option = New DS_KatOut.DT_OptionDataTable
                dt_ElePattern = New DS_KatOut.kh_ele_patternDataTable
                dt_VolStd = New DS_KatOut.kh_std_voltage_mstDataTable
                dt_Stroke = New DS_KatOut.kh_strokeDataTable
                If ds_table Is Nothing Then
                    Using da As New DS_KatOutTableAdapters.DT_OptionTableAdapter
                        dt_Option = da.GetOptnameData(strSeries, strKeyKata, Now, "en", "ja")
                    End Using
                    Using da As New DS_KatOutTableAdapters.kh_ele_patternTableAdapter
                        dt_ElePattern = da.GetElePatternData(strSeries, strKeyKata, Now)
                    End Using
                    Using da As New DS_KatOutTableAdapters.kh_std_voltage_mstTableAdapter
                        dt_VolStd = da.GetDataBy(strSeries, strKeyKata, Now)
                    End Using
                    Using da As New DS_KatOutTableAdapters.kh_strokeTableAdapter
                        dt_Stroke = da.GetDataBy(strSeries, strKeyKata, Now)
                    End Using
                Else
                    Dim dr() As DataRow = Nothing
                    Dim strWhere As String = "series_kataban='" & strSeries & "' AND key_kataban='" & strKeyKata & "'"
                    dr = ds_table.Tables("dt_Option").Select(strWhere)
                    For intj As Integer = 0 To dr.Length - 1
                        dt_Option.ImportRow(dr(intj))
                    Next
                    dr = ds_table.Tables("dt_ElePattern").Select(strWhere)
                    For intj As Integer = 0 To dr.Length - 1
                        dt_ElePattern.ImportRow(dr(intj))
                    Next
                    dr = ds_table.Tables("dt_VolStd").Select(strWhere)
                    For intj As Integer = 0 To dr.Length - 1
                        dt_VolStd.ImportRow(dr(intj))
                    Next
                    dr = ds_table.Tables("dt_Stroke").Select(strWhere)
                    For intj As Integer = 0 To dr.Length - 1
                        dt_Stroke.ImportRow(dr(intj))
                    Next
                End If
                dt_Option.TableName = "Option"
                dt_ElePattern.TableName = "ElePattern"
                dt_VolStd.TableName = "dt_VolStd"
                dt_Stroke.TableName = "dt_Stroke"
                DS_Tab.Tables.Add(dt_Option)
                DS_Tab.Tables.Add(dt_ElePattern)
                DS_Tab.Tables.Add(dt_VolStd)
                DS_Tab.Tables.Add(dt_Stroke)

                '項目毎にﾁｪｯｸする
                If ItemCheck(strSeries, strKeyKata, 1, DS_Tab, Divlist, strItem, strItemNo, strNowKata, dtStrcEle) = True Then
                    dr_ItemName = dt_ItemName.Select("series_kataban='" & dr_Series("series_kataban").ToString _
                            & "' AND key_Kataban='" & dr_Series("key_Kataban").ToString & "'")
                    ReDim strItemName(dr_ItemName.Length - 1)
                    For intj As Integer = 0 To dr_ItemName.Length - 1
                        If dr_ItemName(intj)("ktbn_strc_nm").ToString = "電圧" Then
                            strItemName(intj) = "電　　圧"
                        Else
                            strItemName(intj) = dr_ItemName(intj)("ktbn_strc_nm").ToString
                        End If
                    Next
                    strFirseHyphen = dr_Series("hyphen_div").ToString
                    strSpecNo = dr_Series("spec_no").ToString
                    strPriceNo = dr_Series("price_no").ToString
                    strKataName = dr_Series("series_nm").ToString
                    Separator = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try

        strNowKata = Nothing
        strItemNo = Nothing
        dt_Series = Nothing
        dt_ItemName = Nothing
        dt_Hyphen = Nothing
        dr_Hyphen = Nothing
        dr_Series = Nothing
        strNewHyphen = Nothing
        Divlist = Nothing
        dt_Series = Nothing
        dt_ItemName = Nothing
        dt_Hyphen = Nothing
        dt_Option = Nothing
        dt_ElePattern = Nothing
        dt_VolStd = Nothing
        dt_Stroke = Nothing
    End Function

    ''' <summary>
    ''' 各項目のチェック
    ''' </summary>
    ''' <param name="strSeries"></param>
    ''' <param name="strKeyKata"></param>
    ''' <param name="intItemSeqNo"></param>
    ''' <param name="DS_Tab"></param>
    ''' <param name="Divlist"></param>
    ''' <param name="strItem"></param>
    ''' <param name="strItemNo"></param>
    ''' <param name="strNowKata"></param>
    ''' <param name="dtStrcEle">オプション候補</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function ItemCheck(ByVal strSeries As String, ByVal strKeyKata As String, _
                                    ByVal intItemSeqNo As Integer, ByVal DS_Tab As DataSet, _
                                    ByVal Divlist As ArrayList, ByVal strItem() As String, ByVal strItemNo() As Integer, _
                                    ByVal strNowKata As String, ByVal dtStrcEle As DataTable) As Boolean
        ItemCheck = False

        Dim strHyphen() As String = Divlist(0)
        Dim strStructure_div() As String = Divlist(1)
        Dim strElement_div() As String = Divlist(2)
        Dim strNewHyphen() As String = Divlist(3)
        Dim dt_View As New DS_KatOut.DT_OptionDataTable
        Dim dr_View() As DataRow
        Dim strItem1(strItem.Length) As String
        Dim dr_Option() As DataRow

        Try
            dr_View = DS_Tab.Tables("Option").Select("ktbn_strc_seq_no='" & (intItemSeqNo).ToString & "'")
            For inti As Integer = 0 To dr_View.Length - 1
                dt_View.ImportRow(dr_View(inti))
            Next

            strItem1(0) = Nothing
            For inti As Integer = 0 To strItem.Length - 1
                If strItem(inti) Is Nothing Then
                    strItem1(inti + 1) = ""
                Else
                    strItem1(inti + 1) = strItem(inti)
                End If
            Next

            Dim obj As New KHOptionCtl
            Dim strListOption(,) As String = Nothing
            Dim objKtbnStrc As New KHKtbnStrc
            objKtbnStrc.strcSelection.strSeriesKataban = strSeries
            objKtbnStrc.strcSelection.strKeyKataban = strKeyKata
            objKtbnStrc.strcSelection.strOpSymbol = strItem1
            Call obj.subOptionList(Nothing, objKtbnStrc, "1", "", "", "ja", intItemSeqNo, strListOption, DS_Tab.Tables("Option"), DS_Tab.Tables("ElePattern"))

            Dim bol_flg As Boolean = False
            For inti As Integer = dt_View.Rows.Count - 1 To 0 Step -1
                bol_flg = False
                For intj As Integer = 1 To UBound(strListOption)
                    If dt_View.Rows(inti)("option_symbol") = strListOption(intj, 1) Then
                        'ADD BY YGY 20150427
                        If strListOption(intj, 1).Equals(String.Empty) AndAlso _
                            strListOption(intj, 2).Equals(String.Empty) Then
                        Else
                            bol_flg = True
                            Exit For
                        End If
                    End If
                Next
                If Not bol_flg Then
                    dt_View.Rows(inti).Delete()
                End If
            Next

            dt_View.AcceptChanges()

            If dt_View.Rows.Count = 0 Then
                If strNowKata.Length <= 0 Then    '分解成功
                    
                    If CheckNecessaryOption(dtStrcEle, strKeyKata, intItemSeqNo, DS_Tab, objKtbnStrc) Then
                        '形番以降に必須オプションがあるかどうかの判断
                        ItemCheck = True
                    Else
                        ItemCheck = False
                    End If

                    Exit Try
                End If
                If intItemSeqNo - 1 < strHyphen.Length - 1 Then   '残り部分の分解
                    intItemSeqNo += 1 '次の項目へ
                    If ItemCheck(strSeries, strKeyKata, intItemSeqNo, DS_Tab, Divlist, _
                                 strItem, strItemNo, strNowKata, dtStrcEle) Then
                        ItemCheck = True
                        Exit Try
                    End If
                Else
                    '最後、戻るしかない
                    If intItemSeqNo - 1 > 0 Then
                        For inti As Integer = intItemSeqNo - 1 To 0 Step -1
                            If Not strItem(inti) Is Nothing AndAlso strItem(inti).ToString.Length > 0 Then
                                strNowKata = strItem(inti) & strNowKata
                                If ItemCheck(strSeries, strKeyKata, inti + 1, DS_Tab, Divlist, _
                                             strItem, strItemNo, strNowKata, dtStrcEle) Then
                                    ItemCheck = True
                                    Exit Try
                                End If
                            End If
                        Next
                    Else
                        Exit Try    '分解失敗
                    End If
                End If
            End If

            'ストローク場合(範囲なのに、分解できるように追加する)
            '例えは：ストローク範囲＝1～300、既存値は25、50、75、100四つだけ。250でも分解できるはず
            If strElement_div(intItemSeqNo - 1).ToString = "3" Then
                Dim strStrock As String = String.Empty
                For inti As Integer = 0 To strElement_div.Length - 1
                    If strElement_div(inti) = "5" AndAlso Not strItem(inti) Is Nothing AndAlso strItem(inti).Length > 0 Then    '口径
                        strStrock = String.Empty
                        For intj As Integer = 0 To strNowKata.Length - 1
                            If IsNumeric(Mid(strNowKata, intj + 1, 1)) Then
                                strStrock &= Mid(strNowKata, intj + 1, 1)
                            Else
                                Exit For
                            End If
                        Next
                        Dim ExitFlag As Boolean = False
                        For intj As Integer = 0 To dt_View.Rows.Count - 1
                            If strStrock = dt_View.Rows(intj)("option_symbol").ToString Then
                                ExitFlag = True
                                Exit For
                            End If
                        Next
                        If ExitFlag Then Exit For '既に存在
                        'なければ、追加する
                        Dim dr_Stroke() As DataRow = DS_Tab.Tables("dt_Stroke").Select("bore_size='" & strItem(inti) & "'")
                        If dr_Stroke.Length <= 0 Then Exit For
                        If dr_Stroke(0)("min_stroke").ToString.Length <= 0 Then Exit For
                        If dr_Stroke(0)("max_stroke").ToString.Length <= 0 Then Exit For
                        If dr_Stroke(0)("max_stroke") <= 0 Then Exit For
                        If strStrock.Length <= 0 Then Exit For

                        '範囲外
                        If CLng(strStrock) < CLng(dr_Stroke(0)("min_stroke")) Then
                            Exit For
                        End If
                        '範囲外
                        If CLng(strStrock) > CLng(dr_Stroke(0)("max_stroke")) Then
                            Exit For
                        End If
                        'Step外
                        If CLng(strStrock) Mod CLng(dr_Stroke(0)("stroke_unit")) <> 0 Then
                            Exit For
                        End If
                        Dim dr As DataRow = dt_View.NewRow
                        For intk As Integer = 0 To dt_View.Columns.Count - 1
                            dr(dt_View.Columns(intk).ColumnName) = dt_View.Rows(0)(dt_View.Columns(intk).ColumnName)
                        Next
                        dr("option_symbol") = strStrock
                        dr("symbol_length") = Len(strStrock)
                        dt_View.Rows.Add(dr)
                    End If
                Next
            End If

            '項目毎にﾁｪｯｸする
            Select Case strSeries
                Case "AB21", "AB41", "AB31", "AB42", "AG31", "AG33", "AG34", "AG41", "AG43", "AG44"
                    Select Case intItemSeqNo
                        Case 3
                            Dim dr As DataRow = dt_View.NewRow
                            dr("ktbn_strc_seq_no") = intItemSeqNo
                            dr("option_symbol") = "0"
                            dr("element_div") = ""
                            dr("structure_div") = "1"
                            dr("default_option_nm") = ""
                            dr("symbol_length") = 1
                            dt_View.Rows.Add(dr)
                        Case 4
                            Dim dr As DataRow = dt_View.NewRow
                            dr("ktbn_strc_seq_no") = intItemSeqNo
                            dr("option_symbol") = "00"
                            dr("element_div") = ""
                            dr("structure_div") = "1"
                            dr("default_option_nm") = ""
                            dr("symbol_length") = 1
                            dt_View.Rows.Add(dr)
                    End Select
                Case "GAG31", "GAG33", "GAG34", "GAG35", "GAG41", "GAG43", "GAG44", "GAG45", _
                     "GAB312", "GAB352", "GAB412", "GAB422", "GAB452", "GAB462", "NAB"
                    Select Case intItemSeqNo
                        Case 4
                            Dim dr As DataRow = dt_View.NewRow
                            dr("ktbn_strc_seq_no") = intItemSeqNo
                            dr("option_symbol") = "0"
                            dr("element_div") = ""
                            dr("structure_div") = "1"
                            dr("default_option_nm") = ""
                            dr("symbol_length") = 1
                            dt_View.Rows.Add(dr)
                        Case 5
                            If strSeries = "NAB" Then Exit Select
                            Dim dr As DataRow = dt_View.NewRow
                            dr("ktbn_strc_seq_no") = intItemSeqNo
                            dr("option_symbol") = "00"
                            dr("element_div") = ""
                            dr("structure_div") = "1"
                            dr("default_option_nm") = ""
                            dr("symbol_length") = 1
                            dt_View.Rows.Add(dr)
                    End Select
            End Select
            dr_Option = dt_View.Select("ktbn_strc_seq_no='" & intItemSeqNo & "'", _
                            "ktbn_strc_seq_no,symbol_length DESC")

            If strElement_div(intItemSeqNo - 1).ToString = "1" Then   '電圧
                Dim strVol As String = String.Empty
                Dim VolFlag As Boolean = False
                For inti As Integer = 0 To dr_Option.Length - 1
                    If dr_Option(inti)("default_option_nm").ToString.Length > 0 And _
                        strNowKata.StartsWith(dr_Option(inti)("default_option_nm").ToString) Then
                        VolFlag = False
                        Exit For
                    End If
                    If dr_Option(inti)("default_option_nm") = "Other voltage" Or _
                       dr_Option(inti)("option_symbol") = "Other voltage" Then
                        VolFlag = True
                    End If
                Next
                If VolFlag And (strNowKata.StartsWith("AC") Or strNowKata.StartsWith("DC")) Then
                    If strElement_div.Length = intItemSeqNo Then
                        '最後のオプションがその他電圧の場合
                        If strNowKata.EndsWith("V") Then
                            Dim dr As DataRow = dt_View.NewRow
                            dr("ktbn_strc_seq_no") = intItemSeqNo
                            dr("option_symbol") = strNowKata
                            dr("element_div") = ""
                            dr("structure_div") = "1"
                            dr("default_option_nm") = ""
                            dr("symbol_length") = Len(strNowKata)
                            dt_View.Rows.Add(dr)
                            dr_Option = dt_View.Select("ktbn_strc_seq_no='" & intItemSeqNo & "'", _
                                        "ktbn_strc_seq_no,symbol_length DESC")
                        Else
                        End If
                    Else
                        Dim str() As String = strNowKata.Split("V")
                        If str.Length > 0 Then
                            Dim dr As DataRow = dt_View.NewRow
                            dr("ktbn_strc_seq_no") = intItemSeqNo
                            dr("option_symbol") = str(0) & "V"
                            dr("element_div") = ""
                            dr("structure_div") = "1"
                            dr("default_option_nm") = ""
                            dr("symbol_length") = Len(dr("option_symbol").ToString)
                            dt_View.Rows.Add(dr)
                            dr_Option = dt_View.Select("ktbn_strc_seq_no='" & intItemSeqNo & "'", _
                                        "ktbn_strc_seq_no,symbol_length DESC")
                        End If
                    End If
                End If
            End If

            If dr_Option.Length <= 0 Then Exit Try

            '既存項目毎に比較する
            Dim bolEmpty As Boolean = False
            For intk As Integer = 0 To dr_Option.Length - 1
                If dr_Option(intk)("option_symbol").ToString = String.Empty Then
                    bolEmpty = True '空白可能
                    Exit For
                End If
            Next
            'Optionﾁｪｯｸ
            If OptionCheck(strItemNo(intItemSeqNo - 1), dr_Option, strItem, strItemNo, strNewHyphen, Divlist, intItemSeqNo - 1, _
                           strNowKata, bolEmpty) Then
                If strNowKata.Length <= 0 Then    '分解成功

                    If CheckNecessaryOption(dtStrcEle, strKeyKata, intItemSeqNo, DS_Tab, objKtbnStrc) Then
                        '形番以降に必須オプションがあるかどうかの判断
                        ItemCheck = True
                    Else
                        ItemCheck = False
                    End If

                    Exit Try
                End If
                If intItemSeqNo - 1 < strHyphen.Length - 1 Then   '残り部分の分解
                    intItemSeqNo += 1 '次の項目へ
                    If ItemCheck(strSeries, strKeyKata, intItemSeqNo, DS_Tab, Divlist, _
                                 strItem, strItemNo, strNowKata, dtStrcEle) Then
                        ItemCheck = True
                        Exit Try
                    End If
                Else
                    '最後、戻るしかない
                    If intItemSeqNo - 1 > 0 Then
                        For inti As Integer = intItemSeqNo - 1 To 0 Step -1
                            If Not strItem(inti) Is Nothing AndAlso strItem(inti).ToString.Length > 0 Then
                                strNowKata = strItem(inti) & IIf(strNewHyphen(inti) = "1", "-", "") & strNowKata
                                If ItemCheck(strSeries, strKeyKata, inti + 1, DS_Tab, Divlist, _
                                             strItem, strItemNo, strNowKata, dtStrcEle) Then
                                    ItemCheck = True
                                    Exit Try
                                End If
                            End If
                        Next
                    Else
                        Exit Try    '分解失敗
                    End If
                End If
            Else
                'このItemでの比較が失敗、前Itemへ戻る
                If intItemSeqNo > 0 Then
                    For inti As Integer = intItemSeqNo - 1 To 0 Step -1
                        If Not strNowKata.StartsWith("-") Then
                            strNowKata = IIf(strNewHyphen(inti) = "1", "-", "") & strNowKata
                        End If
                        If Not strItem(inti) Is Nothing AndAlso strItem(inti).ToString.Length > 0 Then
                            strNowKata = strItem(inti) & strNowKata
                            If (strNowKata.StartsWith("A-C") Or strNowKata.StartsWith("D-C")) And strNowKata.EndsWith("V") Then
                                strNowKata = strNowKata.Replace("-", "")
                            End If
                            If ItemCheck(strSeries, strKeyKata, inti + 1, DS_Tab, Divlist, _
                                         strItem, strItemNo, strNowKata, dtStrcEle) Then
                                ItemCheck = True
                                Exit Function
                            End If
                        End If
                    Next
                Else
                    Exit Function    '分解失敗
                End If
            End If
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try

        strHyphen = Nothing
        strStructure_div = Nothing
        strElement_div = Nothing
        strNewHyphen = Nothing
        dt_View = Nothing
        dr_View = Nothing
        strItem1 = Nothing
        dr_Option = Nothing
    End Function

    ''' <summary>
    ''' 形番以降に必須オプションがあるかどうかの判断
    ''' </summary>
    ''' <param name="dtStrcEle"></param>
    ''' <param name="strKeyKata"></param>
    ''' <param name="intItemSeqNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CheckNecessaryOption(ByVal dtStrcEle As DataTable, ByVal strKeyKata As String, _
                                                 ByVal intItemSeqNo As Integer, ByVal DS_Tab As DataSet, _
                                                 ByVal objKtbnStrc As KHKtbnStrc) As Boolean
        Dim options() As DataRow = dtStrcEle.Select("key_kataban='" & strKeyKata & "' AND ktbn_strc_seq_no > " & intItemSeqNo)

        For Each opt In options
            Dim seq As Integer = opt.Item("ktbn_strc_seq_no")
            Dim rows() As DataRow = options.CopyToDataTable.Select("ktbn_strc_seq_no = '" & seq & "'")
            If rows.Count = 1 Then
                If Not rows(0).Item("option_symbol").ToString.Trim.Equals(String.Empty) Then

                    Dim obj As New KHOptionCtl
                    Dim strListOption(,) As String = Nothing
                    '選択候補が表示できるかどうかの判断
                    Call obj.subOptionList(Nothing, objKtbnStrc, "1", "", "", "ja", seq, _
                                           strListOption, DS_Tab.Tables("Option"), DS_Tab.Tables("ElePattern"))

                    If strListOption.Length <= 3 Then
                    Else
                        Return False        '形番以降に必須オプションがある場合
                    End If

                End If
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' オプションチェック
    ''' </summary>
    ''' <param name="intOptionSeq"></param>
    ''' <param name="dr_Option"></param>
    ''' <param name="strItem"></param>
    ''' <param name="strItemNo"></param>
    ''' <param name="strNewHyphen"></param>
    ''' <param name="Divlist"></param>
    ''' <param name="intItemSeqNo"></param>
    ''' <param name="strNowKata"></param>
    ''' <param name="bolEmpty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function OptionCheck(ByVal intOptionSeq As Integer, ByVal dr_Option() As DataRow, _
                                     ByRef strItem() As String, ByRef strItemNo() As Integer, ByRef strNewHyphen() As String, _
                                     ByVal Divlist As ArrayList, ByVal intItemSeqNo As Integer, ByRef strNowKata As String, _
                                     ByVal bolEmpty As Boolean) As Boolean
        OptionCheck = False

        Dim strHyphen() As String = Divlist(0)
        Dim strStructure_div() As String = Divlist(1)
        Dim strKey As String = String.Empty

        Try
            '戻る場合、既に比較した物をクリアする
            If intOptionSeq <> 0 AndAlso intOptionSeq <= strItemNo(intItemSeqNo) Then
                For inti As Integer = intItemSeqNo To strItemNo.Length - 1
                    strItemNo(inti) = 0
                    strItem(inti) = String.Empty
                Next
                If intOptionSeq = dr_Option.Length - 1 Then    '最後
                    If bolEmpty Then                           '空白可能
                        OptionCheck = True
                        Exit Try
                    End If
                End If
            End If

            strKey = String.Empty
            If intOptionSeq >= dr_Option.Length Then Exit Try
            strKey = dr_Option(intOptionSeq)("option_symbol").ToString
            If strNowKata.StartsWith("-") Then
                Dim myFlg As Boolean = False
                For inti As Integer = 0 To dr_Option.Length - 1
                    If dr_Option(inti)("option_symbol").ToString.Length > 0 AndAlso _
                       dr_Option(inti)("option_symbol").ToString.StartsWith("-") AndAlso _
                       strNowKata.StartsWith(dr_Option(inti)("option_symbol").ToString) Then
                        myFlg = True
                        Exit For
                    End If
                Next
                If Not myFlg Then
                    If strHyphen(intItemSeqNo).ToString = "1" Then
                        strNowKata = Mid(strNowKata, 2)
                        strNewHyphen(intItemSeqNo) = "1"
                    ElseIf intItemSeqNo > 0 AndAlso (strItem(intItemSeqNo - 1) Is Nothing And strHyphen(intItemSeqNo - 1).ToString = "1") Then
                        strNowKata = Mid(strNowKata, 2)
                        strNewHyphen(intItemSeqNo - 1) = "1"
                    End If
                End If
            End If
            If strKey.Length = 0 Then    '空白ならば、次へ行く
                If intOptionSeq < dr_Option.Length - 1 Then
                    intOptionSeq += 1
                End If
            End If

            Dim ExitFlag As Boolean = False
            If CInt(strStructure_div(intItemSeqNo).ToString) < 4 Then   '複数じゃない場合
                For intj As Integer = intOptionSeq To dr_Option.Length - 1
                    strKey = dr_Option(intj)("option_symbol").ToString
                    If strNowKata.Split("-").Length > 1 Then
                        If strKey.Length > 0 And strHyphen(intItemSeqNo).ToString = "1" Then strKey &= "-"
                    End If

                    If strNowKata.StartsWith(strKey) Then '同じならば
                        ExitFlag = True
                        strItem(intItemSeqNo) = dr_Option(intj)("option_symbol").ToString
                        strItemNo(intItemSeqNo) = intj + 1
                        strNowKata = Mid(strNowKata, Len(strKey) + 1)
                        strNewHyphen(intItemSeqNo) = IIf(strKey.Contains("-") = False, "0", "1")
                        OptionCheck = True
                        If strNowKata.Length = 0 Then
                            OptionCheck = True
                            Exit Try
                        End If
                        Exit For
                    End If
                Next
            Else                        '複数可能、同じOption中で10回検索する
                For inti As Integer = 1 To 10
                    For intj As Integer = intOptionSeq To dr_Option.Length - 1
                        If dr_Option(intj)("option_symbol").ToString.Length <= 0 Then Continue For
                        If strNowKata.StartsWith(dr_Option(intj)("option_symbol").ToString) Then '同じならば
                            ExitFlag = True
                            If Not strItem(intItemSeqNo) Is Nothing AndAlso strItem(intItemSeqNo).Length > 0 Then
                                strItem(intItemSeqNo) &= "," & dr_Option(intj)("option_symbol").ToString
                            Else
                                strItem(intItemSeqNo) &= dr_Option(intj)("option_symbol").ToString
                            End If
                            strItemNo(intItemSeqNo) = intj + 1
                            strNowKata = Mid(strNowKata, Len(dr_Option(intj)("option_symbol").ToString) + 1)
                            OptionCheck = True
                            If strNowKata.Length = 0 Then
                                OptionCheck = True
                                Exit Try
                            End If
                            Exit For
                        End If
                    Next
                Next
            End If
            If ExitFlag Then
                If strHyphen(intItemSeqNo).ToString = "1" And strNowKata.StartsWith("-") Then
                    strNowKata = Mid(strNowKata, 2)
                    strNewHyphen(intItemSeqNo) = "1"
                End If
                OptionCheck = True
            Else
                If bolEmpty Then                        '空白可能、次のItemデータへ
                    OptionCheck = True
                Else
                    Exit Try
                End If
            End If
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
        strHyphen = Nothing
        strStructure_div = Nothing
        strKey = Nothing
    End Function

    ''' <summary>
    ''' ミックスチェック
    ''' </summary>
    ''' <param name="strSeriesKata"></param>
    ''' <param name="strValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncMixCheck(strSeriesKata As String, ByVal strValue() As String) As Boolean
        Try
            fncMixCheck = False

            '各機種毎にミックス構成が選択されているかチェックする
            Select Case strSeriesKata
                Case "VSKM"
                    If strValue(1).Trim = "Z" Or strValue(2).Trim = "00" Or _
                       strValue(3).Trim = "Z" Or strValue(4).Trim = "CX" Or _
                       strValue(8).Trim = "Z" Or strValue(10).Trim = "Z" Then
                        fncMixCheck = True
                    End If
                Case "VSJM"
                    If strValue(1).Trim = "Z" Or strValue(2).Trim = "00" Or _
                       strValue(3).Trim = "Z" Or strValue(4).Trim = "CX" Or _
                       strValue(10).Trim = "Z" Then
                        fncMixCheck = True
                    End If
                Case "VSNM"
                    If strValue(1).Trim = "Z" Or strValue(2).Trim = "00" Or _
                         strValue(3).Trim = "CX" Or strValue(9).Trim = "Z" Then
                        fncMixCheck = True
                    End If
                Case "VSNPM"
                    If strValue(1).Trim = "CX" Or strValue(6).Trim = "Z" Then
                        fncMixCheck = True
                    End If
                Case "VSXM"
                    If strValue(1).Trim = "Z" Or strValue(2).Trim = "00" Or _
                       strValue(3).Trim = "Z" Or strValue(4).Trim = "CX" Or _
                       strValue(9).Trim = "Z" Then
                        fncMixCheck = True
                    End If
                Case "VSZM"
                    If strValue(1).Trim = "Z" Or strValue(2).Trim = "00" Or _
                       strValue(3).Trim = "Z" Or strValue(4).Trim = "CX" Or _
                       strValue(9).Trim = "Z" Then
                        fncMixCheck = True
                    End If
                Case "VSJPM"
                    If strValue(1).Trim = "Z" Or strValue(2).Trim = "CX" Or _
                       strValue(9).Trim = "Z" Then
                        fncMixCheck = True
                    End If
                Case "VSXPM"
                    If strValue(1).Trim = "Z" Or strValue(2).Trim = "CX" Or _
                       strValue(7).Trim = "Z" Then
                        fncMixCheck = True
                    End If
                Case "VSZPM"
                    If strValue(1).Trim = "CX" Or strValue(6).Trim = "Z" Then
                        fncMixCheck = True
                    End If
            End Select

        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
