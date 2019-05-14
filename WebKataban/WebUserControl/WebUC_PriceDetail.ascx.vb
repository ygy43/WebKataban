Imports WebKataban.KHCodeConstants.CdCst

Public Class WebUC_PriceDetail
    Inherits KHBase

    '単価画面戻るイベント
    Public Event BackToTanka(ByVal intMode As Integer)

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub
        If Me.objUserInfo.UserId Is Nothing Then Exit Sub
        If HidMode.Value.Length <= 0 Then Exit Sub

        Try
            '価格詳細
            Dim dtPriceDetail As New DataTable
            '国コードと出荷場所
            Dim lstShipPlace As New List(Of String)
            '国コードリスト
            Dim lstCountryCd As New List(Of String)

            'タイトルの設定
            subInitPage()

            '価格詳細の作成
            Select Case HidMode.Value
                Case 1
                    '一般の場合
                    '出力可能な出荷場所を取得
                    lstShipPlace = fncGetShipPlace()

                    '出力国コードの取得
                    lstCountryCd = fncGetCountryCd()

                    '価格計算
                    dtPriceDetail = fncNormalPriceDetail(lstCountryCd, lstShipPlace)

                Case 2
                    'ISOの場合
                    '出力可能な出荷場所を取得
                    'lstShipPlace = New List(Of String) From {"JPN"}
                    lstShipPlace = fncGetShipPlace()

                    '出力国コードの取得
                    lstCountryCd = fncGetCountryCd()

                    '価格計算
                    dtPriceDetail = fncISOPriceDetail(lstCountryCd, lstShipPlace)
            End Select

            '結果表示
            grdPriceDetail.DataSource = dtPriceDetail
            grdPriceDetail.DataBind()

            'OKボタンの設定
            Button1.OnClientClick = "Clip_Copy('" & fncCreateClipContent() & "');"

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 結果リストにタイトルの追加
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub grdPriceDetail_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdPriceDetail.RowCreated
        If e.Row.RowType.Equals(DataControlRowType.Header) Then

            Dim HeaderGrid As GridView = CType(sender, GridView)
            Dim HeaderGridRow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim HeaderCell As TableHeaderCell = New TableHeaderCell

            Select Case HidMode.Value
                Case 1
                    '一般
                    '空欄
                    HeaderCell.Text = ""
                    HeaderCell.ColumnSpan = 1
                    HeaderGridRow.Cells.Add(HeaderCell)

                    '現地定価
                    HeaderCell = New TableHeaderCell
                    HeaderCell.Text = fncGetLabelContent("2")
                    HeaderCell.ColumnSpan = 1
                    HeaderGridRow.Cells.Add(HeaderCell)

                    '購入価格
                    HeaderCell = New TableHeaderCell
                    HeaderCell.Text = fncGetLabelContent("3")
                    If HeaderGrid.DataSource IsNot Nothing Then
                        HeaderCell.ColumnSpan = HeaderGrid.DataSource.Columns.Count - 2
                    Else
                        HeaderCell.ColumnSpan = 1
                    End If
                Case 2
                    'ISO
                    '空欄
                    HeaderCell.Text = ""
                    HeaderCell.ColumnSpan = 2
                    HeaderGridRow.Cells.Add(HeaderCell)

                    '現地定価
                    HeaderCell = New TableHeaderCell
                    HeaderCell.Text = fncGetLabelContent("2")
                    HeaderCell.ColumnSpan = 1
                    HeaderGridRow.Cells.Add(HeaderCell)

                    '購入価格
                    HeaderCell = New TableHeaderCell
                    HeaderCell.Text = fncGetLabelContent("3")
                    If HeaderGrid.DataSource IsNot Nothing Then
                        HeaderCell.ColumnSpan = HeaderGrid.DataSource.Columns.Count - 3
                    Else
                        HeaderCell.ColumnSpan = 1
                    End If
            End Select
            

            HeaderGridRow.Cells.Add(HeaderCell)

            grdPriceDetail.Controls(0).Controls.AddAt(0, HeaderGridRow)
        End If
    End Sub

    ''' <summary>
    ''' OKボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnOK_Click(sender As Object, e As EventArgs) Handles Button1.Click

        RaiseEvent BackToTanka(HidMode.Value)
    End Sub

    ''' <summary>
    ''' キャンセルボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancel_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RaiseEvent BackToTanka(HidMode.Value)
    End Sub

    ''' <summary>
    ''' バインドイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub grdPriceDetail_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdPriceDetail.RowDataBound
        '表示設定
        If e.Row.RowType.Equals(DataControlRowType.DataRow) Then

            Select Case HidMode.Value
                Case 1
                    '一般の場合

                    For intColumn As Integer = 0 To e.Row.Cells.Count - 1
                        '列幅の設定
                        e.Row.Cells(intColumn).Width = New Unit("150px")

                        '寄せの設定
                        If intColumn >= 1 Then
                            '価格の場合は右寄せ
                            e.Row.Cells(intColumn).HorizontalAlign = HorizontalAlign.Right
                        Else
                            e.Row.Cells(intColumn).HorizontalAlign = HorizontalAlign.Left
                        End If
                    Next

                Case 2
                    'ISOの場合
                    For intColumn As Integer = 0 To e.Row.Cells.Count - 1
                        '列幅の設定
                        If intColumn = 1 Then
                            e.Row.Cells(intColumn).Width = New Unit("250px")
                        Else
                            e.Row.Cells(intColumn).Width = New Unit("150px")
                        End If

                        '寄せの設定
                        If intColumn >= 2 Then
                            '価格の場合は右寄せ
                            e.Row.Cells(intColumn).HorizontalAlign = HorizontalAlign.Right
                        Else
                            e.Row.Cells(intColumn).HorizontalAlign = HorizontalAlign.Left
                        End If
                    Next
            End Select
        End If
    End Sub

#Region "メソッド"

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Public Sub frmInit(ByVal intMode As Integer)
        HidMode.Value = intMode
    End Sub

    ''' <summary>
    ''' タイトルの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subInitPage()
        '形番
        lblSeriesKat.Text = objKtbnStrc.strcSelection.strFullKataban

        '形番名称
        lblSeriesNm.Text = objKtbnStrc.strcSelection.strGoodsNm

        'Label取得
        Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHPriceDetail, selLang.SelectedValue, Me)
    End Sub

    ''' <summary>
    ''' 現地定価とFOB価格の取得
    ''' </summary>
    ''' <param name="strCountry">国コード</param>
    ''' <param name="strShipPlace">出荷場所</param>
    ''' <param name="strCurrency">通貨</param>
    ''' <param name="strLocalPrice">現地定価</param>
    ''' <param name="strFobPrice">FOB価格</param>
    ''' <remarks></remarks>
    Private Sub subGetFobAndLocalPrice(ByVal strCountry As String, _
                                       ByVal strShipPlace As String, _
                                       ByVal strCurrency As String, _
                                       ByRef strLocalPrice As String, _
                                       ByRef strFobPrice As String)
        '単価情報
        Dim objUnitPrice As New KHUnitPrice
        Dim strPriceList(,) As String = Nothing
        Dim strPriceFCA As String = Nothing
        Dim strPriceFCA2 As String = Nothing
        'マスタ未登録メッセージ
        Dim strMsg As String = ClsCommon.fncGetMsg(selLang.SelectedValue, "I5220")

        '日本の場合国コードへ変換
        If ShipPlaceJapan.Contains(strShipPlace) Then
            strShipPlace = "JPN"
        End If

        '価格情報の計算(フル権限)
        Call objUnitPrice.subPriceListSelect(objConBase, _
                                             strCountry, _
                                             Me.selLang.SelectedValue, _
                                             strCurrency, _
                                             255, _
                                             strPriceList, _
                                             strPriceFCA, _
                                             strPriceFCA2, _
                                             strShipPlace, _
                                             objKtbnStrc)
        If strPriceList.Length > 5 Then
            'FOB価格と現地定価の取得
            For i As Integer = 0 To (strPriceList.Length / 5) - 1
                If strPriceList(i, 4) = KHCodeConstants.CdCst.UnitPrice.FobPrice Then
                    'FOB価格
                    If CType(strPriceList(i, 2), Decimal) = 0 Then
                        strFobPrice = String.Empty
                    Else
                        strFobPrice = strPriceList(i, 2) & Space(1) & strPriceList(i, 3)
                    End If
                ElseIf strPriceList(i, 4) = KHCodeConstants.CdCst.UnitPrice.APrice Then
                    '現地定価
                    If CType(strPriceList(i, 2), Decimal) = 0 Then
                        strLocalPrice = String.Empty
                    Else
                        strLocalPrice = strPriceList(i, 2) & Space(1) & strPriceList(i, 3)
                    End If
                End If
            Next
        End If

    End Sub

    ''' <summary>
    ''' 現地定価とFOB価格の取得(ISO)
    ''' </summary>
    ''' <param name="strCountry">国コード</param>
    ''' <param name="strShipPlace">出荷場所</param>
    ''' <param name="strCurrency">通貨</param>
    ''' <param name="strLocalPrice">現地定価</param>
    ''' <param name="strFobPrice">FOB価格</param>
    ''' <remarks></remarks>
    Private Sub subGetFobAndLocalPriceISO(ByVal drCompData As DataRow, _
                                          ByVal strCountry As String, _
                                          ByVal strShipPlace As String, _
                                          ByVal strCurrency As String, _
                                          ByRef strLocalPrice As String, _
                                          ByRef strFobPrice As String)
        Dim dtCompData As New DataTable
        Dim strPriceList(,) As String = Nothing
        Dim objUnitPrice As New KHUnitPrice
        Dim intSglPrice(5) As Decimal

        '日本の場合国コードへ変換
        If ShipPlaceJapan.Contains(strShipPlace) Then
            strShipPlace = "JPN"
        End If

        With drCompData
            intSglPrice(0) = .Item("ls_price")
            intSglPrice(1) = .Item("rg_price")
            intSglPrice(2) = .Item("ss_price")
            intSglPrice(3) = .Item("bs_price")
            intSglPrice(4) = .Item("gs_price")
            intSglPrice(5) = .Item("ps_price")

            objUnitPrice.subISOPriceListSelect(objCon, _
                                               objConBase, _
                                               objKtbnStrc, _
                                               Me.objUserInfo.UserId, _
                                               Me.objLoginInfo.SessionId, _
                                               strCountry, _
                                               selLang.SelectedValue, _
                                               strCurrency, 255, _
                                               .Item("option_kataban"), _
                                               .Item("kataban_check_div"), _
                                               intSglPrice, strShipPlace, _
                                               strPriceList)
        End With

        'FOB価格と現地定価の取得
        For i As Integer = 0 To (strPriceList.Length / 5) - 1
            If strPriceList(i, 4) = KHCodeConstants.CdCst.UnitPrice.FobPrice Then
                'FOB価格
                If CType(strPriceList(i, 2), Decimal) = 0 Then
                    strFobPrice = String.Empty
                Else
                    strFobPrice = strPriceList(i, 2) & Space(1) & strPriceList(i, 3)
                End If
            ElseIf strPriceList(i, 4) = KHCodeConstants.CdCst.UnitPrice.APrice Then
                '現地定価
                If CType(strPriceList(i, 2), Decimal) = 0 Then
                    strLocalPrice = String.Empty
                Else
                    strLocalPrice = strPriceList(i, 2) & Space(1) & strPriceList(i, 3)
                End If
            End If
        Next

    End Sub

    ''' <summary>
    ''' 一般の価格詳細
    ''' </summary>
    ''' <param name="lstCountryCd"></param>
    ''' <param name="lstShipPlace"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncNormalPriceDetail(ByVal lstCountryCd As List(Of String), ByVal lstShipPlace As List(Of String)) As DataTable
        Dim dtPriceDetail As New DataTable
        '通貨
        Dim dtCountryMst As New DataTable
        '国名称
        Dim dt_Country As DataTable = KHCountry.fncGetAllCountryName(objConBase)

        '結果テーブルの作成
        dtPriceDetail = fncCreatePriceTable(lstShipPlace)

        'カントリマスタの取得
        dtCountryMst = MasterBLL.fncGetAllCountryMst(objConBase)

        '価格の計算
        For Each strCountryCd As String In lstCountryCd
            Dim strAPrice As String = String.Empty                    '現地定価
            Dim drPriceDetail As DataRow

            drPriceDetail = dtPriceDetail.NewRow

            For Each strShipPlace As String In lstShipPlace

                Dim strLocalPrice As String = String.Empty
                Dim strFobPrice As String = String.Empty
                Dim strCurrency As String = String.Empty

                '通貨の取得
                strCurrency = dtCountryMst.Select("country_cd = '" & strCountryCd & "'")(0).Item("currency_cd")

                '価格の取得
                subGetFobAndLocalPrice(strCountryCd, strShipPlace, strCurrency, strLocalPrice, strFobPrice)

                '現地定価の設定
                If strAPrice.Equals(String.Empty) Then
                    strAPrice = strLocalPrice
                Else
                    If Not strAPrice.Equals(strLocalPrice) Then
                        'エラーメッセージ
                    End If
                End If

                '購入価格の設定
                drPriceDetail.Item(fncGetColumnName(strShipPlace)) = strFobPrice
            Next

            '国名称
            drPriceDetail.Item(fncGetLabelContent("1")) = fncGetCounntryName(strCountryCd, dt_Country)

            '現地定価
            drPriceDetail.Item("-") = strAPrice

            dtPriceDetail.Rows.Add(drPriceDetail)
        Next

        Return dtPriceDetail
    End Function

    ''' <summary>
    ''' ISOの価格詳細
    ''' </summary>
    ''' <param name="lstCountryCd"></param>
    ''' <param name="lstShipPlace"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncISOPriceDetail(ByVal lstCountryCd As List(Of String), ByVal lstShipPlace As List(Of String)) As DataTable
        Dim dtPriceDetail As New DataTable
        '通貨
        Dim dtCountryMst As New DataTable
        '国名称
        Dim dt_Country As DataTable = KHCountry.fncGetAllCountryName(objConBase)

        Dim dtCompData As New DataTable
        'データ取得
        dtCompData = TankaISOBLL.fncSQL_GetCompData(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

        '結果テーブルの作成
        dtPriceDetail = fncCreatePriceTableISO(lstShipPlace)

        'カントリマスタの取得
        dtCountryMst = MasterBLL.fncGetAllCountryMst(objConBase)

        'オプション名称データ取得()
        Dim strOpNm As New List(Of String)
        Dim dtLabel As DataTable = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHISOTanka, selLang.SelectedValue)

        Dim lblContent = (From content In dtLabel
                          Where content.Field(Of String)("label_div") = "L"
                          Select content.Field(Of String)("label_content"))
        strOpNm = lblContent.ToList

        '価格計算
        If dtCompData IsNot Nothing Then

            For Each strCountryCd As String In lstCountryCd

                Dim strAPrice As String = String.Empty                    '現地定価
                Dim strLocalPrice As String = String.Empty                '現地定価
                Dim strFobPrice As String = String.Empty                  '購入価格
                Dim strCurrency As String = String.Empty                  '通貨

                '通貨の取得
                strCurrency = dtCountryMst.Select("country_cd = '" & strCountryCd & "'")(0).Item("currency_cd")


                '価格の取得
                For Each drCompData As DataRow In dtCompData.Rows
                    Dim drPriceDetail As DataRow
                    drPriceDetail = dtPriceDetail.NewRow

                    If fncISOShowOrNot(strOpNm, drCompData.Item("spec_strc_seq_no")) Then
                        '価格出荷場所により計算
                        For Each strShipPlace As String In lstShipPlace

                            '価格計算
                            subGetFobAndLocalPriceISO(drCompData, strCountryCd, strShipPlace, strCurrency, strLocalPrice, strFobPrice)

                            '国コード
                            drPriceDetail.Item(fncGetLabelContent("1")) = fncGetCounntryName(strCountryCd, dt_Country)

                            '形番
                            drPriceDetail.Item(fncGetLabelContent("4")) = drCompData.Item("option_kataban")

                            '現地定価の設定
                            drPriceDetail.Item("-") = strLocalPrice

                            '購入価格の設定
                            drPriceDetail.Item(fncGetColumnName(strShipPlace)) = strFobPrice

                        Next

                        dtPriceDetail.Rows.Add(drPriceDetail)
                    End If
                Next
            Next


        End If

        Return dtPriceDetail
    End Function

    ''' <summary>
    ''' 出荷場所コードの取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetShipPlace() As List(Of String)
        Dim lstShipPlace As New List(Of String)
        Dim lstShipPlaceTmp As New List(Of String)

        lstShipPlaceTmp = Session("ShipPlaces")

        For Each strShipPlace As String In lstShipPlaceTmp
            Dim strPlace As String = strShipPlace

            Select Case strShipPlace
                '日本の場合国コードへ変換
                Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
                    strPlace = "JPN"
            End Select

            If Not lstShipPlace.Contains(strPlace) Then
                lstShipPlace.Add(strPlace)
            End If
        Next

        Return lstShipPlace

    End Function

    ''' <summary>
    ''' 国コードの取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetCountryCd() As List(Of String)

        Dim lstResult As New List(Of String)
        Dim dtDisplayableCountry As New DS_PriceDetail.kh_country_displayable_price_mstDataTable

        'ユーザー自分の国コードを追加
        lstResult.Add(objUserInfo.CountryCd)

        '表示可能国マスタに登録した国コードを追加
        Using da As New DS_PriceDetailTableAdapters.kh_country_displayable_price_mstTableAdapter
            da.FillByCountryCd(dtDisplayableCountry, objUserInfo.CountryCd)
        End Using

        For Each dr As DataRow In dtDisplayableCountry

            Dim strCountryCd As String = dr.Item("disp_country_cd").ToString

            lstResult.Add(strCountryCd)

        Next
        
        Return lstResult
    End Function

    ''' <summary>
    ''' 国名の取得
    ''' </summary>
    ''' <param name="strCountryCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetCounntryName(ByVal strCountryCd As String, ByVal dt_Country As DataTable) As String

        Dim strResult As String = String.Empty
        Dim drcountry() As DataRow = dt_Country.Select("country_cd='" & strCountryCd & "' AND language_cd='" & Me.selLang.SelectedValue & "'")

        If drcountry.Count <= 0 Then
            drcountry = dt_Country.Select("country_cd='" & strCountryCd & "' AND language_cd='" & CdCst.LanguageCd.DefaultLang & "'")
            strResult = drcountry(0).Item("country_nm")
        Else
            strResult = drcountry(0).Item("country_nm")
        End If

        Return strResult
    End Function

    ''' <summary>
    ''' 出力テーブルの作成
    ''' </summary>
    ''' <param name="lstShipPlace"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreatePriceTable(ByVal lstShipPlace As List(Of String)) As DataTable

        Dim dtResult As New DataTable

        '国コード
        dtResult.Columns.Add(fncGetLabelContent("1"))
        dtResult.Columns.Add("-")

        For Each strShipPlace As String In lstShipPlace
            dtResult.Columns.Add(fncGetColumnName(strShipPlace))
        Next

        Return dtResult

    End Function

    ''' <summary>
    ''' ISO価格詳細
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreatePriceTableISO(ByVal lstShipPlace As List(Of String)) As DataTable

        Dim dtResult As New DataTable

        '国コード
        dtResult.Columns.Add(fncGetLabelContent("1"))

        '形番
        dtResult.Columns.Add(fncGetLabelContent("4"))

        '現地定価
        dtResult.Columns.Add("-")

        '日本出荷品
        'dtResult.Columns.Add(fncGetColumnName("JPN"))
        For Each strShipPlace As String In lstShipPlace
            dtResult.Columns.Add(fncGetColumnName(strShipPlace))
        Next

        Return dtResult

    End Function

    ''' <summary>
    ''' 形番表示の判断
    ''' </summary>
    ''' <param name="strOpNm"></param>
    ''' <param name="intSpecStrcSeqNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncISOShowOrNot(ByVal strOpNm As List(Of String), ByVal intSpecStrcSeqNo As Integer) As Boolean
        Dim strOptionNm As String = String.Empty

        Select Case objKtbnStrc.strcSelection.strSeriesKataban
            Case "CMF", "GMF"
                Select Case intSpecStrcSeqNo
                    Case 1
                        strOptionNm = strOpNm(0)  'ベース
                    Case 2, 3, 4, 5, 6, 7
                        strOptionNm = strOpNm(1)  '電磁弁形式
                    Case 13, 14
                        strOptionNm = strOpNm(2)  '給気スペーサ
                    Case 15, 16
                        strOptionNm = strOpNm(3)  '排気スペーサ
                    Case 17, 18
                        strOptionNm = strOpNm(4)  'パイロットチェック弁
                    Case 19, 20, 21, 22
                        strOptionNm = strOpNm(5)  'スペーサ形減圧弁
                    Case 23, 24
                        strOptionNm = strOpNm(6)  '流露遮蔽板
                    Case Else
                        strOptionNm = String.Empty
                End Select
            Case "LMF0"
                Select Case intSpecStrcSeqNo
                    Case 1
                        strOptionNm = strOpNm(0)
                    Case 2, 3, 4, 5, 6, 7
                        strOptionNm = strOpNm(1)
                    Case 13, 14
                        strOptionNm = strOpNm(2)
                    Case 15, 16
                        strOptionNm = strOpNm(3)
                    Case 17
                        strOptionNm = strOpNm(4)
                    Case 18, 19
                        strOptionNm = strOpNm(6)
                    Case Else
                        strOptionNm = String.Empty
                End Select
        End Select

        If Not strOptionNm.Equals(String.Empty) Then
            Return True
        Else
            Return False
        End If

    End Function

    ''' <summary>
    ''' 出荷場所コードにより出荷場所名称を作成
    ''' </summary>
    ''' <param name="strShipPlace"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetColumnName(ByVal strShipPlace As String) As String

        Dim strResult As String = String.Empty
        Dim strCountryName As String = String.Empty
        Dim dt_Country As DataTable = KHCountry.fncGetAllCountryName(objConBase)

        If ShipPlaceJapan.Contains(strShipPlace) Then
            '日本出荷の場合
            strResult = fncGetLabelContent("10")
            strShipPlace = "JPN"
        ElseIf strShipPlace.Equals("THF") Then
            strResult = fncGetLabelContent("10")
        Else
            strResult = fncGetLabelContent("7")
        End If

        '国名の入れ替え
        strCountryName = dt_Country.Select("country_cd='" & strShipPlace & "' AND language_cd='" & Me.selLang.SelectedValue & "'")(0).Item("country_nm").ToString
        strResult = strResult.Replace("[1]", strCountryName)

        '特殊的な出荷場所の設定
        If strShipPlace.Equals("KTA") OrElse _
            strShipPlace.Equals("TYO") OrElse _
            strShipPlace.Equals("MDN") OrElse _
            strShipPlace.Equals("OMA") OrElse _
            strShipPlace.Equals("CJA") Then
            'Made in -> Made by 　　　 「China」に含まれる「in」が変換されないように「 in 」スペースを追加
            If Me.selLang.SelectedValue.Equals("en") Then
                If strResult.Contains(" in ") Then
                    strResult = strResult.Replace(" in ", " by ")
                End If
            End If
        End If

        Return strResult

    End Function

    ''' <summary>
    ''' 価格詳細リストのタイトルを取得
    ''' </summary>
    ''' <param name="strLabelKBN">
    ''' 1:国コード
    ''' 2:現地定価
    ''' 3:購入価格
    ''' 4:形番
    ''' 7:[1]生産品
    ''' 9:[1]生産品の可能性があります。発注先を確認してください。
    ''' 10:[1]出荷品
    ''' 11:[1]生産品の可能性あり
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetLabelContent(ByVal strLabelKBN As String) As String
        Dim strResult As String = String.Empty

        If selLang IsNot Nothing Then
            Select Case strLabelKBN
                Case "1"
                    '国コード
                    strResult = LabelBLL.fncSelectLabelById(objCon, "KHRateMstMnt", selLang.SelectedValue, "L", 6)
                Case "2"
                    '現地定価
                    strResult = LabelBLL.fncSelectLabelById(objCon, "KHRateMstMnt", selLang.SelectedValue, "R", 1)
                Case "3"
                    '購入価格
                    strResult = LabelBLL.fncSelectLabelById(objCon, "KHRateMstMnt", selLang.SelectedValue, "R", 2)
                Case "4"
                    '形番
                    strResult = LabelBLL.fncSelectLabelById(objCon, "KHSiyou", selLang.SelectedValue, "L", 2)
                Case "7"
                    '[1]生産品
                    strResult = LabelBLL.fncSelectLabelById(objCon, "KHTanka", selLang.SelectedValue, "L", 7)
                Case "9"
                    '[1]生産品の可能性があります。発注先を確認してください。
                    strResult = LabelBLL.fncSelectLabelById(objCon, "KHTanka", selLang.SelectedValue, "L", 9)
                Case "10"
                    '[1]出荷品
                    strResult = LabelBLL.fncSelectLabelById(objCon, "KHTanka", selLang.SelectedValue, "L", 10)
                Case "11"
                    '[1]生産品の可能性あり
                    strResult = LabelBLL.fncSelectLabelById(objCon, "KHTanka", selLang.SelectedValue, "L", 11)
            End Select
        End If

        Return strResult
    End Function

    ''' <summary>
    ''' コピー内容の作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreateClipContent() As String
        Dim strResult As String = String.Empty
        Dim intStartColumn As Integer = 0                               '価格が始まる列番号

        Select Case HidMode.Value
            Case 1
                '一般
                '現地定価
                strResult &= "\t" & fncGetLabelContent("2") & "\t\t"
                intStartColumn = 2
            Case 2
                'ISO
                '現地定価
                strResult &= "\t\t" & fncGetLabelContent("2") & "\t\t"
                intStartColumn = 3
        End Select
        
        '購入価格
        strResult &= fncGetLabelContent("3") & "\r\n"

        For intHeader As Integer = 0 To grdPriceDetail.HeaderRow.Cells.Count - 1
            Dim cell As TableCell = grdPriceDetail.HeaderRow.Cells(intHeader)

            If intHeader = grdPriceDetail.HeaderRow.Cells.Count - 1 Then
                '最後の一列
                strResult &= cell.Text.Trim & "\r\n"
            ElseIf intHeader >= intStartColumn - 1 Then
                strResult &= cell.Text.Trim & "\t\t"
            Else
                strResult &= cell.Text.Trim & "\t"
            End If

        Next

        'ClipBoardコピー内容の作成
        For Each row As GridViewRow In grdPriceDetail.Rows
            For intCell As Integer = 0 To row.Cells.Count - 1
                Dim cell As TableCell = row.Cells(intCell)

                If intCell = row.Cells.Count - 1 Then
                    '最後の一列
                    If cell.Text.Equals("&nbsp;") Then
                        strResult &= "\r\n"
                    Else
                        '価格と単位を分ける
                        Dim txtPriceUnit() As String = cell.Text.Trim.Split(Space(1))

                        strResult &= txtPriceUnit(0) & "\t" & txtPriceUnit(1) & "\r\n"
                    End If
                ElseIf intCell >= intStartColumn - 1 Then
                    If cell.Text.Equals("&nbsp;") Then
                        strResult &= "\t"
                    Else
                        '価格と単位を分ける
                        Dim txtPriceUnit() As String = cell.Text.Trim.Split(Space(1))

                        strResult &= txtPriceUnit(0) & "\t" & txtPriceUnit(1) & "\t"
                    End If
                Else
                    If cell.Text.Equals("&nbsp;") Then
                        strResult &= "\t"
                    Else
                        '価格と単位を分ける
                        strResult &= cell.Text.Trim & "\t"
                    End If

                End If
            Next
        Next

        Return strResult
    End Function
#End Region
End Class