Imports WebKataban.ClsCommon

Public Class WebUC_Type
    Inherits KHBase

#Region "プロパティ"
    Public Event GotoYouso()
    Public Event GotoTanka()
    Public Event GotoKatOut()
    Public Event GotoKatsepchk()
    Public Event Goto100test()

    Private ListKey As New ArrayList                    '画面上の機種情報
    Private bllType As New TypeBLL
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
        '初期値の設定
        txtKataban.Text = String.Empty
        RadioButtonList1.SelectedIndex = 0
    End Sub

    ''' <summary>
    ''' 外部からの呼出
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        '前画面から戻ってきた場合、機種を入力して検索
        Dim strKisyu As String = String.Empty
        Dim intSearchDiv As String = -1

        '選択履歴
        If Not Session("KisyuInfo") Is Nothing Then
            strKisyu = Session.Item("KisyuInfo")(0)
            intSearchDiv = Session.Item("KisyuInfo")(1)
            Session.Remove("KisyuInfo")
        End If

        '画面をロードする
        Me.OnLoad(Nothing)

        '機種で検索
        If Not strKisyu.Equals(String.Empty) Then
            txtKataban.Text = strKisyu
            If Not intSearchDiv = -1 Then
                RadioButtonList1.SelectedIndex = intSearchDiv
            End If
            SearchType(1)
        End If
        'テキストボックスの設定
        txtKataban.Attributes.Add("onfocus", "this.select();")
        txtKataban.Focus()
    End Sub

    ''' <summary>
    ''' ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub
        If Me.objUserInfo.UserId Is Nothing Then Exit Sub
        If selLang Is Nothing Then Exit Sub
        If Not FormIDCheck() Then Exit Sub
        GVDetail.Visible = False
        Try
            'セッションのクリア
            If Not Me.Session("KtbnStrc") Is Nothing Then Me.Session.Remove("KtbnStrc")

            '画面ラベル設定
            Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHModelSelection, selLang.SelectedValue, Me)
            bllType.subDeleteSelKtbnInfo(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

            'ShiireSearchModeが0の場合は仕入品非表示
            If My.Settings.ShiireSearchMode = 0 Then
                RadioButtonList1.Items(2).Attributes.Add("style", "display:none")
                RadioButtonList1.Width = Nothing
            End If

            'フォントの設定
            Call SetAllFontName(Me)

            'ボタン表示の設定
            Button3.Visible = False    '次のページ
            Button2.Visible = False    '前のページ
            Panel7.Visible = False
            If Me.objUserInfo.UserClass = CdCst.UserClass.InfoSysForceSysAdmin Then
                btnKatOut.Visible = True
                btnKatsepchk.Visible = True
                btn100Test.Visible = True
            Else
                btnKatOut.Visible = False
                btnKatsepchk.Visible = False
                btn100Test.Visible = False
            End If

            'フォカスのセット
            'Me.txtKataban.Focus()

            'Javascriptの設定
            subSetInitScript()

            '国内代理店はメッセージを表示させる
            Select Case Me.objUserInfo.UserClass
                Case CdCst.UserClass.DmAgentRs, CdCst.UserClass.DmAgentSs, CdCst.UserClass.DmAgentBs, CdCst.UserClass.DmAgentGs, CdCst.UserClass.DmAgentPs
                    '国内代理店はメッセージを表示させる
                    Me.ImgFixedMessage1.Visible = True
                Case Else
                    Me.ImgFixedMessage1.Visible = False
            End Select

            'マニホールドテスト専用
            If Not Me.Session("ManifoldSeriesKey") Is Nothing Then
                If Me.Session("TestFlag") Is Nothing Then
                    If Me.txtKataban.Text.ToString.Trim.Length > 0 Then
                        Me.Session("TestFlag") = True
                        Call BtnSearch_Click(Me, Nothing)
                    End If
                End If
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 検索ボタンを押す
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles Button4.Click
        HidSelPage.Value = String.Empty
        HidSelRowID.Value = String.Empty
        '検索
        If fncValidateSearch() Then Call SearchType(1)

        'マニホールドテスト専用
        If Not Me.Session("ManifoldSeriesKey") Is Nothing Then
            '選択したデータを探す
            ListKey = Me.ViewState.Item(CdCstType.strDTList)
            Dim str() As String = Nothing
            Dim strSession() As String = Me.Session("ManifoldSeriesKey").ToString.Split(",")
            For inti As Integer = 0 To ListKey.Count - 1
                str = ListKey(inti).ToString.Split("_")
                If str.Length < 5 Then Continue For
                If str(0) = strSession(0) And str(1) = strSession(1) Then
                    Dim intStartID As Integer = CInt(Strings.Right(GVDetail.Rows(0).ClientID, 2)) + inti
                    Me.HidSelRowID.Value = intStartID
                    Me.Session("TestFlag") = Nothing
                    Call chkbtnOK()
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 検索
    ''' </summary>
    ''' <param name="intMode">1:検索ボタン、2:前ページ、3：次ページ</param>
    ''' <remarks></remarks>
    Public Sub SearchType(intMode As Integer)
        Dim strKata As String = StrConv(Me.txtKataban.Text.Trim, VbStrConv.Narrow)
        Dim strMinKata As String = String.Empty
        Button3.Visible = False
        Button2.Visible = False
        Panel7.Visible = False

        Try
            Dim clsSeriesSearch As KHSeriesSearch
            Dim dsKataban As DataSet

            If Not Me.ViewState.Item(CdCstType.strPageFirstKisyu) Is Nothing Then
                Select Case intMode
                    Case 1 '1:検索ボタン
                        Me.HidRowCount.Value = String.Empty
                        Me.ViewState.Remove(CdCstType.strPageFirstKisyu)   '検索時、履歴をクリアする
                    Case 2 '2:前ページ
                        ListKey = Me.ViewState.Item(CdCstType.strPageFirstKisyu)
                        If ListKey.Count > 2 Then
                            If Me.HidRowCount.Value < 16 Then
                                strMinKata = ListKey(ListKey.Count - 2).ToString
                            Else
                                strMinKata = ListKey(ListKey.Count - 3).ToString
                                ListKey.RemoveAt(ListKey.Count - 1)
                                Me.ViewState.Item(CdCstType.strPageFirstKisyu) = ListKey
                            End If

                        ElseIf ListKey.Count = 2 Then
                            If Me.HidRowCount.Value < 16 Then
                                strMinKata = ListKey(ListKey.Count - 2).ToString
                            End If
                            ListKey.RemoveAt(ListKey.Count - 1)
                            Me.ViewState.Item(CdCstType.strPageFirstKisyu) = ListKey
                        Else
                            Me.ViewState.Remove(CdCstType.strPageFirstKisyu)   '検索時、履歴をクリアする
                        End If
                    Case 3 '3：次ページ
                        Me.HidRowCount.Value = String.Empty
                        ListKey = Me.ViewState.Item(CdCstType.strPageFirstKisyu)
                        If ListKey.Count > 0 Then
                            strMinKata = ListKey(ListKey.Count - 1).ToString
                        End If
                End Select
            End If

            '検索
            Dim lstWhereSeries As New ArrayList
            If Me.objUserInfo.UserClass = CdCst.UserClass.OsAgentLs Then
                If Me.objUserInfo.PersonCd = "EC" Then
                    lstWhereSeries.AddRange(New String() {"AX1", "AX2", "AX4", "AX6", "AX8", _
                                                          "ETS", "ECS", "ETV", "ERL2", "ESD2", _
                                                          "KBZ", "KBB", "KSA", "ESSD", "ELCR"})
                End If
            End If

            '検索
            clsSeriesSearch = New KHSeriesSearch(RadioButtonList1.SelectedValue, selLang.SelectedValue, _
                                              strKata, CdCstType.intMaxRowCnt + 1, strMinKata, Me.objUserInfo.CountryCd)
            dsKataban = bllType.fncSearch(objCon, clsSeriesSearch, lstWhereSeries)

            '検索結果を記録
            If Not dsKataban Is Nothing AndAlso Not dsKataban.Tables("KatabanTbl") Is Nothing Then
                If dsKataban.Tables("KatabanTbl").Rows.Count <= 0 Then
                    Select Case RadioButtonList1.SelectedValue
                        Case "3" '仕入品
                            If clsSeriesSearch.strResultTypeCdValue = KHSeriesSearch.ResultType.MaxCountOver Then
                                '検索結果が2000件を超える場合
                                AlertMessage("I0160")
                                Me.txtKataban.Focus()
                                Exit Sub
                            Else
                                '検索データがない場合
                                AlertMessage("I0020")
                                Me.txtKataban.Focus()
                                Exit Sub
                            End If
                        Case Else
                            '検索データがない場合
                            AlertMessage("I0020")
                            Me.txtKataban.Focus()
                            Exit Sub
                    End Select
                End If

                '件数を記録
                Me.HidRowCount.Value = dsKataban.Tables("KatabanTbl").Rows.Count

                '次ページ始まる機種を記録
                If intMode <> 2 Or Me.ViewState.Item(CdCstType.strPageFirstKisyu) Is Nothing Then

                    ListKey = New ArrayList

                    If Not Me.ViewState.Item(CdCstType.strPageFirstKisyu) Is Nothing Then
                        ListKey = Me.ViewState.Item(CdCstType.strPageFirstKisyu)
                    End If

                    If dsKataban.Tables("KatabanTbl").Rows.Count > 15 Then
                        ListKey.Add(dsKataban.Tables("KatabanTbl").Rows(CdCstType.intMaxRowCnt - 1)("sortKey").ToString)
                    End If
                    Me.ViewState.Add(CdCstType.strPageFirstKisyu, ListKey)
                End If

                '画面情報をViewStateに保存する
                Dim dt_list As New ArrayList
                Dim dr As DataRow = Nothing
                Dim strKeyKatas As String = String.Empty
                '通貨
                Dim strCurrency As String = String.Empty

                For inti As Integer = 0 To dsKataban.Tables("KatabanTbl").Rows.Count - 1
                    dr = dsKataban.Tables("KatabanTbl").Rows(inti)
                    dt_list.Add(dr("series_kataban") & "_" & dr("key_kataban") & "_" & dr("disp_kataban") & "_" & _
                                dr("disp_name") & "_" & dr("division") & "_" & dr("currency_cd"))
                    strKeyKatas = strKeyKatas & dr("key_kataban") & CdCst.Sign.Delimiter.Comma
                    strCurrency = strCurrency & dr("currency_cd") & CdCst.Sign.Delimiter.Comma
                Next
                Me.ViewState.Add(CdCstType.strDTList, dt_list)
                'キー形番を記録
                HidKeyKatabans.Value = strKeyKatas
                HidCurrency.Value = strCurrency
                ListKey = Nothing

                'ボタンの設定
                Dim intPageID As Integer = 0
                Select Case intMode
                    Case 1
                        HidSelPage.Value = 1
                    Case 2
                        intPageID = HidSelPage.Value
                        If intPageID > 1 Then HidSelPage.Value = intPageID - 1
                        If HidSelPage.Value > 1 Then Button2.Visible = True
                    Case 3
                        intPageID = HidSelPage.Value
                        HidSelPage.Value = intPageID + 1
                        Button2.Visible = True
                End Select
                HidSelRowID.Value = String.Empty

                '結果リストの作成
                Call CreatTextCell(1, dsKataban.Tables("KatabanTbl"))

                'ボタン表示
                If dsKataban.Tables("KatabanTbl").Rows.Count > 0 Then
                    Panel7.Visible = True
                End If

                '次ページボタン
                If dsKataban.Tables("KatabanTbl").Rows.Count > 15 Then
                    Button3.Visible = True
                End If

                dsKataban = Nothing


            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 次のページ
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnNext_Click(sender As Object, e As System.EventArgs) Handles Button3.Click
        SearchType(3)
    End Sub

    ''' <summary>
    ''' 前のページ
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnPrev_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        SearchType(2)
    End Sub

    ''' <summary>
    ''' 結果リストの作成
    ''' </summary>
    ''' <param name="intPageID"></param>
    ''' <param name="dtResult"></param>
    ''' <remarks></remarks>
    Private Sub CreatTextCell(intPageID As Integer, dtResult As DataTable)
        Try
            GVDetail.Visible = True

            '検索結果データを取得
            Dim dt_view As DataTable = dtResult.Clone
            Dim dr_view As DataRow = Nothing
            If Not dtResult Is Nothing Then
                For inti As Integer = 0 To dtResult.Rows.Count - 1
                    If inti >= 15 Then Exit For
                    dr_view = dt_view.NewRow
                    dr_view.ItemArray = dtResult.Rows(inti).ItemArray
                    dt_view.Rows.Add(dr_view)
                Next
            End If

            'GridViewの構築
            Dim dt_Title As DataTable = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHModelSelection, selLang.SelectedValue)
            Dim dr() As DataRow = dt_Title.Select("label_seq='3' AND label_div='L'")
            GVDetail.Columns.Clear()
            Dim col As New Web.UI.WebControls.BoundField
            If dr.Length > 0 Then col.HeaderText = dr(0)("label_content").ToString
            col.DataField = "disp_kataban"
            col.ItemStyle.Wrap = True
            col.ItemStyle.HorizontalAlign = HorizontalAlign.Left
            col.ItemStyle.Width = WebControls.Unit.Percentage(35)
            col.ItemStyle.Height = WebControls.Unit.Pixel(20)
            GVDetail.Columns.Add(col)

            dr = dt_Title.Select("label_seq='2' AND label_div='L'")
            col = New Web.UI.WebControls.BoundField
            If dr.Length > 0 Then col.HeaderText = dr(0)("label_content").ToString
            col.DataField = "disp_name"
            col.ItemStyle.Wrap = True
            col.ItemStyle.HorizontalAlign = HorizontalAlign.Left
            col.ItemStyle.Width = WebControls.Unit.Percentage(65)
            col.ItemStyle.Height = WebControls.Unit.Pixel(20)
            GVDetail.Columns.Add(col)
            GVDetail.DataSource = dt_view
            GVDetail.DataBind()

            'フォカスを第一行に設定する
            If dt_view.Rows.Count > 0 Then
                Dim strName As String = Me.ClientID & "_"
                Dim intStartID As Integer = CInt(Strings.Right(GVDetail.Rows(0).ClientID, 2))

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "TestValue",
                         "fncGridClick('" & strName & "','" & GVDetail.Rows(0).ClientID & "','" & intStartID & "',0);", True)
                Me.GVDetail.Rows(0).Cells(0).Focus()
                Me.HidSelRowID.Value = intStartID
            Else
                Me.txtKataban.Focus()
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' OKボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnOK_Click(sender As Object, e As System.EventArgs) Handles btnOK.Click
        Debug.WriteLine(My.Settings.LogFolder & "負荷テスト\新.txt", "形番：" & txtKataban.Text.Trim & ControlChars.Tab)

        Call chkbtnOK()
    End Sub

    ''' <summary>
    ''' OKボタンを押す
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function chkbtnOK() As Boolean
        chkbtnOK = False
        Try
            If HidSelPage.Value = String.Empty OrElse HidSelRowID.Value = String.Empty OrElse _
                Me.ViewState.Item(CdCstType.strDTList) Is Nothing Then
                AlertMessage("W0090") '先に検索してください。
                Me.txtKataban.Focus()
            Else
                '検索情報
                Dim strKisyuInfo As New List(Of String)
                '選択した行番号
                Dim intStartID As Integer = CInt(Strings.Right(Me.GVDetail.Rows(0).ClientID, 2))
                '選択データ
                Dim strAllInfo() As String
                Dim strSeries As String = String.Empty
                Dim strKeyKataban As String = String.Empty
                Dim strFullKataban As String = String.Empty
                Dim strGoodsNum As String = String.Empty
                Dim strCurrency As String = String.Empty

                '検索情報の保存
                strKisyuInfo.Add(txtKataban.Text.Trim)
                strKisyuInfo.Add(RadioButtonList1.SelectedIndex)
                Session.Add("KisyuInfo", strKisyuInfo)

                '選択したデータの取得
                ListKey = Me.ViewState.Item(CdCstType.strDTList)
                strAllInfo = ListKey(CInt(HidSelRowID.Value) - intStartID).ToString.Split("_")

                If strAllInfo.Length < 5 Then
                    '情報に誤りがある場合
                Else
                    '機種
                    strSeries = strAllInfo(0).ToString
                    'キー形番
                    strKeyKataban = strAllInfo(1).ToString
                    'フル形番
                    strFullKataban = strAllInfo(2).ToString
                    '表示名
                    strGoodsNum = strAllInfo(3).ToString
                    '通貨
                    strCurrency = strAllInfo(5).ToString

                    '選択したデータを保存する
                    Select Case strAllInfo(4).ToString
                        Case "1"
                            '引当シリーズ形番追加(機種)
                            Call bllType.subInsertSelSrsKtbnMdl(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                strSeries, strKeyKataban, strGoodsNum, strCurrency)
                            'ページ遷移(形番引当画面)
                            RaiseEvent GotoYouso()
                        Case "2"
                            '引当シリーズ形番追加(フル形番)
                            Call bllType.subInsertSelSrsKtbnFull(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                strFullKataban, strGoodsNum, strCurrency)
                            'ページ遷移(単価見積画面)
                            RaiseEvent GotoTanka()
                        Case "3"
                            '引当シリーズ形番追加(仕入品)
                            Call bllType.subInsertSelSrsKtbnShiire(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                strFullKataban, strSeries, strCurrency)
                            'ページ遷移(単価見積画面)
                            RaiseEvent GotoTanka()

                    End Select
                    GC.Collect()
                    chkbtnOK = True
                End If
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' データバインドイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GVDetail_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVDetail.RowDataBound
        If e.Row.RowIndex < 0 Then Exit Sub
        Try
            Dim strName As String = Me.ClientID & "_"
            Dim intStartID As Integer = 0
            If e.Row.RowIndex = 0 Then
                intStartID = CInt(Strings.Right(e.Row.ClientID, 2))
            Else
                intStartID = CInt(Strings.Right(GVDetail.Rows(0).ClientID, 2))
            End If
            e.Row.TabIndex = e.Row.RowIndex + 4
            'If (e.Row.RowIndex + 1) Mod 2 = 0 Then
            '    e.Row.BackColor = Drawing.Color.FromArgb(204, 204, 255)
            '    'e.Row.BackColor = Drawing.Color.FromArgb(173, 205, 207)
            'Else
            '    e.Row.BackColor = Drawing.Color.White
            'End If

            e.Row.Attributes.Add(CdCst.JavaScript.OnClick, "fncGridClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',0);")
            'Firefox対応するために
            'e.Row.Attributes.Add(CdCst.JavaScript.OnKeyUp, "fncGrid_OnKeyup('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',0);")
            e.Row.Attributes.Add(CdCst.JavaScript.OnKeyUp, "fncGrid_OnKeyup(event, '" & strName & "','" & e.Row.ClientID & "','" & intStartID & "', 0);")
            e.Row.Attributes.Add(CdCst.JavaScript.OnDblClick, "TypeDblClick('" & strName & "');")
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' JavaScript生成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInitScript()
        Dim strJSGVDetail As String
        Dim strJSTxtKataban As String

        Try
            '機種テキストイベントの設定
            txtKataban.Attributes.Add("onKeyDown", "KatabanKeyDown(event," & "'" & strParent & Me.ID & "_', '" & txtKataban.ID & "');")
            txtKataban.Style.Add("text-transform", "uppercase")

            'GridViewのEnterKeyイベントの設定
            strJSGVDetail = "if (event.keyCode == 13){document.getElementById('" & Me.ClientID & "_btnOK" & "').focus();return false;}else{return true;}"
            Me.GVDetail.Attributes.Add(CdCst.JavaScript.OnKeyDown, strJSGVDetail)

            'テキストエリアのEnterKeyイベントの設定
            strJSTxtKataban = "if (event.keyCode == 13){document.getElementById('" & Me.ClientID & "_Button4" & "').focus();return false;}else{return true;}"
            Me.txtKataban.Attributes.Add(CdCst.JavaScript.OnKeyDown, strJSTxtKataban)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' 入力情報のチェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncValidateSearch() As Boolean
        fncValidateSearch = False
        Try
            If Me.txtKataban.Text.Trim.Length <= 0 Then
                AlertMessage("W0070") '検索条件を入力してください。
                Me.txtKataban.Focus()
                Exit Function
            End If
            If ClsCommon.fncCnvNarrow(Me.txtKataban.Text.Trim) = False Then
                AlertMessage("W0060")
                Me.txtKataban.Focus()
                Exit Function
            End If
            'W0080 検索区分を指定してください。
            fncValidateSearch = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

#Region "自動テスト関連"
    ''' <summary>
    ''' 組合せ出力画面
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnKatOut_Click(sender As Object, e As EventArgs) Handles btnKatOut.Click
        If HidSelPage.Value = String.Empty OrElse HidSelRowID.Value = String.Empty OrElse _
                Me.ViewState.Item(CdCstType.strDTList) Is Nothing Then
            AlertMessage("W0090") '先に検索してください。
            Me.txtKataban.Focus()
            Exit Sub
        End If

        Dim intStartID As Integer = CInt(Strings.Right(Me.GVDetail.Rows(0).ClientID, 2))
        '選択したデータを探す
        ListKey = Me.ViewState.Item(CdCstType.strDTList)
        Dim str() As String = ListKey(CInt(HidSelRowID.Value) - intStartID).ToString.Split("_")
        If str.Length < 5 Then Exit Sub

        Select Case str(4).ToString
            Case "1"
                '引当シリーズ形番追加(機種)
                '通貨追加
                If objKtbnStrc.strcSelection.strCurrency Is Nothing Then
                    objKtbnStrc.strcSelection.strCurrency = "JPY"
                End If
                Call bllType.subInsertSelSrsKtbnMdl(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                    str(0).ToString, str(1).ToString, str(3).ToString, objKtbnStrc.strcSelection.strCurrency)
                Dim bolReturn As Boolean = False

                'ページ遷移(形番引当画面)
                '引当情報取得
                Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
                Dim strcCompData As New YousoBLL.CompData
                strcCompData.strSeriesKataban = objKtbnStrc.strcSelection.strSeriesKataban
                strcCompData.strKeyKataban = objKtbnStrc.strcSelection.strKeyKataban
                strcCompData.strFullKataban = objKtbnStrc.strcSelection.strFullKataban
                strcCompData.strGoodsNm = objKtbnStrc.strcSelection.strGoodsNm
                strcCompData.strHyphen = objKtbnStrc.strcSelection.strHyphen
                strcCompData.strOpSymbol = objKtbnStrc.strcSelection.strOpSymbol

                bolReturn = YousoBLL.fncKatabanStrcSelect(objCon, strcCompData, selLang.SelectedValue) '形番構成取得
                bolReturn = YousoBLL.subKtbnStrcEleSelect(objCon, strcCompData)                        '形番構成要素取得

                For intLoopCnt = 1 To strcCompData.strElementDiv.Length - 1
                    Call objKtbnStrc.subSelKtbnStrcIns(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                                       intLoopCnt, strcCompData.strElementDiv(intLoopCnt), _
                                                       strcCompData.strStructureDiv(intLoopCnt), _
                                                       strcCompData.strAdditionDiv(intLoopCnt), _
                                                       strcCompData.strHyphenDiv(intLoopCnt), _
                                                       strcCompData.strKtbnStrcNm(intLoopCnt), 0)
                Next

                RaiseEvent GotoKatOut()
            Case "2"
                AlertMessage("W0090") '先に検索してください。
        End Select
    End Sub

    ''' <summary>
    ''' 形番分解画面
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnKatsepchk_Click(sender As Object, e As EventArgs) Handles btnKatsepchk.Click
        RaiseEvent GotoKatsepchk()
    End Sub

    ''' <summary>
    ''' 100万件テスト
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btn100Test_Click(sender As Object, e As EventArgs) Handles btn100Test.Click
        RaiseEvent Goto100test()
    End Sub
#End Region

End Class