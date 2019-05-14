Public Class WebUC_KatOut
    Inherits KHBase

    Public Event BackToType()

    Private strcCompData As YousoBLL.CompData
    Private intStrWidth As Integer = 13                       '1文字の幅
    Private intHypenWidth As Integer = 20                     'ハイフン幅
    Private intVolStrcnt As Integer = 11                      '要素区分「1(電圧)」の文字数
    Private intStrokeStrcnt As Integer = 4                    '要素区分「3(ストローク)」の文字数
    Private HT_Sel As Hashtable

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        If objKtbnStrc Is Nothing Then
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
        End If
        Me.Session.Remove("HT_Sel")

        Me.HidGotID.Value = String.Empty   '初期化する
        Me.HidLostID.Value = String.Empty
        Call subGetCompData()       '構成データ取得
        Me.lblKataban.Text = objKtbnStrc.strcSelection.strGoodsNm

        Me.HidKey.Value = objKtbnStrc.strcSelection.strSeriesKataban & "_" & objKtbnStrc.strcSelection.strKeyKataban

        Me.HidGVStartID.Value = String.Empty
        For inti As Integer = 1 To 35
            If Not Me.PnlText.FindControl("txt" & inti) Is Nothing Then
                CType(Me.PnlText.FindControl("txt" & inti), TextBox).TabIndex = inti
            End If
        Next

        '選択欄の生成
        For inti As Integer = PnlText.Controls.Count - 1 To 0 Step -1
            PnlText.Controls(inti).Visible = False
        Next
        Call CreatTextBox()

        'Call Page_Load(Me, Nothing)

        If txt1.Visible = True Then txt1.Focus()
        GVDetail.DataSource = New DataTable
        GVDetail.DataBind()
    End Sub

    ''' <summary>
    ''' ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub
        If Me.objUserInfo.UserId Is Nothing Then Exit Sub
        If Me.HidKey.Value.ToString.Length <= 0 Then Exit Sub

        Try
            GVDetail.Visible = False
            If Me.GVDetail.Visible AndAlso Me.GVDetail.Rows.Count > 0 Then
                Me.HidGVStartID.Value = Me.GVDetail.Rows(0).ClientID.ToString
            End If

            '初期処理
            Call SetGridView()

            Call SetAllFontName(Me)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' オプション作成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetGridView()
        '初期処理
        If HidLostID.Value.Length > 0 AndAlso Not Me.PnlText.FindControl("txt" & HidLostID.Value) Is Nothing Then
            CType(Me.PnlText.FindControl("txt" & HidLostID.Value), TextBox).BackColor = DefaultColor
        End If
        If HidGotID.Value.Length > 0 AndAlso Not Me.PnlText.FindControl("txt" & HidGotID.Value) Is Nothing Then
            Me.PnlText.FindControl("txt" & HidGotID.Value).Focus()
            CType(Me.PnlText.FindControl("txt" & HidGotID.Value), TextBox).BackColor = Drawing.ColorTranslator.FromHtml("#FFCC33")

            If Me.HidDblClick.Value = "1" Then '選択（追加）
                If Not Me.Session("HT_Sel") Is Nothing Then
                    HT_Sel = Me.Session("HT_Sel")
                    If Not HT_Sel(HidGotID.Value) Is Nothing Then
                        If HT_Sel(HidGotID.Value).ToString.Length > 0 Then
                            If Me.HidListValue.Value.ToString.Length > 0 Then
                                HT_Sel(HidGotID.Value) &= "," & Me.HidListValue.Value
                            End If
                            Me.HidListValue.Value = HT_Sel(HidGotID.Value)
                        Else
                            HT_Sel(HidGotID.Value) = Me.HidListValue.Value
                        End If
                    Else
                        HT_Sel.Add(HidGotID.Value, Me.HidListValue.Value)
                    End If
                Else
                    HT_Sel = New Hashtable
                    HT_Sel.Add(HidGotID.Value, Me.HidListValue.Value.ToString)
                End If
                Me.Session("HT_Sel") = HT_Sel
            ElseIf Me.HidDblClick.Value = "0" Or Me.HidDblClick.Value = "" Then  '要素選択
                If Not Me.Session("HT_Sel") Is Nothing Then
                    HT_Sel = Me.Session("HT_Sel")
                    If Not HT_Sel(HidGotID.Value) Is Nothing Then
                        Me.HidListValue.Value = HT_Sel(HidGotID.Value)
                    Else
                        Me.HidListValue.Value = String.Empty
                        HT_Sel.Add(HidGotID.Value, "")
                    End If
                Else
                    Me.HidListValue.Value = String.Empty
                    HT_Sel = New Hashtable
                    HT_Sel.Add(HidGotID.Value, "")
                    Me.Session("HT_Sel") = HT_Sel
                End If
            ElseIf Me.HidDblClick.Value = "2" Then '選択（削除）
                If Not Me.Session("HT_Sel") Is Nothing Then
                    HT_Sel = Me.Session("HT_Sel")
                    If Not HT_Sel(HidGotID.Value) Is Nothing Then
                        If HT_Sel(HidGotID.Value).ToString.Length > 0 Then
                            If Me.HidListValue.Value.ToString.Length > 0 Then
                                Dim str() As String = HT_Sel(HidGotID.Value).ToString.Split(",")
                                Dim strNew As String = String.Empty
                                For inti As Integer = 0 To str.Length - 1
                                    If str(inti) <> Me.HidListValue.Value Then
                                        If strNew.Length <= 0 Then
                                            strNew = str(inti)
                                        Else
                                            strNew &= "," & str(inti)
                                        End If
                                    End If
                                Next
                                HT_Sel(HidGotID.Value) = strNew
                            End If
                            Me.HidListValue.Value = HT_Sel(HidGotID.Value)
                        End If
                    End If
                End If
                Me.Session("HT_Sel") = HT_Sel
            ElseIf Me.HidDblClick.Value = "3" Then '出力
                Me.HidDblClick.Value = ""

                ' 組合せ出力処理が実行可能かチェックする
                Dim intSeq As Integer = -1
                HT_Sel = Me.Session("HT_Sel")
                If Not HT_Sel Is Nothing Then
                    Dim strMsgCd As String = String.Empty
                    Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
                    If KHCombinationOut.ExecuteCheck(objKtbnStrc, HT_Sel, intSeq, strMsgCd) = False Then
                        Call AlertMessage(strMsgCd)
                    Else
                        Call CreatFile()
                    End If
                End If
                Exit Sub
            End If

            If Me.HidListValue.Value.ToString.Length > 0 Then
                Dim str() As String = Me.HidListValue.Value.ToString.Split(",")
                Dim dt As New DataTable
                Dim dc As New DataColumn("option_symbol")
                dt.Columns.Add(dc)
                Dim dr As DataRow = Nothing
                For inti As Integer = 0 To str.Length - 1
                    If str(inti).ToString.Trim.Length <= 0 Then Continue For
                    dr = dt.NewRow
                    dr("option_symbol") = str(inti)
                    dt.Rows.Add(dr)
                Next
                Me.GVSelect.DataSource = dt
                Me.GVSelect.DataBind()
            Else
                Me.GVSelect.DataSource = New DataTable
                Me.GVSelect.DataBind()
            End If

            '形番構成要素取得
            Dim dt_Option As New DS_KatOut.DT_OptionNameDataTable
            Using da As New DS_KatOutTableAdapters.DT_OptionNameTableAdapter
                Dim str() As String = Me.HidKey.Value.ToString.Split("_")
                If str.Length <> 2 Then Exit Sub
                dt_Option = da.GetData(Now, "ja", str(0).Trim, str(1).Trim, CLng(HidGotID.Value))

                Dim dt_title As New DS_KatSep.kh_ktbn_strc_nm_mstDataTable
                Using da_title As New DS_KatSepTableAdapters.kh_ktbn_strc_nm_mstTableAdapter
                    da_title.FillBy(dt_title, str(0).Trim, str(1).Trim, CLng(HidGotID.Value))
                End Using
                If dt_title.Rows.Count > 0 Then
                    Me.GVDetail.Columns(0).HeaderText = dt_title.Rows(0)("ktbn_strc_nm").ToString
                    Me.GVDetail.Columns(1).HeaderText = dt_title.Rows(0)("ktbn_strc_nm").ToString
                End If

                If Me.HidListValue.Value.ToString.Length > 0 Then
                    Dim str_del() As String = Me.HidListValue.Value.ToString.Split(",")
                    Dim dr() As DataRow
                    For inti As Integer = 0 To str_del.Length - 1
                        If str_del(inti) <> "無記号" Then
                            dr = dt_Option.Select("option_symbol='" & str_del(inti) & "'")
                        Else
                            dr = dt_Option.Select("option_symbol=''")
                        End If
                        If dr.Length > 0 Then dr(0).Delete()
                    Next
                End If
                dt_Option.AcceptChanges()
                GVDetail.Visible = True
                GVDetail.DataSource = dt_Option
                GVDetail.DataBind()
            End Using
        End If

        Me.HidListValue.Value = String.Empty
    End Sub

    ''' <summary>
    ''' TextBoxの作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreatTextBox() As Boolean
        CreatTextBox = False

        Dim intAllLen As Integer = 20
        Dim intTop As Integer = 10
        Dim intFirstLeft As Integer = 0
        Try
            '機種設定
            Dim objlbl As System.Web.UI.WebControls.Label = Me.PnlText.FindControl("txt0")
            objlbl.Text = strcCompData.strSeriesKataban
            objlbl.Font.Bold = True
            objlbl.Font.Size = myFontSize
            objlbl.EnableViewState = True
            objlbl.Visible = True
            intAllLen += objlbl.Width.Value

            If strcCompData.strHyphen = CdCst.HyphenDiv.Necessary Then
                objlbl = Me.PnlText.FindControl("H0")
                objlbl.Text = "－"
                objlbl.Width = intHypenWidth
                objlbl.Font.Bold = True
                objlbl.Font.Size = myFontSize
                objlbl.EnableViewState = True
                objlbl.Visible = True
                intAllLen += objlbl.Width.Value
            End If
            intFirstLeft = intAllLen

            Dim intLen As Integer = 0
            Dim intMax As Integer = 0
            For inti As Integer = 1 To strcCompData.strKtbnStrcNm.Length - 1
                intLen = 22
                intMax = 2
                '一行入れない、二行にします
                If intAllLen + intLen >= PnlText.Width.Value - 40 Then
                    intTop += 50
                    intAllLen = intFirstLeft    '前行の位置と合わせる
                End If

                Dim objtxt As System.Web.UI.WebControls.TextBox = Me.PnlText.FindControl("txt" & inti.ToString)
                objtxt.Width = intLen
                objtxt.TabIndex = inti + 1
                objtxt.Font.Bold = True
                objtxt.Font.Size = myFontSize
                intLen = objtxt.Width.Value
                objtxt.BackColor = DefaultColor
                objtxt.MaxLength = intMax
                objtxt.Text = String.Empty
                objtxt.EnableViewState = True
                objtxt.AutoPostBack = False
                objtxt.Visible = True
                objtxt.Style.Add("text-transform", "uppercase")
                objtxt.Attributes.Add("onBlur", "KatOutLostFocus('" & Me.ClientID & "_', '" & objtxt.ID & "');")
                objtxt.Attributes.Add("onFocus", "KatOutGotFocus('" & Me.ClientID & "_', '" & objtxt.ID & "');")

                'ハイフン設定
                If strcCompData.strHyphenDiv(inti).ToString = "1" Then     'ハイフンあり
                    objlbl = Me.PnlText.FindControl("H" & inti.ToString)
                    objlbl.Text = "－"
                    objlbl.Width = intHypenWidth
                    objlbl.Font.Bold = True
                    objlbl.Font.Size = myFontSize
                    objlbl.EnableViewState = True
                    objlbl.Visible = True
                    intLen += objlbl.Width.Value
                End If
                intAllLen += intLen
                objtxt = Nothing
                objlbl = Nothing
            Next
            intLen = Nothing
            intMax = Nothing
            CreatTextBox = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try

    End Function

    ''' <summary>
    ''' 構成情報取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subGetCompData()
        Dim bolReturn As Boolean
        Try
            '引当情報取得
            strcCompData.strSeriesKataban = objKtbnStrc.strcSelection.strSeriesKataban
            strcCompData.strKeyKataban = objKtbnStrc.strcSelection.strKeyKataban
            strcCompData.strFullKataban = objKtbnStrc.strcSelection.strFullKataban
            strcCompData.strGoodsNm = objKtbnStrc.strcSelection.strGoodsNm
            strcCompData.strHyphen = objKtbnStrc.strcSelection.strHyphen
            strcCompData.strOpSymbol = objKtbnStrc.strcSelection.strOpSymbol

            bolReturn = YousoBLL.fncKatabanStrcSelect(objCon, strcCompData, "ja") '形番構成取得
            bolReturn = YousoBLL.subKtbnStrcEleSelect(objCon, strcCompData)                        '形番構成要素取得
            HidMaxSelCount.Value = strcCompData.strStructureDiv.Length

            objKtbnStrc.strcSelection.strOpAdditionDiv = strcCompData.strAdditionDiv
            objKtbnStrc.strcSelection.strOpElementDiv = strcCompData.strElementDiv
            objKtbnStrc.strcSelection.strOpHyphenDiv = strcCompData.strHyphenDiv
            objKtbnStrc.strcSelection.strOpKtbnStrcNm = strcCompData.strKtbnStrcNm
            objKtbnStrc.strcSelection.strOpStructureDiv = strcCompData.strStructureDiv
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 戻るボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        RaiseEvent BackToType()
    End Sub

    ''' <summary>
    ''' 全ての機種を出力
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnOutPutAll_Click(sender As Object, e As EventArgs) Handles btnOutPutAll.Click
        Dim dt_all As New DS_KatOut.SeriesKatabanDataTable
        Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
        Using da As New DS_KatOutTableAdapters.SeriesKatabanTableAdapter
            da.FillAll(dt_all, Now, objKtbnStrc.strcSelection.strSeriesKataban)

            If objKtbnStrc.strcSelection.strKeyKataban.ToString.Trim.Length > 0 Then
                Dim dr_del() As DataRow = dt_all.Select("series_kataban='" & objKtbnStrc.strcSelection.strSeriesKataban & "'")
                If dr_del.Length > 0 Then
                    For inti As Integer = 0 To dr_del.Length - 1
                        If dr_del(inti)("key_kataban").ToString.Trim <> objKtbnStrc.strcSelection.strKeyKataban.ToString.Trim Then
                            dr_del(inti).Delete()
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
            dt_all.AcceptChanges()

            Dim dr As DataRow = Nothing
            Dim dt_vol As New DS_KatOut.kh_volAllDataTable
            Using da_vol As New DS_KatOutTableAdapters.kh_volAllTableAdapter
                dt_vol = da_vol.GetAllData(Now)
                If dt_vol.Rows.Count > 0 Then objKtbnStrc.strcSelection.dt_vol = dt_vol
            End Using

            Dim dt_stroke As New DS_KatOut.kh_strokeAllDataTable
            Using da_stroke As New DS_KatOutTableAdapters.kh_strokeAllTableAdapter
                dt_stroke = da_stroke.GetAllData(Now)
                If dt_stroke.Rows.Count > 0 Then objKtbnStrc.strcSelection.dt_Stroke = dt_stroke
            End Using

            For inti As Integer = 0 To dt_all.Rows.Count - 1
                dr = dt_all.Rows(inti)
                Me.Session.Remove("HT_Sel")
                objKtbnStrc = New KHKtbnStrc
                'ADD BY YGY 20140609 　　　初期化
                Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
                'objKtbnStrc.strcSelection.strRodEndOption = ""
                objKtbnStrc.strcSelection.dt_vol = dt_vol
                objKtbnStrc.strcSelection.dt_Stroke = dt_stroke
                objKtbnStrc.strcSelection.strSeriesKataban = dr("series_kataban").ToString.Trim.ToUpper
                objKtbnStrc.strcSelection.strKeyKataban = dr("key_kataban").ToString.Trim.ToUpper
                objKtbnStrc.strcSelection.strHyphen = dr("hyphen_div").ToString.Trim.ToUpper
                objKtbnStrc.strcSelection.strPriceNo = dr("price_no").ToString.Trim.ToUpper
                objKtbnStrc.strcSelection.strSpecNo = dr("spec_no").ToString.Trim.ToUpper
                objKtbnStrc.strcSelection.strGoodsNm = dr("series_nm").ToString.Trim

                'ADD BY YGY 20140609    マニホールドが対象外です    ↓↓↓↓↓↓
                If Not objKtbnStrc.strcSelection.strSpecNo.Equals(String.Empty) Then
                    Continue For
                End If
                'ADD BY YGY 20140609    ↑↑↑↑↑↑
                'ADD BY YGY 20140616    キー形番が'*'の場合は対象外です    ↓↓↓↓↓↓
                If objKtbnStrc.strcSelection.strKeyKataban.Equals("*") Then
                    Continue For
                End If
                'ADD BY YGY 20140616    ↑↑↑↑↑↑

                Me.HidGotID.Value = String.Empty   '初期化する
                Me.HidLostID.Value = String.Empty
                Call subGetCompData()       '構成データ取得
                Me.HidKey.Value = objKtbnStrc.strcSelection.strSeriesKataban & "_" & objKtbnStrc.strcSelection.strKeyKataban
                Me.HidGVStartID.Value = String.Empty
                '選択欄の生成
                For intj As Integer = PnlText.Controls.Count - 1 To 0 Step -1
                    PnlText.Controls(intj).Visible = False
                Next
                Call CreatTextBox()

                '初期処理
                Call SetGridView()
                '一括出力、メッセージ非表示
                Call fncOutPut(1)
            Next
        End Using
    End Sub

    'ファイル出力
    Protected Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Call fncOutPut()
    End Sub

    ''' <summary>
    ''' 機種出力
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub fncOutPut(Optional intMode As Integer = 0)
        '空白入力欄のすべて要素を出力対象になる
        Dim txt As TextBox = Nothing
        Dim dr() As DataRow = Nothing

        '形番構成要素取得
        Dim dt_Option As New DS_KatOut.DT_OptionNameDataTable
        Using da As New DS_KatOutTableAdapters.DT_OptionNameTableAdapter
            Dim str() As String = Me.HidKey.Value.ToString.Split("_")
            If str.Length <> 2 Then Exit Sub
            dt_Option = da.GetAllData(Now, "ja", str(0).Trim, str(1).Trim)
        End Using

        For inti As Integer = 1 To 35
            If Not Me.PnlText.FindControl("txt" & inti) Is Nothing AndAlso Me.PnlText.FindControl("txt" & inti).Visible Then
                txt = CType(Me.PnlText.FindControl("txt" & inti), TextBox)
                If txt.Text.ToString.Trim.Length <= 0 Then '未選択
                    dr = dt_Option.Select("ktbn_strc_seq_no='" & inti.ToString & "'")
                    'If dr.Length > 0 Then      DELETE BY YGY 20140604
                    txt.Text = dr.Length

                    Dim strKey As String = String.Empty
                    Dim strValue As String = String.Empty
                    For intj As Integer = 0 To dr.Length - 1
                        strValue = dr(intj)("option_symbol").ToString.Trim
                        If strValue.Trim = "" Then strValue = "無記号"
                        If strKey.Length <= 0 Then
                            strKey = strValue
                        Else
                            strKey &= "," & strValue
                        End If
                    Next

                    If Not Me.Session("HT_Sel") Is Nothing Then
                        HT_Sel = Me.Session("HT_Sel")
                        If Not HT_Sel(inti.ToString) Is Nothing Then
                            HT_Sel(inti.ToString) = strKey
                        Else
                            HT_Sel.Add(inti.ToString, strKey)
                        End If
                    Else
                        HT_Sel = New Hashtable
                        HT_Sel.Add(inti.ToString, strKey)
                    End If
                    Me.Session("HT_Sel") = HT_Sel
                    'End If                    DELETE BY YGY 20140604
                End If
            End If
        Next

        If intMode = 0 Then
            ' 確認メッセージ出力
            Dim sbScript As New StringBuilder
            Dim strMessage As String = "選択したオプションで組合せ可能な形番情報を全て生成しファイルに出力します。" & _
                      "選択したオプション数によっては多少時間が掛かりますがよろしいですか？"
            sbScript.Append("KatOutConfirm('" & strMessage & "','" & Me.ClientID & "_');")
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "SetValue", sbScript.ToString, True)
        Else
            Me.HidDblClick.Value = ""

            ' 組合せ出力処理が実行可能かチェックする
            Dim intSeq As Integer = -1
            HT_Sel = Me.Session("HT_Sel")
            If Not HT_Sel Is Nothing Then
                Dim strMsgCd As String = String.Empty
                If KHCombinationOut.ExecuteCheck(objKtbnStrc, HT_Sel, intSeq, strMsgCd) = False Then
                    Dim strOutSeries As String = objKtbnStrc.strcSelection.strSeriesKataban & _
                        IIf(objKtbnStrc.strcSelection.strKeyKataban.Length > 0, "_" & _
                            objKtbnStrc.strcSelection.strKeyKataban, "")
                    Dim strPath As String = My.Settings.LogFolder & strOutSeries & ".txt"
                    System.IO.File.WriteAllText(strPath, strMsgCd, System.Text.Encoding.UTF8)
                Else
                    Call CreatFile(intMode)
                End If
            End If
            Exit Sub
        End If
    End Sub

    ''' <summary>
    ''' バインド
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GVDetail_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVDetail.RowDataBound
        Try
            If e.Row.RowIndex < 0 Then
                e.Row.Cells(0).ColumnSpan = 2
                e.Row.Cells.RemoveAt(1)
                Exit Sub
            End If

            Dim strName As String = Me.ClientID & "_"
            Dim intStartID As Integer = 0
            If e.Row.RowIndex = 0 Then
                intStartID = CInt(Strings.Right(e.Row.ClientID, 2))
            Else
                intStartID = CInt(Strings.Right(GVDetail.Rows(0).ClientID, 2))
            End If

            e.Row.TabIndex = e.Row.RowIndex + 36
            e.Row.Attributes.Add(CdCst.JavaScript.OnClick, "fncGridClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',1);")
            e.Row.Attributes.Add(CdCst.JavaScript.OnKeyUp, "fncGrid_OnKeyup(event, '" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',1);")
            e.Row.Attributes.Add(CdCst.JavaScript.OnDblClick, "KatOutDblClick('" & strName & "','" & e.Row.ClientID & "');")
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' バインド
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GVSelect_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVSelect.RowDataBound
        Try
            If e.Row.RowIndex < 0 Then Exit Sub
            Dim strName As String = Me.ClientID & "_"
            Dim intStartID As Integer = 0
            If e.Row.RowIndex = 0 Then
                intStartID = CInt(Strings.Right(e.Row.ClientID, 2))
            Else
                intStartID = CInt(Strings.Right(GVSelect.Rows(0).ClientID, 2))
            End If

            e.Row.TabIndex = e.Row.RowIndex + 36
            e.Row.Attributes.Add(CdCst.JavaScript.OnClick, "fncGridClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',1);")
            e.Row.Attributes.Add(CdCst.JavaScript.OnKeyUp, "fncGrid_OnKeyup(event, '" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',1);")
            e.Row.Attributes.Add(CdCst.JavaScript.OnDblClick, "KatOutSelDblClick('" & strName & "','" & e.Row.ClientID & "');")
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' ファイル作成
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub CreatFile(Optional intMode As Integer = 0)
        Dim HTResult As New ArrayList
        Dim HTItem As New ArrayList
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
        Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban

        Dim DS_Tab As New DataSet
        DS_Tab = KHCombinationOut.GetCacheTable(strSeriesKata, strKeyKata, objKtbnStrc.strcSelection.strPriceNo.ToString.Trim)

        Dim HTOut As New ArrayList
        Dim HTCheck As New Hashtable
        Dim HTCheckOp As New Hashtable
        HTOut.Add("disp_kataban")
        If chkKata.Checked Then HTOut.Add("kataban_check_div")
        If chkPlace.Checked Then HTOut.Add("place_cd")
        If chkName.Checked Then HTOut.Add("disp_name")
        If chkls.Checked Then HTOut.Add("ls_price")
        If chkrgs.Checked Then HTOut.Add("rg_price")
        If chkss.Checked Then HTOut.Add("ss_price")
        If chkbs.Checked Then HTOut.Add("bs_price")
        If chkgs.Checked Then HTOut.Add("gs_price")
        If chkps.Checked Then HTOut.Add("ps_price")

        ' 形番を生成する
        Dim ItemCode() As String = Nothing
        Dim timStart As Date = Now
        Dim strOutSeries As String = strSeriesKata & IIf(strKeyKata.Length > 0, "_" & strKeyKata, "")

        Dim intAll As Long = 0
        Dim strPath As String = My.Settings.LogFolder & strOutSeries & ".txt"
        'フル形番生成
        HT_Sel = New Hashtable
        If Not Me.Session("HT_Sel") Is Nothing Then HT_Sel = Me.Session("HT_Sel")
        System.IO.File.WriteAllText(strPath, "", System.Text.Encoding.UTF8)

        Call KHCombinationOut.Kataban_Deployment(objCon, objKtbnStrc, 1, HT_Sel, ItemCode, HTResult, HTItem, _
                                          DS_Tab, HTCheck, HTCheckOp, HTOut, strPath, intAll)

        '単価取得
        Dim objUnitPrice As New KHUnitPrice
        Dim strOutFile As String = String.Empty
        For inti As Integer = 0 To HTResult.Count - 1
            objKtbnStrc.strcSelection.strFullKataban = HTResult(inti)
            objKtbnStrc.strcSelection.strOpSymbol = HTItem(inti)
            objKtbnStrc.strcSelection.intListPrice = 0
            objKtbnStrc.strcSelection.intRegPrice = 0
            objKtbnStrc.strcSelection.intSsPrice = 0
            objKtbnStrc.strcSelection.intBsPrice = 0
            objKtbnStrc.strcSelection.intGsPrice = 0
            objKtbnStrc.strcSelection.intPsPrice = 0
            objKtbnStrc.strcSelection.strKatabanCheckDiv = ""
            objKtbnStrc.strcSelection.strPlaceCd = ""

            Call objUnitPrice.subPriceInfoSet_ForkatOut(objCon, objKtbnStrc, "JPN", "", DS_Tab)

            Try
                strOutFile &= objKtbnStrc.strcSelection.strFullKataban & ControlChars.Tab

                If HTOut.Contains("kataban_check_div") Then
                    strOutFile &= "Z" & objKtbnStrc.strcSelection.strKatabanCheckDiv & ControlChars.Tab
                End If
                If HTOut.Contains("place_cd") Then
                    strOutFile &= objKtbnStrc.strcSelection.strPlaceCd & ControlChars.Tab
                End If
                If HTOut.Contains("disp_name") Then
                    strOutFile &= objKtbnStrc.strcSelection.strGoodsNm & ControlChars.Tab
                End If
                If HTOut.Contains("ls_price") Then
                    strOutFile &= CInt(objKtbnStrc.strcSelection.intListPrice) & ControlChars.Tab
                End If
                If HTOut.Contains("rg_price") Then
                    strOutFile &= CInt(objKtbnStrc.strcSelection.intRegPrice) & ControlChars.Tab
                End If
                If HTOut.Contains("ss_price") Then
                    strOutFile &= CInt(objKtbnStrc.strcSelection.intSsPrice) & ControlChars.Tab
                End If
                If HTOut.Contains("bs_price") Then
                    strOutFile &= CInt(objKtbnStrc.strcSelection.intBsPrice) & ControlChars.Tab
                End If
                If HTOut.Contains("gs_price") Then
                    strOutFile &= CInt(objKtbnStrc.strcSelection.intGsPrice) & ControlChars.Tab
                End If
                If HTOut.Contains("ps_price") Then
                    strOutFile &= CInt(objKtbnStrc.strcSelection.intPsPrice) & ControlChars.NewLine
                End If

                strOutFile &= ControlChars.NewLine
                
            Catch ex As Exception
            End Try
        Next
        If strOutFile.Length > 0 Then System.IO.File.AppendAllText(strPath, strOutFile, System.Text.Encoding.UTF8)
        strOutFile = String.Empty
        intAll += HTResult.Count

        HTResult.Clear()
        HTItem.Clear()
        HTCheck.Clear()
        HTCheckOp.Clear()

        DS_Tab = Nothing
        Dim timEnd As Date = Now
        Dim myAllTime As TimeSpan = timEnd - timStart

        intAll += HTResult.Count

        HTResult = Nothing
        HTCheck = Nothing
        GC.Collect()

        If intMode = 0 Then
            ' 確認メッセージ出力
            Dim sbScript As New StringBuilder
            Dim strMessage As String = "機種「" & strSeriesKata & "」の組合せ出力が完了しました。" & _
                      "出力件数：" & intAll & "件。"
            sbScript.Append("alert('" & strMessage & "');")
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "alert", sbScript.ToString, True)
        End If
    End Sub

End Class