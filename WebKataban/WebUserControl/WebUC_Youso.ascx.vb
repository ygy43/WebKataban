Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports WebKataban.YousoBLL
Imports WebKataban.KHCodeConstants.CdCst

Public Class WebUC_Youso
    Inherits KHBase

#Region "プロパティ"
    'オプション情報
    Private strcCompData As CompData
    Public Event BtnOKGo(intMode As Integer)
    Public Event GotoRodEnd()
    Public Event GotoOutOfOption()
    Public Event GotoStopper()
    Public Event GotoMotor()       '取付モータ仕様画面
#End Region

    ''' <summary>
    ''' 外部呼出し初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Try
            '全ての情報をクリアする
            subInitInfomations()

            '構成データ取得
            Call subGetCompData()

            '形番構成情報の更新(DB)
            Call DeleteSelectInfo()

            'テキストボックスの作成
            subSetAllTextBox()

            '表示非表示の設定
            Call subSetVisibility()

            '画面をロードする
            Me.OnLoad(Nothing)

        Catch ex As Exception
            AlertMessage(ex)
        End Try
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
        If Not FormIDCheck() Then Exit Sub
        If CInt(HidOptionNumber.Value) <= 0 Then Exit Sub

        Try
            'ページの基本設定
            Call subInitPage()

            '引当情報を取得
            Call GetHikiateInfo()

            '選択したボックスの情報を配列に入れて、更新する
            Call SaveInputData()

            If HidOKClick.Value = "0" Then
                'ページ情報の設定
                Call subSetPage()

                '特殊設定
                Call subSetSpecial()

                'マニホールドテスト専用
                Call ManifoldTest_Youso()
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 画面の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subInitPage()

        'ラベル表示名の設定
        Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHYouso, selLang.SelectedValue, Me) 'Label取得

        'フォントの設定
        Call SetAllFontName(Me)

        '自動Submitボタンを無効
        Call subSetInitScript()

    End Sub

    ''' <summary>
    ''' ページ情報の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetPage()
        Dim dtList As New DataTable                             '候補リスト
        Dim strTitle As String = String.Empty                   '結果タイトル
        Dim intCurrent As Integer                               '現在位置
        Dim strArrayOption() As String = Nothing                '選択したオプション
        Dim strOptionComma As String = String.Empty             '選択したオプション(カンマ区切)
        Dim intStrKbn As Integer                                '複数選択区分

        '現在位置の取得
        If Me.HidCurrentFocus.Value.Trim.Equals(String.Empty) Then
            intCurrent = 1
        Else
            intCurrent = CInt(Me.HidCurrentFocus.Value)
        End If

        '複数選択区分
        intStrKbn = CInt(objKtbnStrc.strcSelection.strOpStructureDiv(intCurrent))

        '前位置によりPlaceLevelの判断とオプションチェック
        If intCurrent > 0 Then
            If Session("yousoCheckResult") IsNot Nothing AndAlso _
                intCurrent = (CInt(HidOptionNumber.Value) - 1) Then
                '最後のオプションがエラーが出る場合は
                Session.Remove("yousoCheckResult")
            Else
                Session.Remove("yousoCheckResult")
                For inti As Integer = 1 To intCurrent
                    Dim intPlaceLevel As Integer

                    'オプションのPlaceLevelを取得
                    intPlaceLevel = funGetPlaceLevels(inti, strArrayOption, strOptionComma)

                    'オプション情報の保存
                    subSaveOptionInfo(inti, strOptionComma, intPlaceLevel)

                    'オプションチェック
                    If Not fncCheckOption(inti, strArrayOption, strOptionComma) Then
                        '次の画面に行くかどうかを記録
                        Session.Add("yousoCheckResult", False)

                        intCurrent = inti
                    End If
                Next
            End If

        End If

        '@@@@現在位置により画面タイトルと候補リストの設定(intCurrent)
        If intCurrent > 0 Then

            Dim intNext As Integer = intCurrent

            'オプションタイトルの取得
            strTitle = fncGetOptionTitle(intCurrent, CType(Me.PnlText.FindControl("txt" & intCurrent), TextBox).Text.Trim.ToUpper)
            '候補リストの取得
            '初期化する場合
            dtList = subListMake(intCurrent)

            If CInt(objKtbnStrc.strcSelection.strOpStructureDiv(intCurrent)) >= 4 Then
                CType(Me.FindControl("txt" & intCurrent), TextBox).Text = String.Empty
            End If

            If Session("yousoCheckResult") IsNot Nothing Then
                'オプションが複数選択できるかどうかの設定
                Call SetHidMultiple()
                '初期処理(フォカスの設定)
                'If intCurrent = CInt(HidOptionNumber.Value) Then
                '    intCurrent = intCurrent - 1
                'End If
                Call subSetListAndFocus(dtList, strTitle, intCurrent)
                Exit Sub
            End If

            While dtList.Rows.Count <= 1
                HidSelRowID.Value = String.Empty
                'オプションの最後まで
                If intNext + 1 < CInt(HidOptionNumber.Value) Then
                    '候補が一つしかいない場合オプションを設定
                    If dtList.Rows.Count = 1 Then
                        CType(Me.FindControl("txt" & intNext), TextBox).Text = dtList.Rows(0)("dispKataban")
                        objKtbnStrc.strcSelection.strOpSymbol(intNext) = dtList.Rows(0)("dispKataban")
                    End If

                    '次のオプションタイトルの取得
                    intNext += 1
                    strTitle = fncGetOptionTitle(intNext, CType(Me.PnlText.FindControl("txt" & intCurrent), TextBox).Text.Trim.ToUpper)

                    '次の候補リストの取得
                    dtList = subListMake(intNext)

                    '次の候補が複数選択できるの場合はクリアする
                    If CInt(objKtbnStrc.strcSelection.strOpStructureDiv(intNext)) >= 4 Then
                        CType(Me.FindControl("txt" & intNext), TextBox).Text = String.Empty
                    Else
                        If dtList.Rows.Count = 0 Then
                            CType(Me.FindControl("txt" & intNext), TextBox).Text = String.Empty
                        Else
                            Dim strNowSelection As String = CType(Me.FindControl("txt" & intNext), TextBox).Text.Trim
                            Dim blnHasFlg As Boolean = False

                            If Not strNowSelection.Equals(String.Empty) Then
                                Dim strAllOption As New List(Of String)
                                Dim strOptions = dtList.AsEnumerable.Select(Function(x) x.Item("dispkataban")).ToList()

                                For Each opt In strOptions
                                    If opt.ToString.Equals(strNowSelection) Then
                                        blnHasFlg = True
                                        Exit For
                                    End If
                                Next
                            End If

                            If blnHasFlg = False Then
                                CType(Me.FindControl("txt" & intNext), TextBox).Text = String.Empty
                            End If
                        End If
                    End If
                Else
                    Exit While
                End If
            End While

            'フォカスの変更
            If Not intNext = intCurrent Then
                Me.HidCurrentFocus.Value = intNext
                intCurrent = intNext
            End If
        End If

        'オプションが複数選択できるかどうかの設定
        Call SetHidMultiple()

        '@@@@初期処理(フォカスの設定)
        Call subSetListAndFocus(dtList, strTitle, intCurrent)
    End Sub

    ''' <summary>
    ''' 特殊設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetSpecial()
        '[X]オプションの設定
        Call SetXOption()

        'その他電圧
        Call SetOtherVoltage()
    End Sub

    ''' <summary>
    ''' フォカスの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetListAndFocus(ByVal dtList As DataTable, ByVal strTitle As String, ByVal intCurrent As Integer)

        If dtList.Rows.Count > 1 Then
            '候補リストの設定
            subSetList(dtList, intCurrent, strTitle)

            'キーダウンイベントの設定
            'Call subSetOptionKeyDown()
            'フォカスの設定
            subSetFocus(intCurrent)

        ElseIf dtList.Rows.Count = 1 Then
            '最後のオプションの場合は
            CType(Me.FindControl("txt" & intCurrent), TextBox).Text = dtList.Rows(0)("dispKataban")
            labelTitle.Text = String.Empty
            GVDetail.DataSource = New DataTable
            GVDetail.DataBind()
            Me.btnOK.Focus()
        Else
            '最後のオプションの場合はリストクリア
            labelTitle.Text = String.Empty
            GVDetail.DataSource = dtList
            GVDetail.DataBind()
            Me.btnOK.Focus()
        End If

    End Sub

    ''' <summary>
    ''' TextBoxの色を設定
    ''' </summary>
    ''' <param name="intCurrent"></param>
    ''' <remarks></remarks>
    Private Sub subSetFocus(ByVal intCurrent As Integer)
        Dim intOptionNumber As Integer = 0

        intOptionNumber = CInt(HidOptionNumber.Value) - 1
        For inti As Integer = 1 To intOptionNumber
            Dim txtTmp As TextBox

            If intCurrent = inti Then
                txtTmp = CType(Me.FindControl("txt" & inti), TextBox)
                txtTmp.Focus()
                txtTmp.BackColor = Drawing.Color.FromArgb(255, 204, 51)
            Else
                txtTmp = CType(Me.FindControl("txt" & inti), TextBox)
                txtTmp.BackColor = Drawing.Color.FromArgb(255, 255, 192)
            End If
        Next
    End Sub

    ''' <summary>
    ''' 候補がある場合リストを作成
    ''' </summary>
    ''' <param name="dtList"></param>
    ''' <param name="intCurrent"></param>
    ''' <param name="strTitle"></param>
    ''' <remarks></remarks>
    Private Sub subSetList(ByVal dtList As DataTable, ByVal intCurrent As Integer, _
                                      ByVal strTitle As String)
        '候補がある場合リストを作成
        If strTitle.Length > 0 Then
            CreatTextCell(dtList, strTitle)
        Else
            CreatTextCell(dtList, objKtbnStrc.strcSelection.strOpKtbnStrcNm(intCurrent))
        End If
    End Sub

    ''' <summary>
    ''' 情報の削除
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DeleteSelectInfo()
        Try
            Dim objConDel As SqlConnection = New SqlClient.SqlConnection(My.Settings.connkhdb)
            objConDel.Open()

            Call subSpecInfoDelete(objConDel)    '仕様書情報削除
            Call subRodInfoDelete(objConDel)     'ロッド先端特注情報削除
            Call subOutofOpInfoDelete(objConDel) 'オプション外特注情報削除
            Call subSetKtbnStrc(objConDel)       'DBに引当情報の保存

            If Not objConDel Is Nothing Then If Not objConDel.State = ConnectionState.Closed Then objConDel.Close()
            objConDel = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 選択した情報をObjectに保存
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SaveInputData()
        Dim strArray As New ArrayList
        strArray = GetInputData(1)

        If strArray.Count > 1 Then
            For inti As Integer = 0 To strArray.Count - 1
                If strArray(inti) Is Nothing Then Continue For
                If inti + 1 >= objKtbnStrc.strcSelection.strOpSymbol.Length Then Exit For
                objKtbnStrc.strcSelection.strOpSymbol(inti + 1) = strArray(inti)
            Next
            Me.Session.Add("KtbnStrc", objKtbnStrc)
        End If

        'If Me.GVDetail.Visible AndAlso Me.GVDetail.Rows.Count > 0 Then
        '    Me.HidGVStartID.Value = Me.GVDetail.Rows(0).ClientID.ToString
        'End If
        'Me.HidDblClick.Value = String.Empty
    End Sub

    ''' <summary>
    ''' キーダウンイベントの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetOptionKeyDown()
        Dim intOptionNumber As Integer = CInt(HidOptionNumber.Value)
        Dim strJS As String = "if (event.keyCode == 13){return false;}else{return true;}"

        For inti As Integer = 1 To intOptionNumber - 1
            Dim txtTmp As TextBox

            txtTmp = CType(PnlText.FindControl("txt" & inti), TextBox)
            txtTmp.Attributes.Add("onKeyDown", "YousoKeyDown(event," & "'" & Me.ClientID & "_', '" & GVDetail.Rows(0).ClientID & "');" & strJS)
        Next
    End Sub

    ''' <summary>
    ''' 初期設定
    ''' </summary>
    ''' <param name="objConDel"></param>
    ''' <remarks></remarks>
    Private Sub subSetKtbnStrc(Optional objConDel As SqlConnection = Nothing)
        Dim intLoopCnt As Integer
        Dim InputFlg As String = String.Empty

        Try
            If objConDel Is Nothing Then objConDel = objCon
            If YousoBLL.fncSelKtbnStrcCheck(objConDel, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId) Then
                InputFlg = "1" '単価画面から遷移してきた場合
            Else
                InputFlg = "0" '機種選択画面から遷移してきた場合
                '引当形番構成追加
                For intLoopCnt = 1 To strcCompData.strElementDiv.Length - 1
                    Call objKtbnStrc.subSelKtbnStrcIns(objConDel, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                                       intLoopCnt, strcCompData.strElementDiv(intLoopCnt), _
                                                       strcCompData.strStructureDiv(intLoopCnt), _
                                                       strcCompData.strAdditionDiv(intLoopCnt), _
                                                       strcCompData.strHyphenDiv(intLoopCnt), _
                                                       strcCompData.strKtbnStrcNm(intLoopCnt), 0)
                Next
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

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

            bolReturn = YousoBLL.fncKatabanStrcSelect(objCon, strcCompData, selLang.SelectedValue) '形番構成取得
            bolReturn = YousoBLL.subKtbnStrcEleSelect(objCon, strcCompData)                        '形番構成要素取得

            '四月バージョンアップ
            Call ClearAddition_Div(objKtbnStrc)

            'オプション個数
            HidOptionNumber.Value = strcCompData.strStructureDiv.Length
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当仕様書情報削除
    ''' </summary>
    ''' <param name="objConDel"></param>
    ''' <remarks></remarks>
    Private Sub subSpecInfoDelete(objConDel As SqlConnection)
        Dim objManifold As New KHManifold(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
        Dim intReturn As Integer
        Try
            '引当仕様書クリア
            intReturn = objManifold.fncSPSelSpecDel(objConDel)
            '引当仕様書構成クリア
            intReturn = objManifold.fncSPSpecStrcDel(objConDel)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当ロッド先端特注情報削除
    ''' </summary>
    ''' <param name="objConDel"></param>
    ''' <remarks></remarks>
    Private Sub subRodInfoDelete(objConDel As SqlConnection)
        Dim objRod As KHRodEndCstm
        Dim dalKtbnStrc As New KtbnStrcDAL
        Dim bolReturn As Boolean

        Try
            'ロッド先端特注クラスインスタンス作成
            objRod = New KHRodEndCstm(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                      strcCompData.strSeriesKataban, strcCompData.strKeyKataban)
            bolReturn = objRod.fncSPSelRodDel(objConDel)
            '引当シリーズ形番更新(オプション情報)
            Call dalKtbnStrc.subSelSrsKtbnOptionUpd(objConDel, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, "")
        Catch ex As Exception
            AlertMessage(ex)
        Finally
            objRod = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 引当オプション外特注情報削除
    ''' </summary>
    ''' <param name="objConDel"></param>
    ''' <remarks></remarks>
    Private Sub subOutofOpInfoDelete(objConDel As SqlConnection)
        Dim obj As KHOutOfOptionCstm
        Dim bolReturn As Boolean
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            'オプション外特注クラスインスタンス作成
            obj = New KHOutOfOptionCstm(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                       selLang.SelectedValue, strcCompData.strSeriesKataban, _
                                       strcCompData.strKeyKataban)
            bolReturn = obj.fncSPSelOutOpDel(objConDel)
            '引当シリーズ形番更新(オプション情報)
            Call dalKtbnStrc.subSelSrsKtbnOptionUpd(objConDel, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, "")
        Catch ex As Exception
            AlertMessage(ex)
        Finally
            obj = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' テキストボックスの作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreatTextBox() As Boolean
        CreatTextBox = False

        Dim intAllLen As Integer = 20
        Dim intTop As Integer = 10
        Dim intFirstLeft As Integer = 0
        Dim intLen As Integer = 0
        Dim intMax As Integer = 0
        Dim objtxt As System.Web.UI.WebControls.TextBox = Nothing
        Dim objlbl As System.Web.UI.WebControls.Label = Nothing


        Dim intOptionNumber As Integer = CInt(HidOptionNumber.Value)
        Try
            '機種設定
            txt0.Text = strcCompData.strSeriesKataban
            txt0.Font.Bold = True
            txt0.Font.Size = myFontSize
            txt0.EnableViewState = True
            txt0.Visible = True
            intAllLen += txt0.Width.Value

            If strcCompData.strHyphen = CdCst.HyphenDiv.Necessary Then
                H0.Text = "－"
                H0.Width = CdCstYouso.intHypenWidth
                H0.Font.Bold = True
                H0.Font.Size = myFontSize
                H0.EnableViewState = True
                H0.Visible = True
                intAllLen += H0.Width.Value
            End If
            intFirstLeft = intAllLen

            For inti As Integer = 1 To intOptionNumber - 1
                'テキストボックスの大きさの調整をする
                intLen = 0
                intMax = 0

                Select Case strcCompData.strElementDiv(inti).ToString
                    Case "1"
                        intLen = CdCstYouso.intVolStrcnt '電圧
                    Case "3"
                        intLen = CdCstYouso.intStrokeStrcnt 'ストローク
                    Case Else
                        If strcCompData.strStructureDiv(inti) >= 4 Then   '複数選択有れば
                            For intj As Integer = inti To (strcCompData.strKtbnStrcEle.Length / 4) - 1
                                If strcCompData.strKtbnStrcEle(intj, 1) = inti.ToString Then
                                    intLen = intLen + strcCompData.strKtbnStrcEle(intj, 2).ToString.Trim.Length
                                ElseIf strcCompData.strKtbnStrcEle(intj, 1).ToString.Length > 0 AndAlso _
                                   CLng(strcCompData.strKtbnStrcEle(intj, 1)) > inti Then
                                    Exit For
                                End If
                            Next
                        Else
                            For intj As Integer = inti To (strcCompData.strKtbnStrcEle.Length / 4) - 1     '最大桁数を検索
                                If strcCompData.strKtbnStrcEle(intj, 1) = inti.ToString Then
                                    If intLen < strcCompData.strKtbnStrcEle(intj, 2).ToString.Trim.Length Then
                                        intLen = strcCompData.strKtbnStrcEle(intj, 2).ToString.Trim.Length
                                    End If
                                ElseIf strcCompData.strKtbnStrcEle(intj, 1).ToString.Length > 0 AndAlso _
                                    CLng(strcCompData.strKtbnStrcEle(intj, 1)) > inti Then
                                    Exit For
                                End If
                            Next
                        End If
                End Select
                intMax = intLen

                If intLen = 0 Then
                    intLen = (intLen + 2) * CdCstYouso.intStrWidth
                Else
                    intLen = (intLen + 1) * CdCstYouso.intStrWidth
                End If
                If intLen > 100 Then intLen = 100

                '一行入れない、二行にします
                If intAllLen + intLen >= PnlText.Width.Value - 40 Then
                    intTop += 50
                    intAllLen = intFirstLeft    '前行の位置と合わせる
                End If

                objtxt = Me.PnlText.FindControl("txt" & inti.ToString)
                objtxt.Width = intLen
                objtxt.TabIndex = inti + 1
                'objtxt.Font.Bold = True
                'objtxt.Font.Size = myFontSize
                intLen = objtxt.Width.Value
                'objtxt.BackColor = DefaultColor
                objtxt.MaxLength = intMax
                objtxt.Text = String.Empty
                objtxt.EnableViewState = True
                objtxt.AutoPostBack = False
                objtxt.Visible = True
                'objtxt.Style.Add("text-transform", "uppercase")
                'objtxt.Attributes.Add("onBlur", "YousoLostFocus('" & Me.ClientID & "_', '" & objtxt.ID & "');")
                objtxt.Attributes.Add("onFocus", "YousoGotFocus('" & Me.ClientID & "_', '" & objtxt.ID & "');")
                'ハイフン設定
                If strcCompData.strHyphenDiv(inti).ToString = "1" Then     'ハイフンあり
                    objlbl = Me.PnlText.FindControl("H" & inti.ToString)
                    objlbl.Text = "－"
                    objlbl.Width = CdCstYouso.intHypenWidth
                    'objlbl.Font.Bold = True
                    'objlbl.Font.Size = myFontSize
                    objlbl.EnableViewState = True
                    objlbl.Visible = True
                    intLen += objlbl.Width.Value
                End If
                intAllLen += intLen
                objtxt = Nothing
                objlbl = Nothing
            Next
            Me.btnOK.TabIndex = strcCompData.strKtbnStrcNm.Length
            CreatTextBox = True
        Catch ex As Exception
            AlertMessage(ex)
        Finally
            intAllLen = Nothing
            intTop = Nothing
            intFirstLeft = Nothing
            intLen = Nothing
            intMax = Nothing
            objtxt = Nothing
            objlbl = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 初期処理(候補リストの作成)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function subInitDetail(ByVal intCurrent As Integer, ByVal intNext As Integer, ByRef strTitle As String) As DataTable

        Dim strArrayOption() As String = Nothing
        Dim strOptionComma As String = String.Empty
        Dim PlaceLvl As Long = 0

        subInitDetail = New DataTable
        Try

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' 選択したオプションのPlaceLevelを取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Function funGetPlaceLevels(ByVal intCurrent As Integer, ByRef strOptions() As String, ByRef strOptionComma As String) As Integer
        Dim strOption As String = String.Empty
        Dim objOption As New KHOptionCtl
        Dim intPlaceLvl As Integer = -1

        '選択オプションの内容を取得
        strOption = CType(Me.PnlText.FindControl("txt" & intCurrent), TextBox).Text.Trim.ToUpper

        'オプションのPlaceLevelを取得
        If CInt(objKtbnStrc.strcSelection.strOpStructureDiv(intCurrent)) >= 4 AndAlso _
            (Not strOption.Trim.Equals(String.Empty)) Then
            '複数選択可能な場合、オプション分解
            strOptions = objOption.fncOptionResolution(objCon, objKtbnStrc, Me.objUserInfo.UserId, _
                                                           Me.objLoginInfo.SessionId, _
                                                           selLang.SelectedValue, _
                                                           intCurrent, strOption)
            '分解したオプションをカンマ区切りの文字列にする
            For intLoopCnt = 1 To strOptions.Length - 1
                If intLoopCnt = 1 Then
                    strOptionComma = strOptions(intLoopCnt)
                    Call YousoBLL.subGetPlacelvl(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                        objKtbnStrc.strcSelection.strKeyKataban, strOptionComma, intCurrent, intPlaceLvl)
                Else
                    strOptionComma = strOptionComma & CdCst.Sign.Comma & strOptions(intLoopCnt)
                    'より小さいデータを取得する
                    Dim PlaceLvl_New As Long = 0
                    Call YousoBLL.subGetPlacelvl(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                        objKtbnStrc.strcSelection.strKeyKataban, strOptions(intLoopCnt), intCurrent, PlaceLvl_New)
                    If intPlaceLvl > PlaceLvl_New Then intPlaceLvl = PlaceLvl_New
                End If
            Next
        Else
            '複数選択不可な場合
            strOptionComma = strOption
            ReDim strOptions(1)
            strOptions(1) = strOptionComma
            Call YousoBLL.subGetPlacelvl(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                objKtbnStrc.strcSelection.strKeyKataban, strOptionComma, intCurrent, intPlaceLvl)
        End If

        Return intPlaceLvl
    End Function

    ''' <summary>
    ''' オプション情報を保存
    ''' </summary>
    ''' <param name="intCurrent"></param>
    ''' <param name="strOptionComma"></param>
    ''' <param name="PlaceLvl"></param>
    ''' <remarks></remarks>
    Private Sub subSaveOptionInfo(ByVal intCurrent As Integer, ByVal strOptionComma As String, _
                                  ByVal PlaceLvl As Integer)
        ''引当形番構成更新処理（DB側、第二スレッドに入れる、非同期更新）
        'strOptionComma_upd = strOptionComma
        'strPlaceLvl_upd = PlaceLvl
        'NewThread_upd.ThreadStart("1")

        '画面のobjKtbnStrcを更新する
        If objKtbnStrc.strcSelection.strOpSymbol.Length > intCurrent Then
            objKtbnStrc.strcSelection.strOpSymbol(intCurrent) = strOptionComma
        End If
        If objKtbnStrc.strcSelection.strOpCountryDiv.Length > intCurrent Then
            objKtbnStrc.strcSelection.strOpCountryDiv(intCurrent) = PlaceLvl
        End If
        Me.Session.Add("KtbnStrc", objKtbnStrc)
    End Sub

    ''' <summary>
    ''' オプションをチェック
    ''' </summary>
    ''' <param name="intCurrent"></param>
    ''' <param name="strArrayOption"></param>
    ''' <param name="strOptionComma"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCheckOption(ByVal intCurrent As Integer, ByVal strArrayOption() As String, _
                                    ByVal strOptionComma As String) As Boolean
        Dim objOption As New KHOptionCtl
        Dim bolReturn As Boolean = True
        Dim strMessageCd As String = Nothing
        Dim txtOption As TextBox = CType(Me.PnlText.FindControl("txt" & intCurrent), TextBox)

        'オプションチェック
        For intLoopCnt = 1 To strArrayOption.Length - 1
            If Len(Trim(strArrayOption(intLoopCnt))) <> 0 Then
                If objOption.fncOptionCheck(objCon, "0", Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                            intCurrent, strArrayOption(intLoopCnt), objKtbnStrc, _
                                            strMessageCd) = False Then
                    bolReturn = False
                    Exit For
                End If
            Else
                bolReturn = True
            End If
        Next

        If Not bolReturn Then
            If intCurrent = CInt(HidOptionNumber.Value) - 1 Then
                Me.HidCurrentFocus.Value = String.Empty
            Else
                Me.HidCurrentFocus.Value = intCurrent
            End If

            'エラーメッセージをセット
            If strMessageCd Is Nothing Then strMessageCd = "W0100"
            txtOption.Focus()
            'subInitDetail = 9
            AlertMessage(strMessageCd)
        Else
            'オプション順序チェック(複数選択の場合)
            If CInt(objKtbnStrc.strcSelection.strOpStructureDiv(intCurrent)) >= 4 Then
                If objOption.fncOptSeqCheck(strOptionComma, txtOption.Text.Trim.ToUpper, strMessageCd) = False Then
                    Me.HidCurrentFocus.Value = intCurrent
                    'エラーメッセージをセット
                    If strMessageCd Is Nothing Then strMessageCd = "W0100"
                    txtOption.Focus()
                    'subInitDetail = 9
                    AlertMessage(strMessageCd)
                    bolReturn = False
                End If
            End If
        End If

        Return bolReturn
    End Function

    ''' <summary>
    ''' オプションタイトルの取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetOptionTitle(ByVal intCurrent As Integer, ByVal strOption As String) As String
        Dim strTitle As String = String.Empty

        'Strock範囲を表示する
        If objKtbnStrc.strcSelection.strOpElementDiv(intCurrent) = CdCst.ElementDiv.Stroke Then
            Dim intBoreSize As Integer = 0
            Dim intMinStrock As Integer = 0
            Dim intMaxStrock As Integer = 0
            Dim intUnitStrock As Integer = 0

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOpElementDiv.Length - 1
                If objKtbnStrc.strcSelection.strOpElementDiv(intLoopCnt) = CdCst.ElementDiv.Port Then
                    Dim strPort As String = objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                    If intCurrent = intLoopCnt Then strPort = strOption
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "CAC"
                            strPort = strPort.Replace("N", "")
                    End Select
                    If IsNumeric(strPort) Then
                        intBoreSize = CInt(strPort)
                        If fncGetStroke(objCon, objKtbnStrc, intBoreSize, intMinStrock, intMaxStrock, intUnitStrock) Then
                            Dim dt_label As DataTable = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHYouso, selLang.SelectedValue)
                            Dim dr() As DataRow = dt_label.Select("label_div='L' and label_seq ='6'")
                            If dr.Length > 0 Then
                                strTitle = dr(0)("label_content").ToString
                                strTitle = strTitle.Replace("[1]", intMinStrock).Replace("[2]", intMaxStrock)
                                strTitle = objKtbnStrc.strcSelection.strOpKtbnStrcNm(intCurrent) & strTitle
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next
        End If

        Return strTitle
    End Function

    ''' <summary>
    ''' 引当画面のリストを作成する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function subListMake(ByVal intCurrent As Integer) As DataTable
        Dim objOption As New KHOptionCtl
        '新しい選択された複数オプション
        Dim strSelVal As String = fncGetCurrentSelectedMulti(intCurrent)
        '結果テーブル
        Dim dt_detail As New DataTable
        Dim strColumnNames As New List(Of String) From {"dispKataban", "dispName"}
        dt_detail = fncCreateTableByColumnNames(strColumnNames)
        'オプション複数選択フラグ
        Dim intMultiFlg As Integer = CInt(objKtbnStrc.strcSelection.strOpStructureDiv(intCurrent))

        Try
            If Me.HidSelRowID.Value.ToString.Equals(String.Empty) Then
                'フォカスをゲットする時

                '複数選択項目の記録を初期化
                'Me.HidSelectedMultiOptions.Value = String.Empty

                'オプションリスト取得
                dt_detail = GetCandidateList(intCurrent)

                '複数選択できる場合は全ての候補をHiddenFieldに記録
                If intMultiFlg >= 4 Then
                    '複数グループの設定
                    Dim CandidateByGroup As ArrayList = GetCandidateGroup(intCurrent)

                    '初期化
                    For inti As Integer = 0 To CandidateByGroup.Count - 1
                        If inti = 0 Then
                            HidAllMultiOptions.Value = CandidateByGroup(inti)
                        Else
                            HidAllMultiOptions.Value &= ";" & CandidateByGroup(inti)
                        End If
                    Next

                    'HidSelectedMultiOptions.Value = String.Empty
                End If
            Else
                'ダブルクリックでロードする時
                'If blnJumpFlg = False Then
                If intMultiFlg < 4 Then
                    Return dt_detail
                End If
            End If

            '    ''項目が複数選択可能な場合
            '    ''HiddenFieldの設定
            '    'SetHiddenField(strSelVal)

            '    ''グループごとに選択候補の取得
            '    'Dim CandidateByGroup As ArrayList = GetCandidateGroup(intCurrent)
            '    ''複数オプション
            '    'Dim strAllMultiOption As String = HidSelectedMultiOptions.Value
            '    'Dim strMultiOptions() As String = strAllMultiOption.Split(",")

            '    ''オプションリスト取得
            '    'dt_detail = GetCandidateList(intCurrent)

            '    ''不要な候補の削除
            '    ''If strMultiOptions.Count = 0 Then
            '    ''    FiltOption(dt_detail, strSelVal, CandidateByGroup)
            '    ''Else
            '    'If strSelVal = String.Empty Then
            '    '    '選択終了の場合
            '    '    dt_detail.Clear()
            '    '    dt_detail.AcceptChanges()
            '    'Else
            '    '    '選択終了以外を選択した場合
            '    '    For Each strMultiOption As String In strMultiOptions
            '    '        FiltOption(dt_detail, strMultiOption, CandidateByGroup)
            '    '    Next
            '    'End If
            '    ''End If
            'End If
            ''    Else
            ''    '遷移する場合は複数選択オプションをクリア
            ''    HidMultiOption.Value = String.Empty
            ''    '次に行くとき
            ''    If intMultiFlg >= 4 Then
            ''        '項目が複数選択可能な場合

            ''        'グループごとに選択候補の取得
            ''        Dim CandidateByGroup As ArrayList = GetCandidateGroup(intCurrent)

            ''        'オプションリスト取得
            ''        dt_detail = GetCandidateList(intCurrent)

            ''        '不要な候補の削除
            ''        FiltOption(dt_detail, strSelVal, CandidateByGroup, blnJumpFlg)

            ''        'HiddenFieldの設定
            ''        SetHiddenField(strSelVal)
            ''    Else
            ''        '項目が複数選択不可の場合
            ''        Dim intOptionNumber As Integer = objKtbnStrc.strcSelection.strOpStructureDiv.Length

            ''        If intOptionNumber.Equals(intCurrent) Then
            ''            '最後のオプションの場合
            ''        Else
            ''            'オプションリスト取得
            ''            dt_detail = GetCandidateList(intCurrent)
            ''        End If
            ''    End If

            ''    CType(Me.FindControl("txt" & intCurrent), TextBox).Text = String.Empty
            ''End If
            'End If

        Catch ex As Exception
            AlertMessage(ex)
        End Try

        Return dt_detail
    End Function

    ''' <summary>
    ''' 新しい選択された複数オプション
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetCurrentSelectedMulti(ByVal intCurrent As Integer) As String
        Dim strResult As String = String.Empty
        Dim strAll As String = CType(Me.PnlText.FindControl("txt" & intCurrent), TextBox).Text
        Dim intLength As Integer = strAll.Length - Me.HidSelectedMultiOptions.Value.ToString.Replace(",", "").ToString.Length

        If intLength > 0 Then strResult = Strings.Right(strAll, intLength)

        Return strResult
    End Function

    ''' <summary>
    ''' オプションリストの取得
    ''' </summary>
    ''' <param name="intFocusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCandidateList(ByVal intFocusNo As Integer) As DataTable
        Dim dt_detail As New DataTable
        Dim dr_detail As DataRow
        Dim strArrayOption(,) As String = Nothing
        Dim objOption As New KHOptionCtl
        'データテーブル初期化
        dt_detail.Columns.Add("dispKataban")
        dt_detail.Columns.Add("dispName")
        'オプションリストの取得
        Call objOption.subOptionList(objCon, objKtbnStrc, "1", Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                             selLang.SelectedValue, intFocusNo, strArrayOption)
        dt_detail.Rows.Clear()

        'データテーブルへの変換
        For inti = 1 To UBound(strArrayOption)
            dr_detail = dt_detail.NewRow
            dr_detail("dispKataban") = strArrayOption(inti, 1).ToString
            dr_detail("dispName") = strArrayOption(inti, 2).ToString
            dt_detail.Rows.Add(dr_detail)
        Next

        Return dt_detail
    End Function

    ''' <summary>
    ''' グループごとの選択候補の取得
    ''' </summary>
    ''' <param name="intFocusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCandidateGroup(ByVal intFocusNo As Integer) As ArrayList
        '複数選択のグループを取得
        Dim CandidateByGroup As New ArrayList
        Dim intConSeqNoBr As New ArrayList
        Dim strConOpSymbol As New ArrayList
        Dim strManyOpt As String = String.Empty

        If YousoBLL.fncElePtnSelect(objCon, objKtbnStrc, intFocusNo, intConSeqNoBr, strConOpSymbol) Then
            For inti As Integer = 0 To intConSeqNoBr.Count - 1
                If inti = 0 Then
                    strManyOpt = strConOpSymbol(0).ToString
                Else
                    If intConSeqNoBr(inti - 1) = intConSeqNoBr(inti) Then
                        strManyOpt = strManyOpt & CdCst.Sign.Comma & strConOpSymbol(inti).ToString
                    Else
                        CandidateByGroup.Add(strManyOpt)
                        strManyOpt = strConOpSymbol(inti).ToString
                    End If
                End If
            Next
            If strManyOpt.Length > 0 Then CandidateByGroup.Add(strManyOpt)
        End If

        Return CandidateByGroup
    End Function

    ''' <summary>
    ''' 選択候補から不要なものを取り除く
    ''' </summary>
    ''' <param name="dt_detail">すべての選択肢候補</param>
    ''' <param name="strSelVal">選択された候補</param>
    ''' <param name="Allval">すべての選択肢候補</param>
    ''' <remarks></remarks>
    Private Sub FiltOption(ByRef dt_detail As DataTable, ByVal strSelVal As String, _
                           ByVal Allval As ArrayList)
        If strSelVal.Trim.Equals(String.Empty) Then
            'If blnJumpFlg Then
            '    '前のオプションから遷移してきた場合はそのまま候補を返す
            'Else
            '    '選択終了を選択した場合
            '    dt_detail.Clear()
            '    dt_detail.AcceptChanges()
            'End If
        Else

            dt_detail.DefaultView.RowFilter = ("")
            dt_detail.DefaultView.ToTable()
            '自分自身と前のデータを削除
            If strSelVal.Trim.Length > 0 Then
                For intL As Integer = dt_detail.Rows.Count - 1 To 0 Step -1
                    If dt_detail.Rows(intL)("dispKataban") = strSelVal Then
                        For inti As Integer = intL - 1 To 0 Step -1
                            If dt_detail.Rows(inti)("dispKataban").ToString.Length > 0 Then
                                dt_detail.Rows(inti).Delete()
                            End If
                        Next
                        dt_detail.AcceptChanges()
                        Exit For
                    End If
                Next
            End If

            '自分自身を削除
            If strSelVal.Trim.Length > 0 Then
                For intL As Integer = dt_detail.Rows.Count - 1 To 0 Step -1
                    If dt_detail.Rows(intL)("dispKataban") = strSelVal Then
                        dt_detail.Rows(intL).Delete()
                        dt_detail.AcceptChanges()
                        Exit For
                    End If
                Next
            End If

            '同じグループのデータを削除
            For inti As Integer = 0 To Allval.Count - 1
                Dim strval() As String = Allval(inti).ToString.Split(CdCst.Sign.Comma)
                For intj As Integer = 0 To strval.Length - 1
                    If strSelVal = strval(intj) Then          'グループを決める
                        For intk As Integer = 0 To strval.Length - 1
                            For intL As Integer = dt_detail.Rows.Count - 1 To 0 Step -1  '同じグループのデータを削除
                                If dt_detail.Rows(intL)("dispKataban") = strval(intk) Then
                                    dt_detail.Rows(intL).Delete()
                                End If
                            Next
                            dt_detail.AcceptChanges()
                        Next
                    End If
                Next
            Next



            '選択終了を削除
            If dt_detail.Rows.Count = 1 And dt_detail.Rows(0)("dispKataban").Equals(String.Empty) Then
                dt_detail.Clear()
                dt_detail.AcceptChanges()
            End If
        End If
    End Sub

    ''' <summary>
    ''' 画面上のテキストボックス値の設定
    ''' </summary>
    ''' <param name="dt_detail"></param>
    ''' <param name="intBlurNo"></param>
    ''' <param name="intFocusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetOptionText(ByVal dt_detail As DataTable, ByVal intBlurNo As Integer, ByVal intFocusNo As Integer) As Boolean
        SetOptionText = True
        Select Case dt_detail.Rows.Count
            Case 0
                If intBlurNo <> intFocusNo Then
                    CType(Me.PnlText.FindControl("txt" & intFocusNo), TextBox).Text = String.Empty
                End If
                SetOptionText = False
            Case 1
                SetOptionText = False
                If dt_detail.Rows(0)("dispKataban").ToString.Length > 0 Then
                    CType(Me.PnlText.FindControl("txt" & intFocusNo), TextBox).Text = dt_detail.Rows(0)("dispKataban").ToString.ToUpper.Trim
                Else
                    If intBlurNo <> intFocusNo Then
                        CType(Me.PnlText.FindControl("txt" & intFocusNo), TextBox).Text = String.Empty
                    End If
                End If
        End Select
    End Function

    ''' <summary>
    ''' テキストバインド
    ''' </summary>
    ''' <param name="dt_detail"></param>
    ''' <param name="strTitle"></param>
    ''' <remarks></remarks>
    Private Sub CreatTextCell(dt_detail As DataTable, strTitle As String)
        Try
            'ラベルタイトル設置
            labelTitle.Text = strTitle
            GVDetail.DataSource = dt_detail
            GVDetail.DataBind()
            Panel5.Visible = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 入力情報の取得
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInputData(Optional intMode As Integer = 0) As ArrayList
        GetInputData = New ArrayList
        Try
            If intMode = 0 Then
                For inti As Integer = 1 To 35
                    If Not Me.PnlText.FindControl("txt" & inti) Is Nothing AndAlso _
                       Me.PnlText.FindControl("txt" & inti).Visible Then
                        'CHANGED BY YGY 20140822
                        '自動にテストの場合
                        'GetInputData.Add(CType(Me.PnlText.FindControl("txt" & inti), TextBox).Text.Trim().ToUpper)
                        GetInputData.Add(CType(Me.PnlText.FindControl("txt" & inti), TextBox).Text.Trim().Replace(CdCst.Sign.Comma, String.Empty).ToUpper)
                    End If
                Next
            Else
                '正常の場合
                Dim strSelOptions() As String = Me.HidSelectedOptions.Value.ToString.ToString.Split("|")
                Dim strHidMultiple() As String = Me.HidMultiplcation.Value.Split(",")

                If strHidMultiple.Count >= strSelOptions.Count Then
                    For inti As Integer = 0 To strSelOptions.Length - 1
                        Dim strSelOption As String = strSelOptions(inti).Trim.ToUpper

                        '複数選択可能な場合、分解する
                        If strHidMultiple.Count > 0 AndAlso strSelOptions.Count > 0 AndAlso _
                            (Not strHidMultiple(inti).Equals(String.Empty)) AndAlso (Not strSelOption.Equals(String.Empty)) Then
                            If CType(strHidMultiple(inti), Integer) >= 4 Then
                                Dim objOption As New KHOptionCtl
                                Dim strArrayOption() As String
                                Dim strComma As String = String.Empty
                                strArrayOption = objOption.fncOptionResolution(objCon, objKtbnStrc, Me.objUserInfo.UserId, _
                                                                                   Me.objLoginInfo.SessionId, _
                                                                                   selLang.SelectedValue, _
                                                                                   inti + 1, strSelOption)
                                For intj As Integer = 1 To strArrayOption.Count - 1
                                    If intj = strArrayOption.Count - 1 Then
                                        strComma &= strArrayOption(intj)
                                    Else
                                        strComma &= strArrayOption(intj) & ","
                                    End If
                                Next
                                GetInputData.Add(strComma.Trim().ToUpper)
                            Else
                                GetInputData.Add(strSelOption)
                            End If
                        Else
                            GetInputData.Add(strSelOption)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' OKボタン押したら
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Dim objOption As New KHOptionCtl
        Dim strArray As New ArrayList
        Dim strArrayOption() As String
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Dim intLoopCnt3 As Integer
        Dim intWrongBox As Integer
        Dim strOptionComma As String
        Dim intKtbnStrcSeqNo As Integer = Nothing
        Dim strOptionSymbol As String = Nothing
        Dim strMessageCd As String = Nothing
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim objUserInfo As KHSessionInfo.UserInfo
        Dim dt_kh_sel_ktbn_strc As New DS_Youso.kh_sel_ktbn_strcDataTable
        Dim da As New DS_YousoTableAdapters.kh_sel_ktbn_strcTableAdapter

        Dim bolOK As Boolean = True

        Try
            'OKボタン押すフラグをリセット
            HidOKClick.Value = "0"

            If pnlAMD0X.Visible Then
                If txtAMD0X.Text.Trim.Length <> 5 Then
                    bolOK = False
                    'エラーメッセージ設定
                    AlertMessage("W8950") '動作区分でX(ミックス)を選択した場合は、数値5桁を入力してください。
                    txtAMD0X.Focus()
                    Exit Sub
                End If
            End If

            '選択したボックスの情報を配列に入れて、更新する
            If Not Me.Session("ManifoldItemKey") Is Nothing Then
                strArray = GetInputData(0)
            Else
                strArray = GetInputData(1)
                'ADD BY YGY 20141112
                '複数オプションの場合チェックする前にカンマを削除すること
                For inti As Integer = 0 To strArray.Count - 1
                    If strArray(inti).ToString.Contains(",") Then
                        strArray(inti) = strArray(inti).ToString.Replace(",", String.Empty)
                    End If
                Next
            End If

            '生産国データを取得する
            Dim dt_Placelvl As DataTable = YousoBLL.fncGetPlacelvl(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                                                                   objKtbnStrc.strcSelection.strKeyKataban)

            '形番引当情報全部取得し、処理完了すると一括更新する
            Dim dr_kh_sel_ktbn_strc() As DataRow = Nothing
            da.Fill(dt_kh_sel_ktbn_strc, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

            Dim strPort As String = String.Empty          '口径 Add by Zxjike 2013/09/11
            Dim dt_AllCountryLevel As DataTable = Nothing 'Add by Zxjike 2013/09/11
            Dim intStrockCount As Integer = 0             'ストローク番号 Add by Zxjike 2013/11/22
            For intLoopCnt1 = 0 To strArray.Count - 1
                If CInt(objKtbnStrc.strcSelection.strOpStructureDiv(intLoopCnt1 + 1)) >= 4 Then
                    If (objKtbnStrc.strcSelection.strSeriesKataban = "ADK21" And intLoopCnt1 = 4) Or _
                       (objKtbnStrc.strcSelection.strSeriesKataban = "APK21" And intLoopCnt1 = 4) Then
                        strOptionComma = strArray(intLoopCnt1)
                    Else
                        '複数選択項目の場合はオプションを分解する
                        'オプション分解
                        strArrayOption = objOption.fncOptionResolution(objCon, objKtbnStrc, Me.objUserInfo.UserId, _
                                         Me.objLoginInfo.SessionId, selLang.SelectedValue, _
                                         intLoopCnt1 + 1, strArray(intLoopCnt1))

                        strOptionComma = String.Empty
                        For intLoopCnt2 = 1 To strArrayOption.Length - 1
                            If intLoopCnt2 = 1 Then
                                strOptionComma = strArrayOption(intLoopCnt2)
                            Else
                                strOptionComma = strOptionComma & CdCst.Sign.Comma & strArrayOption(intLoopCnt2)
                            End If
                        Next
                    End If
                Else
                    If objKtbnStrc.strcSelection.strOpElementDiv(intLoopCnt1 + 1) = "1" Then
                        '特定機種の電圧記号を変更する
                        strOptionComma = fncChangeVlt(objKtbnStrc.strcSelection.strSeriesKataban, _
                                                      objKtbnStrc.strcSelection.strKeyKataban, _
                                                      strArray(intLoopCnt1), intLoopCnt1)
                    Else
                        strOptionComma = strArray(intLoopCnt1)
                    End If
                End If

                '生産国レベルを取得して更新する   Add by Zxjike 2013/09/05
                Dim intPlacelvl As Long = 0
                Dim intPlacelvl_new As Long = 0
                Dim str() As String = strOptionComma.Split(",")
                Dim dr() As DataRow = Nothing

                'RM1808***_生産国レベルロジック変更（複数選択項目時修正）
                Dim bolPlacelvl_1 As Boolean = True
                Dim bolPlacelvl_2 As Boolean = True
                Dim bolPlacelvl_4 As Boolean = True
                Dim bolPlacelvl_8 As Boolean = True
                For inti As Integer = 0 To str.Length - 1
                    If inti = 0 Then
                        dr = dt_Placelvl.Select("ktbn_strc_seq_no='" & intLoopCnt1 + 1 & "' AND option_symbol='" & str(0).ToString & "'")
                        If dr.Length > 0 Then
                            intPlacelvl = CLng(dr(0)("place_lvl").ToString)

                            Dim lstLevel As List(Of Integer) = ClsCommon.fncSeperatePlaceLevel(intPlacelvl)

                            If Not lstLevel.Contains(8) Then
                                bolPlacelvl_8 = False
                            End If
                            If Not lstLevel.Contains(4) Then
                                bolPlacelvl_4 = False
                            End If
                            If Not lstLevel.Contains(2) Then
                                bolPlacelvl_2 = False
                            End If
                            If Not lstLevel.Contains(1) Then
                                bolPlacelvl_1 = False
                            End If

                        Else
                            Dim drOtherVoltage As DataRow() = dt_Placelvl.Select("ktbn_strc_seq_no='" & intLoopCnt1 + 1 & "' AND option_symbol like '" & OtherVoltage.English & "%'")

                            If drOtherVoltage.Length > 0 Then
                                'その他電圧の場合
                                intPlacelvl = CLng(drOtherVoltage(0)("place_lvl").ToString)
                            End If
                        End If
                    Else
                        dr = dt_Placelvl.Select("ktbn_strc_seq_no='" & intLoopCnt1 + 1 & "' AND option_symbol='" & str(inti).ToString & "'")
                        If dr.Length > 0 Then intPlacelvl_new = CLng(dr(0)("place_lvl").ToString)
                        intPlacelvl = 0
                        'If intPlacelvl > intPlacelvl_new Then intPlacelvl = intPlacelvl_new
                        Dim lstLevel As List(Of Integer) = ClsCommon.fncSeperatePlaceLevel(intPlacelvl_new)

                        If Not lstLevel.Contains(8) Then
                            bolPlacelvl_8 = False
                        End If
                        If Not lstLevel.Contains(4) Then
                            bolPlacelvl_4 = False
                        End If
                        If Not lstLevel.Contains(2) Then
                            bolPlacelvl_2 = False
                        End If
                        If Not lstLevel.Contains(1) Then
                            bolPlacelvl_1 = False
                        End If

                    End If
                Next

                '複数選択項目時生産国レベル再セット
                If str.Length > 1 Then
                    If bolPlacelvl_8 = True Then
                        intPlacelvl += 8
                    End If
                    If bolPlacelvl_4 = True Then
                        intPlacelvl += 4
                    End If
                    If bolPlacelvl_2 = True Then
                        intPlacelvl += 2
                    End If
                    If bolPlacelvl_1 = True Then
                        intPlacelvl += 1
                    End If
                End If

                'Stroke範囲判断する Add by Zxjike 2013/09/11
                Select Case objKtbnStrc.strcSelection.strOpElementDiv(intLoopCnt1 + 1)
                    Case CdCst.ElementDiv.Port
                        strPort = strArray(intLoopCnt1).ToString
                        dt_AllCountryLevel = YousoBLL.fncGetAllCountryLevel(objConBase)
                    Case CdCst.ElementDiv.Stroke
                        If strPort.Length > 0 AndAlso IsNumeric(strPort) AndAlso _
                            strOptionComma.Length > 0 AndAlso IsNumeric(strOptionComma) Then

                            'ストロークにより生産国レベルの取得
                            intPlacelvl = fncGetStrokePlaceLevel(intPlacelvl, strOptionComma, strPort, dt_AllCountryLevel)

                        End If
                End Select

                If intLoopCnt1 = 0 Then    'オプション外又はロット先端あれば、日本のみになる 2013/09/09

                    'ロッド先端の場合は
                    If objKtbnStrc.strcSelection.strRodEndOption.Length > 0 Then

                        Dim lstLevel As List(Of Integer) = ClsCommon.fncSeperatePlaceLevel(intPlacelvl)

                        'RM1708046_インドネシア追加の為変更
                        If lstLevel.Contains(8) Then

                            'タイ生産不可時ロッド先端特注指定でタイ生産品表示される不具合修正
                            If lstLevel.Contains(4) Then
                                '日本、タイ、インドネシア
                                intPlacelvl = 13
                            Else
                                '日本、インドネシア
                                intPlacelvl = 9
                            End If

                        ElseIf lstLevel.Contains(4) Then

                            '日本、タイ
                            intPlacelvl = 5

                        Else

                            'それ以外の場合は日本のみ
                            intPlacelvl = 1

                        End If

                    End If

                    'オプション外の場合は日本のみ
                    If objKtbnStrc.strcSelection.strOtherOption.Length > 0 Then

                        'オプション外情報設定を保存しているテーブルを参照し、値を取得するよう変更  2017/04/11
                        Dim dt_Outofop_Placelvl As DataTable = YousoBLL.fncGetOutofopPlacelvl(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

                        '2018/7/26_タイ生産不可の場合、オプション外を選択すると一部でタイ生産品表示される不具合修正
                        Dim lstLevel As List(Of Integer) = ClsCommon.fncSeperatePlaceLevel(intPlacelvl)
 
                        'タイ生産品
                        If lstLevel.Contains(4) Then
                            If dt_Outofop_Placelvl.Rows.Count > 0 Then
                                intPlacelvl = dt_Outofop_Placelvl.Rows(0)("place_lvl")
                            End If
                        Else
                            intPlacelvl = 1
                        End If

                    End If

                End If

                ''RM1808***_メキシコログイン時SCW*2でタイ生産品オプション以外選択時エラー
                'If Me.objUserInfo.CountryCd = "MEX" And intPlacelvl <> 0  Then
                '    Dim lstLevel As List(Of Integer) = ClsCommon.fncSeperatePlaceLevel(intPlacelvl)
                '    If lstLevel.Contains(4) = False Then
                '        bolOK = False
                '        'エラーメッセージ設定
                '        AlertMessage("W9300")
                '        Me.FindControl("txt" & intLoopCnt1 + 1).Focus()
                '        Exit Sub
                '    End If
                'End If

                '引当形番構成更新処理
                'Call objKtbnStrc.subSelKtbnStrcUpd(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                '                                   intLoopCnt1 + 1, strOptionComma, intPlacelvl, objCon)
                dr_kh_sel_ktbn_strc = dt_kh_sel_ktbn_strc.Select("ktbn_strc_seq_no='" & intLoopCnt1 + 1 & "'")
                If Not dr_kh_sel_ktbn_strc Is Nothing AndAlso dr_kh_sel_ktbn_strc.Length > 0 Then
                    dr_kh_sel_ktbn_strc(0)("option_symbol") = strOptionComma
                    dr_kh_sel_ktbn_strc(0)("place_lvl") = intPlacelvl
                End If
                '画面のobjKtbnStrcを更新する
                If objKtbnStrc.strcSelection.strOpSymbol.Length > (intLoopCnt1 + 1) Then
                    objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1 + 1) = strOptionComma
                End If
                If objKtbnStrc.strcSelection.strOpCountryDiv.Length > (intLoopCnt1 + 1) Then
                    objKtbnStrc.strcSelection.strOpCountryDiv(intLoopCnt1 + 1) = intPlacelvl
                End If
            Next
            'Me.Session.Add("KtbnStrc", objKtbnStrc)
            dt_Placelvl = Nothing

            '形番構成要素取得 Add by Zxjike 2013/09/10
            Dim dt_KataStrcEleSel As New DataTable
            dt_KataStrcEleSel = YousoBLL.subKataStrcEleSel(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                                objKtbnStrc.strcSelection.strKeyKataban, CdCst.LanguageCd.DefaultLang)

            '要素パターン取得処理 Add by Zxjike 2013/09/10
            Dim dt_ElePatternSel As New DataTable
            dt_ElePatternSel = YousoBLL.subElePatternSel(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                                objKtbnStrc.strcSelection.strKeyKataban)

            For intLoopCnt1 = 1 To objKtbnStrc.strcSelection.strOpElementDiv.Length - 1
                'オプション分割(複数選択項目用)
                strArrayOption = Split(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1), CdCst.Sign.Delimiter.Comma)

                For intLoopCnt2 = 0 To strArrayOption.Length - 1
                    'オプションチェック
                    If objOption.fncOptionCheck(objCon, "0", Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                                intLoopCnt1, strArrayOption(intLoopCnt2), objKtbnStrc, _
                                                strMessageCd, dt_KataStrcEleSel, dt_ElePatternSel) = False Then
                        intWrongBox = intLoopCnt1
                        If strMessageCd Is Nothing Then strMessageCd = "W0100"
                        bolOK = False
                        'エラーメッセージ設定
                        AlertMessage(strMessageCd)
                        If intLoopCnt1 > 0 Then
                            If Not Me.FindControl("txt" & intLoopCnt1) Is Nothing Then
                                Me.FindControl("txt" & intLoopCnt1).Focus()
                            End If
                        End If
                        Exit Sub
                    End If
                Next

                'オプション順序チェック(複数選択の場合)
                If CInt(objKtbnStrc.strcSelection.strOpStructureDiv(intLoopCnt1)) >= 4 Then
                    If objOption.fncOptSeqCheck(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1), _
                                                strArray(intLoopCnt1 - 1), strMessageCd) = False Then
                        intWrongBox = intLoopCnt1
                        If strMessageCd Is Nothing Then strMessageCd = "W0100"
                        bolOK = False
                        'エラーメッセージ設定
                        AlertMessage(strMessageCd)
                        If intLoopCnt1 > 0 Then
                            If Not Me.FindControl("txt" & intLoopCnt1) Is Nothing Then
                                Me.FindControl("txt" & intLoopCnt1).Focus()
                            End If
                        End If
                        Exit Sub
                    End If
                End If
            Next
            dt_KataStrcEleSel = Nothing 'Add by Zxjike 2013/09/10
            dt_ElePatternSel = Nothing 'Add by Zxjike 2013/09/10

            '特殊対応(付加区分にてオプション記号を強制的に変更する)
            For intLoopCnt1 = 1 To objKtbnStrc.strcSelection.strOpAdditionDiv.Length - 1
                If objKtbnStrc.strcSelection.strOpAdditionDiv(intLoopCnt1).Trim >= "2" Then
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "AB21" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "00B" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Then
                                ''引当形番構成更新処理
                                'Call objKtbnStrc.subSelKtbnStrcUpd(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                '                                   3, "0", objKtbnStrc.strcSelection.strOpCountryDiv(intLoopCnt1))
                                '画面のobjKtbnStrcを更新する
                                objKtbnStrc.strcSelection.strOpSymbol(3) = "0"
                                dr_kh_sel_ktbn_strc = dt_kh_sel_ktbn_strc.Select("ktbn_strc_seq_no='3'")
                                If Not dr_kh_sel_ktbn_strc Is Nothing AndAlso dr_kh_sel_ktbn_strc.Length > 0 Then
                                    dr_kh_sel_ktbn_strc(0)("option_symbol") = "0"
                                End If
                            End If
                        End If
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1).Trim <> "" Then
                            For intLoopCnt2 = intLoopCnt1 - 1 To 1 Step -1
                                If objKtbnStrc.strcSelection.strOpAdditionDiv(intLoopCnt2).Trim >= "1" And _
                                   objKtbnStrc.strcSelection.strOpAdditionDiv(intLoopCnt2).Trim < objKtbnStrc.strcSelection.strOpAdditionDiv(intLoopCnt1).Trim Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt2).Trim = "" Then
                                        Dim intMaxLength As Integer = 0
                                        Dim strAryOption(,) As String = Nothing

                                        'オプションリスト取得
                                        Call objOption.subOptionList(objCon, objKtbnStrc, "1", _
                                                                     Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                                                     selLang.SelectedValue, intLoopCnt2, strAryOption)

                                        For intLoopCnt3 = 1 To UBound(strAryOption)
                                            If Len(strAryOption(intLoopCnt3, 1)) > intMaxLength Then
                                                intMaxLength = Len(strAryOption(intLoopCnt3, 1))
                                            End If
                                        Next

                                        ''引当形番構成更新処理
                                        'Call objKtbnStrc.subSelKtbnStrcUpd(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                        '                                   intLoopCnt2, _
                                        '                                   objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt2).PadRight(intMaxLength, "0"), _
                                        '                                   objKtbnStrc.strcSelection.strOpCountryDiv(intLoopCnt2))
                                        dr_kh_sel_ktbn_strc = dt_kh_sel_ktbn_strc.Select("ktbn_strc_seq_no='" & intLoopCnt2 & "'")
                                        If Not dr_kh_sel_ktbn_strc Is Nothing AndAlso dr_kh_sel_ktbn_strc.Length > 0 Then
                                            dr_kh_sel_ktbn_strc(0)("option_symbol") = objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt2).PadRight(intMaxLength, "0")
                                        End If
                                        '画面のobjKtbnStrcを更新する
                                        objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt2) = objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt2).PadRight(intMaxLength, "0")
                                    Else
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            Next

            da.Update(dt_kh_sel_ktbn_strc) '一括更新
            da = Nothing
            dt_kh_sel_ktbn_strc = Nothing
            Me.Session.Add("KtbnStrc", objKtbnStrc)

            'OKボタン時のオプションチェック
            If objOption.fncOtherOptionCheck(objKtbnStrc, intKtbnStrcSeqNo, objKtbnStrc.strcSelection.strRodEndOption, _
                                             strMessageCd) = False Then
                bolOK = False
                'エラーメッセージ設定
                AlertMessage(strMessageCd)
                Me.FindControl("txt" & intKtbnStrcSeqNo).Focus()
                Exit Sub
            End If

            '引当シリーズ形番情報更新
            Call objKtbnStrc.subFullKatabanCreate(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, Me.txtAMD0X.Text.Trim.ToUpper)

            '単価表示回数をカウントアップしてセッションに保持させる
            objUserInfo = httpCon.Session(CdCst.SessionInfo.Key.UserInfo)
            objUserInfo.TnkDispCnt = objUserInfo.TnkDispCnt + 1
            httpCon.Session(CdCst.SessionInfo.Key.UserInfo) = objUserInfo

            Dim intMode As Integer = YousoBLL.GetNextFormMode(objKtbnStrc, objOption)
            If bolOK Then RaiseEvent BtnOKGo(intMode) 'ページ遷移
            GC.Collect()
        Catch ex As Exception
            'エラー画面に遷移する
            AlertMessage(ex)
        Finally
            objOption = Nothing
            objKtbnStrc = Nothing
            strArray = Nothing
            strArrayOption = Nothing
            da = Nothing
            dt_kh_sel_ktbn_strc = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' ロッド先端特注
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '画面入力項目チェック
        If fncRodEndOpenCheck() Then
            Call SaveKatahiki()
            RaiseEvent GotoRodEnd()
        End If
    End Sub

    ''' <summary>
    ''' オプション外
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '画面入力項目チェック
        If fncOptionOpenCheck() Then
            Call SaveKatahiki()
            RaiseEvent GotoOutOfOption()
        End If
    End Sub

    ''' <summary>
    ''' ストッパ位置
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call SaveKatahiki()
        RaiseEvent GotoStopper()
    End Sub

    ''' <summary>
    ''' 取付モータ仕様
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        RaiseEvent GotoMotor()
    End Sub

    ''' <summary>
    ''' 取付モータ仕様(ETS)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RaiseEvent GotoMotor()
    End Sub

    ''' <summary>
    ''' ポート位置(IAVB)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        RaiseEvent GotoMotor()
    End Sub

    'RM1804032_画像表示追加
    ''' <summary>
    ''' 手配本数一覧(EKS)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Session("LabelClick7") = True
        RaiseEvent GotoMotor()
    End Sub

    ''' <summary>
    ''' DBに情報を保存
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SaveKatahiki()
        Dim dt_kh_sel_ktbn_strc As New DS_Youso.kh_sel_ktbn_strcDataTable
        Dim da As New DS_YousoTableAdapters.kh_sel_ktbn_strcTableAdapter
        Try
            '形番引当情報全部取得し、処理完了すると一括更新する
            Dim dr_kh_sel_ktbn_strc() As DataRow = Nothing
            da.Fill(dt_kh_sel_ktbn_strc, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
            objKtbnStrc = Me.Session("KtbnStrc")
            For inti As Integer = 0 To objKtbnStrc.strcSelection.strOpSymbol.Length - 1
                dr_kh_sel_ktbn_strc = dt_kh_sel_ktbn_strc.Select("ktbn_strc_seq_no='" & inti + 1 & "'")
                If Not dr_kh_sel_ktbn_strc Is Nothing AndAlso dr_kh_sel_ktbn_strc.Length > 0 Then
                    dr_kh_sel_ktbn_strc(0)("option_symbol") = objKtbnStrc.strcSelection.strOpSymbol(inti + 1)
                    dr_kh_sel_ktbn_strc(0)("place_lvl") = objKtbnStrc.strcSelection.strOpCountryDiv(inti + 1)
                End If
            Next
            da.Update(dt_kh_sel_ktbn_strc) '一括更新
        Catch ex As Exception
            AlertMessage(ex)
        Finally
            da = Nothing
            dt_kh_sel_ktbn_strc = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' ロッド先端画面を表示するか否かをチェックする
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncRodEndOpenCheck() As Boolean
        fncRodEndOpenCheck = False

        'オプションチェックが失敗する場合
        If Session("yousoCheckResult") IsNot Nothing Then
            Session.Remove("yousoCheckResult")
            Exit Function
        End If

        objKtbnStrc = Me.Session("KtbnStrc")
        If objKtbnStrc Is Nothing Then Exit Function
        Try
            '口径取得
            Dim strBoreSize As String = String.Empty
            strBoreSize = YousoBLL.subBoreSizeSelect(objCon, objKtbnStrc)

            '口径チェック
            If strBoreSize.Trim.Length = 0 Then
                AlertMessage("W8530")
                Exit Function
            End If

            Dim strOpSymbol() As String = objKtbnStrc.strcSelection.strOpSymbol
            '機種毎にチェック
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "SSD"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case CdCst.Sign.Blank
                            If InStr(strOpSymbol(1), "B") <> 0 Or InStr(strOpSymbol(1), "Y") <> 0 Or _
                               ((InStr(strOpSymbol(1), "G") <> 0 Or InStr(strOpSymbol(1), "G2") <> 0 Or _
                                InStr(strOpSymbol(1), "G3") <> 0) And (InStr(strOpSymbol(1), "G1") = 0) And _
                                (InStr(strOpSymbol(1), "G4") = 0) And (InStr(strOpSymbol(1), "G5") = 0)) Then
                                AlertMessage("W8540")
                                Exit Function
                            End If
                            If strOpSymbol(20).Trim = "FA" Or strOpSymbol(20).Trim = "LB2" Then
                                AlertMessage("W8540")
                            ElseIf InStr(strOpSymbol(19), "N") <> 0 Then
                                AlertMessage("W8550")
                                Exit Function
                            End If
                        Case "K"
                            If InStr(strOpSymbol(1), "B") <> 0 Or InStr(strOpSymbol(1), "Y") <> 0 Or _
                               ((InStr(strOpSymbol(1), "G") <> 0 Or InStr(strOpSymbol(1), "G2") <> 0 Or _
                                InStr(strOpSymbol(1), "G3") <> 0) And (InStr(strOpSymbol(1), "G1") = 0) And _
                                (InStr(strOpSymbol(1), "G4") = 0) And (InStr(strOpSymbol(1), "G5") = 0)) Then
                                AlertMessage("W8540")
                                Exit Function
                            End If
                            If strOpSymbol(18).Trim = "FA" Or strOpSymbol(18).Trim = "LB2" Then
                                AlertMessage("W8540")
                                Exit Function
                            ElseIf InStr(strOpSymbol(17).Trim, "N") <> 0 Then
                                AlertMessage("W8550")
                                Exit Function
                            End If
                        Case "D"
                            If InStr(strOpSymbol(11), "N") <> 0 Then
                                AlertMessage("W8550")
                                Exit Function
                            End If
                    End Select
                Case "SCA2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "B", "C"
                            If InStr(strOpSymbol(1), "B") <> 0 Then
                                AlertMessage("W0360")
                                Exit Function
                            End If
                    End Select
                Case "JSC3"

                Case "SCS", "SCS2"
                    If InStr(strOpSymbol(1), "B") <> 0 Then
                        AlertMessage("W0360")
                        Exit Function
                    Else
                        If InStr(strOpSymbol(18), "IY") <> 0 Then
                            AlertMessage("W8560")
                            Exit Function
                        End If
                    End If
                Case "CMK2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case CdCst.Sign.Blank, "5"
                            If InStr(strOpSymbol(1), "B") <> 0 Or InStr(strOpSymbol(1), "SR") <> 0 Or _
                               ((InStr(strOpSymbol(1), "G") <> 0 Or InStr(strOpSymbol(1), "G2") <> 0 Or _
                                InStr(strOpSymbol(1), "G3") <> 0) And (InStr(strOpSymbol(1), "G1") = 0) And _
                                (InStr(strOpSymbol(1), "G4") = 0)) Then
                                ' バリエーション「*B*」「*SR*」の場合は選択不可
                                ' バリエーション「G」「G2」「G3」を含む場合は選択不可(「G1」,「G4」はOK)
                                AlertMessage("W8660")
                                Exit Function
                            Else ' オプション「J」,「L」選択の場合は選択不可
                                If InStr(strOpSymbol(15), "J") <> 0 Or InStr(strOpSymbol(15), "L") <> 0 Then
                                    AlertMessage("W8660")
                                    Exit Function
                                End If
                            End If
                        Case "D", "E" ' オプション「J」,「L」選択の場合は選択不可
                            If InStr(strOpSymbol(10), "J") <> 0 Or InStr(strOpSymbol(10), "L") <> 0 Then
                                AlertMessage("W8660")
                                Exit Function
                            Else
                                If InStr(strOpSymbol(11), "IY") <> 0 Then ' 付属品「IY」選択時は選択不可
                                    AlertMessage("W8560")
                                    Exit Function
                                End If
                            End If
                    End Select
            End Select
            fncRodEndOpenCheck = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    '''  画面を表示するか否かをチェックする
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncOptionOpenCheck() As Boolean
        fncOptionOpenCheck = False

        'オプションチェックが失敗する場合
        If Session("yousoCheckResult") IsNot Nothing Then
            Session.Remove("yousoCheckResult")
            Exit Function
        End If

        objKtbnStrc = Me.Session("KtbnStrc")
        If objKtbnStrc Is Nothing Then Exit Function
        Try
            '口径取得
            Dim strBoreSize As String = String.Empty
            strBoreSize = YousoBLL.subBoreSizeSelect(objCon, objKtbnStrc)

            '口径チェック
            If strBoreSize.Trim.Length = 0 Then
                AlertMessage("W0860")
                Exit Function
            End If
            fncOptionOpenCheck = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' バインドイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GVDetail_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVDetail.RowDataBound
        Try
            If e.Row.RowIndex < 0 Then
                Exit Sub
            End If

            Dim strName As String = Me.ClientID & "_"
            Dim intStartID As Integer = 0
            Dim intRowCount As Integer = CType(GVDetail.DataSource, DataTable).Rows.Count

            If e.Row.RowIndex = 0 Then
                intStartID = CInt(Strings.Right(e.Row.ClientID, 2))
            Else
                intStartID = CInt(Strings.Right(GVDetail.Rows(0).ClientID, 2))
            End If

            e.Row.TabIndex = e.Row.RowIndex + 36
            e.Row.Attributes.Add(CdCst.JavaScript.OnClick, "YousoGridClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "','" & e.Row.RowIndex & "');")


            '次へボタンの作成
            Dim imgButton As New ImageButton
            imgButton.ImageUrl = "~/KHImage/arrow.png"
            imgButton.Width = "15"

            'CHANGED BY YGY 20140611    複数選択できる場合は最後の要素候補を選択したら次の要素に飛ばせるため　　　↓↓↓↓↓↓
            If HidCurrentFocus.Value Is Nothing OrElse HidCurrentFocus.Value.Equals(String.Empty) Then
                imgButton.OnClientClick = "ArrowClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "');"

                e.Row.Attributes.Add(CdCst.JavaScript.OnDblClick, "YousoDblClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "');")
                e.Row.Attributes.Add(CdCst.JavaScript.OnKeyDown, "fncGrid_OnKeyup(event, '" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',1);")
            Else
                '複数選択できる要素に最後の候補以外が選択された場合

                imgButton.OnClientClick = "ArrowClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "','" & e.Row.RowIndex & "');"

                e.Row.Attributes.Add(CdCst.JavaScript.OnDblClick, "YousoDblClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "','" & e.Row.RowIndex & "');")
                e.Row.Attributes.Add(CdCst.JavaScript.OnKeyDown, "fncGrid_OnKeyup(event, '" & strName & "','" & e.Row.ClientID & "','" & intStartID & "', 1 ,'" & e.Row.RowIndex & "');")
            End If
            'CHANGED BY YGY 20140611    ↑↑↑↑↑↑

            e.Row.Cells(2).Controls.Add(imgButton)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' JavaScript生成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInitScript()
        ''TextBoxのEnterキーイベント
        Dim strName As String = Me.ClientID & "_"
        Dim intOptionNumber As Integer = CInt(HidOptionNumber.Value)
        Dim strJS As String = "if (event.keyCode == 13){return false;}else{return true;}"

        For inti As Integer = 1 To intOptionNumber - 1
            Dim txtTmp As TextBox

            txtTmp = CType(PnlText.FindControl("txt" & inti), TextBox)
            txtTmp.Attributes.Add(CdCst.JavaScript.OnKeyDown, "YousoKeyDown(event," & "'" & strName & "', '2','" & inti & "');" & strJS)
        Next

        'GridViewのEnterキーを無効にする
        GVDetail.Attributes.Add(CdCst.JavaScript.OnKeyDown, strJS)
        Me.btnOK.UseSubmitBehavior = False
        Me.Button2.UseSubmitBehavior = False
        Me.Button3.UseSubmitBehavior = False
        Me.Button4.UseSubmitBehavior = False
        Me.Button5.UseSubmitBehavior = False
        Me.Button6.UseSubmitBehavior = False
        Me.Button7.UseSubmitBehavior = False
    End Sub

    ''' <summary>
    ''' マニホールドテスト時
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ManifoldTest_Youso()
        If Not Me.Session("ManifoldItemKey") Is Nothing Then
            If Me.Session("TestFlag") Is Nothing Then
                If Not Me.Session("ManifoldSeriesKey") Is Nothing Then
                    Dim strSeries() As String = Me.Session("ManifoldSeriesKey").split(",")
                    If strSeries(2).Length > 0 Then 'GAMD0
                        pnlAMD0X.Visible = True
                        Me.txtAMD0X.Text = strSeries(2).ToString
                    End If
                End If

                Dim str() As String = Me.Session("ManifoldItemKey")
                For inti As Integer = 0 To str.Length - 1
                    If Not Me.PnlText.FindControl("txt" & (inti + 1)) Is Nothing Then
                        If str(inti) Is Nothing Then str(inti) = String.Empty
                        CType(Me.PnlText.FindControl("txt" & (inti + 1)), TextBox).Text = str(inti).ToString.Trim
                    End If
                Next
                Me.Session("TestFlag") = True
                Call btnOK_Click(Me, Nothing)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 単価画面から戻る時に、ある項目の「0、00」を削除する（画面で選択できず、単価画面のフル形番生成する時に追加したもの）
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <remarks></remarks>
    Private Sub ClearAddition_Div(ByVal objKtbnStrc As KHKtbnStrc)
        Dim intMax As Integer = 0
        For inti As Integer = 1 To objKtbnStrc.strcSelection.strOpAdditionDiv.Length - 1
            If objKtbnStrc.strcSelection.strOpAdditionDiv(inti).ToString.Length > 0 AndAlso _
                CLng(objKtbnStrc.strcSelection.strOpAdditionDiv(inti)) > intMax Then
                intMax = CLng(objKtbnStrc.strcSelection.strOpAdditionDiv(inti))
            End If
        Next
        If intMax > 0 Then   '0追加あるかもしれない
            Dim strAddition As String = String.Empty
            For inti As Integer = 1 To objKtbnStrc.strcSelection.strOpAdditionDiv.Length - 1
                strAddition = objKtbnStrc.strcSelection.strOpAdditionDiv(inti).ToString
                If strAddition.Length > 0 AndAlso CLng(strAddition) < intMax AndAlso CLng(strAddition) > 0 Then
                    If IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(inti)) AndAlso _
                        CLng(objKtbnStrc.strcSelection.strOpSymbol(inti)) = 0 Then
                        '自分自身＝0且つ自分よりレベル高い項目は空白でわない時に、0を削除する
                        Dim bolExit As Boolean = False
                        For intj As Integer = 1 To objKtbnStrc.strcSelection.strOpAdditionDiv.Length - 1
                            If objKtbnStrc.strcSelection.strOpAdditionDiv(intj).ToString.Length > 0 AndAlso _
                                CLng(objKtbnStrc.strcSelection.strOpAdditionDiv(intj).ToString) > CLng(strAddition) Then
                                If objKtbnStrc.strcSelection.strOpSymbol(intj).ToString.Length > 0 Then
                                    bolExit = True
                                    Exit For
                                End If
                            End If
                        Next
                        If bolExit Then
                            objKtbnStrc.strcSelection.strOpSymbol(inti) = String.Empty
                            '引当形番構成更新処理
                            Call KtbnStrcDAL.subSelKtbnStrcUpd(Me.objUserInfo.UserId, _
                                                               Me.objLoginInfo.SessionId, _
                                                               inti, "", objKtbnStrc.strcSelection.strOpCountryDiv(inti))
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 全ての情報をクリアする
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subInitInfomations()
        'HiddenFiledのクリア
        Me.HidCurrentFocus.Value = 1
        Me.HidSelRowID.Value = String.Empty
        Me.HidSelectedOptions.Value = String.Empty
        Me.HidOptionNumber.Value = String.Empty
        Me.HidMultiplcation.Value = String.Empty
        Me.HidAllMultiOptions.Value = String.Empty
        Me.HidSelectedMultiOptions.Value = String.Empty

        'テキストボックスのクリア
        subClearTexts()
        Me.txtAMD0X.Text = String.Empty

        'テーブルの初期化
        GVDetail.DataSource = New DataTable
        GVDetail.DataBind()
        GVDetail.Font.Name = GetFontName(selLang.SelectedValue)

        'セッションの削除
        Me.Session.Remove("KtbnStrc")

        '前回の選択情報を削除
        'NewThread.ThreadStart("strConn")

    End Sub

    ''' <summary>
    ''' 画面の表示非表示を設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetVisibility()
        '名称設定
        Me.lblKataban.Text = strcCompData.strGoodsNm
        Me.labelTitle.Font.Name = GetFontName(selLang.SelectedValue)
        Me.lblKataban.Font.Name = GetFontName(selLang.SelectedValue)
        Me.Label10.Visible = False 'RM1609018 2016/09/14 K.Ohwaki　Append　

        'ラベル表示設定
        Select Case strcCompData.strSeriesKataban
            Case "SCA2", "SCG", "SCG-D", "SCG-G", "SCG-G2", "SCG-G3", _
                 "SCG-G4", "SCG-M", "SCG-O", "SCG-Q", "SCG-U"
                Me.Label1.Visible = True
                Me.Label2.Visible = True
                Me.Label3.Visible = True
                Me.Label5.Visible = False
                Me.Label9.Visible = False
            Case "SSD", "SSD2", "CMK2", "SCM", "STS-B", "STS-M", "STL-B", "STL-M", "JSC3", "SCA2"
                Me.Label1.Visible = True
                Me.Label2.Visible = True
                Me.Label3.Visible = False
                Me.Label5.Visible = False
                Me.Label9.Visible = False
            Case "JSC4"
                If strcCompData.strKeyKataban = "2" Then
                    Me.Label1.Visible = True
                    Me.Label2.Visible = True
                    Me.Label3.Visible = False
                    Me.Label5.Visible = False
                    Me.Label9.Visible = False
                Else
                    Me.Label1.Visible = False
                    Me.Label2.Visible = False
                    Me.Label3.Visible = False
                    Me.Label4.Visible = False
                    Me.Label5.Visible = False
                    Me.Label9.Visible = False
                End If
            Case "SMD2", "SMD2-L", "SMD2-XL", "SMD2-YL", "SMD2-X", "SMD2-Y", "SMD2-M", "SMD2-ML"
                Me.Label1.Visible = False
                Me.Label2.Visible = False
                Me.Label3.Visible = False
                Me.Label4.Visible = True
                Me.Label5.Visible = False
                Me.Label9.Visible = False
            Case "SCS"
                If strcCompData.strKeyKataban <> "4" Then
                    Me.Label1.Visible = True
                    Me.Label2.Visible = True
                    Me.Label3.Visible = False
                    Me.Label4.Visible = False
                    Me.Label5.Visible = True
                    Me.Label9.Visible = False
                Else
                    Me.Label1.Visible = True
                    Me.Label2.Visible = True
                    Me.Label3.Visible = False
                    Me.Label4.Visible = False
                    Me.Label5.Visible = True
                    Me.Label9.Visible = False
                End If
            Case "SNP", "V3301"
                Me.Label1.Visible = False
                Me.Label2.Visible = False
                Me.Label3.Visible = False
                Me.Label4.Visible = False
                Me.Label5.Visible = False
                Me.Label9.Visible = True
            Case "PPD3" 'RM1609018 2016/09/14 K.Ohwaki　Append　Start
                Me.Label1.Visible = False
                Me.Label2.Visible = False
                Me.Label3.Visible = False
                Me.Label4.Visible = False
                Me.Label5.Visible = False
                Me.Label9.Visible = False
                Me.Label10.Visible = True
                'RM1609018 2016/09/14 K.Ohwaki　Append　End
            Case Else
                Me.Label1.Visible = False
                Me.Label2.Visible = False
                Me.Label3.Visible = False
                Me.Label4.Visible = False
                Me.Label5.Visible = False
                Me.Label9.Visible = False
        End Select

        Button1.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        '取付モータ仕様
        Button4.Visible = False
        Button5.Visible = False
        Button6.Visible = False
        Button7.Visible = False

        Select Case strcCompData.strSeriesKataban
            Case "SSD"
                Select Case strcCompData.strKeyKataban
                    Case "L", "4", "P", "E", "R", "S"
                    Case Else
                        Button1.Visible = True
                End Select
            Case "SCA2"
                Button1.Visible = True
                Button2.Visible = True
            Case "SCS2"
                Select Case strcCompData.strKeyKataban
                    Case "4"
                    Case Else
                        Button1.Visible = True
                        Button2.Visible = True
                End Select
            Case "JSC4"
                Select Case strcCompData.strKeyKataban
                    Case "2"
                        Button1.Visible = True
                        Button2.Visible = True
                    Case Else
                End Select
            Case "JSC3"
                Select Case strcCompData.strKeyKataban
                    Case "R", "S"
                    Case Else
                        Button1.Visible = True
                        Button2.Visible = True
                End Select
            Case "CMK2"
                Select Case strcCompData.strKeyKataban
                    Case "4"
                    Case Else
                        Button1.Visible = True
                End Select
            Case "LCG", "LCG-Q", "LCR", "LCR-Q"
                Button3.Visible = True
            Case "ETV", "ECS", "ECV", "ESM"
                '取付モータ仕様
                Button4.Visible = True
            Case "ETS", "EBS", "EBR"    'RM1803042_EBS,EBR追加
                Select Case objKtbnStrc.strcSelection.strKeyKataban
                    Case "A", "B", "C", "D"
                        Button5.Visible = True
                    Case Else
                        '取付モータ仕様
                        Button4.Visible = True
                End Select

            Case "SCS"
                Select Case strcCompData.strKeyKataban
                    Case "2"
                    Case Else
                        Button1.Visible = True
                        Button2.Visible = True
                End Select
                'ポート位置　RM1610026
            Case "IAVB"
                Button6.Visible = True
                'RM1804032_画像表示追加
            Case "EKS"
                Button7.Visible = True
                Button4.Visible = True
        End Select
    End Sub

    ''' <summary>
    ''' テキストボックスの作成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetAllTextBox()
        '各オプションを非表示に
        For inti As Integer = PnlText.Controls.Count - 1 To 0 Step -1
            PnlText.Controls(inti).Visible = False
        Next

        'テキストボックスのTab順
        For inti As Integer = 1 To 35
            If Not Me.PnlText.FindControl("txt" & inti) Is Nothing Then
                CType(Me.PnlText.FindControl("txt" & inti), TextBox).TabIndex = inti
            End If
        Next
        Panel5.Visible = False

        '選択欄の生成
        Call CreatTextBox()

    End Sub

    ''' <summary>
    ''' TextBoxのクリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subClearTexts()
        For inti As Integer = 1 To 35
            If Not Me.PnlText.FindControl("txt" & inti) Is Nothing Then
                CType(Me.PnlText.FindControl("txt" & inti), TextBox).Text = String.Empty
                CType(Me.PnlText.FindControl("txt" & inti), TextBox).BackColor = DefaultColor
            End If
        Next
    End Sub

    ''' <summary>
    ''' 引当情報を取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GetHikiateInfo()
        If Not Me.Session("KtbnStrc") Is Nothing Then
            objKtbnStrc = Me.Session("KtbnStrc")
        Else
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
            Do While (objKtbnStrc.strcSelection.strOpStructureDiv.Length) <> CLng(HidOptionNumber.Value)
                Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
            Loop
            Me.Session.Add("KtbnStrc", objKtbnStrc)
        End If
    End Sub

    ''' <summary>
    ''' オプションが複数選択できるかどうかの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetHidMultiple()
        Dim strMultiple As String = String.Empty
        If Not objKtbnStrc.strcSelection.strOpStructureDiv Is Nothing Then
            For inti As Integer = 1 To objKtbnStrc.strcSelection.strOpStructureDiv.Length - 1
                If Not objKtbnStrc.strcSelection.strOpStructureDiv(inti) Is Nothing Then
                    strMultiple &= objKtbnStrc.strcSelection.strOpStructureDiv(inti).ToString & ","
                Else
                    strMultiple &= ","
                End If
            Next
        End If
        Me.HidMultiplcation.Value = strMultiple
    End Sub

    ''' <summary>
    ''' Xオプションの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetXOption()
        Select Case objKtbnStrc.strcSelection.strSeriesKataban
            Case "GAMD0"
                If Not Me.PnlText.FindControl("txt1") Is Nothing Then
                    If CType(Me.PnlText.FindControl("txt1"), TextBox).Text.Trim.ToUpper = "X" Then
                        Dim strUserClass As String = Me.objUserInfo.UserClass.ToString.Trim
                        Dim intUserClass As Integer
                        '整数に変換
                        If Not Integer.TryParse(strUserClass, intUserClass) Then
                            intUserClass = -1
                        End If

                        If intUserClass = 21 OrElse intUserClass >= 45 Then
                            '国内営業の場合
                            pnlAMD0X.Visible = True
                        Else
                            Me.txtAMD0X.Text = String.Empty
                            pnlAMD0X.Visible = False
                        End If
                    Else
                        Me.txtAMD0X.Text = String.Empty
                        pnlAMD0X.Visible = False
                    End If
                End If
        End Select
    End Sub

    ''' <summary>
    ''' その他電圧の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetOtherVoltage()
        Me.lblOtherVol.Visible = False
        If Me.HidCurrentFocus.Value.Length > 0 AndAlso objKtbnStrc.strcSelection.strOpElementDiv.Length > Me.HidCurrentFocus.Value Then
            'その他電圧の場合
            If objKtbnStrc.strcSelection.strOpElementDiv(Me.HidCurrentFocus.Value) = CdCst.ElementDiv.Voltage Then
                'その他電圧は表示しないように
                'Me.lblOtherVol.Visible = True
                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    Case "PKW", "PKS", "PKA", "PVS", "PDVE4"
                        Me.lblOtherVol.Text = "その他電圧についてはお問い合わせください"
                    Case Else
                        Me.lblOtherVol.Text = "その他電圧"
                End Select
            End If
        End If
    End Sub

    ''' <summary>
    ''' ストロークにより生産国レベルの取得
    ''' </summary>
    ''' <param name="intPlacelvl">
    ''' kh_kataban_strc_eleテーブルに登録されたplace_lvl
    ''' </param>
    ''' <param name="strOptionComma">画面に入力したストロークオプション</param>
    ''' <param name="strPort">画面に入力したポート</param>
    ''' <param name="dt_AllCountryLevel">全ての国コードに対応する生産レベル</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetStrokePlaceLevel(ByVal intPlacelvl As Integer, _
                                            ByVal strOptionComma As String, _
                                            ByVal strPort As String, _
                                            ByVal dt_AllCountryLevel As DataTable) As Integer
        Dim intResult As Integer = 0

        If intPlacelvl = 0 Then
            '中間ストロークの場合は
            'ストローク範囲により生産国レベルを取得

            Dim lstCountryOfStroke As New List(Of String)

            '生産国ごとのストローク範囲の取得
            Dim dt_strockCountry As DataTable = YousoBLL.fncGetStrokeCountry(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                                            objKtbnStrc.strcSelection.strKeyKataban, strPort)

            For Each dr As DataRow In dt_strockCountry.Rows
                '国コード
                Dim strCountryCd As String = dr.Item("country_cd").ToString
                '国レベルの取得
                Dim intCountryLvl As Integer = fncGetCountryPlaceLevel(dt_AllCountryLevel, strCountryCd)
                If intCountryLvl < 0 Then Continue For

                'ストローク範囲
                Dim drStroke As DataRow = dt_strockCountry.Select("country_cd='" & strCountryCd & "'")(0)

                If drStroke.Item("min_stroke") > 0 OrElse drStroke.Item("max_stroke") > 0 Then
                    'ストローク範囲がある場合
                    Dim drTemp() As DataRow

                    '選択されたストロークが範囲内にあるかどうか
                    drTemp = dt_strockCountry.Select("country_cd='" & strCountryCd & "' AND min_stroke <='" & _
                            strOptionComma & "' AND max_stroke>='" & strOptionComma & "'")

                    If drTemp.Count > 0 Then
                        '特殊的なストロークロジック
                        subSetPlaceLevel(strPort, strOptionComma, strCountryCd, intCountryLvl, intResult)
                    End If
                Else
                    'ストローク範囲がない場合
                    '標準ストロークかを判断
                    Dim dt_stdStroke As DataTable = YousoBLL.fncGetStdStroke(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                                        objKtbnStrc.strcSelection.strKeyKataban, strPort, strOptionComma)

                    If Not dt_stdStroke Is Nothing AndAlso dt_stdStroke.Rows.Count > 0 Then
                        '標準ストロークの場合
                        '特殊的なストロークロジック
                        subSetPlaceLevel(strPort, strOptionComma, strCountryCd, intCountryLvl, intResult)
                    End If
                End If
            Next
        Else
            '標準ストロークかの判断
            Dim dt_stdStroke As DataTable = YousoBLL.fncGetStdStroke(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                    objKtbnStrc.strcSelection.strKeyKataban, strPort, strOptionComma)

            If dt_stdStroke IsNot Nothing AndAlso dt_stdStroke.Rows.Count > 0 Then
                '標準ストロークの場合はkh_kataban_strc_eleに登録された生産レベルを追加

                'PlcaeLevelにより取得した生産国レベル
                Dim lstCountryOfPlaceLevel As New List(Of String)

                '生産国ごとのストローク範囲の取得
                Dim dt_strockCountry As DataTable = YousoBLL.fncGetStrokeCountry(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                                                objKtbnStrc.strcSelection.strKeyKataban, strPort)

                '生産国レベルの分解
                lstCountryOfPlaceLevel = KHCountry.fncGetStroke_Logic(intPlacelvl).Split(",").ToList

                For Each strPlaceLevel As String In lstCountryOfPlaceLevel
                    Dim intCountryLvl As Integer = CInt(strPlaceLevel)
                    Dim strCountryCd As String = String.Empty

                    '国コード
                    strCountryCd = dt_AllCountryLevel.Select("place_lvl='" & intCountryLvl & "'")(0).Item("place_div").ToString()

                    'ストローク生産可能範囲
                    Dim drStrokes() As DataRow = dt_strockCountry.Select("country_cd='" & strCountryCd & "'")

                    If drStrokes.Count > 0 Then

                        Dim drStroke As DataRow = drStrokes(0)

                        If drStroke.Item("max_stroke") = 0 Then
                            '生産範囲Max_Stroke=0の場合は標準ストロークなら生産可能

                            '特殊的なストロークロジック
                            subSetPlaceLevel(strPort, strOptionComma, strCountryCd, intCountryLvl, intResult)
                        Else
                            '生産範囲Max_Stroke<>0の場合は範囲により判断
                            Dim drTemp() As DataRow

                            drTemp = dt_strockCountry.Select("country_cd='" & strCountryCd & "' AND min_stroke <='" & _
                                                             strOptionComma & "' AND max_stroke>='" & strOptionComma & "'")

                            If drTemp.Count > 0 Then
                                '特殊的なストロークロジック
                                subSetPlaceLevel(strPort, strOptionComma, strCountryCd, intCountryLvl, intResult)
                            End If

                        End If
                    End If
                Next
            Else
                '標準ストローク以外の場合はkh_strokeにより生産可能レベルを追加

                '生産国ごとのストローク範囲の取得
                Dim dt_strockCountry As DataTable = YousoBLL.fncGetStrokeCountry(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                                                objKtbnStrc.strcSelection.strKeyKataban, strPort)

                For Each dr As DataRow In dt_strockCountry.Rows
                    '国コード
                    Dim strCountryCd As String = dr.Item("country_cd").ToString
                    '国レベルの取得
                    Dim intCountryLvl As Integer = fncGetCountryPlaceLevel(dt_AllCountryLevel, strCountryCd)
                    If intCountryLvl < 0 Then Continue For

                    'ストローク生産可能範囲
                    Dim drStrokes() As DataRow = dt_strockCountry.Select("country_cd='" & strCountryCd & "'")

                    If drStrokes.Count > 0 Then

                        Dim drStroke As DataRow = drStrokes(0)

                        If drStroke.Item("min_stroke") > 0 OrElse drStroke.Item("max_stroke") > 0 Then
                            'ストローク範囲がある場合
                            Dim drTemp() As DataRow

                            '選択されたストロークが範囲内にあるかどうか
                            drTemp = dt_strockCountry.Select("country_cd='" & strCountryCd & "' AND min_stroke <='" & _
                                    strOptionComma & "' AND max_stroke>='" & strOptionComma & "'")

                            If drTemp.Count > 0 Then
                                '特殊的なストロークロジック
                                subSetPlaceLevel(strPort, strOptionComma, strCountryCd, intCountryLvl, intResult)
                            End If
                        End If
                    End If
                Next

            End If

        End If

        Return intResult

    End Function

    ''' <summary>
    ''' 国コードに対応する生産レベルの取得
    ''' </summary>
    ''' <param name="dt_AllCountryLevel"></param>
    ''' <param name="strCountryCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetCountryPlaceLevel(ByVal dt_AllCountryLevel As DataTable, _
                                             ByVal strCountryCd As String) As Integer
        '国レベル
        Dim intResult As Integer
        Dim drCountryLevel() As DataRow

        drCountryLevel = dt_AllCountryLevel.Select("place_div='" & strCountryCd & "'")
        If drCountryLevel.Count = 0 Then
            intResult = -1
        Else
            intResult = drCountryLevel(0).Item("place_lvl")
        End If

        Return intResult

    End Function

    ''' <summary>
    ''' 特殊的なストロークロジック
    ''' </summary>
    ''' <param name="strPort">口径</param>
    ''' <param name="strOptionComma">ストローク</param>
    ''' <param name="strCountryCd">追加したい国コード</param>
    ''' <param name="intCountryLvl">追加したい国の生産レベル</param>
    ''' <param name="intResult">生産レベル</param>
    ''' <remarks></remarks>
    Private Sub subSetPlaceLevel(ByVal strPort As String, _
                                        ByVal strOptionComma As String, _
                                        ByVal strCountryCd As String, _
                                        ByVal intCountryLvl As Integer, _
                                        ByRef intResult As Integer)

        If strCountryCd.Equals("PRC") Then
            '中国の場合は特殊的なストロークロジックで判断
            If YousoBLL.fncGetStroke_Logic(objKtbnStrc, strPort, strOptionComma) Then
                intResult += intCountryLvl
            End If
        Else
            intResult += intCountryLvl
        End If
    End Sub
End Class
