Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports System.Net

Public Class WebUC_ISOTanka
    Inherits KHBase

#Region "プロパティ"
    '価格コピー画面へ
    Public Event GotoCopyPrice()
    '価格詳細画面へ
    Public Event GotoPriceDetail()
    Public Event SiyouFileOutput(objKtbnStrc As KHKtbnStrc, strSiyou As String)
    Public Event IFFileOutput(objKtbnStrc As KHKtbnStrc, strName As String, strNewPlace As String)
    Public Event FileOutput(objKtbnStrc As KHKtbnStrc, strName As String, strOrder As String, strPriceList As String, intMode As Integer)
    'EDIに戻るイベント
    Public Event EDIReturn()

    Private CST_COMMA As String = CdCst.Sign.Comma
    Private CST_PIPE As String = CdCst.Sign.Delimiter.Pipe
    Private CST_BLANK As String = ""
    Private strDispDiv() As String            'チェック区分/出荷場所表示可否情報
    Private EditDiv As String                 '小数点区分
    Private EditDivOpt As String              '小数点区分オプション("," or ".")
    Private strPriceList(,) As String         '単価リスト
    Private bllTanka As New TankaBLL          'ビジネスロジック
    Private dt_Addinfo As DataTable = Nothing



#End Region

#Region "イベント"

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()

        '画面をクリア
        subClearPage()

        'Javascriptの初期化
        subSetInitScript()

        '共通項目の設定
        subSetCommonItems()

        '画面ロード
        Me.OnLoad(Nothing)
    End Sub

    ''' <summary>
    ''' 画面をクリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subClearPage()
        HidPriceForFile.Value = String.Empty
        txt_AmtPrice.Text = String.Empty
        txt_AmtTax.Text = String.Empty
        txt_SumTotal.Text = String.Empty
        txt_PrpAmt.Text = String.Empty
        txt_ELPrd.Text = String.Empty

        '言語項目の初期化
        selLang = Me.Parent.Parent.Parent.Parent.FindControl("ContentTitle").FindControl("selLang")

        '価格詳細画面
        Me.HidPriceDetail.Value = String.Empty

        'セッションクリア
        Me.Session("TestFlag") = Nothing                                'テストフラグ
        Me.Session("possiblePlace") = Nothing                          '生産可能場所
        Me.Session("ShipPlaces") = Nothing                             '全ての出荷場所
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
        Dim strPlace As String = String.Empty
        '受注EDI
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current

        Try
            If Me.HidShiftD.Value = "2" Then
                '価格積上げ画面
                Me.HidShiftD.Value = "1"

                RaiseEvent GotoCopyPrice()
            ElseIf Me.HidPriceDetail.Value = "2" Then
                '価格詳細画面へ遷移
                Me.HidPriceDetail.Value = "1"

                '出荷場所の保存
                subSetPlace()

                RaiseEvent GotoPriceDetail()
            Else
                '各価格リストの設定
                subSetEachList()

            End If

            'インドユーザー注意メッセージ
            If (objKtbnStrc.strcSelection.strFullKataban IsNot Nothing) AndAlso _
                (Not objKtbnStrc.strcSelection.strFullKataban.Equals(String.Empty)) Then

                Select Case Me.objUserInfo.CountryCd

                    Case "IND"

                        If KHKataban.fncSpecialUserMessage(objCon, objKtbnStrc.strcSelection.strFullKataban.Split("-")(0), Me.objUserInfo.CountryCd, "IND") Then
                            Me.pnlIndMessage.Visible = True
                        Else
                            Me.pnlIndMessage.Visible = False
                        End If

                    Case "E90", "EUR"

                        If KHKataban.fncSpecialUserMessage(objCon, objKtbnStrc.strcSelection.strFullKataban.Split("-")(0), Me.objUserInfo.CountryCd, "E90") Then
                            Me.pnlEurMessage.Visible = True
                        Else
                            Me.pnlEurMessage.Visible = False
                        End If

                    Case "MEX"      'RM1707049_2017/7/26_CZ対応

                        If KHKataban.fncSpecialUserMessage(objCon, objKtbnStrc.strcSelection.strFullKataban.Split("-")(0), Me.objUserInfo.CountryCd, "MEX") Then
                            Me.pnlIndMessage.Visible = True
                        Else
                            Me.pnlIndMessage.Visible = False
                        End If

                End Select
                
            End If

            'RM1707049_2017/7/26_CZ対応
            'RM170****_2017/8/24_メキシコメッセージ対応
            Select Case objUserInfo.CountryCd
                Case "IND"
                    Label32.Text = Label32.Text.Replace("[1]", "CKD India CZ17101")     'インド
                Case "MEX"
                    Label32.Text = Label32.Text.Replace("[1]", "New price from September 1, 2017 (CZ17102)")         'メキシコ
            End Select

            strPlace = Me.cmbPlace.SelectedItem.Value
            Select Case strPlace
                Case "JPN", "1001", "1002", "1003", "1004", "1005"
                    strPlace = "JPN"
            End Select

            '特価決裁Noメッセージ表示（購入価格を表示するユーザーのみメッセージ表示）
            If strPlace = "JPN" And objUserInfo.PriceDispLvl > 63 Then
                Label41.Visible = True
                Label41.Text = objKtbnStrc.strcSelection.strAuthorizationNo
            Else
                Label41.Visible = False
                Label41.Text = Nothing
            End If

            '受注EDIボタン
            If strPlace = "JPN" And httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) IsNot Nothing Then
                Button5.Visible = True
            Else
                Button5.Visible = False
            End If

            'マニホールドテスト専用
            Call subManifoldTest()

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 仕様出力
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '画面に表示されたフル形番により再設定
        If Not lblSeriesKat.Text.Trim.Equals(objKtbnStrc.strcSelection.strFullKataban) Then
            objKtbnStrc.strcSelection.strFullKataban = lblSeriesKat.Text.Trim
        End If
        RaiseEvent SiyouFileOutput(objKtbnStrc, String.Empty)
    End Sub

    ''' <summary>
    ''' I/F
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        RaiseEvent IFFileOutput(objKtbnStrc, Me.ClientID, String.Empty)
    End Sub

    ''' <summary>
    ''' 受注EDI登録
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        'Dim CommonDbService As New CommonDbService.CommonDbServiceClient
        Dim objEdiInfo As KHSessionInfo.EdiInfo
        Dim objKataban As New KHKataban
        'Dim result As WebKataban.CommonDbService.DbProcessResult
        Dim clsKHSBOInerfaceResult As New KHSBOInterface
        Dim strFobPrice As String = 0
        Dim strSessionIDFob As String = String.Empty
        Dim strCurrencyCode As String = String.Empty
        Dim blCZFlag As Boolean = False

        Try

            '海外生産品の場合は受注EDI連携不可
            Select Case cmbPlace.SelectedValue
                Case "1001", "1002", "1003", "1004", "1005", "JPN"
                Case Else
                    AlertMessage("W0960") '海外生産品の場合は使用できません。
                    Exit Sub
            End Select

            If httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) IsNot Nothing Then

            Else
                AlertMessage("E9999") 'システムエラー
                Exit Sub
            End If

            'セッション情報をオブジェクト変数にセット
            objEdiInfo = httpCon.Session(CdCst.SessionInfo.Key.EdiInfo)

            'セッション名の取得
            Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                Case "05", "06"
                    'ISOの場合
                    strSessionIDFob = "strPriceListFobISO"
                Case Else
                    strSessionIDFob = "strPriceListFob"
            End Select

            'セッション情報の取得(購入価格）
            If Not (Session(strSessionIDFob) Is Nothing) Then
                strFobPrice = Session(strSessionIDFob)
                Session.Remove(strSessionIDFob)
                strCurrencyCode = Session("strCurrencyCode")
            Else
                strFobPrice = 0
            End If

            'CZ特価フラグ
            If Label41.Text <> Nothing Then
                blCZFlag = True
            Else
                blCZFlag = False
            End If

            '連携データクラスにデータセット
            clsKHSBOInerfaceResult = KHSBOInterface.fncJutyuEdiInterfaceGet(objCon, objKtbnStrc, Me.objUserInfo.OfficeCd, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, strFobPrice, strCurrencyCode, objEdiInfo.KeyInfo, blCZFlag, intItemRow.Value)

            '受注EDI情報送信Webサービスを実行する。
            'CommonDbService = New CommonDbService.CommonDbServiceClient
            'result = CommonDbService.AddKatahikiInfoIso(clsKHSBOInerfaceResult.clKatahikiInfoDtoIso)

            RaiseEvent EDIReturn() '送信した後、引当システムを閉じる（ログオフ）
        Catch ex As Exception
            AlertMessage(ex) 'エラー画面に遷移する
        End Try

    End Sub

    ''' <summary>
    ''' ファイル出力
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'ファイル出力
        RaiseEvent FileOutput(objKtbnStrc, Me.ClientID, Me.HidPriceForFile.Value, Me.HidPriceList.Value, 2)

    End Sub

#End Region

#Region "メッソド"

    ''' <summary>
    ''' 共通項目の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetCommonItems()

        Dim dtCompData As New DataTable

        '価格計算
        Call subGetUnitPrice()

        'シリーズと形番の設定
        Call subInitScreen()

        'ボタンの表示設定
        Call ShowButton()

        'データ取得
        dtCompData = TankaISOBLL.fncSQL_GetCompData(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

        '出荷場所の設定
        Call subSetShipPlace(dtCompData)

        '営業本部、情報システム部ユーザーのみ価格積上げ表示画面を表示する
        If Me.objUserInfo.UserClass >= CdCst.UserClass.BizHeadquarters Then
            Me.HidShiftD.Value = "1"
        End If

    End Sub

    ''' <summary>
    ''' 各価格リストの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetEachList()
        Dim dtCompData As New DataTable

        '引当情報取得
        Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)

        'データ取得
        dtCompData = TankaISOBLL.fncSQL_GetCompData(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

        'ラベルタイトル設置
        Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHUnitPrice, selLang.SelectedValue, Me)
        'ラベルタイトル設置
        Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHISOTanka, selLang.SelectedValue, Me)

        '権限により共通項目の表示設定
        Call subDispSet()

        '生産品ラベル制御
        If cmbPlace.Items.Count > 1 Then
            'ドロップダウンボックス選択変更時に生産品ラベルも更新(初期化する時に実行しない)
            Call SetPlaceMark(cmbPlace.SelectedItem.Value, cmbPlace.Items.Count)
        End If

        '各価格リストの作成
        Call subMakeDetail(dtCompData)

    End Sub

    ''' <summary>
    ''' 出荷場所の表示設定
    ''' </summary>
    ''' <param name="strCountry"></param>
    ''' <param name="lstPlaceIDCount"></param>
    ''' <remarks></remarks>
    Private Sub SetPlaceMark(ByVal strCountry As String, ByVal lstPlaceIDCount As Integer)
        Dim lstCountry_Key As New ArrayList
        Dim bolMaybe As Boolean = False

        If Not Me.Session("possiblePlace") Is Nothing Then lstCountry_Key = Me.Session("possiblePlace")

        '生産可能場所が選択された場合
        If lstCountry_Key.Contains(cmbPlace.SelectedValue) Then bolMaybe = True

        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False
        Label19.Visible = False

        Select Case strCountry
            Case "THA"
                If bolMaybe Then
                    Label19.Visible = True
                    Label19.Width = Unit.Percentage(80)
                    Label19.BackColor = Drawing.Color.LemonChiffon
                    Label19.ForeColor = Drawing.Color.HotPink
                    Me.Label17.Visible = True
                Else
                    Label15.Visible = True
                    Label15.Width = Unit.Percentage(80)
                    Label15.BackColor = Drawing.Color.Yellow
                    Label15.ForeColor = Drawing.Color.Red
                End If
            Case "JPN"
                Label15.Visible = True
                Label15.Width = Unit.Percentage(80)
                Label15.BackColor = Drawing.Color.White
                Label15.ForeColor = Drawing.Color.Red
        End Select

        '複数生産拠点
        If lstPlaceIDCount > 1 Then
            Label16.Visible = True
        End If

        '出荷場所名
        subSetShipPlaceName()
    End Sub

    ''' <summary>
    ''' 出荷場所の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetShipPlace(ByVal dtCompData As DataTable)
        Dim dt_place As New DataTable
        Dim strAsean As List(Of String) = CdCst.strAseanCode
        Dim blnAsean As Boolean = False
        Dim dtStrageEvaluation As New DataTable

        'ログインユーザーがASEANであるかどうか
        If strAsean.Contains(Me.objUserInfo.CountryCd) Then
            'ASEANの場合
            blnAsean = True
        Else
            'ASEAN以外の場合
            blnAsean = False
        End If

        '出荷場所リストの作成
        dt_place = fncSetPlaceCdList(dtCompData, blnAsean, dtStrageEvaluation)

        Me.cmbPlace.Items.Clear()
        Me.cmbPlace.DataTextField = "PlaceName"
        Me.cmbPlace.DataValueField = "PlaceID"
        Me.cmbPlace.DataSource = dt_place
        Me.cmbPlace.DataBind()

        'Me.cmbStrageEvaluation.Items.Clear()
        'Me.cmbStrageEvaluation.DataTextField = "StrageEvaluationName"
        'Me.cmbStrageEvaluation.DataValueField = "StrageEvaluationID"
        'Me.cmbStrageEvaluation.DataSource = dt_place
        'Me.cmbStrageEvaluation.DataBind()

        If dt_place.Rows.Count > 1 Then
            PnlProductionPlace.Visible = blnAsean
        Else
            PnlProductionPlace.Visible = False
        End If

    End Sub

    ''' <summary>
    ''' 出荷場所名の取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetShipPlaceName()
        '全ての国コードと国名
        Dim strPlace As String = String.Empty
        Dim drTmp() As DataRow
        Dim dt_country As DataTable = KHCountry.fncGetAllCountryName(objConBase)

        strPlace = Me.cmbPlace.SelectedItem.Value

        '対応する出荷場所名を取得
        drTmp = dt_country.Select("country_cd='" & strPlace & "' AND language_cd='" & Me.selLang.SelectedValue & "'")

        If drTmp IsNot Nothing AndAlso drTmp.Count > 0 Then
            Label15.Text = Label15.Text.Replace("[1]", drTmp(0)("country_nm").ToString)
            Label19.Text = Label19.Text.Replace("[1]", drTmp(0)("country_nm").ToString)
            Label16.Text = Label16.Text.Replace("[1]", drTmp(0)("country_nm").ToString)
            Label17.Text = Label17.Text.Replace("[1]", drTmp(0)("country_nm").ToString)
        End If
    End Sub

    ''' <summary>
    ''' 出荷場所リストの作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSetPlaceCdList(ByVal dtCompData As DataTable, _
                                       ByVal blnAsean As Boolean, _
                                       ByRef dtStrageEvaluation As DataTable) As DataTable
        Dim dt_country As New DataTable                               '国マスタ
        Dim dtResult As New DataTable                                 '処理結果
        Dim lstCountry_Key As New ArrayList                           '出荷場所リスト
        Dim lstPossiblePlace As New ArrayList                         '生産可能場所

        '結果テーブル初期化
        dtResult = fncCreateTableByColumnNames(New List(Of String) From {"PlaceName", "PlaceID"})
        dtStrageEvaluation = fncCreateTableByColumnNames(New List(Of String) From {"StrageEvaluationName", "StrageEvaluationID"})

        '全ての国コードと国名
        dt_country = KHCountry.fncGetCountryName(objConBase)

        '一番目の部品の出荷場所を追加
        If dtCompData.Rows.Count > 0 Then
            Dim strPlaceCd As String = dtCompData.Rows(0).Item("place_cd")

            lstCountry_Key.Add(strPlaceCd)
        End If

        'ASEANユーザーの場合は生産可能場所を追加
        If blnAsean Then
            'フル形番の第1ハイフンにより生産可能場所の追加
            lstPossiblePlace = KHCountry.fncCountryKeyGet(objConBase, objKtbnStrc.strcSelection.strFullKataban)

            If lstPossiblePlace.Count > 0 Then
                lstCountry_Key.AddRange(lstPossiblePlace)
                'セッションに保存
                Session.Add("possiblePlace", lstPossiblePlace)
            End If
        End If

        '表示順番の調整
        lstCountry_Key = fncSetOrder(lstCountry_Key)

        '生産場所名の取得
        For Each strplace As String In lstCountry_Key
            Dim drPlace As DataRow
            Dim drTmp() As DataRow

            drPlace = dtResult.NewRow

            '日本国コードの変換
            'subChangeToJapanesePlaceCd(strplace)

            '対応する出荷場所名を取得
            If selLang.SelectedValue = "ja" Then
                drTmp = dt_country.Select("country_cd='" & strplace & "' AND language_cd='ja'")
            Else
                drTmp = dt_country.Select("country_cd='" & strplace & "' AND language_cd='en'")
            End If

            If drTmp.Length > 0 Then
                drPlace("PlaceID") = strplace
                drPlace("PlaceName") = "CKD " & drTmp(0)("country_nm").ToString
            Else
                drPlace("PlaceName") = strplace
                subChangeToJapaneseCountryCd(strplace)
                drPlace("PlaceID") = strplace
            End If

            dtResult.Rows.Add(drPlace)
        Next

        Return dtResult

    End Function

    ''' <summary>
    ''' 出荷場所表示順番の調整
    ''' </summary>
    ''' <param name="lstCountry_Key"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSetOrder(ByVal lstCountry_Key As ArrayList) As ArrayList
        Dim lstResult As New ArrayList
        Dim lstCountriesByCountryCd As New ArrayList     'ユーザの国コードにより表示可能な「国コード」

        'ユーザの国コードにより表示可能な「国コード」を取得
        lstCountriesByCountryCd = KHCountry.fncCountryTradeGet(objConBase, objUserInfo.CountryCd)

        For Each strCountryKey As String In lstCountriesByCountryCd
            For Each strDispCountry As String In lstCountry_Key
                Dim strPlaceID As String = strDispCountry

                '日本コードの変換
                subChangeToJapaneseCountryCd(strPlaceID)

                If strPlaceID.Equals(strCountryKey) Then
                    lstResult.Add(strDispCountry)
                End If
            Next
        Next

        Return lstResult
    End Function

    ''' <summary>
    ''' 日本の出荷場所変換
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subChangeToJapaneseCountryCd(ByRef strPlaceID As String)
        Select Case strPlaceID
            Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
                strPlaceID = "JPN"
        End Select
    End Sub

    ''' <summary>
    ''' ボタン表示設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShowButton()
        Dim strUseFncInfo() As String = Nothing
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        Try
            '権限にI/Fボタンの表示を制限する
            'Call KHKataban.subUseFncInfoGet(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
            '                                   selLang.SelectedValue, Me.objUserInfo.UseFunctionLvl, _
            '                                   strUseFncInfo, objKtbnStrc)
            '仕様出力ボタン表示
            Button2.Visible = True

            'If strUseFncInfo(2) Then
            '    Button3.Visible = True
            'Else
            Button3.Visible = False
            'End If

            '受注EDI T.Y セッションが有効であれば[EDI]ボタンを表示し、[ファイル出力]ボタンを非表示する。
            If httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) IsNot Nothing Then
                  Select cmbPlace.SelectedValue
                    Case "1002", "1003", "1004", "1005"
                        Button5.Visible = True
                    Case Else
                        Button5.Visible = False
                End Select

                Button6.Visible = False
            Else
                Button6.Visible = True
                Button5.Visible = False
            End If

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 画面設定
    ''' </summary>
    ''' <remarks>Series_Kataban=M（M4SA1など）のときに形番の最後に仕様書情報を付加し、標準納期を計算するように改修</remarks>
    Private Sub subGetUnitPrice()
        Dim objUnitPrice As New KHUnitPrice
        Try
            '単価取得
            Call objUnitPrice.subPriceInfoSet(objCon, objKtbnStrc, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                              Me.objUserInfo.CountryCd, Me.objUserInfo.OfficeCd, String.Empty, String.Empty)
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)
            If Me.objUserInfo.EditDiv = CdCst.EditDiv.Normal Then
                Me.txt_EditNormal.Text = CdCst.EditDiv.Normal
            Else
                Me.txt_EditNormal.Text = CdCst.EditDiv.Other
            End If
        Catch ex As Exception
            AlertMessage(ex)
        Finally
            objUnitPrice = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subInitScreen()
        Try
            '名称設定
            lblSeriesNm.Text = objKtbnStrc.strcSelection.strGoodsNm
            lblSeriesKat.Text = objKtbnStrc.strcSelection.strFullKataban

            Select Case Me.objUserInfo.UserClass
                Case CdCst.UserClass.DmAgentRs, CdCst.UserClass.DmAgentSs, CdCst.UserClass.DmAgentBs, _
                    CdCst.UserClass.DmAgentGs, CdCst.UserClass.DmAgentPs
                    '国内代理店はメッセージを表示させる
                    'Me.imgFixedMessage1.Visible = True
                    Panel2.Visible = True
                Case Else
                    'Me.imgFixedMessage1.Visible = False
                    Panel2.Visible = False
            End Select
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 初期画面設定
    ''' </summary>
    ''' <remarks>消費税(txt_AmtTax)を追加 </remarks>
    Private Sub subDispSet()
        Dim objKataban As New KHKataban
        Dim objKtbnStrc As New KHKtbnStrc
        Dim objUnitPrice As New KHUnitPrice
        Dim intLoopCnt As Integer
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current

        Try
            '引当情報取得
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)

            '小数点区分(EditDiv)取得
            EditDiv = TankaISOBLL.fncDecPointDivSelect(objConBase, Me.objUserInfo.UserId)
            If EditDiv = "0" Then
                EditDivOpt = CdCst.Sign.Dot
            Else
                EditDivOpt = CdCst.Sign.Comma
            End If

            'EditDivを変更する
            Me.txt_AmtPrice.EditDiv = EditDiv
            Me.txt_SumTotal.EditDiv = EditDiv
            Me.txt_AmtTax.EditDiv = EditDiv


            '受注EDI セッションが有効のときの処理
            If httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) IsNot Nothing Then
                '価格情報の非表示
                Me.Label9.Visible = False
                Me.Label14.Visible = False
                Me.Label10.Visible = False
                Me.txt_AmtPrice.Visible = False
                Me.txt_AmtTax.Visible = False
                Me.txt_SumTotal.Visible = False
                Me.strHiddenKbn.Value = "1"
            End If
            '海外代理店は合計と消費税を非表示にする
            If Me.objUserInfo.UserClass = CdCst.UserClass.OsAgentCs Then
                Me.Label14.Visible = False
                Me.Label10.Visible = False
                Me.txt_AmtTax.Visible = False
                Me.txt_SumTotal.Visible = False
                Me.strHiddenKbn.Value = "2"
            End If

            'チェック区分/出荷場所表示情報
            ReDim strDispDiv(4)
            For intLoopCnt = 0 To UBound(strDispDiv)
                strDispDiv(intLoopCnt) = False
            Next

            '権限によって形番チェックと出荷場所の表示を制限する
            dt_Addinfo = subAddInfoDispGet(objCon, Me.objUserInfo.UserId, _
                            Me.objLoginInfo.SessionId, selLang.SelectedValue, Me.objUserInfo.AddInformationLvl, _
                            lblSeriesKat.Text, objKtbnStrc)

            Dim strKeylvl As String = "1024,512,256,128,64,32,16,8,4,2,1"
            Dim strLevel() As String = strKeylvl.Split(",")
            Dim ccd_flg As Boolean = False 'FRL白色チェック区分１対応

            For inti As Integer = 0 To strLevel.Length - 1
                If dt_Addinfo Is Nothing Then Exit For
                Dim dr_display() As DataRow = dt_Addinfo.Select("strLevel='" & CInt(strLevel(inti)) & "'")

                If dr_display.Count > 0 Then
                    Select Case CInt(strLevel(inti))
                        Case 128 '中国輸出不可
                        Case 64 'EL品情報
                            Me.txt_ELPrd.Text = String.Empty
                            If dr_display.Length > 0 Then
                                If dr_display(0)("strDisplay") = True Then
                                    Me.txt_ELPrd.Text = dr_display(0)("strValue").ToString
                                End If
                            End If
                            Me.Label13.Visible = True
                            Me.txt_ELPrd.Visible = True
                        Case 32 '販売数量単位
                        Case 16 '標準納期
                            Me.txt_DelDate.Text = String.Empty
                            If dr_display.Length > 0 Then
                                If dr_display(0)("strDisplay") = True Then
                                    Dim str() As String = dr_display(0)("strValue").ToString.Split(CdCst.Sign.Delimiter.Pipe)
                                    If str.Length = 2 Then
                                        ccd_flg = False 'FRL白色 標準納期0日対応
                                        ccd_flg = KHKataban.subJapanChinaAmount(objKtbnStrc.strcSelection.strFullKataban)
                                        If ccd_flg = True Then
                                            If selLang.SelectedValue = "ja" Then
                                                Me.txt_DelDate.Text = "0日間(実稼働日)"
                                            Else
                                                Me.txt_DelDate.Text = "0day(the work days)"
                                            End If
                                        Else
                                            Me.txt_DelDate.Text = str(0)
                                        End If
                                        Me.txt_PrpAmt.Text = str(1)
                                    End If
                                End If
                            End If
                            Me.Label11.Visible = True
                            Me.Label12.Visible = True
                            Me.txt_DelDate.Visible = True
                            Me.txt_PrpAmt.Visible = True
                        Case 8 '担当者情報
                        Case 4 '在庫情報
                        Case 2 '出荷場所
                            strDispDiv(1) = True
                        Case 1 '形番チェック区分
                            strDispDiv(0) = True
                    End Select
                End If
            Next

            Select Case Me.objUserInfo.UserClass
                Case CdCst.UserClass.DmAgentRs, CdCst.UserClass.DmAgentSs, CdCst.UserClass.DmAgentBs, _
                    CdCst.UserClass.DmAgentGs, CdCst.UserClass.DmAgentPs
                    'チェック区分の説明
                    strDispDiv(2) = True        '国内代理店
                Case CdCst.UserClass.OsAgentCs
                    strDispDiv(3) = True        '海外代理店
            End Select
            If httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) IsNot Nothing Then
                strDispDiv(4) = True            '受注EDI連携
            End If

            'セッションにFob価格をセット
            'If (Session("strPriceListFobISO") Is Nothing) Then
            Session("strPriceListFobISO") = Nothing
            Session("strCountryCod") = Me.objUserInfo.CountryCd
            'End If

        Catch ex As Exception
            AlertMessage(ex)
        Finally
            objKataban = Nothing
            objUnitPrice = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 明細作成
    ''' </summary>
    ''' <param name="dtCompData">価格データ</param>
    ''' <remarks></remarks>
    Private Sub subMakeDetail(ByVal dtCompData As DataTable)
        'アプリケーションパス
        Dim strAppPath As String = System.Web.HttpContext.Current.Request.ApplicationPath
        Dim objMsg As New ClsCommon
        Dim strOptionNm As String = ""

        '詳細リスト番号
        Dim intDataNo As Integer = 0
        '各項目の数量
        Dim strItemCnt As String = String.Empty
        '価格リスト
        Dim strPriceLst() As String = Nothing
        Dim script As New System.Text.StringBuilder

        Try
            script.Append("<script language=""JavaScript"">")
            Me.PnlTankaList.Controls.Clear()

            '全項目数の取得
            Dim intTtlCnt As Integer = fncGetTotalCount(dtCompData)

            '選択した情報の取得
            Dim selectInfo As List(Of String()) = fncGetSelectInfo()

            '価格詳細の作成
            For intCnt As Integer = 0 To dtCompData.Rows.Count - 1

                Dim intSpecStrcSeqNo As Integer = dtCompData.Rows(intCnt).Item("spec_strc_seq_no")

                'オプション名の取得
                strOptionNm = fncGetOptionName(intSpecStrcSeqNo)

                If strOptionNm <> CST_BLANK Then
                    intDataNo += 1

                    'ISOユーザーコントロール
                    Dim objCtrl As New UC_ISOTanka

                    '詳細リストの作成
                    objCtrl = fncCreateUCISO(dtCompData.Rows(intCnt), intDataNo, intTtlCnt, strOptionNm, selectInfo, strItemCnt, strPriceLst)

                    Me.PnlTankaList.Controls.Add(objCtrl)

                    'ログの出力
                    subOutputLog(dtCompData.Rows(intCnt))

                    ' 詳細リストのJavascriptの設定
                    subSetUCScript(objCtrl, intDataNo, script)
                End If
            Next

            '単価情報を保存する
            If Me.Session("TestMode") Is Nothing Then
                subInsertHistory(dtCompData)
            End If

            script.Append("</script>")
            ScriptManager.RegisterStartupScript(Page, Page.GetType, "scrollTop", script.ToString, False)

            If Me.PnlTankaList.Controls.Count <= 2 Then
                Me.PnlTankaList.Height = WebControls.Unit.Pixel(420)
            Else
                Me.PnlTankaList.Height = WebControls.Unit.Pixel(Me.PnlTankaList.Controls.Count * 170)
            End If

            '項目数
            Me.intItemRow.Value = intTtlCnt
            '項目の数量(例："|1|3|1")
            Me.intItemCnt.Value = strItemCnt
            'リスト行
            If strPriceLst IsNot Nothing Then
                Me.intListRowCnt.Value = UBound(strPriceLst)
            Else
                Me.intListRowCnt.Value = 0
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 選択した情報の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetSelectInfo() As List(Of String())
        Dim result As New List(Of String())
        Dim strFileInfo As String = Me.HidPriceForFile.Value

        If Not strFileInfo.Equals(String.Empty) Then
            Dim strAll As List(Of String)

            strAll = strFileInfo.Split("_").ToList

            For Each strItem In strAll
                result.Add(strItem.Split("|"))
            Next
        End If

        Return result

    End Function

    ''' <summary>
    ''' 全項目数の取得
    ''' </summary>
    ''' <param name="dtCompData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetTotalCount(ByVal dtCompData As DataTable) As Integer
        Dim intResult As Integer = 0

        For intCnt As Integer = 0 To dtCompData.Rows.Count - 1
            Dim intSpecStrcSeqNo As Integer = dtCompData.Rows(intCnt)("spec_strc_seq_no")
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "CMF", "GMF"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            intResult = intResult + 1
                        Case 2, 3, 4, 5, 6, 7
                            intResult = intResult + 1
                        Case 13, 14
                            intResult = intResult + 1
                        Case 15, 16
                            intResult = intResult + 1
                        Case 17, 18
                            intResult = intResult + 1
                        Case 19, 20, 21, 22
                            intResult = intResult + 1
                        Case 23, 24
                            intResult = intResult + 1
                        Case Else
                    End Select
                Case "LMF0"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            intResult = intResult + 1
                        Case 2, 3, 4, 5, 6, 7
                            intResult = intResult + 1
                        Case 13, 14
                            intResult = intResult + 1
                        Case 15, 16
                            intResult = intResult + 1
                        Case 17
                            intResult = intResult + 1
                        Case 18, 19
                            intResult = intResult + 1
                        Case Else
                    End Select
            End Select
        Next

        Return intResult
    End Function

    ''' <summary>
    ''' オプション名の取得
    ''' </summary>
    ''' <param name="intSpecStrcSeqNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetOptionName(ByVal intSpecStrcSeqNo As Integer) As String
        Dim strResult As String = String.Empty
        Dim strOpNm As New ArrayList

        'オプション名称データ取得
        Dim dt_Title As DataTable = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHISOTanka, selLang.SelectedValue)

        For inti As Integer = 0 To dt_Title.Rows.Count - 1
            If dt_Title.Rows(inti)("label_div").ToString = "L" Then
                strOpNm.Add(dt_Title.Rows(inti)("label_content").ToString)
            End If
        Next

        Select Case objKtbnStrc.strcSelection.strSeriesKataban
            Case "CMF", "GMF"
                Select Case intSpecStrcSeqNo
                    Case 1
                        strResult = strOpNm(0)  'ベース
                    Case 2, 3, 4, 5, 6, 7
                        strResult = strOpNm(1)  '電磁弁形式
                    Case 13, 14
                        strResult = strOpNm(2)  '給気スペーサ
                    Case 15, 16
                        strResult = strOpNm(3)  '排気スペーサ
                    Case 17, 18
                        strResult = strOpNm(4)  'パイロットチェック弁
                    Case 19, 20, 21, 22
                        strResult = strOpNm(5)  'スペーサ形減圧弁
                    Case 23, 24
                        strResult = strOpNm(6)  '流露遮蔽板
                    Case Else
                        strResult = CST_BLANK
                End Select
            Case "LMF0"
                Select Case intSpecStrcSeqNo
                    Case 1
                        strResult = strOpNm(0)
                    Case 2, 3, 4, 5, 6, 7
                        strResult = strOpNm(1)
                    Case 13, 14
                        strResult = strOpNm(2)
                    Case 15, 16
                        strResult = strOpNm(3)
                    Case 17
                        strResult = strOpNm(4)
                    Case 18, 19
                        strResult = strOpNm(6)
                    Case Else
                        strResult = CST_BLANK
                End Select
        End Select

        Return strResult
    End Function

    ''' <summary>
    ''' 価格詳細リストの作成
    ''' </summary>
    ''' <param name="drCompData">データ情報</param>
    ''' <param name="intDataNo">詳細番号</param>
    ''' <param name="intTtlCnt">項目数</param>
    ''' <param name="strOptionNm">項目名</param>
    ''' <param name="strItemCnt">数量</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreateUCISO(ByVal drCompData As DataRow, _
                                    ByVal intDataNo As Integer, _
                                    ByVal intTtlCnt As Integer, _
                                    ByVal strOptionNm As String, _
                                    ByVal strSelectInfo As List(Of String()), _
                                    ByRef strItemCnt As String, _
                                    ByRef strPriceLst() As String) As UC_ISOTanka

        Dim ucResult As New UC_ISOTanka
        Dim objUnitPrice As New KHUnitPrice
        Dim intSglPrice(5) As Decimal

        '出荷場所の取得
        Dim strCountryCode As String = Me.cmbPlace.SelectedValue

        '価格リストの作成
        ucResult = LoadControl("UC_ISOTanka.ascx")
        ucResult.LangCd = Me.selLang.SelectedValue
        ucResult.objKtbnStrc = Me.objKtbnStrc
        ucResult.objCon = Me.objCon
        ucResult.objConBase = Me.objConBase

        '一番目のリストの数量項目だけが入力できる
        If intDataNo = 1 Then
            ucResult.IsFirst = True
        End If

        '単価リスト取得
        intSglPrice(0) = drCompData.Item("ls_price")
        intSglPrice(1) = drCompData.Item("rg_price")
        intSglPrice(2) = drCompData.Item("ss_price")
        intSglPrice(3) = drCompData.Item("bs_price")
        intSglPrice(4) = drCompData.Item("gs_price")
        intSglPrice(5) = drCompData.Item("ps_price")

        '現地定価とFOB価格の取得(ISO)
        objUnitPrice.subISOPriceListSelect(objCon, objConBase, objKtbnStrc, _
                                           Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                           Me.objUserInfo.CountryCd, selLang.SelectedValue, _
                                           Me.objUserInfo.CurrencyCd, Me.objUserInfo.PriceDispLvl, _
                                           drCompData.Item("option_kataban"), _
                                           drCompData.Item("kataban_check_div"), _
                                           intSglPrice, strCountryCode, strPriceList)

        For intCnt1 As Integer = 0 To UBound(strPriceList) - 1
            If intCnt1 = 0 Then
                ReDim strPriceLst(0)
            Else
                ReDim Preserve strPriceLst(UBound(strPriceLst) + 1)
            End If
            If strPriceList(intCnt1 + 1, 2) <= 0 Then
                strPriceLst(intCnt1) = strPriceList(intCnt1 + 1, 1) & CST_PIPE & _
                                       ClsCommon.fncGetMsg(selLang.SelectedValue, "I5220") & CST_PIPE & _
                                       CST_PIPE & _
                                       strPriceList(intCnt1 + 1, 4)
            Else
                strPriceLst(intCnt1) = strPriceList(intCnt1 + 1, 1) & CST_PIPE & _
                                       strPriceList(intCnt1 + 1, 2) & CST_PIPE & _
                                       strPriceList(intCnt1 + 1, 3) & CST_PIPE & _
                                       strPriceList(intCnt1 + 1, 4)
            End If
        Next

        'セッションにFob価格をセット
        For i As Integer = 0 To (strPriceList.Length / 5) - 1
            If strPriceList(i, 4) = "FobPrice" Then
                If (Session("strPriceListFobISO") Is Nothing) Or Session("strPriceListFobISO") = "" Then
                    Session("strPriceListFobISO") = strPriceList(i, 2)
                    Session("strCurrencyCode") = strPriceList(i, 3)
                Else
                    Dim strstrPriceListFob As String
                    Session("strPriceListFobISO") = Session("strPriceListFobISO") & "," & strPriceList(i, 2)
                    strstrPriceListFob = Session("strPriceListFobISO")
                    Session("strCurrencyCode") = Session("strCurrencyCode") & "," & strPriceList(i, 3)
                End If

                '価格詳細画面用
                SetPriceDetailControl()
            End If
        Next

        'プロパティ設定
        With ucResult
            'ID設定
            .ID = "ISODetail" & CStr(intDataNo)
            '言語区分
            .LangCd = selLang.SelectedValue
            'オプション名称
            .OptionNm = strOptionNm
            'オプション形番形番
            .OptionKtbn = drCompData.Item("option_kataban")
            '表示単価リスト
            If strPriceLst IsNot Nothing Then
                .PriceLst = strPriceLst.Clone
            Else
                .PriceLst = Nothing
            End If
            '編集区分
            .EditDiv = Me.objUserInfo.EditDiv
            '形番チェック
            .KtbnChk = drCompData.Item("kataban_check_div")
            '出荷場所
            '.ShipPlace = drCompData.Item("place_cd")
            .ShipPlace = Me.cmbPlace.SelectedItem.Text
            '行数計
            .TtlCnt = intTtlCnt
            '行No.
            .DataNo = intDataNo
            '各項目の数量をhiddenエリアに入れる(例："|1|3|1")
            strItemCnt = strItemCnt & CST_PIPE & drCompData.Item("quantity")
            '形番チェック/出荷場所表示情報
            .DispDiv = strDispDiv

            If strSelectInfo.Count > 0 AndAlso strSelectInfo.Count >= intDataNo - 1 Then
                '既に入力した場合
                If strSelectInfo(intDataNo - 1).Count >= 6 Then
                    '掛率
                    .Rate = strSelectInfo(intDataNo - 1)(0)
                    '単価
                    .UnitPrice = strSelectInfo(intDataNo - 1)(1)
                    '掛単価
                    '.RatePrice = CDec(strSelectInfo(intDataNo - 1)(0)) * CDec(strSelectInfo(intDataNo - 1)(1))
                    '数量
                    .Quantity = strSelectInfo(intDataNo - 1)(2)
                    '金額
                    .Price = strSelectInfo(intDataNo - 1)(3)
                    '消費税
                    .Tax = strSelectInfo(intDataNo - 1)(4)
                    '合計
                    .Total = strSelectInfo(intDataNo - 1)(5)
                End If
            Else
                '初期値の設定
                If .EditDiv = "0" Then
                    .Rate = CdCst.UnitPrice.DefaultNmlRate
                    .RatePrice = CdCst.UnitPrice.DefaultNmlRateUnitPrice
                Else
                    .Rate = CdCst.UnitPrice.DefaultOtrRate
                    .RatePrice = CdCst.UnitPrice.DefaultOtrRateUnitPrice
                End If
                .UnitPrice = CdCst.UnitPrice.DefaultUnitPrice
            End If
            
        End With

        '画面設定
        'subSetInitScreen(ucResult)

        Return ucResult

    End Function

    ''' <summary>
    ''' 詳細リストのJavascriptの設定
    ''' </summary>
    ''' <param name="objCtrl"></param>
    ''' <param name="intDataNo"></param>
    ''' <param name="Script"></param>
    ''' <remarks></remarks>
    Private Sub subSetUCScript(ByVal objCtrl As UC_ISOTanka, ByVal intDataNo As Integer, ByRef Script As StringBuilder)
        'JS
        objCtrl.AttTxtPrc(CdCst.JavaScript.OnChange) = "fncUntPrcOnchange('" & objCtrl.ClientID & "_')"
        objCtrl.AttTxtAmnt(CdCst.JavaScript.OnChange) = "fncAmountOnchange('" & objCtrl.ClientID & "_')"
        objCtrl.AttTxtRate(CdCst.JavaScript.OnChange) = "fncRateOnchange('" & objCtrl.ClientID & "_')"
        objCtrl.AttChkUnitList(CdCst.JavaScript.OnClick) = "ISOTanka_ChkUnitList('" & objCtrl.ClientID & "_','" & intDataNo & "')"

        If intDataNo > 1 Then
            objCtrl.txt_Amount.BackColor = Drawing.Color.White
        Else
            objCtrl.txt_Amount.BackColor = Drawing.Color.FromArgb(255, 255, 204)
        End If

        objCtrl.txt_Price.BackColor = Drawing.Color.White
        objCtrl.txt_Total.BackColor = Drawing.Color.White
        objCtrl.txt_Tax.BackColor = Drawing.Color.White
        objCtrl.txt_DtlPrc.BackColor = Drawing.Color.White
        objCtrl.txt_KtbnChk.BackColor = Drawing.Color.White
        objCtrl.txt_Place.BackColor = Drawing.Color.White
        Script.Append("if(document.getElementById('" & objCtrl.ClientID & "_pnlPrice')){")
        Script.Append("document.getElementById('" & objCtrl.ClientID & "_pnlPrice').scrollTop = '60';}")
    End Sub

    ''' <summary>
    ''' 履歴登録
    ''' </summary>
    ''' <param name="dtCompData"></param>
    ''' <remarks></remarks>
    Private Sub subInsertHistory(ByVal dtCompData As DataTable)
        Dim dt_history As New DS_History.kh_price_historyDataTable
        'DBに単価情報を保存する
        Dim dr_history As DataRow = dt_history.NewRow
        'ホスト名
        Dim hostname As String = "******"

        Try
            For Each drCompData As DataRow In dtCompData.Rows
                Dim strStart As Date = Now
                Dim strEnd As New Date
                Dim strTime As New TimeSpan
                Dim intCnt As Integer = dtCompData.Rows.IndexOf(drCompData)

                dr_history("ELFlag") = String.Empty
                dr_history("KataPlace") = drCompData.Item("place_cd")
                dr_history("KataCheck") = drCompData.Item("kataban_check_div")
                '単価リスト
                For intLoopCnt1 = 1 To UBound(strPriceList)
                    Select Case strPriceList(intLoopCnt1, 4)
                        Case CdCst.UnitPrice.ListPrice
                            dr_history("ListPrice") = strPriceList(intLoopCnt1, 2)
                        Case CdCst.UnitPrice.RegPrice
                            dr_history("RegPrice") = strPriceList(intLoopCnt1, 2)
                        Case CdCst.UnitPrice.SsPrice
                            dr_history("SSPrice") = strPriceList(intLoopCnt1, 2)
                        Case CdCst.UnitPrice.BsPrice
                            dr_history("BSPrice") = strPriceList(intLoopCnt1, 2)
                        Case CdCst.UnitPrice.GsPrice
                            dr_history("GSPrice") = strPriceList(intLoopCnt1, 2)
                        Case CdCst.UnitPrice.PsPrice
                            dr_history("PSPrice") = strPriceList(intLoopCnt1, 2)
                    End Select
                Next

                strEnd = Now
                strTime = strEnd - strStart

                dr_history("UpdateDate") = Now
                dr_history("UpdateComputer") = Right(hostname.PadRight(10), 10)
                dr_history("UpdateUser") = Me.objUserInfo.UserId
                dr_history("MFNo") = intCnt + 1
                dr_history("Kataban_Title") = objKtbnStrc.strcSelection.strGoodsNm
                dr_history("Kataban") = drCompData.Item("option_kataban")
                dr_history("Runtime") = strTime.Milliseconds
                dt_history.Rows.Add(dr_history)
            Next

            If dt_history.Rows.Count > 0 Then
                Using da As New DS_HistoryTableAdapters.kh_price_historyTableAdapter
                    da.Update(dt_history)
                End Using
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' ログの出力
    ''' </summary>
    ''' <param name="drCompData"></param>
    ''' <remarks></remarks>
    Private Sub subOutputLog(ByVal drCompData As DataRow)
        'ログ出力   DBに保存しました、ファイル出力を廃棄してもいい

        '形番情報取得
        Dim strKatabanInfo(3) As String        '形番情報

        strKatabanInfo(1) = drCompData.Item("option_kataban")
        strKatabanInfo(2) = drCompData.Item("kataban_check_div")
        strKatabanInfo(3) = drCompData.Item("place_cd")

        If dt_Addinfo Is Nothing Then
            '権限によって形番チェックと出荷場所の表示を制限する
            dt_Addinfo = subAddInfoDispGet(objCon, Me.objUserInfo.UserId, _
                         Me.objLoginInfo.SessionId, selLang.SelectedValue, _
                         Me.objUserInfo.AddInformationLvl, lblSeriesKat.Text, objKtbnStrc)
        End If

        'テキスト出力(比較のため、１ヶ月削除保留)
        Call bllTanka.subLogFileOutput(objCon, strPriceList, dt_Addinfo, Me.objUserInfo.UserId, _
                               Me.objLoginInfo.SessionId, Me.objUserInfo.CountryCd, selLang.SelectedValue, strKatabanInfo)
        Call bllTanka.subLogOutput(objConBase, objCon, strPriceList, dt_Addinfo, Me.objUserInfo.UserId, _
                               Me.objLoginInfo.SessionId, Me.objUserInfo.CountryCd, selLang.SelectedValue)

    End Sub

    ''' <summary>
    ''' JavaScript生成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInitScript()

        Try
            'ボタン自動サブミットを無効
            Me.Button2.UseSubmitBehavior = False
            Me.Button3.UseSubmitBehavior = False
            Me.Button5.UseSubmitBehavior = False
            Me.Button6.UseSubmitBehavior = False
            Me.Button9.UseSubmitBehavior = False

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 表示付加情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strSelectLang">選択言語</param>
    ''' <param name="intAddinfoDispLvl">付加情報レベル</param>
    ''' <param name="strFullKatabanSiyouNo">付加情報</param>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks>表示付加情報レベルを元に表示付加情報を取得する</remarks>
    Public Shared Function subAddInfoDispGet(objCon As SqlConnection, ByVal strUserId As String, _
                                 ByVal strSessionId As String, ByVal strSelectLang As String, _
                                 ByVal intAddinfoDispLvl As Integer, ByVal strFullKatabanSiyouNo As String, _
                                 ByVal objKtbnStrc As KHKtbnStrc) As DataTable
        Dim objKataban As New KHKataban
        Dim objKHStdDlv As New KHStdDlv
        Dim intAddinfoLvl As Integer
        Dim strQtyUnitNm As String = Nothing
        Dim strStdDlvDt As String = Nothing
        Dim strQuantity As String = Nothing
        Dim strSalesUnit As String = Nothing
        Dim strSapBaseUnit As String = Nothing
        Dim strQuantityPerSalesUnit As String = Nothing
        Dim strOrderLot As String = Nothing

        Dim ccd_flg As Boolean = False
        subAddInfoDispGet = New DataTable
        Try
            '配列初期化
            Dim dc As New DataColumn("strLevel")
            subAddInfoDispGet.Columns.Add(dc)
            dc = New DataColumn("strDisplay")
            subAddInfoDispGet.Columns.Add(dc)
            dc = New DataColumn("strValue")
            subAddInfoDispGet.Columns.Add(dc)

            '表示付加情報レベル設定
            intAddinfoLvl = intAddinfoDispLvl

            Dim strKey As String = "1024,512,256,128,64,32,16,8,4,2,1"
            Dim strLevel() As String = strKey.Split(",")

            For inti As Integer = 0 To strLevel.Length - 1
                If intAddinfoLvl >= CInt(strLevel(inti)) Then
                    Dim dr As DataRow = subAddInfoDispGet.NewRow
                    dr("strLevel") = CInt(strLevel(inti))
                    dr("strDisplay") = True
                    dr("strValue") = ""
                    Select Case CInt(strLevel(inti))
                        Case 128 '中国輸出不可
                            '表示のみ
                        Case 64 'EL品情報
                            If objKataban.fncELKatabanCheck(objCon, objKtbnStrc.strcSelection.strFullKataban, "1") Then
                                dr("strValue") = "O"
                            End If
                        Case 32 '販売数量単位
                            If objKataban.fncQtyUnitInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, _
                                                         strSelectLang, strQtyUnitNm, objKtbnStrc) Then
                                dr("strValue") = strQtyUnitNm
                            End If
                        Case 16 '標準納期
                            '標準納期取得
                            'それ以外のときは影響を受けないようにしている
                            If strFullKatabanSiyouNo <> "" Then
                                Call objKHStdDlv.subStdDlvDtInfo(objCon, strFullKatabanSiyouNo, _
                                                                 strSelectLang, strStdDlvDt, strQuantity)
                            Else
                                Call objKHStdDlv.subStdDlvDtInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, _
                                                                 strSelectLang, strStdDlvDt, strQuantity)
                            End If
                            ccd_flg = KHKataban.subJapanChinaAmount(objKtbnStrc.strcSelection.strFullKataban)
                            If ccd_flg = True Then strQuantity = ""
                            dr("strValue") = strStdDlvDt & CdCst.Sign.Delimiter.Pipe & strQuantity
                        Case 8 '担当者情報
                            '表示のみ
                        Case 4 '在庫情報
                            If objKataban.fncStockInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, strSelectLang, _
                                objKtbnStrc.strcSelection.strPlaceCd, 0, 0, "") Then
                                dr("strDisplay") = True
                            End If
                        Case 2 '出荷場所
                            dr("strValue") = objKtbnStrc.strcSelection.strPlaceCd
                        Case 1 '形番チェック区分
                            dr("strDisplay") = True
                            If objKtbnStrc.strcSelection.strCostCalcNo = CdCst.CostCalcNo.C5 Then
                                dr("strValue") = objKtbnStrc.strcSelection.strKatabanCheckDiv & "(" & objKtbnStrc.strcSelection.strCostCalcNo & ")"
                            Else
                                dr("strValue") = objKtbnStrc.strcSelection.strKatabanCheckDiv
                            End If
                    End Select
                    intAddinfoLvl -= CInt(strLevel(inti))
                    subAddInfoDispGet.Rows.Add(dr)
                End If
            Next

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objKataban = Nothing
            objKHStdDlv = Nothing
        End Try

    End Function

    ''' <summary>
    ''' 価格詳細ボタンとHiddenFieldの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetPriceDetailControl()
        Button9.Visible = True
        Button9.OnClientClick = "f_ShowPriceDetail('" & Me.ClientID & "_" & "')"
        HidPriceDetail.Value = "1"
    End Sub

    ''' <summary>
    ''' 全ての国の出荷場所を取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetPlace()
        Dim lstPlace As New List(Of String)

        If cmbPlace.Visible = True Then
            '選択できる場合
            For Each item As ListItem In cmbPlace.Items
                lstPlace.Add(item.Value)
            Next
        Else
            '選択できない場合
            lstPlace.Add("JPN")
        End If

        Session.Add("ShipPlaces", lstPlace)

    End Sub

    ''' <summary>
    ''' マニホールドテスト専用
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subManifoldTest()
        If Not Me.Session("ManifoldKataban") Is Nothing Then
            If Me.Session("TestFlag") Is Nothing Then
                Me.Session("TestFlag") = True
                Dim lngLoop As Long = CLng(Me.Session("ManifoldKatabanLoop"))
                Dim listKataban As ManifoldKataban = Me.Session("ManifoldKataban")
                Dim strKataban As String = listKataban.KATABAN
                Dim strSiyou As String = listKataban.SIYOUSYO

                Dim strPath As String = My.Settings.LogFolder & "MFShiyouTestISO.txt"

                If Me.Session("TestMode").ToString <> "2" Then
                    'WEBサービスを呼び出してIF出力内容の作成
                    Dim strFobPrice As String = String.Empty
                    Dim strCountryCd As String = String.Empty

                    If Session("strPriceListFobISO") IsNot Nothing Then
                        strFobPrice = Session("strPriceListFobISO")
                    End If

                    If Session("strCountryCod") IsNot Nothing Then
                        strCountryCd = Session("strCountryCod")
                    End If


                    Dim strOutputText As String = KHSBOInterface.fncSBOInterfaceGet(objCon, objKtbnStrc, strFobPrice, strCountryCd, Me.objUserInfo.OfficeCd, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
                    System.IO.File.AppendAllText(strPath, strOutputText)

                    '仕様書の作成
                    RaiseEvent SiyouFileOutput(objKtbnStrc, strSiyou)

                    Session("EventEndFlg") = True
                    GC.Collect()
                End If

            End If
        End If
    End Sub
#End Region

End Class