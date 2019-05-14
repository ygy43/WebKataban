Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports System.Net

Public Class WebUC_Tanka
    Inherits KHBase

#Region "プロパティ"
    Private EditDivOpt As String = String.Empty
    Private dt_Addinfo As DataTable = Nothing
    Private NewPlace As Boolean = False
    Private bllTanka As New TankaBLL

    '価格一覧画面へイベント
    Public Event GotoCopyPrice()
    '価格詳細画面へイベント
    Public Event GotoPriceDetail()
    '仕様書出力画面へイベント
    Public Event SiyouFileOutput(objKtbnStrc As KHKtbnStrc, strSiyou As String)
    'I/F出力イベント
    Public Event IFFileOutput(objKtbnStrc As KHKtbnStrc, strName As String, strNewPlace As String)
    'ファイル出力イベント
    Public Event FileOutput(objKtbnStrc As KHKtbnStrc, strName As String, strOrder As String, strPriceList As String, intMode As Integer)
    'JSONファイル出力イベント
    Public Event JSONFileOutput(objKtbnStrc As KHKtbnStrc, strName As String)
    'EDIに戻るイベント
    Public Event EDIReturn()
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        If Not Me.Session("lstCountry_Key") Is Nothing Then Me.Session.Remove("lstCountry_Key")
        If Not Me.Session("ShipPlaces") Is Nothing Then Me.Session.Remove("ShipPlaces")
        'セッションとHiddenFieldの初期化
        Me.Session("TestFlag") = Nothing
        Me.HidShiftD.Value = String.Empty
        Me.HdnSetYFlg.Value = String.Empty
        Me.HidSelRowID.Value = String.Empty
        Me.SelUnitValue.Value = String.Empty
        Me.HidPriceForFile.Value = String.Empty
        Me.SelCurrValue.Value = String.Empty
        Me.HidNewPlace.Value = String.Empty
        '価格詳細画面
        Me.HidPriceDetail.Value = String.Empty


        selLang = Me.Parent.Parent.Parent.Parent.FindControl("ContentTitle").FindControl("selLang")

        'Me.OnLoad(Nothing)

        Me.txt_Rate.Text = String.Empty
        Me.TextUnitPrice.Text = String.Empty
        Me.TextCnt.Text = String.Empty
        Me.lblQtyUnit.Text = String.Empty
        Me.lblQtyUnit1.Text = String.Empty

        'フォントの設定
        'Call SetFontAlign(lblCheck, TextAlign.Center) 
        Call SetFontAlign(lblCheck, TextAlign.Left)
        Call SetFontAlign(lblCheckZ, TextAlign.Right)
        Call SetFontAlign(lblEL, TextAlign.Center)
        Call SetFontAlign(lblKosuu, TextAlign.Center)
        Call SetFontAlign(lblNouki, TextAlign.Center)
        Call SetFontAlign(txtSelKosu, TextAlign.Center)
        Call SetFontAlign(txtSelNoki, TextAlign.Center)

        lblSeriesNm.Font.Name = GetFontName(selLang.SelectedValue)
        lblSeriesKat.Font.Name = GetFontName(selLang.SelectedValue)
        GVPrice.Font.Name = GetFontName(selLang.SelectedValue)

        'スタイルの設定
        Call SetAttributes(txt_Rate, 0)
        Call SetAttributes(TextUnitPrice, 0)
        Call SetAttributes(TextCnt, 0)
        Call SetAttributes(TextRateUnitPrice, 1)
        Call SetAttributes(TextMoney, 1)
        Call SetAttributes(TextTax, 1)
        Call SetAttributes(TextAmount, 1)

        'ロード
        '画面の設定
        Call subDispSet(String.Empty)

        'ボタンの設定
        Call subBtnView()

        'Javascriptの設定
        Call subSetInit()

        'テキストをクリア
        Call TxtClear()


        Me.OnLoad(Nothing)


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

        Try
            If Me.HidShiftD.Value = "2" Then
                '価格積上げ画面へ
                Me.HidShiftD.Value = "1"
                RaiseEvent GotoCopyPrice()

            ElseIf Me.HidPriceDetail.Value = "2" Then
                '価格詳細画面へ遷移
                Me.HidPriceDetail.Value = "1"

                '出荷場所の保存
                subSetPlace()

                RaiseEvent GotoPriceDetail()

            Else
                Me.txt_EditNormal.Text = Me.objUserInfo.EditDiv
                'フォントの設定
                Call SetAllFontName(Me)

                '生産品ラベル制御
                If cmbPlace.Items.Count > 1 Then
                    'ドロップダウンボックス選択変更時に生産品ラベルも更新(初期化する時に実行しない)
                    Call SetPlaceMark(cmbPlace.SelectedItem.Value, cmbPlace.Items.Count)

                    Select Case cmbPlace.SelectedItem.Value
                        '日本製を選んだ時のみFRL注意メッセージを表示
                        Case "P", "S", "K", "C", "JPN", "P55", "C55", "S55", "K55", "1001", "1002", "1003", "1004", "1005"
                            Dim strFrlMessage As String = fncFRLMessage()
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "ShowFrlMessage", "fncShowFrlMessage('" & strFrlMessage & "');", True)
                    End Select

                End If

                '引当情報取得
                Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)

                '画面表示項目制御
                Call ChangeTitle(cmbPlace.SelectedValue, Me.lblCheck.Text)

                '共通ラベルタイトル設置
                Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHUnitPrice, selLang.SelectedValue, Me)
                Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHTanka, selLang.SelectedValue, Me)

                '価格情報を取得して表示する
                Call subPriceListMake(objKtbnStrc)

                '営業本部、情報システム部ユーザーのみ価格積上げ表示画面を表示する
                If Me.objUserInfo.UserClass >= CdCst.UserClass.DmSalesOffice Then
                    Me.HidShiftD.Value = "1"
                End If

                If Not NewPlace Then
                    Dim strKey As String = TextRateUnitPrice.ClientID & "," & TextMoney.ClientID & "," & TextTax.ClientID & "," & TextAmount.ClientID
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "SetValue", "fncDisableText('" & strKey & "');", True)

                    If SelCurrValue.Value.Length > 0 And SelUnitValue.Value.Length > 0 And HidSelRowID.Value.Length > 0 Then
                        Dim strID As String = Strings.Left(Me.GVPrice.Rows(0).ClientID, Me.GVPrice.Rows(0).ClientID.Length - 2)
                        strID = strID & HidSelRowID.Value.PadLeft(2, "0")
                        For inti As Integer = 0 To Me.GVPrice.Rows.Count - 1
                            If strID = Me.GVPrice.Rows(inti).ClientID Then
                                Me.GVPrice.Rows(inti).BackColor = System.Drawing.ColorTranslator.FromHtml("#003C80")
                                Me.GVPrice.Rows(inti).ForeColor = Drawing.Color.White
                            End If
                        Next
                    End If
                Else
                    NewPlace = False
                End If
            End If

            '単価画面特殊ラベル設定
            Call SetCountryName()

            '3D CAD Jsonデータの作成
            Call Create3DJsonData()

            'マニホールドテスト専用
            Call ManifoldTanka()

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

#Region "Create Json Data"

    ''' <summary>
    ''' Jsonデータの作成
    ''' </summary>
    Private Sub Create3DJsonData()

        'Jsonデータの作成
        Dim strJsonData As String = subMakeJSONData(objKtbnStrc)
        
        'Jsonデータエンコード
        Dim strEncodedJsonData As String = Uri.EscapeDataString(strJsonData)
        
        'Cadenas表示設定
        Dim strLanguage as String = "english"

        Select Case selLang.SelectedValue
            Case CdCst.LanguageCd.Japanese
                strLanguage = "japanese"
        End Select

        Button12.OnClientClick = "call_cadenas('" & strLanguage & "', '" & strEncodedJsonData & "')"

    End Sub
    
    ''' <summary>
    '''     ＪＳＯＮファイル出力データを作成する
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function subMakeJSONData(objKtbnStrc As KHKtbnStrc) As String

        Dim strTab(4) As String
        Dim sbFileData As New StringBuilder
        Dim intLoopCnt As Integer
        Dim intLoopCnt2 As Integer
        Dim intLoopMax_Option As Integer
        Dim intLoopMax_Station As Integer
        strTab(1) = CdCst.Sign.Delimiter.Tab
        strTab(2) = strTab(1) & CdCst.Sign.Delimiter.Tab
        strTab(3) = strTab(2) & CdCst.Sign.Delimiter.Tab
        strTab(4) = strTab(3) & CdCst.Sign.Delimiter.Tab
        subMakeJSONData = String.Empty


        Try

            intLoopMax_Option = CdCst.Siyou_04.Spacer4   '電磁弁～スペーサ最終行までの最大（2018/9時点では、SpecNo.04のみ対応）

            '属性記号D1～D4の合計分ステーション数
            For intLoopCnt = 1 To intLoopMax_Option
                Select Case objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).ToString
                    Case "D1", "D2", "D3", "D4"
                        intLoopMax_Station += CInt(objKtbnStrc.strcSelection.intQuantity(intLoopCnt))
                End Select
            Next

            With sbFileData

                .AppendLine("{")
                .AppendLine(strTab(1) & ClsCommon.fncAddQuote("ASSY") & ": {")
                .AppendLine(
                    strTab(2) & ClsCommon.fncAddQuote("ASSY-ON") & ": " &
                    ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.strFullKataban) & CdCst.Sign.Delimiter.Comma)
                .AppendLine(strTab(2) & ClsCommon.fncAddQuote("MANIFOLD-BASE") & ": {")
                If objKtbnStrc.strcSelection.decDinRailLength = 0 Then
                    '.AppendLine(strTab(3) & ClsCommon.fncAddQuote("BASE-RAIL-ON") & ": " & ClsCommon.fncAddQuote("BAA") & CdCst.Sign.Delimiter.Comma)
                    '.AppendLine(strTab(3) & ClsCommon.fncAddQuote("BASE-RAIL-LEN") & ": " &
                    '            ClsCommon.fncAddQuote("87.5"))
                Else
                    .AppendLine(
                        strTab(3) & ClsCommon.fncAddQuote("BASE-RAIL-ON") & ": " & ClsCommon.fncAddQuote("BAA") &
                        CdCst.Sign.Delimiter.Comma)
                    .AppendLine(
                        strTab(3) & ClsCommon.fncAddQuote("BASE-RAIL-LEN") & ": " &
                        ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.decDinRailLength.ToString("#.#")))
                End If
                .AppendLine(strTab(2) & "}" & CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 1 To intLoopMax_Station
                    Dim strValveKataban As String = String.Empty
                    Dim strSpacerKataban As String = String.Empty
                    Dim strAport As String = String.Empty
                    Dim strBport As String = String.Empty

                    .AppendLine(strTab(2) & ClsCommon.fncAddQuote("STATION-" & intLoopCnt.ToString("00")) & ": {")

                    For intLoopCnt2 = 1 To intLoopMax_Option
                        If Mid(objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt2).ToString, intLoopCnt, 1) = "1" _
                            Then
                            If objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt2).Trim <> "" Then
                                strAport = objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt2).Trim
                                strBport = objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt2).Trim
                            End If
                            Select Case objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt2).ToString
                                Case "D1", "D2"
                                    'VALVE
                                    strValveKataban = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt2).Trim
                                Case "GB", "S7", "S3", "S2"
                                    'SPACER
                                    strSpacerKataban = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt2).Trim
                                Case "D3", "D4"
                                    'MASPLATE
                                    .AppendLine(
                                        strTab(3) & ClsCommon.fncAddQuote("MASKING-PLATE-ON") & ": " &
                                        ClsCommon.fncAddQuote(
                                            objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt2).Trim) &
                                        CdCst.Sign.Delimiter.Comma)
                                    .AppendLine(strTab(3) & ClsCommon.fncAddQuote("OPTIONS-ON") & ": [")
                                    .AppendLine(
                                        strTab(4) &
                                        ClsCommon.fncAddQuote(
                                            objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt2).Trim))
                                    .AppendLine(strTab(3) & "]")
                                Case Else
                            End Select
                        End If
                    Next
                    'VALVE
                    If strValveKataban <> "" Then
                        .AppendLine(strTab(3) & ClsCommon.fncAddQuote("VALVE") & ": {")
                        .AppendLine(
                            strTab(4) & ClsCommon.fncAddQuote("VALVE-ON") & ": " &
                            ClsCommon.fncAddQuote(strValveKataban))
                        .AppendLine(strTab(3) & "}" & CdCst.Sign.Delimiter.Comma)
                        'SPACER
                        If strSpacerKataban <> "" Then
                            .AppendLine(
                                strTab(3) & ClsCommon.fncAddQuote("SPACER-ON") & ": " &
                                ClsCommon.fncAddQuote(strSpacerKataban) & CdCst.Sign.Delimiter.Comma)
                            .AppendLine(strTab(3) & ClsCommon.fncAddQuote("OPTIONS-ON") & ": [")
                            .AppendLine(strTab(4) & ClsCommon.fncAddQuote(strSpacerKataban) & CdCst.Sign.Delimiter.Comma)
                            .AppendLine(strTab(4) & ClsCommon.fncAddQuote(strValveKataban))
                        Else
                            .AppendLine(strTab(3) & ClsCommon.fncAddQuote("OPTIONS-ON") & ": [")
                            .AppendLine(strTab(4) & ClsCommon.fncAddQuote(strValveKataban))
                        End If
                        'CX
                        If strAport <> "" Then
                            .AppendLine(strTab(3) & "]" & CdCst.Sign.Delimiter.Comma)
                            .AppendLine(strTab(3) & ClsCommon.fncAddQuote("ST-A-PORT") & ": {")
                            .AppendLine(
                                strTab(4) & ClsCommon.fncAddQuote("ST-PIPING-SIZE") & ": " &
                                ClsCommon.fncAddQuote(strAport))
                            .AppendLine(strTab(3) & "}" & CdCst.Sign.Delimiter.Comma)
                            .AppendLine(strTab(3) & ClsCommon.fncAddQuote("ST-B-PORT") & ": {")
                            .AppendLine(
                                strTab(4) & ClsCommon.fncAddQuote("ST-PIPING-SIZE") & ": " &
                                ClsCommon.fncAddQuote(strBport))
                            .AppendLine(strTab(3) & "}")
                        Else
                            .AppendLine(strTab(3) & "]")
                        End If
                    End If
                    If intLoopCnt = intLoopMax_Station Then
                        .AppendLine(strTab(2) & "}")
                    Else
                        .AppendLine(strTab(2) & "}" & CdCst.Sign.Delimiter.Comma)
                    End If
                Next
                .AppendLine(strTab(1) & "}")
                .Append("}")

            End With
            subMakeJSONData = sbFileData.ToString
        Catch ex As Exception
            'Call ShowErrPage(ex.Message) 'エラー画面に遷移する
        End Try
    End Function

#End Region

#Region "画面の設定"
    ''' <summary>
    ''' 画面設定
    ''' </summary>
    ''' <param name="strDBCountry"></param>
    ''' <remarks>
    ''' Series_Kataban=M（M4SA1など）のときに形番の最後に仕様書情報を付加し、
    ''' 標準納期を計算するように改修
    ''' </remarks>
    Private Sub subDispSet(ByRef strDBCountry As String)

        Dim objUnitPrice As New KHUnitPrice
        Dim dt_place As New DataTable
        Dim dt_strage_evaluation As New DataTable

        Dim lstCountry_Display As New ArrayList '表示国（順番なし）
        Dim lstCountry_Key As New ArrayList     '画面表示用（順番あり）
        Dim strFullKatabanSiyouNo As String = String.Empty
        Dim strCheck As String = String.Empty   'チェック区分
        Dim strStorageLocation As String = String.Empty     '保管場所
        Dim strEvaluationType As String = String.Empty       '評価タイプ

        Try
            '単価取得
            Call objUnitPrice.subPriceInfoSet(objCon, objKtbnStrc, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                              Me.objUserInfo.CountryCd, Me.objUserInfo.OfficeCd, strStorageLocation, strEvaluationType)

            '引当情報取得
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)

            '形番設定
            lblSeriesKat.Text = subSetKataban()

            '形番+仕様書Noセット 標準納期計算用
            strFullKatabanSiyouNo = lblSeriesKat.Text

            '名称設定
            If objKtbnStrc.strcSelection.strDivision = "3" Then
                '仕入品は名称固定
                Select Case selLang.SelectedValue.Trim
                    Case "ja", String.Empty
                        lblSeriesNm.Text = CdCst.GoogsName_Shiire.ja
                    Case "en"
                        lblSeriesNm.Text = CdCst.GoogsName_Shiire.en
                    Case "ko"
                        lblSeriesNm.Text = CdCst.GoogsName_Shiire.ko
                    Case "tw"
                        lblSeriesNm.Text = CdCst.GoogsName_Shiire.tw
                    Case "zh"
                        lblSeriesNm.Text = CdCst.GoogsName_Shiire.zh
                End Select
            Else
                lblSeriesNm.Text = objKtbnStrc.strcSelection.strGoodsNm
            End If

            '権限によって形番チェックと出荷場所の表示を制限する
            dt_Addinfo = subAddInfoDispGet(objCon, Me.objUserInfo.UserId, _
                    Me.objLoginInfo.SessionId, selLang.SelectedValue, Me.objUserInfo.AddInformationLvl, _
                    strFullKatabanSiyouNo, objKtbnStrc)



            '出荷場所の設定
            dt_place = fncSetPlaceCdList(lstCountry_Display, lstCountry_Key, Me.objUserInfo.CountryCd, dt_strage_evaluation, , strStorageLocation, strEvaluationType)
            Me.cmbPlace.Items.Clear()
            Me.cmbPlace.DataSource = dt_place
            Me.cmbPlace.DataBind()

            '保管場所と評価タイプの設定
            Me.cmbStrageEvaluation.Items.Clear()
            Me.cmbStrageEvaluation.DataSource = dt_strage_evaluation
            Me.cmbStrageEvaluation.DataBind()

            'セッションに保存
            If lstCountry_Key.Count > 0 Then Me.Session.Add("lstCountry_Key", lstCountry_Key)

            'ラベルの設定(価格情報)
            Call subSetLabel(strCheck)

            'ラベルの設定(注意事項)
            Call subSetInfo(strCheck, lstCountry_Display)

            '入力書式の設定
            Call subSetInputFormat()

            'セッションにFob価格をセット
            'If (Session("strPriceListFob") Is Nothing) Then
            Session("strPriceListFob") = ""
            Session("strCountryCod") = Me.objUserInfo.CountryCd
            'End If

            '匿名ユーザーの場合は数量入力エリアを非表示にする
            If Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
                PnlInput.Visible = False
            End If
        Catch ex As Exception
            AlertMessage(ex)
        Finally
            objUnitPrice = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 形番の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function subSetKataban() As String
        Dim intPositionInfo() As Integer
        Dim strResult As String = String.Empty

        '簡易マニホールドの判断
        If KHKataban.fncJudgeSimpleSpec(objCon, objKtbnStrc, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId) = True Then
            intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
            strResult = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen

            For intLoopCnt3 = 1 To UBound(intPositionInfo)
                Select Case objKtbnStrc.strcSelection.strSpecNo.ToString.Trim
                    Case "S", "T"
                        If intPositionInfo(intLoopCnt3) >= 10 Then
                            strResult = strResult & intPositionInfo(intLoopCnt3).ToString
                        Else
                            strResult = strResult & "0" & intPositionInfo(intLoopCnt3).ToString
                        End If
                    Case Else
                        'バルブの選択
                        If intPositionInfo(intLoopCnt3) >= 10 Then
                            Dim strKey As String = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,"
                            If intPositionInfo(intLoopCnt3) >= 10 And intPositionInfo(intLoopCnt3) <= 25 Then
                                strResult = strResult & strKey.Split(",")(intPositionInfo(intLoopCnt3) - 10)
                            Else
                                strResult = strResult & " "
                            End If
                        Else
                            strResult = strResult & intPositionInfo(intLoopCnt3).ToString
                        End If
                End Select
            Next
            If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
                strResult = Replace(strResult, "-ST", "")
                strResult = strResult & "-ST"
            End If
        Else
            strResult = objKtbnStrc.strcSelection.strFullKataban
        End If

        Return strResult
    End Function

    ''' <summary>
    ''' 出荷場所の取得
    ''' </summary>
    ''' <param name="lstCountry_Display">表示国（順番なし）</param>
    ''' <param name="lstCountry_Key">画面表示用（順番あり）</param>
    ''' <param name="strUserCountryCd">ユーザー国コード</param>
    ''' <param name="blnKBN">
    ''' 単価画面用区分　True：単価画面用　False：価格詳細画面用
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSetPlaceCdList(ByRef lstCountry_Display As ArrayList, _
                                       ByRef lstCountry_Key As ArrayList, _
                                       ByVal strUserCountryCd As String, _
                                       ByRef dtStrageEvaluation As DataTable, _
                                       Optional ByVal blnKBN As Boolean = True, _
                                       Optional ByVal strStorageLocation As String = "", _
                                       Optional ByVal strEvaluationType As String = "") As DataTable
        '出荷場所コードと名称
        Dim dtPlace As New DataTable

        '出荷場所変換に関するフラグ
        Dim retCtryItm As Boolean = False

        '結果テーブル初期化
        dtPlace = fncCreateTableByColumnNames(New List(Of String) From {"PlaceName", "PlaceID"})
        dtStrageEvaluation = fncCreateTableByColumnNames(New List(Of String) From {"StrageEvaluationName", "StrageEvaluationID"})

        Try
            '出荷場所候補
            Dim lstPlaceIDs As New ArrayList
            'ユーザの国コードにより表示可能な「国コード」
            Dim lstCountriesByCountryCd As New ArrayList

            'ユーザの国コードにより表示可能な「国コード」を取得
            lstCountriesByCountryCd = KHCountry.fncCountryTradeGet(objConBase, strUserCountryCd)

            'フル形番により出荷場所候補の取得
            lstPlaceIDs = fncGetPlaceIDByFullKataban(lstCountriesByCountryCd, retCtryItm, strUserCountryCd)

            '生産国レベルにより国出荷場所候補の取得
            lstPlaceIDs = fncGetPlaceIDByLevel(lstPlaceIDs, lstCountriesByCountryCd, strUserCountryCd)

            If lstPlaceIDs.Count > 1 Then
                'マニホールドの場合は選択した電磁弁の生産レベルにより生産国を検証する
                If objKtbnStrc.strcSelection.strSpecNo.Trim.Equals(String.Empty) OrElse _
                    objKtbnStrc.strcSelection.strSpecNo.Trim.Equals("00") Then
                Else

                    '生産国コードの検証
                    lstPlaceIDs = fncFilterCountryByByPlaceLevel(lstPlaceIDs)

                End If
            End If

            '第一ハイフン前により国出荷場所候補の取得
            lstCountry_Key = fncGetPlaceIDByHyphen(lstPlaceIDs, lstCountriesByCountryCd)

            '選択された国コードにより出荷場所コードを追加
            lstPlaceIDs = fncGetSelectPlaceID(lstPlaceIDs)

            '単価画面の場合は実行（価格詳細画面の場合は実行しない）
            If blnKBN Then

                '仕入品の場合CommonDBServiceから取得した保管場所と評価タイプをセット
                If objKtbnStrc.strcSelection.strDivision = "3" Then
                    subSetStrageEvaluation(strStorageLocation.ToString.PadRight(5) & strEvaluationType.PadRight(3), dtStrageEvaluation)
                End If

                '価格マスタにより出荷場所の追加
                subSetPlaceIDByOptions(lstPlaceIDs, lstCountriesByCountryCd)

                '出荷場所変換通知
                'If lstPlaceIDs.Count = 1 Then
                dtPlace = fncChangePlace(dtPlace, lstPlaceIDs, retCtryItm, dtStrageEvaluation)

                'End If

                '出荷場所名称の取得
                dtPlace = fncGetPlaceName(dtPlace, lstPlaceIDs, lstCountriesByCountryCd)

                '表示順番を調整
                If lstPlaceIDs.Count > 1 Then
                    dtPlace = fncSetOrder(dtPlace, lstCountriesByCountryCd, lstPlaceIDs, dtStrageEvaluation)
                End If

                'FRL判定用のメッセージを取得  2017/02/22 追加
                If dtPlace.Rows.Count > 0 Then

                    Select Case dtPlace.Rows(0).Item("PlaceID").ToString
                        '日本製を選んだ時のみFRL注意メッセージを表示
                        Case "P", "S", "K", "C", "JPN", "P55", "C55", "S55", "K55", "1001", "1002", "1003", "1004", "1005"
                            Dim strFrlMessage As String = fncFRLMessage()
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "ShowFrlMessage", "fncShowFrlMessage('" & strFrlMessage & "');", True)
                    End Select

                End If

            End If

            lstCountry_Display = lstPlaceIDs
        Catch ex As Exception
            Throw ex
        End Try

        Return dtPlace
    End Function

    ''' <summary>
    ''' ラベルの設定(価格情報)
    ''' </summary>
    ''' <param name="strCheck"></param>
    ''' <remarks></remarks>
    Private Sub subSetLabel(ByRef strCheck As String)
        Dim objKataban As New KHKataban
        'ラベル表示制御
        Dim strKeylvl As String = "1024,512,256,128,64,32,16,8,4,2,1"
        Dim strLevel() As String = strKeylvl.Split(",")
        Dim ccd_flg As Boolean = False 'FRL白色チェック区分１対応

        For inti As Integer = 0 To strLevel.Length - 1
            If dt_Addinfo Is Nothing Then Exit For
            Dim dr_display() As DataRow = dt_Addinfo.Select("strLevel='" & CInt(strLevel(inti)) & "'")
            Select Case CInt(strLevel(inti))
                Case 1 '形番チェック区分
                    If dr_display.Length > 0 Then
                        If dr_display(0)("strDisplay") = True Then
                            ccd_flg = KHKataban.subJapanChinaAmount(objKtbnStrc.strcSelection.strFullKataban)
                            If ccd_flg = True Then
                                Me.lblCheck.Text = "1"
                            Else
                                Me.lblCheck.Text = Replace(dr_display(0)("strValue").ToString, "Z", "")
                            End If
                        End If
                        strCheck = Me.lblCheck.Text
                    Else
                        Me.Label1.Visible = False
                        'チェック区分非表示
                        Me.lblCheck.Visible = False
                        Me.lblCheckZ.Visible = False
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), lblCheckZ.Text, _
                           "fnclblCheck('" & "ClsChk" & "');", True)

                    End If
                Case 2 '出荷場所
                    If dr_display.Length > 0 Then
                        If dr_display(0)("strDisplay") = True Then
                            Me.cmbPlace.Visible = True
                            '保管場所、評価タイプ
                            Me.cmbStrageEvaluation.Visible = True
                        Else
                            Me.cmbPlace.Visible = False
                            '保管場所、評価タイプ
                            Me.cmbStrageEvaluation.Visible = False
                        End If
                    Else
                        Me.cmbPlace.Visible = False
                        Me.Label2.Visible = False
                        '保管場所、評価タイプ
                        Me.cmbStrageEvaluation.Visible = False
                    End If
                    '表示のみ
                Case 4 '在庫情報
                Case 8 '担当者情報
                Case 16 '標準納期
                    If dr_display.Length > 0 Then
                        If dr_display(0)("strDisplay") = True Then
                            Dim str() As String = dr_display(0)("strValue").ToString.Split(CdCst.Sign.Delimiter.Pipe)
                            If str.Length = 2 Then
                                ccd_flg = False 'FRL白色 標準納期0日対応
                                ccd_flg = KHKataban.subJapanChinaAmount(objKtbnStrc.strcSelection.strFullKataban)
                                If ccd_flg = True Then
                                    If selLang.SelectedValue = "ja" Then
                                        Me.lblNouki.Text = "0日間(実稼働日)"
                                    Else
                                        Me.lblNouki.Text = "0day(the work days)"
                                    End If
                                Else
                                    Me.lblNouki.Text = str(0)
                                End If
                                Me.lblKosuu.Text = str(1)
                            End If
                        End If
                    Else
                        Me.lblNouki.Visible = False
                        Me.lblKosuu.Visible = False
                        Me.Label3.Visible = False
                        Me.Label4.Visible = False
                    End If
                Case 32 '販売数量単位 　※仕入品は表示しない
                    If dr_display.Length > 0 Then
                        If dr_display(0)("strDisplay") = True Then
                            '  If objKtbnStrc.strcSelection.strKatabanCheckDiv <> "5" Then
                            If objKtbnStrc.strcSelection.strDivision <> "3" Then
                                lblQtyUnit.Visible = True
                                Dim strQty() As String = dr_display(0)("strValue").ToString.Split(",")
                                If strQty.Length >= 1 Then Me.lblQtyUnit.Text = strQty(0)
                                If strQty.Length >= 2 Then
                                    lblQtyUnit1.Visible = True
                                    Me.lblQtyUnit1.Text = strQty(1)
                                End If
                            End If
                        End If
                    Else
                        Me.lblQtyUnit.Visible = False
                        Me.lblQtyUnit1.Visible = False
                    End If
                Case 64 'EL品情報
                    If dr_display.Length > 0 Then
                        If dr_display(0)("strDisplay") = True Then
                            Me.lblEL.Text = dr_display(0)("strValue").ToString
                        End If
                    Else
                        Me.lblEL.Visible = False
                        Me.Label5.Visible = False
                    End If
                Case 128 '中国輸出不可　　※仕入品は表示しない

                    Me.Label27.Visible = False

                    If dr_display.Length > 0 Then

                        If dr_display(0)("strDisplay") = True Then

                            '中国輸出不可設定
                            If objKataban.fncELKatabanCheck(objCon, objKtbnStrc.strcSelection.strFullKataban, "0") Then
                                'If objKtbnStrc.strcSelection.strKatabanCheckDiv <> "4" And objKtbnStrc.strcSelection.strKatabanCheckDiv <> "5" Then
                                If objKtbnStrc.strcSelection.strKatabanCheckDiv <> "4" And objKtbnStrc.strcSelection.strDivision <> "3" Then
                                    '中国生産品の場合は表示しない
                                    If Session("ChinaExportDisplay") Is Nothing Then
                                        Me.Label27.Visible = True
                                    Else
                                        Me.Label27.Visible = False
                                        Session.Remove("ChinaExportDisplay")
                                    End If
                                End If

                            End If

                        End If

                    End If

            End Select

        Next

        'RM1707001　中国対応メッセージ追加　2017/07/05
        '中国対応メッセージの表示設定
        subSetChinaMessage()

        '欧州対応メッセージの表示設定
        subSetEuropeMessage()

    End Sub

    ''' <summary>
    ''' ラベルの設定(注意事項など)
    ''' </summary>
    ''' <param name="lstCountry_Display"></param>
    ''' <remarks></remarks>
    Private Sub subSetInfo(ByRef strCheck As String, ByRef lstCountry_Display As ArrayList)

        '表示項目制御
        Call ChangeTitle(lstCountry_Display(0), strCheck)

        'ラベル表示設定
        If lstCountry_Display.Count > 1 Then '複数拠点生産品ラベルの表示制御
            If Me.cmbPlace.Visible = True Then
                Label8.Visible = True
            Else
                Label8.Visible = False
            End If

            Me.cmbPlace.Enabled = True
            Me.cmbStrageEvaluation.Enabled = True
        Else
            Label8.Visible = False
            Me.cmbPlace.Enabled = False
            Me.cmbPlace.ForeColor = Drawing.Color.Black
            Me.cmbPlace.BackColor = Drawing.Color.White

            Me.cmbStrageEvaluation.Enabled = False
            Me.cmbStrageEvaluation.ForeColor = Drawing.Color.Black
            Me.cmbStrageEvaluation.BackColor = Drawing.Color.White

        End If

        Select Case Left(objKtbnStrc.strcSelection.strFullKataban, 3) '仕入れ品のため掛率を変更しないで下さい.
            Case "PCU", "AHB"
                Me.Label6.Visible = True
            Case Else
                Me.Label6.Visible = False
        End Select

        '仕入れ品のため掛率を変更しないで下さい.
        Select Case objKtbnStrc.strcSelection.strDivision
            Case "3"
                If objKtbnStrc.strcSelection.strKatabanCheckDiv = "5" Then
                    Me.Label6.Visible = True
                End If
        End Select

        'RM1804032_注意喚起メッセージ追加
        Select Case Left(objKtbnStrc.strcSelection.strFullKataban, 3)
            Case "EKS"
                Me.Label42.Visible = True
            Case Else
                Me.Label42.Visible = False
        End Select

        Select Case Me.objUserInfo.UserClass '国内代理店はメッセージを表示させる
            Case CdCst.UserClass.DmAgentRs, CdCst.UserClass.DmAgentSs, CdCst.UserClass.DmAgentBs, _
                CdCst.UserClass.DmAgentGs, CdCst.UserClass.DmAgentPs
                Me.lblAction.Visible = True
            Case Else
                Me.lblAction.Visible = False
        End Select
    End Sub

    ''' <summary>
    ''' 入力書式の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInputFormat()
        If Me.objUserInfo.EditDiv = 0 Then
            EditDivOpt = CdCst.Sign.Dot
        Else
            EditDivOpt = CdCst.Sign.Comma
        End If

        'EditDivを変更する
        Me.txt_Rate.EditDiv = Me.objUserInfo.EditDiv
        Me.TextUnitPrice.EditDiv = Me.objUserInfo.EditDiv
        'Me.TextRateUnitPrice.EditDiv = Me.objUserInfo.EditDiv
        Me.TextCnt.EditDiv = Me.objUserInfo.EditDiv

        '小数点以下桁数を設定
        Me.txt_Rate.DecLen = 4
        Me.TextUnitPrice.DecLen = 2
        Me.TextCnt.DecLen = 0

        'カンマ編集を設定
        Me.txt_Rate.AllowMinus = False
        Me.TextUnitPrice.AllowMinus = False
        Me.TextCnt.AllowMinus = False

        'ゼロ可否を設定
        Me.txt_Rate.AllowZero = False
        Me.TextUnitPrice.AllowZero = False
        Me.TextCnt.AllowZero = False

        'マイナス可否を設定
        Me.txt_Rate.DispComma = True
        Me.TextUnitPrice.DispComma = True
        Me.TextCnt.DispComma = True
    End Sub

    ''' <summary>
    ''' 単価リスト作成
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <remarks>単価リストを作成する/ログファイルを出力する</remarks>
    Private Sub subPriceListMake(ByVal objKtbnStrc As KHKtbnStrc)

        '受注EDI T.Y
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current

        Dim objUnitPrice As New KHUnitPrice
        Dim strPriceList(,) As String = Nothing
        Dim objMsg As New ClsCommon
        Dim htPriceInfo As Hashtable = Nothing
        Dim strPriceFCA As String = Nothing
        Dim strPriceFCA2 As String = Nothing
        Dim dt_price As New DataTable

        Try
            '処理開始時間
            Dim processStartTime As Date = Now

            '単価表示情報取得
            Dim strCountryCode As String = Me.cmbPlace.SelectedValue

            Select Case strCountryCode
                '日本の場合国コードへ変換
                Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
                    strCountryCode = "JPN"
            End Select
            Call objUnitPrice.subPriceListSelect(objConBase, Me.objUserInfo.CountryCd, _
                                                 Me.selLang.SelectedValue, _
                                                 Me.objUserInfo.CurrencyCd, _
                                                 Me.objUserInfo.PriceDispLvl, _
                                                 strPriceList, strPriceFCA, _
                                                 strPriceFCA2, strCountryCode, objKtbnStrc)

            '2014/06/16 FOB対応
            For i As Integer = 0 To (strPriceList.Length / 5) - 1
                If strPriceList(i, 4) = KHCodeConstants.CdCst.UnitPrice.FobPrice Then
                    'I/F出力用
                    Session("strPriceListFob") = strPriceList(i, 2)
                    Session("strCurrencyCode") = strPriceList(i, 3)
                    '価格詳細画面用
                    SetPriceDetailControl()
                End If
            Next
            Session("strCountryCod") = Me.objUserInfo.CountryCd

            '単価情報をテーブルに保存
            dt_price = fncSavePriceInfoToTable(strPriceList)

            '価格リストの作成
            Call fncCreatTextCell(dt_price)

            'ログ出力(DBに保存しました、ファイル出力を廃棄してもいい)
            Call subOutputLog(strPriceList)

            ''DBに保存
            'If Me.Session("TestMode") Is Nothing Then
            '    'ホスト名
            '    Dim strHostName As String = String.Empty
            '    Dim strIP As String = Request.UserHostAddress
            '    Dim IPhostname As IPHostEntry = System.Net.Dns.GetHostEntry(strIP)
            '    strHostName = IPhostname.HostName
            '    If strHostName.Length > 0 Then strHostName = Left(strHostName, InStr(1, strHostName, ".") - 1)

            '    'Historyテーブルに保存
            '    Call bllTanka.subInsertPriceInfoToHistoryTable(strPriceList, dt_Addinfo, objKtbnStrc, _
            '                                                 Me.objUserInfo.UserId, strHostName, processStartTime)
            'End If

            'テキストの表示可否を設定
            Call subSetTextDisplay(httpCon, strPriceList)

        Catch ex As Exception
            AlertMessage(ex)
        Finally
            objUnitPrice = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 価格情報をテーブルに保存
    ''' </summary>
    ''' <remarks></remarks>
    Private Function fncSavePriceInfoToTable(ByVal strPriceList(,) As String) As DataTable
        '単価テーブル
        Dim dtResult As New DataTable
        '列名
        Dim strColumnNames As List(Of String) = New List(Of String) From {"ColumnKBN", "Kubun", "ViewPrice", "Price", "HdnPrice", "Tanni"}
        Dim intRowSt As Integer = 1
        Dim intRowEd As Integer = UBound(strPriceList)

        '価格テーブルの作成
        dtResult = fncCreateTableByColumnNames(strColumnNames)

        '列を追加する
        For intRow As Integer = intRowSt To intRowEd
            'RM1808066_一部例外GS表示
            '↓ 生産国は日本以外場合：定価、登録店、SS店、BS店、GS店、PS店を非表示にする
            Select Case objKtbnStrc.strcSelection.strMadeCountry
                Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
                Case Else
                    If Left(objKtbnStrc.strcSelection.strFullKataban, 5) = "SCWP2" Or _
                        Left(objKtbnStrc.strcSelection.strFullKataban, 5) = "SCWT2" Or _
                        Left(objKtbnStrc.strcSelection.strFullKataban, 4) = "SCWR" Or _
                        Left(objKtbnStrc.strcSelection.strFullKataban, 4) = "SCWS" Or _
                        Left(objKtbnStrc.strcSelection.strFullKataban, 4) = "RCS2" Or _
                        Left(objKtbnStrc.strcSelection.strFullKataban, 4) = "M4RD" Or _
                        Left(objKtbnStrc.strcSelection.strFullKataban, 4) = "M4RE" Or _
                        Left(objKtbnStrc.strcSelection.strFullKataban, 3) = "4RD" Or _
                        Left(objKtbnStrc.strcSelection.strFullKataban, 3) = "4RE" Or _
                        Left(objKtbnStrc.strcSelection.strFullKataban, 3) = "LCV" Then
                        Select Case strPriceList(intRow, 4)
                            Case CdCst.UnitPrice.ListPrice, CdCst.UnitPrice.RegPrice, CdCst.UnitPrice.SsPrice, _
                                 CdCst.UnitPrice.BsPrice, CdCst.UnitPrice.PsPrice
                                Continue For
                        End Select
                    Else
                        Select Case strPriceList(intRow, 4)
                            Case CdCst.UnitPrice.ListPrice, CdCst.UnitPrice.RegPrice, CdCst.UnitPrice.SsPrice, _
                                 CdCst.UnitPrice.BsPrice, CdCst.UnitPrice.GsPrice, CdCst.UnitPrice.PsPrice
                                Continue For
                        End Select
                    End If
            End Select
            '↑ 生産国は日本以外場合：定価、登録店、SS店、BS店、GS店、PS店を非表示にする 

            Dim dr As DataRow = dtResult.NewRow

            '項目区分
            dr("ColumnKBN") = fncConvertColumnKBN(strPriceList(intRow, 4))

            'タイトル
            dr("Kubun") = strPriceList(intRow, 1)

            '価格
            If Me.objUserInfo.EditDiv = "0" Then
                Dim str() As String = strPriceList(intRow, 2).Split(".")
                If str.Length >= 2 Then
                    dr("Price") = FormatNumber(strPriceList(intRow, 2), str(1).Length)
                Else
                    dr("Price") = FormatNumber(strPriceList(intRow, 2), 0)
                End If
            Else
                dr("Price") = ClsCommon.fncPriceDot(strPriceList(intRow, 2))
            End If

            '表示価格
            If dr("Price").ToString.Length <= 0 OrElse CDbl(dr("Price")) = 0D Then
                dr("Tanni") = String.Empty
                dr("ViewPrice") = ClsCommon.fncGetMsg(selLang.SelectedValue, "I5220")
            Else
                dr("Tanni") = strPriceList(intRow, 3)
                If dr("Price").ToString.Length > 0 Then dr("ViewPrice") = dr("Price").ToString & strPriceList(intRow, 3).ToString.PadLeft(5, " ")
            End If

            dr("HdnPrice") = strPriceList(intRow, 2)

            dtResult.Rows.Add(dr)
        Next

        '国内代理店用に追加
        Select Case Me.objUserInfo.UserClass
            Case CdCst.UserClass.DmAgentRs, CdCst.UserClass.DmAgentSs, CdCst.UserClass.DmAgentBs, _
                CdCst.UserClass.DmAgentGs, CdCst.UserClass.DmAgentPs
                Dim dr As DataRow = dtResult.NewRow
                dr("ViewPrice") = CdCst.FixedMessage.PriceJPY
                dr("Kubun") = ""
                dr("Price") = 0
                dr("Tanni") = ""
                dr("HdnPrice") = 0
                dtResult.Rows.Add(dr)
        End Select

        Return dtResult
    End Function

    ''' <summary>
    ''' 比較のためログファイルを出力
    ''' </summary>
    ''' <param name="strPriceList"></param>
    ''' <remarks></remarks>
    Private Sub subOutputLog(ByVal strPriceList(,) As String)
        If dt_Addinfo Is Nothing Then
            '権限によって形番チェックと出荷場所の表示を制限する
            dt_Addinfo = subAddInfoDispGet(objCon, Me.objUserInfo.UserId, _
                    Me.objLoginInfo.SessionId, selLang.SelectedValue, Me.objUserInfo.AddInformationLvl, _
                     lblSeriesKat.Text, objKtbnStrc)
        End If
        'テキスト出力(比較のため、１ヶ月削除保留)
        Call bllTanka.subLogFileOutput(objCon, strPriceList, dt_Addinfo, Me.objUserInfo.UserId, _
                                       Me.objLoginInfo.SessionId, Me.objUserInfo.CountryCd, selLang.SelectedValue)
        Call bllTanka.subLogOutput(objConBase, objCon, strPriceList, dt_Addinfo, Me.objUserInfo.UserId, _
                                   Me.objLoginInfo.SessionId, Me.objUserInfo.CountryCd, selLang.SelectedValue)
    End Sub

    ''' <summary>
    ''' テキストの表示可否を設定
    ''' </summary>
    ''' <param name="httpCon"></param>
    ''' <param name="strPriceList"></param>
    ''' <remarks></remarks>
    Private Sub subSetTextDisplay(ByVal httpCon As System.Web.HttpContext, ByVal strPriceList(,) As String)
        Me.Label14.Visible = True
        Me.Label15.Visible = True
        Me.Label16.Visible = True
        Me.Label17.Visible = True
        Me.Label18.Visible = True
        Me.Label19.Visible = True
        Me.txt_Rate.Visible = True
        Me.TextUnitPrice.Visible = True
        Me.TextRateUnitPrice.Visible = True
        Me.TextCnt.Visible = True
        Me.TextMoney.Visible = True
        Me.TextTax.Visible = True
        Me.TextAmount.Visible = True

        If httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) IsNot Nothing Then
            'T.Y受注EDI連携用データ取得
            'Call Me.subJuchuEdiWSDbIO(Me.objUserInfo.UserId, objKtbnStrc.strcSelection.strFullKataban, _

            '価格情報の非表示
            Me.Label14.Visible = False
            Me.Label15.Visible = False
            Me.Label16.Visible = False
            Me.Label17.Visible = False
            Me.Label18.Visible = False
            Me.Label19.Visible = False
            Me.txt_Rate.Visible = False
            Me.TextUnitPrice.Visible = False
            Me.TextRateUnitPrice.Visible = False
            Me.TextCnt.Visible = False
            Me.TextMoney.Visible = False
            Me.TextTax.Visible = False
            Me.TextAmount.Visible = False
        End If

        '海外代理店は金額と消費税を非表示にする
        '海外代理店「E-con」を追加
        Select Case Me.objUserInfo.UserClass
            Case CdCst.UserClass.OsAgentCs, CdCst.UserClass.OsAgentLs
                Me.Label17.Visible = False
                Me.Label19.Visible = False
                Me.TextTax.Visible = False
                Me.TextAmount.Visible = False
        End Select

        If Me.objUserInfo.CountryCd <> "JPN" Then
            Me.Label17.Visible = False
            Me.Label19.Visible = False
            Me.TextTax.Visible = False
            Me.TextAmount.Visible = False
        End If
    End Sub

    ''' <summary>
    ''' 価格リストの設定
    ''' </summary>
    ''' <param name="dt_price"></param>
    ''' <remarks></remarks>
    Private Sub fncCreatTextCell(dt_price As DataTable)
        Try
            'ラベルタイトル設置
            Dim dt_Title As DataTable = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHTanka, selLang.SelectedValue)
            Dim dr() As DataRow = dt_Title.Select("label_seq='12' AND label_div='L'")

            'タイトルの設定
            GVPrice.Columns(0).HeaderText = dr(0)("label_content").ToString
            dr = dt_Title.Select("label_seq='13' AND label_div='L'")
            GVPrice.Columns(1).HeaderText = dr(0)("label_content").ToString

            GVPrice.DataSource = dt_price
            GVPrice.DataBind()

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 画面上のチェック区分、出荷場所、標準納期、適用個数、E/L該当品区分の表示制御
    ''' </summary>
    ''' <param name="strCountry"></param>
    ''' <param name="strCheck"></param>
    ''' <remarks></remarks>
    Private Sub ChangeTitle(strCountry As String, strCheck As String)
        Dim strKeylvl As String = "1024,512,256,128,64,32,16,8,4,2,1"
        Dim strLevel() As String = strKeylvl.Split(",")
        Dim TitleLevel As Long = Me.objUserInfo.AddInformationLvl
        '受注EDI
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim objKataban As New KHKataban

        Select Case strCountry
            Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
                '権限によて画面項目をすべて表示する
                For inti As Integer = 0 To strLevel.Length - 1
                    If TitleLevel >= CInt(strLevel(inti)) Then
                        Select Case CInt(strLevel(inti))
                            Case 128 '中国輸出不可
                            Case 64 'EL品情報
                                Me.lblEL.Visible = True
                                Me.Label5.Visible = True
                            Case 32 '販売数量単位
                                Me.lblQtyUnit.Visible = True
                                Me.lblQtyUnit1.Visible = True
                            Case 16 '標準納期
                                Me.lblNouki.Visible = True
                                Me.lblKosuu.Visible = True
                                Me.Label3.Visible = True
                                Me.Label4.Visible = True
                            Case 8 '担当者情報
                                Me.Button10.Visible = True
                            Case 4 '在庫情報
                            Case 2 '出荷場所
                                Me.cmbPlace.Visible = True
                                Me.Label2.Visible = True

                                '保管場所＆評価タイプ 
                                Me.cmbStrageEvaluation.Visible = True

                                '在庫検索
                                If Me.Button10.Visible = True Then
                                    If objKataban.fncStockInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, selLang.SelectedValue, _
                                        objKtbnStrc.strcSelection.strPlaceCd, 0, 0, "") Then
                                        Me.Button11.Visible = True
                                    Else
                                        Me.Button11.Visible = False
                                    End If
                                End If

                            Case 1 '形番チェック区分
                                Me.Label1.Visible = True
                                Me.lblCheck.Visible = True
                                Me.lblCheckZ.Visible = True
                        End Select
                        TitleLevel -= CInt(strLevel(inti))
                    End If
                Next

                '受注EDIボタン
                If httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) IsNot Nothing Then
                    Button5.Visible = True
                End If
            Case Else
                '権限ある場合、形番チェックと出荷場所のみ表示する
                For inti As Integer = 0 To strLevel.Length - 1
                    If TitleLevel >= CInt(strLevel(inti)) Then
                        Select Case CInt(strLevel(inti))
                            Case 128 '中国輸出不可
                            Case 64 'EL品情報
                                Me.lblEL.Visible = False
                                Me.Label5.Visible = False
                            Case 32 '販売数量単位
                                Me.lblQtyUnit.Visible = True
                                Me.lblQtyUnit1.Visible = True
                            Case 16 '標準納期
                                Me.lblNouki.Visible = False
                                Me.lblKosuu.Visible = False
                                Me.Label3.Visible = False
                                Me.Label4.Visible = False
                            Case 8 '担当者情報
                                Me.Button10.Visible = True
                            Case 4 '在庫情報
                            Case 2 '出荷場所
                                Me.cmbPlace.Visible = True
                                Me.Label2.Visible = True

                                '保管場所＆評価タイプ　表示しない 
                                Me.cmbStrageEvaluation.Visible = False

                                '在庫検索（海外生産品は出力しない）
                                Me.Button11.Visible = False

                            Case 1 '形番チェック区分
                                Me.Label1.Visible = True
                                Me.lblCheck.Visible = True
                                Me.lblCheckZ.Visible = True
                                Select Case Me.objUserInfo.UserClass
                                    Case CdCst.UserClass.DmAgentRs, CdCst.UserClass.DmAgentSs, _
                                        CdCst.UserClass.DmAgentBs, CdCst.UserClass.DmAgentGs, _
                                        CdCst.UserClass.DmAgentPs '国内代理店　形番チェック説明
                                        Select Case strCheck
                                            Case CdCst.KatabanChackDiv.Stock
                                                Me.lblCheck.Text = CdCst.KatabanChackDivName.Stock
                                            Case CdCst.KatabanChackDiv.Standard
                                                Me.lblCheck.Text = CdCst.KatabanChackDivName.Standard
                                            Case CdCst.KatabanChackDiv.Special
                                                Me.lblCheck.Text = CdCst.KatabanChackDivName.Special
                                            Case CdCst.KatabanChackDiv.Parts
                                                Me.lblCheck.Text = CdCst.KatabanChackDivName.Parts
                                        End Select
                                End Select
                        End Select
                        TitleLevel -= CInt(strLevel(inti))
                    End If
                Next

                '受注EDIボタン
                Button5.Visible = False

        End Select

        objKataban = Nothing

        Me.PnlSelect.Visible = False
        'Add by Zxjike 2014/03/25 ↓
        Me.Pnl10.Visible = False
        Me.Pnl11.Visible = False
        Me.Pnl14.Visible = False
        Me.Pnl15.Visible = False
        Me.Pnl17.Visible = False
        Me.Pnl18.Visible = False
        Me.Pnl19.Visible = False
        Me.Pnl20.Visible = False
        Me.Pnl21.Visible = False
        Me.Pnl22.Visible = False
        Me.Pnl23.Visible = False
        Me.Pnl24.Visible = False
        'Add by Zxjike 2014/03/25 ↑

        '国内ユーザーのみ
        If Me.objUserInfo.CountryCd = "JPN" And Me.objUserInfo.OfficeCd <> "II2" Then
            Dim htSelInfo As Hashtable = Nothing
            Dim selectFlg As String

            selectFlg = 0
            Select Case Left(objKtbnStrc.strcSelection.strFullKataban, 5)
                Case "MN4GB"
                    'For intLoopCnt = 1 To 18
                    '    '仕様書形番が選択されていること、かつ、仕様書使用数が入っていること
                    '    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length <> 0 And _
                    '       objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    '        'セレクト対象の仕様書形番が選択されているか
                    '        SelectM4G = KHKataban.fncSelectCatalogInfo4G(objCon, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt))
                    '        'セレクト対象の仕様書形番が選択されていないとき
                    '        If SelectM4G = False Then
                    '            selectFlg = 1 'セレクト対象品除外フラグをたてる
                    '        End If
                    '    End If
                    'Next
                    ''セレクト対象品除外フラグがたっていないとき
                    'If selectFlg = 0 Then
                    If KHKataban.fncSelectCatalogInfo4G(objCon, objKtbnStrc.strcSelection.strOptionKataban, objKtbnStrc.strcSelection.intQuantity) Then
                        'セレクト対応品表示対応
                        'セレクト品マスタ検索
                        If KHKataban.fncSelectCatalogInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, htSelInfo) Then
                            Me.PnlSelect.Visible = True
                            If htSelInfo("MsgKbn").ToString.Equals("1") Then
                                Me.Pnl14.Visible = True
                                Me.Pnl15.Visible = True
                            Else
                                Me.Pnl10.Visible = True
                                Me.Pnl11.Visible = True
                            End If
                            Me.txtSelKosu.Text = htSelInfo("DispKosu").ToString
                            Me.txtSelNoki.Text = htSelInfo("DispNoki").ToString
                        End If
                    End If
                Case Else
                    'セレクト対応品表示対応
                    'セレクト品マスタ検索
                    If KHKataban.fncSelectCatalogInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, htSelInfo) Then
                        Me.PnlSelect.Visible = True
                        Dim strMsgKbn As String = htSelInfo("MsgKbn").ToString
                        Select Case strMsgKbn
                            Case "1"
                                Me.Pnl14.Visible = True
                                Me.Pnl15.Visible = True
                            Case "2"
                                Me.Pnl17.Visible = True
                                Me.Pnl18.Visible = True
                            Case "3"
                                Me.Pnl19.Visible = True
                                Me.Pnl20.Visible = True
                            Case Else '"0"
                                Me.Pnl10.Visible = True
                                Me.Pnl11.Visible = True
                        End Select
                        Me.txtSelKosu.Text = htSelInfo("DispKosu").ToString
                        Me.txtSelNoki.Text = htSelInfo("DispNoki").ToString
                    End If
            End Select
        Else
            'Add by Zxjike 2014/03/25 海外営業統括部、海外販社（現地採用者）、海外販社（日本駐在員）RM1403074
            dt_Addinfo = subAddInfoDispGet(objCon, Me.objUserInfo.UserId, _
                                           Me.objLoginInfo.SessionId, selLang.SelectedValue, Me.objUserInfo.AddInformationLvl, _
                                           objKtbnStrc.strcSelection.strFullKataban, objKtbnStrc)

            If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                '出荷場所の表示権限がある時のみ
                Dim dt_Select As New DataTable '海外セレクト品を検索（フル形番）
                Dim strCurShipCd As String = objKtbnStrc.strcSelection.strMadeCountry

                strCurShipCd = strCountry
                dt_Select = KHKataban.fncELKatabanCheck_Kaigai(objCon, strCurShipCd, selLang.SelectedValue, objKtbnStrc.strcSelection.strFullKataban)
                If Not dt_Select Is Nothing AndAlso dt_Select.Rows.Count > 0 Then  'データあれば
                    Me.PnlSelect.Visible = True
                    Me.Pnl21.Visible = True
                    Me.Pnl22.Visible = True
                    Me.Pnl23.Visible = True
                    Dim dr As DataRow = dt_Select.Rows(0)
                    Me.txtSelKosu.Text = dr("kosu_nm").ToString.Replace("{0}", dr("kosu").ToString)
                    Me.txtSelNoki.Text = dr("nouki_nm").ToString.Replace("{0}", dr("nouki").ToString)
                End If
            End If

            'ユーザー特殊注意メッセージ
            'RM1706020 メッセージ表示条件変更 テーブル変更に伴う引数追加  2017/06/14 変更 
            If Not objKtbnStrc.strcSelection.strFullKataban.Equals(String.Empty) Then
                Select Case Me.objUserInfo.CountryCd
                    Case "IND"

                        If KHKataban.fncSpecialUserMessage(objCon, objKtbnStrc.strcSelection.strFullKataban.Split("-")(0), Me.objUserInfo.CountryCd, "IND") Then
                            Me.Pnl24.Visible = True
                        Else
                            Me.Pnl24.Visible = False
                        End If

                    Case "E90"

                        If KHKataban.fncSpecialUserMessage(objCon, objKtbnStrc.strcSelection.strFullKataban.Split("-")(0), Me.objUserInfo.CountryCd, "E90") Then
                            Me.Pnl25.Visible = True
                        Else
                            Me.Pnl25.Visible = False
                        End If

                    Case "E08"

                        If KHKataban.fncSpecialUserMessage(objCon, objKtbnStrc.strcSelection.strFullKataban.Split("-")(0), Me.objUserInfo.CountryCd, "E08") Then
                            Me.Pnl26.Visible = True
                        Else
                            Me.Pnl26.Visible = False
                        End If

                    Case "EUR"

                        'pnl25とpnl26の両方について判定を行う
                        If KHKataban.fncSpecialUserMessage(objCon, objKtbnStrc.strcSelection.strFullKataban.Split("-")(0), Me.objUserInfo.CountryCd, "E90") Then
                            Me.Pnl25.Visible = True
                        Else
                            Me.Pnl25.Visible = False
                        End If

                        If KHKataban.fncSpecialUserMessage(objCon, objKtbnStrc.strcSelection.strFullKataban.Split("-")(0), Me.objUserInfo.CountryCd, "E08") Then
                            Me.Pnl26.Visible = True
                        Else
                            Me.Pnl26.Visible = False
                        End If

                    Case "MEX"      'RM1707049_2017/7/26_CZ対応

                        If KHKataban.fncSpecialUserMessage(objCon, objKtbnStrc.strcSelection.strFullKataban.Split("-")(0), Me.objUserInfo.CountryCd, "MEX") Then
                            Me.Pnl24.Visible = True
                        Else
                            Me.Pnl24.Visible = False
                        End If
                End Select

            End If
        End If
    End Sub

    ''' <summary>
    ''' 「***生産品」欄の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetPlaceMark(ByVal strCountry As String, ByVal lstPlaceIDCount As Integer)
        Dim lstCountry_Key As New ArrayList
        Dim bolMaybe As Boolean = False

        If strCountry.Equals(String.Empty) AndAlso cmbPlace.SelectedValue.Equals(String.Empty) Then
            Exit Sub
        End If

        If Not Me.Session("lstCountry_Key") Is Nothing Then lstCountry_Key = Me.Session("lstCountry_Key")
        If lstCountry_Key.Contains(cmbPlace.SelectedValue) Then bolMaybe = True

        Label7.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False

        'Select Case cmbPlace.SelectedValue
        Session.Remove("ChinaExportDisplay")
        Select Case strCountry
            Case "PRC", "CKD China", "CKD 中国"
                If bolMaybe Then
                    Label11.Visible = True
                    Label11.Width = Unit.Percentage(80)
                    Label11.BackColor = Drawing.Color.Pink
                    Label11.ForeColor = Drawing.Color.LightYellow
                    Me.Label9.Visible = True
                Else
                    Label7.Visible = True
                    Label7.Width = Unit.Percentage(80)
                    Label7.BackColor = Drawing.Color.Red
                    Label7.ForeColor = Drawing.Color.Yellow
                    Session.Add("ChinaExportDisplay", False)
                    Label27.Visible = False
                End If
            Case "KTA"
                'ADD BY YGY 20140722    台湾K生産品
                Label7.Visible = True
                Label7.Width = Unit.Percentage(80)
                Label7.BackColor = Drawing.Color.Blue
                Label7.ForeColor = Drawing.Color.White
            Case "TYO"
                'ADD BY YGY 20140804    台湾T生産品
                Label7.Visible = True
                Label7.Width = Unit.Percentage(80)
                Label7.BackColor = Drawing.Color.Blue
                Label7.ForeColor = Drawing.Color.White
            Case "MDN"
                'ADD BY YGY 20140804    台湾M生産品
                Label7.Visible = True
                Label7.Width = Unit.Percentage(80)
                Label7.BackColor = Drawing.Color.Blue
                Label7.ForeColor = Drawing.Color.White
            Case "OMA"
                'ADD BY YGY 20141006    タイOMA生産品
                Label7.Visible = True
                Label7.Width = Unit.Percentage(80)
                Label7.BackColor = Drawing.Color.Orange
                Label7.ForeColor = Drawing.Color.Red
            Case "IDN"
                'ADD BY YGY 20150709
                Label7.Visible = True
                Label7.Width = Unit.Percentage(80)
                Label7.BackColor = Drawing.Color.Purple
                Label7.ForeColor = Drawing.Color.White
            Case "KOR"
                'ADD BY YGY 20151023
                Label7.Visible = True
                Label7.Width = Unit.Percentage(80)
                Label7.BackColor = Drawing.Color.Black
                Label7.ForeColor = Drawing.Color.White
            Case "THA", "CKD Thailand", "CKD タイ", "THF"
                If bolMaybe Then
                    Label11.Visible = True
                    Label11.Width = Unit.Percentage(80)
                    Label11.BackColor = Drawing.Color.LemonChiffon
                    Label11.ForeColor = Drawing.Color.HotPink
                    Me.Label9.Visible = True
                Else
                    Label7.Visible = True
                    Label7.Width = Unit.Percentage(80)
                    Label7.BackColor = Drawing.Color.Yellow
                    Label7.ForeColor = Drawing.Color.Red
                End If
            Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
                Label10.BackColor = Drawing.Color.White
                Label10.ForeColor = Drawing.Color.Red
                'ラベル設定
                '国内ユーザの場合「生産品」欄を表示しない
                If Not objUserInfo.CountryCd.Equals("JPN") Then
                    If lstPlaceIDCount > 1 Then   '複数生産拠点
                        Label10.Visible = True
                        Label10.Width = Unit.Percentage(80)
                    End If
                Else
                    Label10.Visible = False
                End If
            Case "CJA"
                'ADD BY 斉藤 20160708   中国Ｃ生産品
                Label7.Visible = True
                Label7.Width = Unit.Percentage(80)
                Label7.BackColor = Drawing.Color.DeepPink
                Label7.ForeColor = Drawing.Color.Yellow
            Case Else
                Dim dt_country As DataTable = KHCountry.fncGetCountryName(objConBase)
                If Not dt_country Is Nothing Then
                    Dim dr_other() As DataRow = dt_country.Select("country_cd='" & cmbPlace.SelectedValue & "' AND language_cd='en'")
                    If dr_other.Length > 0 Then
                        Label11.BackColor = Drawing.Color.White
                        Label11.ForeColor = Drawing.Color.Red
                        If bolMaybe Then
                            Label11.Visible = True
                            Label11.Width = Unit.Percentage(80)
                            Label9.Visible = True
                        Else
                            Label7.Visible = True
                            Label7.Width = Unit.Percentage(80)
                        End If
                    End If
                End If
        End Select

        '出荷場所が表示しない場合は全部非表示にする
        If Me.cmbPlace.Visible = False Then
            Label7.Visible = False
            Label9.Visible = False
            Label10.Visible = False
            Label11.Visible = False
        End If
    End Sub

    ''' <summary>
    ''' 画面の国名
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetCountryName()
        If Not Me.cmbPlace.SelectedItem Is Nothing Then
            '全ての国コードと国名
            Dim strPlace As String = String.Empty
            Dim drTmp() As DataRow
            Dim dt_country As DataTable = KHCountry.fncGetAllCountryName(objConBase)

            strPlace = Me.cmbPlace.SelectedItem.Value

            Select Case strPlace
                Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
                    strPlace = "JPN"
            End Select

            '対応する出荷場所名を取得
            drTmp = dt_country.Select("country_cd='" & strPlace & "' AND language_cd='" & Me.selLang.SelectedValue & "'")

            If drTmp IsNot Nothing AndAlso drTmp.Count > 0 Then
                Label7.Text = Label7.Text.Replace("[1]", drTmp(0)("country_nm").ToString)
                Label9.Text = Label9.Text.Replace("[1]", drTmp(0)("country_nm").ToString)
                Label10.Text = Label10.Text.Replace("[1]", drTmp(0)("country_nm").ToString)
                Label11.Text = Label11.Text.Replace("[1]", drTmp(0)("country_nm").ToString)
            End If

            'タイ特殊対応
            If strPlace.Equals("THF") Then
                Label7.Visible = False
                Label10.Visible = True
                Label10.Width = Unit.Percentage(100)
                Label10.BackColor = Drawing.Color.Yellow
                Label10.ForeColor = Drawing.Color.Red
            End If

            If objUserInfo.CountryCd.Equals("PRC") Then

            End If

            'RM1707049_2017/7/26_CZ対応
            'RM170****_2017/8/24_メキシコメッセージ対応
            Select Case objUserInfo.CountryCd
                Case "IND"
                    Label35.Text = Label35.Text.Replace("[1]", "CKD India CZ17101")     'インド
                Case "MEX"
                    Label35.Text = Label35.Text.Replace("[1]", "New price from September 1, 2017 (CZ17102)")         'メキシコ
            End Select

            '特価決裁Noメッセージ表示（購入価格を表示するユーザーのみメッセージ表示）
            If strPlace = "JPN" And objUserInfo.PriceDispLvl > 63 Then
                Label41.Visible = True
                Label41.Text = objKtbnStrc.strcSelection.strAuthorizationNo
            Else
                Label41.Visible = False
                Label41.Text = Nothing
            End If

            'Made in -> Made by    「China」に含まれる「in」が変換されないように「 in 」スペースを追加
            If strPlace.Equals("KTA") OrElse strPlace.Equals("TYO") OrElse strPlace.Equals("MDN") OrElse strPlace.Equals("OMA") OrElse strPlace.Equals("CJA") Then
                If Me.selLang.SelectedValue.Equals("en") Then
                    Label7.Text = Label7.Text.Replace(" in ", " by ")
                End If
            End If

            'RM17***** 出荷場所ラベル項目をハイパーリンクに変更
            Label2.Text = "<a href=""WebUserControl/MessagePage.htm?" & selLang.SelectedValue & _
                """  target=""_blank"" onclick="" window.open('WebUserControl/MessagePage.htm?" & _
                selLang.SelectedValue & "', '_blank', 'width=500, height=500'); return false; "">" & Label2.Text & "</a>"

            'RM1808088_注意事項ラベルタイトル変更
            If objUserInfo.CountryCd = CdCst.CountryCd.DefaultCountry Then
                Select Case Left(objKtbnStrc.strcSelection.strFullKataban, 3)
                    Case "KBX"
                        If objKtbnStrc.strcSelection.strKeyKataban = "B" Then
                            Me.Label43.Visible = True
                            Me.Label43.Text = Me.Label43.Text.Replace("[1]", "（Ｗ００）")
                        Else
                            Me.Label43.Visible = False
                        End If
                    Case "ETV", "ECV"
                        Me.Label43.Visible = True
                        Me.Label43.Text = Me.Label43.Text.Replace("[1]", "")
                    Case Else
                        Me.Label43.Visible = False
                End Select
                'Me.Label43 = New Font("MS UI Gothic", 10, FontStyle.Regular)
            Else
                Me.Label43.Visible = False
            End If

        End If
    End Sub



    ''' <summary>
    ''' ボタン表示可否の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subBtnView()

        '受注EDI T.Y
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current

        Dim objOption As New KHOptionCtl
        Dim objKtbnStrc As New KHKtbnStrc
        Dim strUseFncInfo() As String = Nothing
        'Dim bolSpecInput As Boolean
        Dim bolSpecOutput As Boolean
        Dim strPriceFCA As String = Nothing
        Dim strPriceFCA2 As String = Nothing
        Dim strMessageCd As String = Nothing

        Try
            'bolSpecInput = False
            bolSpecOutput = False

            '引当形番情報取得
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

            Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban.Trim
            Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban.Trim

            If Len(objKtbnStrc.strcSelection.strSpecNo.Trim) <> 0 Then
                Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                    Case "00"
                        'ページ遷移(ロッド先端形状オーダーメイド寸法入力画面)
                        'If Len(objKtbnStrc.strcSelection.strRodEndOption) > 0 Then
                        '    bolSpecInput = True
                        'End If
                    Case "01", "02", "03", "04", "05", "06", "07", "08", "10", "11", "13", "14", "15", "16", "96"
                        'bolSpecInput = True
                        bolSpecOutput = True
                    Case "09"
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                            'bolSpecInput = True
                            bolSpecOutput = True
                        End If
                    Case "12", "18", "19", "20", "21", "22", "23", "94", "95"
                        If objOption.fncVaccumMixCheck(objKtbnStrc) Then
                            'bolSpecInput = True
                            bolSpecOutput = True
                        End If
                    Case "17"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Then
                            'bolSpecInput = True
                            bolSpecOutput = True
                        End If
                    Case "64", "66", "68", "70", "72"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            'bolSpecInput = True
                            bolSpecOutput = True
                        End If
                    Case "51"
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                            'bolSpecInput = True
                            bolSpecOutput = True
                        End If
                        'RM1805001_4Rシリーズ追加
                    Case "52", "60", "61", "62", "63", "65", "67", "69", "71", "S", "T", "U", "A4", "A5", "A6", "A7", "A8"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            'bolSpecInput = True
                            bolSpecOutput = True
                        End If
                    Case "53", "73", "74", "75", "76", "77", "78", "79", "80", "81", _
                         "82", "83", "84", "85", "86", "87", "88", "93"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "80" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "81" Then
                            'bolSpecInput = True
                            bolSpecOutput = True
                        End If
                    Case "89", "90", "98"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            'bolSpecInput = True
                            bolSpecOutput = True
                        End If
                    Case "54", "55", "56", "57", "58", "59", "91", "92"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            'bolSpecInput = True
                            bolSpecOutput = True
                        End If
                        'Case "A1", "A2", "A3"    CHANGED BY YGY 20141027
                    Case "A1", "A2", "A3", "A9", "B1", "B2", "B3", "B4"
                        'bolSpecInput = True
                        bolSpecOutput = True
                End Select
            End If

            '権限にI/Fボタンの表示を制限する
            'Call KHKataban.subUseFncInfoGet(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
            '                                   selLang.SelectedValue, Me.objUserInfo.UseFunctionLvl, _
            '                                   strUseFncInfo, objKtbnStrc)
            '仕様出力ボタン表示
            If bolSpecOutput Then
                Button2.Visible = True
            Else
                Button2.Visible = False
            End If

            'If strUseFncInfo(2) Then
            '    Button3.Visible = True
            'Else
            Button3.Visible = False
            'End If

            '受注EDI T.Y セッションが有効であれば[EDI]ボタンを表示し、[ファイル出力]ボタンを非表示する。
            If httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) IsNot Nothing Then
                'プラントが日本以外の場合は受注EDIボタン表示しない
                Select Case cmbPlace.SelectedValue
                    Case "1002", "1003", "1004", "1005"
                        Button5.Visible = True
                    Case Else
                        Button5.Visible = False
                End Select

                Button6.Visible = False
            Else
                Button5.Visible = False
                '匿名ユーザーに公開するためにファイル出力ボタンを非表示にする
                Button6.Visible = False
            End If

            ''RM1809***_jsonファイル出力用ボタン追加
            'If Me.objUserInfo.UserClass = CdCst.UserClass.InfoSysForceSysAdmin And bolSpecOutput And objKtbnStrc.strcSelection.strSpecNo.Trim = "04" Then
            '    Button12.Visible = True
            'Else
            '    Button12.Visible = False
            'End If

            'RM1808098_CE取得有無確認ボタン追加
            If Left(objKtbnStrc.strcSelection.strFullKataban, 5) = "ADK11" And Me.objUserInfo.CountryCd <> "JPN" Then
                Button13.Visible = True
                Button13.ForeColor = Drawing.Color.Red
            Else
                Button13.Visible = False
            End If

            'ロッド価格メッセージ
            Select Case strSeriesKata
                Case "AMD3", "AMD4", "AMD5"
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "3" And _
                       Right(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "R" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "B" And _
                       ((objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "10UP" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "10BUP") Or _
                       (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "25UP" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "25BUP")) Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "3R" Then
                            strMessageCd = "W8930"
                            'エラーメッセージ設定
                            Call AlertMessage(strMessageCd)
                        End If
                    End If
                Case "GAMD3", "GAMD4", "GAMD5", "AMG3", "AMG4", "AMG5"
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "3R" Then
                        strMessageCd = "W8930"
                        'エラーメッセージ設定
                        Call AlertMessage(strMessageCd)
                    End If
                Case Else
            End Select
            '↓2012/07/09　追加(価格算出のみシリーズ（メッセージ表示）)
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "P1100-UN", "P4100-UN"
                    strMessageCd = "W8920"
                    'エラーメッセージ設定
                    Call AlertMessage(strMessageCd)
                    'RM1801***_非表示
                    'Case "SWD", "MWD"
                    '    '201502月次更新
                    '    If Not objKtbnStrc.strcSelection.strKatabanCheckDiv = "4" Then
                    '        strMessageCd = "W9080"
                    '        'エラーメッセージ設定
                    '        Call AlertMessage(strMessageCd)
                    '    End If
                Case "CXU10"
                    If strKeyKata = "7" Or strKeyKata = "8" Or strKeyKata = "9" Then
                        strMessageCd = "W8920"
                        'エラーメッセージ設定
                        Call AlertMessage(strMessageCd)
                    End If
                Case "CXU30"
                    If strKeyKata = "9" Or strKeyKata = "A" Or strKeyKata = "B" Or strKeyKata = "C" Then
                        strMessageCd = "W8920"
                        'エラーメッセージ設定
                        Call AlertMessage(strMessageCd)
                    End If
            End Select
            Select Case Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 7)
                Case "D101-UN", "D401-UN", "B110-UN", "B310-UN", "B410-UN", _
                     "A100-UN", "A400-UN", "A101-UN", "A401-UN"
                    strMessageCd = "W8920"
                    'エラーメッセージ設定
                    Call AlertMessage(strMessageCd)
            End Select
            Select Case Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 11)
                '201502月次更新
                Case "CXU10-TA-UN", "CXU30-TA-UN", "CXU30-VE-UN", "CXU10-MA-UN", "CXU30-MA-UN", _
                     "CXU13-CA-UN", "CXU48-CA-UN"
                    strMessageCd = "W8920"
                    'エラーメッセージ設定
                    Call AlertMessage(strMessageCd)
            End Select
            Select Case Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 13)
                Case "C1000-J100-UN", "C4000-J400-UN"
                    strMessageCd = "W8920"
                    'エラーメッセージ設定
                    Call AlertMessage(strMessageCd)
            End Select
            '追加(価格算出のみシリーズ（メッセージ表示）)

            If Left(strSeriesKata, 3) = "LCR" Or Left(strSeriesKata, 3) = "SSD" Or _
                Left(strSeriesKata, 4) = "SSD2" Or Left(strSeriesKata, 3) = "SSG" Or _
                Left(strSeriesKata, 3) = "STG" Or Left(strSeriesKata, 3) = "STK" Or _
                Left(strSeriesKata, 3) = "STL" Or Left(strSeriesKata, 3) = "STS" Or _
                Left(strSeriesKata, 4) = "USSD" Or Left(strSeriesKata, 3) = "LFC" Or _
                Left(strSeriesKata, 4) = "SCA2" Or Left(strSeriesKata, 4) = "SMD2" Or _
                Left(strSeriesKata, 4) = "STR2" Or Left(strSeriesKata, 3) = "SCM" Or _
                Left(strSeriesKata, 3) = "LCG" Or Left(strSeriesKata, 3) = "SMG" Or _
                Left(strSeriesKata, 3) = "PCC" Or Left(strSeriesKata, 4) = "RCC2" Or _
                (strSeriesKata) = "CK" Or Left(strSeriesKata, 3) = "CKG" Or _
                 Left(strSeriesKata, 3) = "CKA" Or Left(strSeriesKata, 3) = "CKS" Or _
                 Left(strSeriesKata, 3) = "CKF" Or Left(strSeriesKata, 3) = "CKJ" Or _
                 Left(strSeriesKata, 4) = "CKH2" Or Left(strSeriesKata, 5) = "CKLB2" Or _
                 Left(strSeriesKata, 3) = "FH1" Or Left(strSeriesKata, 3) = "HAP" Or _
                 Left(strSeriesKata, 4) = "BSA2" Or Left(strSeriesKata, 3) = "LHA" Or _
                 Left(strSeriesKata, 3) = "HKP" Or Left(strSeriesKata, 3) = "HLA" Or _
                 Left(strSeriesKata, 3) = "HLB" Or Left(strSeriesKata, 3) = "HLD" Or _
                 Left(strSeriesKata, 3) = "HEP" Or Left(strSeriesKata, 3) = "HCP" Or _
                 Left(strSeriesKata, 3) = "HMF" Or Left(strSeriesKata, 3) = "HFP" Or _
                 Left(strSeriesKata, 3) = "HLC" Or Left(strSeriesKata, 3) = "HGP" Or _
                 Left(strSeriesKata, 3) = "FH5" Or Left(strSeriesKata, 3) = "HBL" Or _
                 Left(strSeriesKata, 3) = "HDL" Or Left(strSeriesKata, 3) = "HMD" Or _
                 Left(strSeriesKata, 3) = "HJD" Or Left(strSeriesKata, 3) = "HJL" Then
                Me.HdnSetYFlg.Value = "Y"
            ElseIf Left(strSeriesKata, 4) = "JSC3" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(5)
                    Case "40", "50", "63", "80", "100"
                        Me.HdnSetYFlg.Value = "Y"
                    Case Else
                        Me.HdnSetYFlg.Value = ""
                End Select
            ElseIf Left(strSeriesKata, 3) = "BHG" Or Left(strSeriesKata, 3) = "BHE" Or Left(strSeriesKata, 3) = "BHA" Then
                If strKeyKata = "" Or strKeyKata = "4" Then
                    Me.HdnSetYFlg.Value = "Y"
                Else
                    Me.HdnSetYFlg.Value = ""
                End If
            ElseIf Left(strSeriesKata, 4) = "CKL2" Then
                If strKeyKata = "" Or strKeyKata = "2" Or strKeyKata = "Q" Then
                    Me.HdnSetYFlg.Value = "Y"
                Else
                    Me.HdnSetYFlg.Value = ""
                End If
                '↓RM1401080 2014/01/23
            ElseIf Left(strSeriesKata, 4) = "CMK2" Or Left(strSeriesKata, 4) = "CKV2" Or _
                   Left(strSeriesKata, 3) = "ULK" Or Left(strSeriesKata, 4) = "JSK2" Or _
                   Left(strSeriesKata, 4) = "CMA2" Or Left(strSeriesKata, 4) = "JSM2" Or _
                   Left(strSeriesKata, 3) = "HCA" Or Left(strSeriesKata, 3) = "HCM" Or _
                   Left(strSeriesKata, 4) = "CAV2" Or Left(strSeriesKata, 5) = "COVN2" Or _
                   Left(strSeriesKata, 5) = "COVP2" Or Left(strSeriesKata, 3) = "SCG" Or _
                   Left(strSeriesKata, 3) = "JSG" Or Left(strSeriesKata, 3) = "HRL" Then
                Me.HdnSetYFlg.Value = "Y"
                '↓RM1402099 2014/02/25
            ElseIf Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "FCD" Or Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "FCH" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "FCS-" Or Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "MDC2" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "MDV" Or Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "MSD" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "MVC" Or Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "STM" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "UCA2" Or Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "LCM" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "LCT" Or Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "LCY" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "RPC" Or Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "GRC" Then
                Me.HdnSetYFlg.Value = "Y"
                '↓RM1403045 2014/03/25
            ElseIf Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "SRL3" Or Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "SRG3" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "SRM3" Or Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "MRL2" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "MRG2" Then
                Me.HdnSetYFlg.Value = "Y"
            Else
                Me.HdnSetYFlg.Value = ""
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' テキストのクリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub TxtClear()
        Me.txt_Rate.Text = String.Empty
        Me.TextUnitPrice.Text = String.Empty
        Me.TextCnt.Text = String.Empty
        Me.TextRateUnitPrice.Text = String.Empty
        Me.TextMoney.Text = String.Empty
        Me.TextTax.Text = String.Empty
        Me.TextAmount.Text = String.Empty
    End Sub

    ''' <summary>
    ''' 各ボックスにjavascriptを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInit()
        Dim strJS As String
        Dim strShiftD As String

        Try
            '掛率
            strJS = Me.txt_Rate.Attributes(CdCst.JavaScript.OnKeyUp)
            Me.txt_Rate.Attributes.Add(CdCst.JavaScript.OnKeyUp, strJS & "fncTanka_onKeyUp('" & Me.ClientID & "_','Rate');")
            'キーダウン
            Me.txt_Rate.Attributes.Add(CdCst.JavaScript.OnKeyDown, "fncTanka_onKeyDown('" & Me.ClientID & "_','Rate');")

            If Me.objUserInfo.EditDiv = "0" Then
                strJS = Me.txt_Rate.Attributes(CdCst.JavaScript.OnBlur)
                strJS = strJS & "this.value = fncRemoveComma(this.value);  if ( fncTextTrim(this.value).length != 0 && fncCheckNum( this.value, '5' ,'false', 'true', '" & Me.objUserInfo.EditDiv & "') != false) {"
                strJS = strJS & "f_UnitPriceCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "'); f_MoneyCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "');}"
                Me.txt_Rate.Attributes.Add(CdCst.JavaScript.OnBlur, strJS)
            Else
                strJS = Me.txt_Rate.Attributes(CdCst.JavaScript.OnBlur)
                strJS = strJS & "this.value = fncRemoveDot(this.value); if ( fncTextTrim(this.value).length != 0 && fncCheckNum( this.value, '5' ,'false', 'true', '" & Me.objUserInfo.EditDiv & "') != false) {"
                strJS = strJS & "this.value = this.value.replace(',','.'); f_UnitPriceCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "'); f_MoneyCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "');}"
                Me.txt_Rate.Attributes.Add(CdCst.JavaScript.OnBlur, strJS)
            End If

            '単価
            strJS = Me.TextUnitPrice.Attributes(CdCst.JavaScript.OnKeyUp)
            Me.TextUnitPrice.Attributes.Add(CdCst.JavaScript.OnKeyUp, strJS & "fncTanka_onKeyUp('" & Me.ClientID & "_','UnitPrice');")

            'キーダウン
            Me.TextUnitPrice.Attributes.Add(CdCst.JavaScript.OnKeyDown, "fncTanka_onKeyDown('" & Me.ClientID & "_','UnitPrice');")

            If Me.objUserInfo.EditDiv = "0" Then
                strJS = Me.TextUnitPrice.Attributes(CdCst.JavaScript.OnBlur)
                strJS = strJS & "this.value = fncRemoveComma(this.value); if ( fncTextTrim(this.value).length != 0 && fncCheckNum( this.value, '9' ,'false', 'true', '" & Me.objUserInfo.EditDiv & "') != false ) {"
                strJS = strJS & "f_RateCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "'); f_MoneyCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "');}"
                Me.TextUnitPrice.Attributes.Add(CdCst.JavaScript.OnBlur, strJS)
            Else
                strJS = Me.TextUnitPrice.Attributes(CdCst.JavaScript.OnBlur)
                strJS = strJS & "this.value = fncRemoveDot(this.value); if ( fncTextTrim(this.value).length != 0 && fncCheckNum( this.value, '9' ,'false', 'true', '" & Me.objUserInfo.EditDiv & "') != false ) {"
                strJS = strJS & "f_RateCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "'); f_MoneyCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "');}"
                Me.TextUnitPrice.Attributes.Add(CdCst.JavaScript.OnBlur, strJS)
            End If

            '数量
            strJS = Me.TextCnt.Attributes(CdCst.JavaScript.OnKeyUp)
            Me.TextCnt.Attributes.Add(CdCst.JavaScript.OnKeyUp, strJS & "fncTanka_onKeyUp('" & Me.ClientID & "_','Cnt');")

            'キーダウン
            Me.TextCnt.Attributes.Add(CdCst.JavaScript.OnKeyDown, "fncTanka_onKeyDown('" & Me.ClientID & "_','Cnt');")

            If Me.objUserInfo.EditDiv = "0" Then
                strJS = Me.TextCnt.Attributes(CdCst.JavaScript.OnBlur)
                strJS = strJS & "this.value = fncRemoveComma(this.value); if ( fncTextTrim(this.value).length != 0 && fncCheckNum( this.value, '6' ,'false', 'true', '" & Me.objUserInfo.EditDiv & "') != false ) {"
                strJS = strJS & " f_MoneyCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "');}"
                Me.TextCnt.Attributes.Add(CdCst.JavaScript.OnBlur, strJS)
            Else
                strJS = Me.TextCnt.Attributes(CdCst.JavaScript.OnBlur)
                strJS = strJS & "this.value = fncRemoveDot(this.value); if ( fncTextTrim(this.value).length != 0 && fncCheckNum( this.value, '6' ,'false', 'true', '" & Me.objUserInfo.EditDiv & "') != false ) {"
                strJS = strJS & " f_MoneyCal('" & Me.ClientID & "_','" & Me.objUserInfo.EditDiv & "');}"
                Me.TextCnt.Attributes.Add(CdCst.JavaScript.OnBlur, strJS)
            End If

            'Shift+Dイベント設定
            strShiftD = "if(event.keyCode == 68 && event.shiftKey==true){frmShiftD('ctl00_ContentDetail_WebUC_Tanka_');}"
            Me.TextCnt.Attributes.Add(CdCst.JavaScript.OnKeyDown, strShiftD)

            'ボタン自動サブミットを無効
            Me.Button2.UseSubmitBehavior = False
            Me.Button3.UseSubmitBehavior = False
            Me.Button5.UseSubmitBehavior = False
            Me.Button6.UseSubmitBehavior = False
            Me.Button9.UseSubmitBehavior = False
            Me.Button10.UseSubmitBehavior = False
            Me.Button11.UseSubmitBehavior = False
            Me.Button12.UseSubmitBehavior = False

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub
#End Region

    ''' <summary>
    ''' 受注EDI単価情報引渡
    ''' </summary>
    ''' <param name="strLoginUser">ログインユーザ</param>
    ''' <param name="strFullKataban">フル形番</param>
    ''' <param name="strPriceList">単価リスト</param>
    ''' <param name="dt_display">付加情報</param>
    ''' <remarks>受注EDIに単価情報を返す</remarks>
    Private Sub subJuchuEdiWSDbIO(ByVal strLoginUser As String, ByVal strFullKataban As String, _
                                ByVal strPriceList(,) As String, ByVal dt_display As DataTable)

        Dim cnfAppSet As New System.Configuration.AppSettingsReader
        Dim sbScript As New StringBuilder
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim objSendEdi As New KHSessionInfo.SendEdi
        Try
            With objSendEdi
                .FullKataban = strFullKataban
                .CheckKubun = String.Empty
                .PlaceCode = String.Empty
                .PriceTeika = strPriceList(1, 2)
                .PriceTouroku = strPriceList(2, 2)
                .PriceSS = strPriceList(3, 2)
                .PriceBS = strPriceList(4, 2)
                .PriceGS = strPriceList(5, 2)
                .PricePS = strPriceList(6, 2)
            End With

            Dim dr_display() As DataRow = Nothing
            For inti As Integer = 1 To 2
                dr_display = dt_display.Select("strLevel='" & inti.ToString & "'")
                If dr_display.Length > 0 Then
                    Select Case inti
                        Case 1 '形番チェック区分
                            objSendEdi.CheckKubun = dr_display(0)("strValue").ToString
                        Case 2 '出荷場所
                            objSendEdi.PlaceCode = dr_display(0)("strValue").ToString
                    End Select
                End If
            Next

            'セッション情報に確保
            httpCon.Session(CdCst.SessionInfo.Key.SendEdi) = objSendEdi
        Catch ex As Exception
            'エラー画面に遷移する
            AlertMessage(ex)
        Finally
            cnfAppSet = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 受注EDI単価情報引渡
    ''' </summary>
    ''' <param name="strLoginUser">ログインユーザ</param>
    ''' <param name="strFullKataban">フル形番</param>
    ''' <param name="strPriceList">単価リスト</param>
    ''' <param name="dt_display">付加情報</param>
    ''' <remarks>受注EDIに単価情報を返す</remarks>
    Private Sub subCommonDbWSDbIO(ByVal strFullKataban As String, _
                                ByVal dt_display As DataTable)

        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim objSendEdi As New KHSessionInfo.SendEdi
        Try
            '単価リスト取得
            'Dim strPriceList(,) As String = Session("strPriceList")
            'With objSendEdi
            '    .FullKataban = strFullKataban
            '    .CheckKubun = String.Empty
            '    .PlaceCode = String.Empty
            '    .PriceTeika = strPriceList(1, 2)
            '    .PriceTouroku = strPriceList(2, 2)
            '    .PriceSS = strPriceList(3, 2)
            '    .PriceBS = strPriceList(4, 2)
            '    .PriceGS = strPriceList(5, 2)
            '    .PricePS = strPriceList(6, 2)
            '    .PriceNet = strPriceList(8, 2)
            '    .Currency = strPriceList(8, 3)
            'End With

            'Session.Remove("strPriceList")

            'objSendEdi = KHSBOInterface.fncJutyuEdiInterfaceGet(objCon, objKtbnStrc, Me.objUserInfo.OfficeCd, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

            Dim dr_display() As DataRow = Nothing
            For inti As Integer = 1 To 2
                dr_display = dt_display.Select("strLevel='" & inti.ToString & "'")
                If dr_display.Length > 0 Then
                    Select Case inti
                        Case 1 '形番チェック区分
                            objSendEdi.CheckKubun = "Z" & dr_display(0)("strValue").ToString
                        Case 2 '出荷場所
                            objSendEdi.PlaceCode = dr_display(0)("strValue").ToString
                    End Select
                End If
            Next
            'セッション情報に確保
            httpCon.Session(CdCst.SessionInfo.Key.SendEdi) = objSendEdi
        Catch ex As Exception
            'エラー画面に遷移する
            AlertMessage(ex)
        Finally

        End Try
    End Sub

    ''' <summary>
    ''' データバインドイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GVPrice_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVPrice.RowDataBound
        If e.Row.RowIndex < 0 Then Exit Sub
        Try
            Dim str() As String = e.Row.Cells(1).Text.ToString.Split(" ")

            If (e.Row.RowIndex + 1) Mod 2 = 0 Then
                'e.Row.BackColor = Drawing.Color.FromArgb(173, 205, 207)
                e.Row.BackColor = Drawing.Color.FromArgb(204, 204, 255)
            Else
                e.Row.BackColor = Drawing.Color.White
            End If

            '価格があるものを選択できる
            If str.Length = 3 Then
                'Dim strPrice As String = str(0).ToString.Replace(",", "").Replace(".", "")
                'Dim strCurr As String = str(2).ToString
                Dim strName As String = Me.ClientID & "_"
                Dim intStartID As Integer = 0
                If e.Row.RowIndex = 0 Then
                    intStartID = CInt(Strings.Right(e.Row.ClientID, 2))
                Else
                    intStartID = CInt(Strings.Right(GVPrice.Rows(0).ClientID, 2))
                End If
                e.Row.Attributes.Add(CdCst.JavaScript.OnClick, "fncGridClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',3);")
                e.Row.Attributes.Add(CdCst.JavaScript.OnKeyUp, "fncGrid_OnKeyup(event, '" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',3);")
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 出荷場所の更新
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdatePlace()
        If Me.HidNewPlace.Value.ToString.Trim.Length > 0 Then
            objKtbnStrc.strcSelection.strPlaceCd = Me.HidNewPlace.Value.ToString.Trim
        End If

        '保管場所セット
        objKtbnStrc.strcSelection.strStorageLocation = IIf(Left(Me.cmbStrageEvaluation.SelectedValue.PadRight(8), 4).Equals("A***"), Nothing, Left(Me.cmbStrageEvaluation.SelectedValue.PadRight(8), 4).Trim)
        '評価タイプセット
        objKtbnStrc.strcSelection.strEvaluationType = IIf(Right(Me.cmbStrageEvaluation.SelectedValue.PadRight(8), 3).Trim.Equals(""), Nothing, Right(Me.cmbStrageEvaluation.SelectedValue.PadRight(8), 3).Trim)

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
        RaiseEvent IFFileOutput(objKtbnStrc, Me.ClientID, Me.HidNewPlace.Value.ToString.Trim)
    End Sub

    ''' <summary>
    ''' ファイル出力
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call UpdatePlace()
        RaiseEvent FileOutput(objKtbnStrc, Me.ClientID, HidPriceForFile.Value, HidPriceList.Value, 1)
    End Sub

    ''' <summary>
    ''' ＪＳＯＮファイル出力
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        '画面に表示されたフル形番により再設定
        If Not lblSeriesKat.Text.Trim.Equals(objKtbnStrc.strcSelection.strFullKataban) Then
            objKtbnStrc.strcSelection.strFullKataban = lblSeriesKat.Text.Trim
        End If
        RaiseEvent JSONFileOutput(objKtbnStrc, Me.ClientID)
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
        Dim clsKHSBOInerfaceResult As New KHSBOInterface
        Dim strFobPrice As String = 0
        Dim strSessionIDFob As String = String.Empty
        Dim strCurrencyCode As String = String.Empty
        'Dim result As WebKataban.CommonDbService.DbProcessResult
        Dim blCZFlag As Boolean = False

        Try
            '海外生産品の場合は受注EDI連携不可
            Select Case cmbPlace.SelectedValue
                Case "1001", "1002", "1003", "1004", "1005"
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
                strCurrencyCode = Session("strCurrencyCode")
                Session.Remove(strSessionIDFob)
            Else
                strFobPrice = 0
            End If

            'CZ特価フラグ
            If Label41.Text <> Nothing Then
                blCZFlag = True
            Else
                blCZFlag = False
            End If

            ''連携データクラスにデータセット
            'clsKHSBOInerfaceResult = KHSBOInterface.fncJutyuEdiInterfaceGet(objCon, objKtbnStrc, Me.objUserInfo.OfficeCd, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, strFobPrice, strCurrencyCode, objEdiInfo.KeyInfo, blCZFlag, Nothing)
            'clsKHSBOInerfaceResult.clKatahikiInfoDto.DeliveryPlant = Me.cmbPlace.SelectedValue
            'clsKHSBOInerfaceResult.clKatahikiInfoDto.StorageLocation = IIf(Left(Me.cmbStrageEvaluation.SelectedValue.PadRight(8), 4).Equals("A***"), Nothing, Left(Me.cmbStrageEvaluation.SelectedValue.PadRight(8), 4).Trim)
            'clsKHSBOInerfaceResult.clKatahikiInfoDto.EvaluationType = IIf(Right(Me.cmbStrageEvaluation.SelectedValue.PadRight(8), 3).Trim.Equals(""), Nothing, Right(Me.cmbStrageEvaluation.SelectedValue.PadRight(8), 3).Trim)

            ''↓2018/7/24_仕入品の販売数量単位が引き継がれない不具合対応
            'Dim MkatabanSalesUnit As String
            ''Dim dt As CommonDbService.M_Kataban
            ''dt = CommonDbService.GetMKataban("999", "9999999999", objKtbnStrc.strcSelection.strFullKataban, "JPY")
            ''If dt Is Nothing Then
            ''MkatabanSalesUnit = String.Empty
            ''Else
            ''MkatabanSalesUnit = dt.SalesUnit
            ''End If

            ''kh_qty_unitのSalesUnitがブランクで、マスタのSalesUnitが存在する場合にマスタのSalesUnitをセット
            'If objKtbnStrc.strcSelection.strSalesUnit = "" And MkatabanSalesUnit <> "" Then
            '    clsKHSBOInerfaceResult.clKatahikiInfoDto.SalesUnit = MkatabanSalesUnit
            'End If
            '↑2018/7/24_仕入品の販売数量単位が引き継がれない不具合対応

            '受注EDI情報送信Webサービスを実行する。
            'CommonDbService = New CommonDbService.CommonDbServiceClient
            'result = CommonDbService.AddKatahikiInfo(clsKHSBOInerfaceResult.clKatahikiInfoDto)

            RaiseEvent EDIReturn() '送信した後、引当システムを閉じる（ログオフ）

        Catch ex As Exception
            AlertMessage(ex) 'エラー画面に遷移する
        End Try
    End Sub

    ''' <summary>
    ''' 生産状況/担当者
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        'URL
        Dim strURL As String = String.Empty

        If objUserInfo.UserClass.Equals(CdCst.UserClass.DmSalesOffice) Then
            strURL = ConfigurationManager.AppSettings("PICUrl_DmSales")
        Else
            strURL = ConfigurationManager.AppSettings("PICUrl")
        End If
        ScriptManager.RegisterStartupScript(Page, Page.GetType, "OpenPIC", "window.open('" & strURL & "');", True)

    End Sub

    ''' <summary>
    ''' 在庫検索
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'URL
        Dim strURL As String = String.Empty

        If cmbStrageEvaluation.SelectedValue.Trim = "G000" Then
            strURL = ConfigurationManager.AppSettings("PICUrl_GLC")
        Else
            strURL = String.Format(ConfigurationManager.AppSettings("PICUrl_Zaiko"), "N30", objKtbnStrc.strcSelection.strFullKataban.ToString)
        End If
        ScriptManager.RegisterStartupScript(Page, Page.GetType, "OpenPIC", "window.open('" & strURL & "');", True)

    End Sub

    'RM1808098_ＣＥ取得有無確認
    ''' <summary>
    ''' ＣＥ取得有無確認
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        'URL
        Dim strURL As String = String.Empty

        strURL = ConfigurationManager.AppSettings("PICUrl_CE")

        ScriptManager.RegisterStartupScript(Page, Page.GetType, "OpenPIC", "window.open('" & strURL & "');", True)

    End Sub

    ''' <summary>
    ''' マニホールドテスト専用
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ManifoldTanka()
        Try
            If Not Me.Session("ManifoldKataban") Is Nothing Then
                If Me.Session("TestFlag") Is Nothing Then
                    Me.Session("TestFlag") = True
                    If Me.Session("TestMode").ToString <> "2" Then
                        Dim listKataban As ManifoldKataban = Me.Session("ManifoldKataban")
                        Dim strKataban As String = listKataban.KATABAN
                        Dim strSiyou As String = listKataban.SIYOUSYO
                        Dim strPlace As String = listKataban.KATAPLACE
                        Dim strCheck As String = listKataban.KATACHECK
                        Dim strGSPrice As String = listKataban.GSPRICE

                        Dim strPath As String = My.Settings.LogFolder & "MFShiyouTest.txt"
                        Dim strValue As String = strKataban & ControlChars.Tab & strSiyou & ControlChars.Tab

                        Dim strErr As String = String.Empty
                        If cmbPlace.SelectedValue <> strPlace Then
                            strErr = "出荷場所エラー：新" & cmbPlace.SelectedValue & " 旧;" & strPlace & ControlChars.Tab
                        End If
                        If lblCheck.Text <> strCheck Then
                            strErr &= "ﾁｪｯｸ区分エラー：新" & lblCheck.Text & " 旧;" & strCheck & ControlChars.Tab
                        End If
                        Dim str() As String = Me.GVPrice.Rows(4).Cells(1).Text.ToString.Trim.Replace(",", "").Split(" ")
                        If str.Length = 3 Then
                            If CLng(str(0)) <> strGSPrice Then
                                strErr &= "GS価格エラー：新" & CLng(str(0)) & " 旧;" & strGSPrice & ControlChars.Tab
                            End If
                        Else
                            strErr &= "GS価格計算エラー"
                        End If

                        If strErr.Length <= 0 Then
                            strValue &= "OK"
                        Else
                            strValue &= strErr
                        End If
                        WriteLog(strPath, strValue)

                        '仕様書を出力
                        RaiseEvent SiyouFileOutput(objKtbnStrc, strSiyou)

                        Session("EventEndFlg") = True
                        GC.Collect()
                    Else
                        '仕様テスト
                        Dim strErr As String = String.Empty
                        '出力パス
                        Dim strPath As String = My.Settings.LogFolder & "ShiyouTest_" & Now.ToString("yyyyMMdd") & ".txt"
                        '比較データ
                        Dim drShiyouTest As DS_PriceTest.kh_shiyou_testRow = Me.Session("ManifoldKataban")

                        '計算結果の比較
                        strErr = fncCompareShiyouResult(drShiyouTest)

                        '結果出力
                        WriteLog(strPath, strErr)

                        Session("EventEndFlg") = True
                        GC.Collect()
                    End If
                End If
            End If

        Catch ex As Exception
            Session("EventEndFlg") = True
            GC.Collect()
        End Try
    End Sub

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
    ''' 仕様処理結果の比較
    ''' </summary>
    ''' <param name="drShiyouTest"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCompareShiyouResult(ByVal drShiyouTest As DS_PriceTest.kh_shiyou_testRow) As String
        '計算結果
        Dim decGsPrice As Decimal = IIf(GetpageData("GS").Equals(String.Empty), 0, CDec(GetpageData("GS")))
        Dim decBsPrice As Decimal = IIf(GetpageData("BS").Equals(String.Empty), 0, CDec(GetpageData("BS")))
        Dim decPsPrice As Decimal = IIf(GetpageData("PS").Equals(String.Empty), 0, CDec(GetpageData("PS")))
        Dim decSsPrice As Decimal = IIf(GetpageData("SS").Equals(String.Empty), 0, CDec(GetpageData("SS")))
        Dim decLsPrice As Decimal = IIf(GetpageData("LS").Equals(String.Empty), 0, CDec(GetpageData("LS")))
        Dim decRgPrice As Decimal = IIf(GetpageData("RG").Equals(String.Empty), 0, CDec(GetpageData("RG")))
        Dim strCheckKBN As String = GetpageData("CHECKKBN")
        Dim strShipPlace As String = GetpageData("SHIPPLACE")
        Dim strResult As String = String.Empty

        '形番設定
        strResult &= drShiyouTest.KATABAN & ControlChars.Tab

        'ﾁｪｯｸ区分
        strResult &= GetCompareResultStr("CHECKKBN", strCheckKBN, drShiyouTest)

        '出荷場所
        strResult &= GetCompareResultStr("SHIPPLACE", strShipPlace, drShiyouTest)

        'GS価格の比較
        strResult &= GetCompareResultDec("GSPRICE", decGsPrice, drShiyouTest)

        'BS価格の比較
        strResult &= GetCompareResultDec("BSPRICE", decBsPrice, drShiyouTest)

        'PS価格の比較
        strResult &= GetCompareResultDec("PSPRICE", decPsPrice, drShiyouTest)

        'SS価格の比較
        strResult &= GetCompareResultDec("SSPRICE", decSsPrice, drShiyouTest)

        'LS価格の比較
        strResult &= GetCompareResultDec("LSPRICE", decLsPrice, drShiyouTest)

        'RG価格の比較
        strResult &= GetCompareResultDec("RGPRICE", decRgPrice, drShiyouTest)

        Return strResult
    End Function

    ''' <summary>
    ''' 画面データの取得
    ''' </summary>
    ''' <param name="strItemName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetpageData(ByVal strItemName As String) As String

        Dim strResult As String = String.Empty

        Select Case strItemName
            Case "GS"
                Dim str() As String = Me.GVPrice.Rows(4).Cells(1).Text.ToString.Trim.Replace(",", "").Split(" ")

                If str.Length = 3 Then
                    strResult = str(0)
                End If
            Case "PS"
                Dim str() As String = Me.GVPrice.Rows(5).Cells(1).Text.ToString.Trim.Replace(",", "").Split(" ")

                If str.Length = 3 Then
                    strResult = str(0)
                End If
            Case "BS"
                Dim str() As String = Me.GVPrice.Rows(3).Cells(1).Text.ToString.Trim.Replace(",", "").Split(" ")

                If str.Length = 3 Then
                    strResult = str(0)
                End If
            Case "SS"
                Dim str() As String = Me.GVPrice.Rows(2).Cells(1).Text.ToString.Trim.Replace(",", "").Split(" ")

                If str.Length = 3 Then
                    strResult = str(0)
                End If
            Case "LS"
                Dim str() As String = Me.GVPrice.Rows(0).Cells(1).Text.ToString.Trim.Replace(",", "").Split(" ")

                If str.Length = 3 Then
                    strResult = str(0)
                End If
            Case "RG"
                Dim str() As String = Me.GVPrice.Rows(1).Cells(1).Text.ToString.Trim.Replace(",", "").Split(" ")

                If str.Length = 3 Then
                    strResult = str(0)
                End If

            Case "CHECKKBN"

                strResult = lblCheck.Text

            Case "SHIPPLACE"

                strResult = cmbPlace.SelectedValue

        End Select


        Return strResult

    End Function

    ''' <summary>
    ''' 比較結果の作成
    ''' </summary>
    ''' <param name="strItemName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCompareResultDec(ByVal strItemName As String, _
                                      ByVal decValue As Decimal, _
                                      ByVal drShiyouTest As DS_PriceTest.kh_shiyou_testRow) As String

        Dim strResult As String = String.Empty

        'GS価格の比較
        If IsDBNull(drShiyouTest.Item(strItemName)) Then
            If decValue = 0 Then
                strResult &= "○"
            Else
                strResult &= "WEB版：" & decValue & Space(4) & "NET版：" & drShiyouTest.Item(strItemName).ToString
            End If
        Else
            If decValue.Equals(drShiyouTest.Item(strItemName)) Then
                strResult &= "○"
            Else
                strResult &= "WEB版：" & decValue & Space(4) & "NET版：" & drShiyouTest.Item(strItemName).ToString
            End If
        End If

        Return strResult & ControlChars.Tab

    End Function

    ''' <summary>
    ''' 比較結果の作成
    ''' </summary>
    ''' <param name="strItemName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCompareResultStr(ByVal strItemName As String, _
                                      ByVal strValue As String, _
                                      ByVal drShiyouTest As DS_PriceTest.kh_shiyou_testRow) As String

        Dim strResult As String = String.Empty

        '処理結果の比較
        If IsDBNull(drShiyouTest.Item(strItemName)) Then
            If strValue.Equals(String.Empty) Then
                strResult &= "○"
            Else
                strResult &= "WEB版：" & strValue & Space(4) & "NET版：" & drShiyouTest.Item(strItemName).ToString
            End If
        Else
            If strValue.Equals(drShiyouTest.Item(strItemName)) Then
                strResult &= "○"
            Else
                strResult &= "WEB版：" & strValue & Space(4) & "NET版：" & drShiyouTest.Item(strItemName).ToString
            End If
        End If

        Return strResult & ControlChars.Tab

    End Function

#Region "出荷場所関連"
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

            'RM1606006 販売数量単位は無条件で表示するように修正    ↓↓↓↓↓↓
            Dim drQtyUnitNm As DataRow = subAddInfoDispGet.NewRow

            If objKataban.fncQtyUnitInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, _
                                                         strSelectLang, strQtyUnitNm, objKtbnStrc) Then
                drQtyUnitNm("strLevel") = 32
                drQtyUnitNm("strDisplay") = True
                drQtyUnitNm("strValue") = strQtyUnitNm
            End If
            'RM1606006 販売数量単位は無条件で表示するように修正    ↑↑↑↑↑↑

            subAddInfoDispGet.Rows.Add(drQtyUnitNm)

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
                                dr("strValue") = "○"
                            End If
                            'Case 32 '販売数量単位
                            '    If objKataban.fncQtyUnitInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, _
                            '                                 strSelectLang, strQtyUnitNm) Then
                            '        dr("strValue") = strQtyUnitNm
                            '    End If
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
                            ccd_flg = False
                            'ccd_flg = KHKataban.subJapanChinaAmount(objKtbnStrc.strcSelection.strFullKataban)
                            'If ccd_flg = True Then strQuantity = ""
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
    ''' フル形番により出荷場所コードを取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetPlaceIDByFullKataban(ByVal lstCountriesByCountryCd As ArrayList, _
                                                ByRef retCtryItm As Boolean, _
                                                ByVal strUserCountryCd As String) As ArrayList
        Dim result As New ArrayList

        Dim lstCountriesByFullKataban As New ArrayList   'フル形番により表示可能な「国コード」

        '初期値を設定
        If objKtbnStrc.strcSelection.strMadeCountry = strUserCountryCd Then
            result.Add(strUserCountryCd)
        End If

        '有効期限によりチェックする
        Dim dt_FullKata As New DataTable

        For Each strCountryCd In lstCountriesByCountryCd
            dt_FullKata = KHCountry.fncCountryItmMstChkP(objConBase, objKtbnStrc.strcSelection.strFullKataban, strCountryCd)

            If (Not dt_FullKata Is Nothing) AndAlso (dt_FullKata.Rows.Count > 0) Then

                '出荷場所変換メッセージ表示フラグの設定
                retCtryItm = True

                '結果を表示リストに追加

                If dt_FullKata.Select("country_cd='" & strCountryCd & "'").Length > 0 Then
                    'フル形番が存在する場合は結果リストに追加
                    If Not result.Contains(strCountryCd) Then
                        result.Add(strCountryCd)
                    End If
                End If
            End If
        Next
        Return result
    End Function

    ''' <summary>
    ''' オプションの生産国レベルにより出荷場所コードを取得
    ''' </summary>
    ''' <param name="lstPlaceIDs">出荷場所コード候補</param>
    ''' <param name="lstCountriesByCountryCd">ユーザ表示可能国コードリスト</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetPlaceIDByLevel(ByVal lstPlaceIDs As ArrayList, _
                                          ByVal lstCountriesByCountryCd As ArrayList, _
                                          ByVal strCountryCd As String) As ArrayList
        Dim intPlaceLevel As Integer
        Dim result As New ArrayList

        result = lstPlaceIDs

        'オプションの生産国レベルの取得
        intPlaceLevel = fncGetPlaceLevel(strCountryCd)

        'フル形番と生産国レベルは最優先（**生産品）
        '生産国レベルから国ｺｰﾄﾞを取得する
        If intPlaceLevel <> 1024 And intPlaceLevel > 0 Then
            Dim myPlaceList As New ArrayList
            '生産国レベルに対応する国コードを取得
            myPlaceList = KHCountry.fncGetPlacelvlName(objConBase, intPlaceLevel)

            '出荷場所候補にない場合追加
            For inti As Integer = 0 To myPlaceList.Count - 1
                Dim strPlaceCd As String = myPlaceList(inti).ToString.Split(",")(1)
                '表示可能な場合追加
                If (Not result.Contains(strPlaceCd)) AndAlso (lstCountriesByCountryCd.Contains(strPlaceCd)) Then
                    result.Add(strPlaceCd)
                End If
            Next
        End If

        Return result
    End Function

    ''' <summary>
    ''' 第一ハイフン前より出荷場所コードを取得
    ''' </summary>
    ''' <param name="lstPlaceIDs">出荷場所コード候補</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetPlaceIDByHyphen(ByRef lstPlaceIDs As ArrayList, ByVal lstCountriesByCountryCd As ArrayList) As ArrayList
        '第一ハイフン前より検索する、2番目優先（**生産可能品、lstCountry_Key中身のものを画面で生産可能品を表示する）
        Dim lstCountry_lvl As New ArrayList
        Dim lstCountry_Key As New ArrayList
        Dim result As New ArrayList

        result = lstPlaceIDs

        If lstCountriesByCountryCd.Count > 1 Or lstCountry_lvl.Count > 0 Then    '表示順番マスタに存在する時 2013/07/31 条件追加
            lstCountry_Key = KHCountry.fncCountryKeyGet(objConBase, objKtbnStrc.strcSelection.strFullKataban)
            'フル形番であれば、対象外になる
            If result.Count > 0 Then
                For inti As Integer = lstCountry_Key.Count - 1 To 0 Step -1
                    If result.Contains(lstCountry_Key(inti)) Then
                        lstCountry_Key.RemoveAt(inti)
                    End If
                Next
            End If

            '生産可能の国コードを追加する
            For inti As Integer = 0 To lstCountry_Key.Count - 1
                If Not result.Contains(lstCountry_Key(inti)) Then
                    result.Add(lstCountry_Key(inti))
                End If
            Next
        End If

        Return lstCountry_Key
    End Function

    ''' <summary>
    ''' 価格マスタの出荷場所を追加
    ''' </summary>
    ''' <param name="lstPlaceIDs"></param>
    ''' <remarks></remarks>
    Private Sub subSetPlaceIDByOptions(ByRef lstPlaceIDs As ArrayList, _
                                       ByVal lstCountriesByCountryCd As ArrayList)
        '価格マスタの生産場所
        Dim strAddPlace As String = String.Empty
        'シリーズの国コード
        Dim strSeriesCountry As String = String.Empty

        '機種の時に機種マスタの生産国が「JPN」以外の場合は、価格マスタの出荷場所の代わりに機種マスタの生産国を追加
        If fncUseSeriesCountryOrNot(strSeriesCountry) Then
            strAddPlace = strSeriesCountry
        Else
            If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                strAddPlace = dt_Addinfo.Select("strLevel='2'")(0)("strValue")
            End If
        End If

        '生産場所を国コードに変換
        subChangeToJapaneseCountryCd(strAddPlace)
        '生産場所の追加
        If (Not lstPlaceIDs.Contains(strAddPlace)) AndAlso lstCountriesByCountryCd.Contains(strAddPlace) Then
            '出荷場所表示権限が無い場合は処理しない
            If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                lstPlaceIDs.Add(dt_Addinfo.Select("strLevel='2'")(0)("strValue"))
            End If
        End If

    End Sub

    ''' <summary>
    ''' 選択された国コードにより出荷場所コードを追加
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetSelectPlaceID(ByVal lstPlaceIDs As ArrayList) As ArrayList
        Dim strSelectCountry As String = String.Empty              '選択された表示名に対応する国コード
        Dim result As New ArrayList

        result = lstPlaceIDs

        If HidNewPlace.Value = String.Empty AndAlso lstPlaceIDs.Count > 0 Then
            strSelectCountry = lstPlaceIDs(0)
        Else
            strSelectCountry = HidNewPlace.Value
        End If
        'タイ工場生産品対応
        Dim strCountries As New ArrayList From {"PRC", "THA", "THF"}

        If Not strCountries.Contains(strSelectCountry) Then
            Dim dt_country As DataTable = KHCountry.fncGetCountryName(objConBase) '全ての国コードと国名
            Dim drCountries() As DataRow = dt_country.Select("country_cd='" & strSelectCountry & "'")

            If drCountries.Length <= 0 Then
                strSelectCountry = "JPN"
                If result.Count <= 0 Then
                    strSelectCountry = objKtbnStrc.strcSelection.strMadeCountry   '原産国
                End If
            End If
        End If

        If lstPlaceIDs.Count > 0 Or strSelectCountry <> "JPN" Then
            Call SetPlaceMark(strSelectCountry, lstPlaceIDs.Count)
        End If

        '生産可能の国コードを追加する
        If Not result.Contains(strSelectCountry) Then
            result.Insert(0, strSelectCountry)
        End If

        Return result
    End Function

    ''' <summary>
    ''' KatabanStrcEleにより生産国コードを検証
    ''' </summary>
    ''' <param name="lstPlaceIDs">生産可能国</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncFilterCountryByByPlaceLevel(ByVal lstPlaceIDs As ArrayList) As ArrayList

        Dim strSeries As String = objKtbnStrc.strcSelection.strSeriesKataban.Trim                                          'シリーズ形番
        Dim strKeyKataban As String = objKtbnStrc.strcSelection.strKeyKataban.Trim                                         'キー形番
        Dim lstOpSymbol As List(Of String) = objKtbnStrc.strcSelection.strOpSymbol.ToList                                  '選択された要素
        Dim intSolenoid As Integer = 1                                                                                     '「切替位置区分」要素番号
        Dim intPortSize As Integer = 3                                                                                     '「口径」要素番号

        'M3G*,M4G*,MN3G*,MN4G*の場合
        'If strSeries Like "M[3-4]G*" OrElse strSeries Like "M?[3-4]G*" Then
        If strSeries Like "M[3-4]G*" Then

            If strKeyKataban.Equals("R") OrElse _
                strKeyKataban.Equals("S") OrElse _
                strKeyKataban.Equals("U") OrElse _
                strKeyKataban.Equals("V") Then
                '接続口径の位置調整
                intPortSize = 4
            End If

            If lstOpSymbol.Item(intSolenoid).Equals("8") OrElse _
                lstOpSymbol.Item(intPortSize).Equals("CX") Then

                '電磁弁の形番により生産国コードをフィルタする
                Dim lstAttributeKataban As List(Of String) = objKtbnStrc.strcSelection.strAttributeSymbol.ToList           '属性コード
                Dim lstOptionKataban As List(Of String) = objKtbnStrc.strcSelection.strOptionKataban.ToList                '選択した形番
                Dim lstCXAKataban As List(Of String) = objKtbnStrc.strcSelection.strCXAKataban.ToList                      '選択したCXA形番
                Dim lstCXBKataban As List(Of String) = objKtbnStrc.strcSelection.strCXBKataban.ToList                      '選択したCXB形番

                For intIndex As Integer = 1 To lstAttributeKataban.Count - 1

                    Dim strAttribute As String = lstAttributeKataban.Item(intIndex)

                    'オプション形番が選択されないまたMPの場合は処理しない
                    If (Not String.IsNullOrEmpty(lstOptionKataban(intIndex).Trim)) AndAlso _
                        (Not lstOptionKataban(intIndex).Trim.Contains("-MP")) Then

                        '電磁弁の場合フィルタを掛ける
                        If strAttribute.Trim.Equals("D1") OrElse _
                            strAttribute.Trim.Equals("D2") Then

                            If Not String.IsNullOrEmpty(lstOptionKataban(intIndex).Trim) Then

                                Dim strSeriesChanged As String = strSeries
                                Dim strReplaced As String = strSeriesChanged.TrimStart("M")

                                '機種の変更（M4Gシリーズの場合は3GEと4GE両方とも選択できるから）
                                If Not lstOptionKataban(intIndex).Trim.StartsWith(strReplaced) Then
                                    strSeriesChanged = strSeriesChanged.Replace("M4", "M3")
                                End If

                                'オプション形番の切換位置区分によりフィルタ
                                If lstOpSymbol.Item(intSolenoid).Equals("8") Then

                                    subFilterCountryByOptionKatabanSolenoid(lstPlaceIDs, strSeriesChanged, lstOptionKataban(intIndex).Trim, intSolenoid)

                                End If

                                'オプション形番の接続口径によりフィルタ
                                If lstOpSymbol.Item(intPortSize).Equals("CX") Then

                                    subFilterCountryByOptionKatabanPortSize(lstPlaceIDs, strSeriesChanged, lstOptionKataban(intIndex).Trim, intPortSize)

                                    'CXA形番によりフィルタ
                                    If Not String.IsNullOrEmpty(lstCXAKataban(intIndex).Trim) Then
                                        subFilterCountryByCXKataban(lstPlaceIDs, strSeriesChanged, lstCXAKataban(intIndex).Trim, intPortSize)
                                    End If

                                    'CXB形番によりフィルタ
                                    If Not String.IsNullOrEmpty(lstCXBKataban(intIndex).Trim) Then
                                        subFilterCountryByCXKataban(lstPlaceIDs, strSeriesChanged, lstCXBKataban(intIndex).Trim, intPortSize)
                                    End If

                                End If

                            End If

                        End If

                    End If

                Next

            End If

        End If

        Return lstPlaceIDs

    End Function

    ''' <summary>
    ''' オプション形番の切換位置区分によりフィルタ
    ''' </summary>
    ''' <param name="lstPlaceIDs">生産国コードの候補</param>
    ''' <param name="strSeries">シリーズ形番</param>
    ''' <param name="strOptionKataban">オプション形番</param>
    ''' <param name="intSolenoid">切換位置区分の位置</param>
    ''' <remarks></remarks>
    Private Sub subFilterCountryByOptionKatabanSolenoid(ByRef lstPlaceIDs As ArrayList, _
                                                        ByVal strSeries As String, _
                                                        ByVal strOptionKataban As String, _
                                                        ByVal intSolenoid As Integer)

        Dim strSelectedSolenoid As String = String.Empty                                    '切換位置区分
        Dim strKeyKataban As String = objKtbnStrc.strcSelection.strKeyKataban.Trim
        Dim intSolenoidPlaceLevel As Integer = 0

        '切換位置区分を取得
        strSelectedSolenoid = fncGetSolenoid(strOptionKataban, strSeries)

        '切換位置区分の生産レベルを取得
        If YousoBLL.subGetPlacelvl(objCon, _
                                   strSeries, _
                                   strKeyKataban, _
                                   strSelectedSolenoid, _
                                   intSolenoid, _
                                   intSolenoidPlaceLevel) Then

            '切換位置区分のフィルタ
            lstPlaceIDs = fncDeleteCountryByPlaceLevel(lstPlaceIDs, intSolenoidPlaceLevel)

        End If

    End Sub

    ''' <summary>
    ''' オプション形番の接続口径によりフィルタ
    ''' </summary>
    ''' <param name="lstPlaceIDs">生産国コードの候補</param>
    ''' <param name="strSeries">シリーズ形番</param>
    ''' <param name="strOptionKataban">オプション形番</param>
    ''' <param name="intPortSize">接続口径の位置</param>
    ''' <remarks></remarks>
    Private Sub subFilterCountryByOptionKatabanPortSize(ByRef lstPlaceIDs As ArrayList, _
                                                        ByVal strSeries As String, _
                                                        ByVal strOptionKataban As String, _
                                                        ByVal intPortSize As Integer)

        Dim strSelectedPortSize As String = strOptionKataban.Split("-")(1)                  '接続口径 
        Dim strKeyKataban As String = objKtbnStrc.strcSelection.strKeyKataban.Trim
        Dim intPortSizePlaceLevel As Integer = 0

        '接続口径の生産レベルを取得
        If YousoBLL.subGetPlacelvl(objCon, _
                                   strSeries, _
                                   strKeyKataban, _
                                   strSelectedPortSize, _
                                   intPortSize, _
                                   intPortSizePlaceLevel) Then

            '接続口径接続口径のフィルタ
            lstPlaceIDs = fncDeleteCountryByPlaceLevel(lstPlaceIDs, intPortSizePlaceLevel)

        End If

    End Sub

    ''' <summary>
    ''' CX形番によりフィルタ
    ''' </summary>
    ''' <param name="lstPlaceIDs"></param>
    ''' <param name="strCXKataban"></param>
    ''' <param name="intPortSize"></param>
    ''' <remarks></remarks>
    Private Sub subFilterCountryByCXKataban(ByRef lstPlaceIDs As ArrayList, _
                                            ByVal strSeries As String, _
                                            ByVal strCXKataban As String, _
                                            ByVal intPortSize As Integer)

        Dim intPortSizePlaceLevel As Integer = 0

        '接続口径の生産レベルを取得
        Call YousoBLL.subGetPlacelvl(objCon, _
                                     strSeries, _
                                     objKtbnStrc.strcSelection.strKeyKataban, _
                                     strCXKataban, _
                                     intPortSize, _
                                     intPortSizePlaceLevel)

        '接続口径接続口径のフィルタ
        lstPlaceIDs = fncDeleteCountryByPlaceLevel(lstPlaceIDs, intPortSizePlaceLevel)

    End Sub

    ''' <summary>
    ''' 生産レベルにより生産不可の国を削除
    ''' </summary>
    ''' <param name="lstPlaceIDs"></param>
    ''' <param name="intPlaceLevel"></param>
    ''' <remarks></remarks>
    Private Function fncDeleteCountryByPlaceLevel(ByVal lstPlaceIDs As ArrayList, _
                                                  ByVal intPlaceLevel As Integer) As ArrayList

        Dim strResult As New ArrayList
        Dim strEnableCountries As New List(Of String)

        If intPlaceLevel <> 0 Then
            '生産可能な国レベル
            Dim strEnableLevels As List(Of String) = KHCountry.fncGetStroke_Logic(intPlaceLevel).Split(",").ToList
            '生産レベルと国コード
            Dim dt_AllCountryLevel As DataTable = YousoBLL.fncGetAllCountryLevel(objConBase)

            '生産レベルに対応する国コードを取得
            For Each strLevel As String In strEnableLevels

                Dim drCountryRows() As DataRow = dt_AllCountryLevel.Select("place_lvl= '" & strLevel & "'")

                If drCountryRows.Count > 0 Then

                    Dim strEnableCountry As String = drCountryRows(0).Item("place_div").ToString

                    strEnableCountries.Add(strEnableCountry)

                End If

            Next

            '生産不可の国コードを削除
            For Each placeId As String In lstPlaceIDs

                If strEnableCountries.Contains(placeId) Then

                    If Not strResult.Contains(placeId) Then
                        strResult.Add(placeId)
                    End If

                End If

            Next

        End If

        Return strResult

    End Function

    ''' <summary>
    ''' 選択した形番から切換位置区分を取得
    ''' </summary>
    ''' <param name="strOptionKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetSolenoid(ByVal strOptionKataban As String, _
                                    ByVal strSeries As String) As String

        Dim strResult As String = String.Empty
        Dim strReplaced As String = strSeries.TrimStart("M")
        Dim strKataban As String = strOptionKataban.Split("-")(0).TrimEnd("R")

        strResult = strKataban.Substring(0, strKataban.Length - 1).Replace(strReplaced, String.Empty)

        Return strResult

    End Function

    ''' <summary>
    ''' オプションの生産国レベルを取得
    ''' </summary>
    ''' <param name="strCountryCd">ユーザー国コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetPlaceLevel(ByVal strCountryCd As String) As Integer
        '生産国レベルを取得する(一番小さい数字を取得する)
        Dim intPlacelvl As Integer = 1024

        For inti As Integer = 1 To objKtbnStrc.strcSelection.strOpCountryDiv.Length - 1
            If objKtbnStrc.strcSelection.strOpCountryDiv(inti) <= 0 Then Continue For
            '単純な比較ではない、リストを展開して比較する   Add by Zxjike 2013/11/15
            If intPlacelvl <> CLng(objKtbnStrc.strcSelection.strOpCountryDiv(inti)) Then
                Dim intReal As Integer = 0
                If intPlacelvl <> "1024" Then
                    Dim strMaxOne() As String = KHCountry.fncGetStroke_Logic(objKtbnStrc.strcSelection.strOpCountryDiv(inti)).Split(",")
                    Dim strMinOne() As String = KHCountry.fncGetStroke_Logic(intPlacelvl).Split(",")
                    If strMaxOne.Length > 0 And strMinOne.Length > 0 Then
                        For intl As Integer = 0 To strMaxOne.Length - 1
                            For intk As Integer = 0 To strMinOne.Length - 1
                                If strMaxOne(intl) = strMinOne(intk) Then
                                    intReal += CLng(strMaxOne(intl))
                                End If
                            Next
                        Next
                    End If
                    intPlacelvl = intReal
                Else
                    intPlacelvl = objKtbnStrc.strcSelection.strOpCountryDiv(inti)
                End If
            End If
        Next

        '禁則条件(要素) 中国生産品とタイ生産品 
        If intPlacelvl >= 2 Then
            Dim strMinOne() As String = KHCountry.fncGetStroke_Logic(intPlacelvl).Split(",")
            For inti As Integer = 0 To strMinOne.Length - 1
                If strMinOne(inti) = "2" Then   '中国生産品判断する
                    'If Not KHCountry.fncGetData_Logic_China(objKtbnStrc, Me.objUserInfo.CountryCd) Then intPlacelvl -= 2
                    If Not KHCountry.fncGetData_Logic_China(objKtbnStrc, strCountryCd) Then intPlacelvl -= 2
                    'Exit For
                ElseIf strMinOne(inti) = "4" Then    'タイ生産品と判断する
                    'If Not KHCountry.fncGetData_Logic_Thailand(objKtbnStrc, Me.objUserInfo.CountryCd) Then intPlacelvl -= 4
                    If Not KHCountry.fncGetData_Logic_Thailand(objKtbnStrc, strCountryCd) Then intPlacelvl -= 4
                    'Exit For
                    'RM1801038_インドネシア禁則対応追加
                ElseIf strMinOne(inti) = "8" Then    'インドネシア生産品と判断する
                    If Not KHCountry.fncGetData_Logic_Indonesia(objKtbnStrc, strCountryCd) Then intPlacelvl -= 8
                End If
            Next
        End If

        Return intPlacelvl
    End Function

    ''' <summary>
    ''' 日本の出荷場所変換
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subChangeToJapanesePlaceCd(ByRef strPlaceID As String)
        Select Case objKtbnStrc.strcSelection.strPlaceCd.ToString.Trim
            Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
                If strPlaceID = "JPN" Then
                    strPlaceID = objKtbnStrc.strcSelection.strPlaceCd.ToString.Trim
                End If
        End Select
    End Sub

    ''' <summary>
    ''' 拠点コードを国コードへ変換
    ''' </summary>
    ''' <param name="strPlaceID"></param>
    ''' <remarks></remarks>
    Private Sub subChangeToJapaneseCountryCd(ByRef strPlaceID As String)
        Select Case strPlaceID
            Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
                strPlaceID = "JPN"
        End Select
    End Sub

    ''' <summary>
    ''' 出荷場所変換通知（保管場所＆評価タイプを追加）
    ''' </summary>
    ''' <param name="dtPlace"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncChangePlace(ByVal dtPlace As DataTable, ByVal lstPlaceIDs As ArrayList, ByVal retCtryItm As Boolean, _
                                    ByRef dtStrageEvaluation As DataTable) As DataTable
        Dim dtResult As New DataTable
        Dim dtStockPlace As New DataTable
        Dim bolReturn As Boolean = False                           '変換必要フラグ
        Dim strChangePlaceCd As String = String.Empty              '変換出荷場所

        Dim FRLMsg As String = ""

        Dim strGlcPlaceCdJa As String = String.Empty
        Dim strGlcPlaceCdOs As String = String.Empty

        'RM1808***_GLC在庫品メッセージ出力制御追加
        Dim strMessage_Type As String = String.Empty    '

        Dim strStorageLocation As String = "A***"
        Dim strEvaluationType As String = String.Empty            '評価タイプ
        Dim strSearchDiv As String = String.Empty

        Dim strGLCMsg As String = String.Empty
        Dim strMsg As String = String.Empty
        Dim drStrageEvaluation As DataRow = dtStrageEvaluation.NewRow

        dtResult = dtPlace

        'FRL判定用のメッセージを取得  2017/02/22 追加
        FRLMsg = fncFRLMessage()

        Dim bolExistCountries As Boolean = lstPlaceIDs.Count > 1   '海外生産品が存在するかの判断

        '出荷場所が営業プラントの場合は保管場所、評価タイプ変更しない 仕入品の場合も表示しない
        If objKtbnStrc.strcSelection.strPlaceCd = "1001" Or objKtbnStrc.strcSelection.strDivision = "3" Then

        Else
            '変換必要があるかどうかの判断(出荷場所変換マスタ　中国・タイ仕入品)
            bolReturn = KHCountry.fncPlaceChangeInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, _
                                                        strChangePlaceCd, strEvaluationType, strSearchDiv)

            'GLC在庫品があるかどうか　2017/3/3　
            dtStockPlace = KHCountry.fncStockPlaceInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, _
                                                     strChangePlaceCd)


            'GLC在庫品がある場合は、保管場所＆評価タイプを追加
            If dtStockPlace.Rows.Count > 0 Then
                For i As Integer = 0 To dtStockPlace.Rows.Count - 1

                    Select Case dtStockPlace.Rows(i).Item("evaluation_type").ToString
                        Case "JP1", "   ", ""
                            '日本製
                            strGlcPlaceCdJa = dtStockPlace.Rows(i).Item("storage_Location").ToString & Space(1) & _
                                                dtStockPlace.Rows(i).Item("evaluation_type").ToString
                            strMessage_Type = dtStockPlace.Rows(i).Item("message_type").ToString
                        Case Else
                            '海外製
                            strGlcPlaceCdOs = dtStockPlace.Rows(i).Item("storage_Location").ToString & _
                                                 Space(1) & dtStockPlace.Rows(i).Item("evaluation_type").ToString.PadRight(3)
                            strMessage_Type = dtStockPlace.Rows(i).Item("message_type").ToString
                    End Select
                Next
            End If

            '出荷場所変換マスタ　中国・タイ仕入品がある場合
            If bolReturn Then
                'ドロップダウンリストに追加
                If strSearchDiv <> "4" Then
                    'RM1808***_GLC在庫品メッセージ出力制御追加
                    If strMessage_Type <> "0" Then
                        '日本生産品
                        subSetStrageEvaluation(strStorageLocation & Space(1) & "JP1", dtStrageEvaluation)
                    End If
                End If

                '海外生産品
                subSetStrageEvaluation(strStorageLocation & Space(1) & strEvaluationType, dtStrageEvaluation)

                '出荷場所変換メッセージ
                If Session("TestMode") Is Nothing And strSearchDiv <> "4" Then
                    Select Case strEvaluationType
                        Case "CN1"
                            strMsg = ClsCommon.fncGetMsg(selLang.SelectedValue, "I7040")
                        Case "TH1"
                            strMsg = ClsCommon.fncGetMsg(selLang.SelectedValue, "I7050")
                        Case Else
                            strMsg = ClsCommon.fncGetMsg(selLang.SelectedValue, "I7040")
                    End Select
                End If

                'GLC在庫品がある場合
                If strGlcPlaceCdOs <> "" Then
                    'RM1808***_GLC在庫品メッセージ出力制御追加
                    If strMessage_Type <> "0" Then
                        'GLCメッセージセット
                        strGLCMsg = ClsCommon.fncGetMsg(selLang.SelectedValue, "Q0010")
                    End If
                    'ドロップダウンリストに追加
                    subSetStrageEvaluation(strGlcPlaceCdOs, dtStrageEvaluation)
                End If
            Else
                'ドロップダウンリストに追加　
                '日本生産品
                If objKtbnStrc.strcSelection.strDivision <> "3" Then
                    'RM1808***_GLC在庫品メッセージ出力制御追加
                    If strMessage_Type <> "0" Then
                        '仕入品以外
                        subSetStrageEvaluation(strStorageLocation & Space(4), dtStrageEvaluation)
                    End If
                End If

                'GLC在庫品がある場合
                If strGlcPlaceCdJa <> "" Then
                    'RM1808***_GLC在庫品メッセージ出力制御追加
                    If strMessage_Type <> "0" Then
                        'GLCメッセージセット
                        strGLCMsg = ClsCommon.fncGetMsg(selLang.SelectedValue, "Q0010")
                    End If
                    'ドロップダウンリストに追加　
                    subSetStrageEvaluation(strGlcPlaceCdJa, dtStrageEvaluation)
                End If

                'GLC在庫品がある場合
                If strGlcPlaceCdOs <> "" Then
                    'RM1808***_GLC在庫品メッセージ出力制御追加
                    If strMessage_Type <> "0" Then
                        'GLCメッセージセット
                        strGLCMsg = ClsCommon.fncGetMsg(selLang.SelectedValue, "Q0010")
                    End If
                    'ドロップダウンリストに追加
                    subSetStrageEvaluation(strGlcPlaceCdOs, dtStrageEvaluation)
                End If

            End If

            '海外生産品がある場合
            If bolExistCountries Then

            Else
                FRLMsg = fncFRLMessage()
                '出荷場所変換メッセージ
                If strMsg <> "" Then
                    Select Case strEvaluationType
                        Case "CN1"
                            subShowPlaceMessage(True, strMsg, "A*** CN1", "A*** JP1", dtResult, lstPlaceIDs, FRLMsg, strGlcPlaceCdJa, strGlcPlaceCdOs, strGLCMsg, dtStrageEvaluation)
                        Case "TH1"
                            subShowPlaceMessage(True, strMsg, "A*** TH1", "A*** JP1", dtResult, lstPlaceIDs, FRLMsg, strGlcPlaceCdJa, strGlcPlaceCdOs, strGLCMsg, dtStrageEvaluation)
                        Case Else
                            subShowPlaceMessage(True, strMsg, "A*** CN1", "A*** JP1", dtResult, lstPlaceIDs, FRLMsg, strGlcPlaceCdJa, strGlcPlaceCdOs, strGLCMsg, dtStrageEvaluation)
                    End Select
                ElseIf strMsg = "" And strGLCMsg <> "" Then
                    subShowPlaceMessage(True, String.Empty, String.Empty, "A***", dtResult, lstPlaceIDs, FRLMsg, strGlcPlaceCdJa, strGlcPlaceCdOs, strGLCMsg, dtStrageEvaluation)
                End If

                'RM1808***_A***をリストに追加しないように修正する
                'If strMessage_Type = "0" Then
                '    Dim dtResultStrageEvaluation As New DataTable
                '    Dim intEvaluation As Integer = 0
                '    Dim intGLC As Integer = 0

                '    dtResultStrageEvaluation = dtStrageEvaluation.Clone

                '    For Each dr As DataRow In dtStrageEvaluation.Rows
                '        Dim sortNo As Integer = 0
                '        Dim drNew As DataRow = dtResultStrageEvaluation.NewRow

                '        If dr.Item("StrageEvaluationID").ToString.EndsWith("JP1") Or _
                '            Right(dr.Item("StrageEvaluationID").ToString, 3) = "   " Then
                '            intEvaluation = 4
                '        Else
                '            intEvaluation = 2
                '        End If

                '        If dr.Item("StrageEvaluationID").ToString.StartsWith("A***") Then
                '            intGLC = 0
                '        Else
                '            intGLC = -1
                '        End If

                '        sortNo = intEvaluation + intGLC

                '        drNew("StrageEvaluationID") = sortNo
                '        drNew("StrageEvaluationName") = dr("StrageEvaluationName")
                '        dtResultStrageEvaluation.Rows.Add(drNew)
                '    Next

                '    Dim drarray() As DataRow
                '    drarray = dtResultStrageEvaluation.Select("", "StrageEvaluationID", DataViewRowState.CurrentRows)
                '    dtStrageEvaluation.Clear()

                '    For i = 0 To (drarray.Length - 1)
                '        Dim drNew As DataRow = dtStrageEvaluation.NewRow
                '        drNew("StrageEvaluationID") = drarray(i)("StrageEvaluationName").ToString
                '        drNew("StrageEvaluationName") = drarray(i)("StrageEvaluationName").ToString
                '        dtStrageEvaluation.Rows.Add(drNew)
                '    Next
                'End If
            End If
        End If

        Return dtResult

    End Function

    ''' <summary>
    ''' 出荷場所コードにより名称を取得
    ''' </summary>
    ''' <param name="dtPlace">出荷場所テーブル</param>
    ''' <param name="lstPlaceIDs">出荷場所コード</param>
    ''' <param name="lstCountriesByCountryCd">表示可能な国コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetPlaceName(ByVal dtPlace As DataTable, ByRef lstPlaceIDs As ArrayList, _
                                     ByVal lstCountriesByCountryCd As ArrayList) As DataTable
        Dim dtResult As DataTable = dtPlace

        '国内の場合
        If lstPlaceIDs.Count = 1 Then
            Dim strPlace As String = String.Empty
            Dim drPlace As DataRow = dtResult.NewRow

            strPlace = lstPlaceIDs(0)

            If strPlace = "PRC" Then
                If objLoginInfo.SelectLang = "ja" Then
                    drPlace("PlaceID") = "PRC"
                    drPlace("PlaceName") = "CKD中国"
                Else
                    drPlace("PlaceID") = "PRC"
                    drPlace("PlaceName") = "CKD China"
                End If

                If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                    Call SetPlaceMark(strPlace, lstPlaceIDs.Count)
                Else
                    Label7.Visible = False
                    Label9.Visible = False
                    Label10.Visible = False
                    Label11.Visible = False
                End If
            ElseIf strPlace = "KTA" Then
                'ADD BY YGY 20140722    台湾K生産品
                If objLoginInfo.SelectLang = "ja" Then
                    drPlace("PlaceID") = "KTA"
                    drPlace("PlaceName") = "台湾K"
                Else
                    drPlace("PlaceID") = "KTA"
                    drPlace("PlaceName") = "TaiwanK"
                End If

                If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                    Call SetPlaceMark(strPlace, lstPlaceIDs.Count)
                Else
                    Label7.Visible = False
                    Label9.Visible = False
                    Label10.Visible = False
                    Label11.Visible = False
                End If
            ElseIf strPlace = "TYO" Then
                'ADD BY YGY 20140804    台湾T生産品
                If objLoginInfo.SelectLang = "ja" Then
                    drPlace("PlaceID") = "TYO"
                    drPlace("PlaceName") = "台湾T"
                Else
                    drPlace("PlaceID") = "TYO"
                    drPlace("PlaceName") = "TaiwanT"
                End If

                If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                    Call SetPlaceMark(strPlace, lstPlaceIDs.Count)
                Else
                    Label7.Visible = False
                    Label9.Visible = False
                    Label10.Visible = False
                    Label11.Visible = False
                End If
            ElseIf strPlace = "MDN" Then
                'ADD BY YGY 20140822    台湾M生産品
                If objLoginInfo.SelectLang = "ja" Then
                    drPlace("PlaceID") = "MDN"
                    drPlace("PlaceName") = "台湾M"
                Else
                    drPlace("PlaceID") = "MDN"
                    drPlace("PlaceName") = "TaiwanM"
                End If

                If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                    Call SetPlaceMark(strPlace, lstPlaceIDs.Count)
                Else
                    Label7.Visible = False
                    Label9.Visible = False
                    Label10.Visible = False
                    Label11.Visible = False
                End If
            ElseIf strPlace = "OMA" Then
                'ADD BY YGY 20141006    タイOMA生産品
                If objLoginInfo.SelectLang = "ja" Then
                    drPlace("PlaceID") = "OMA"
                    drPlace("PlaceName") = "タイOMA"
                Else
                    drPlace("PlaceID") = "OMA"
                    drPlace("PlaceName") = "ThailandOMA"
                End If

                If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                    Call SetPlaceMark(strPlace, lstPlaceIDs.Count)
                Else
                    Label7.Visible = False
                    Label9.Visible = False
                    Label10.Visible = False
                    Label11.Visible = False
                End If
            ElseIf strPlace = "CJA" Then
                'ADD BY 斉藤 20160708    中国C生産品
                If objLoginInfo.SelectLang = "ja" Then
                    drPlace("PlaceID") = "CJA"
                    drPlace("PlaceName") = "中国C"
                Else
                    drPlace("PlaceID") = "CJA"
                    drPlace("PlaceName") = "ChinaC"
                End If

                If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                    Call SetPlaceMark(strPlace, lstPlaceIDs.Count)
                Else
                    Label7.Visible = False
                    Label9.Visible = False
                    Label10.Visible = False
                    Label11.Visible = False
                End If
            Else
                '日本国コードの変換
                subChangeToJapanesePlaceCd(strPlace)
                drPlace("PlaceID") = strPlace
                drPlace("PlaceName") = strPlace
            End If

            Dim drs() As DataRow = dtResult.Select("PlaceID='" & drPlace("PlaceID") & "' AND PlaceName='" & drPlace("PlaceName") & "'")
            If drs.Count <= 0 Then
                dtResult.Rows.InsertAt(drPlace, 0)
            End If
        Else
            '海外多出荷場所の場合
            '全ての国コードと国名
            Dim dt_country As DataTable = KHCountry.fncGetCountryName(objConBase)

            For inti As Integer = 0 To lstPlaceIDs.Count - 1
                Dim drTmp() As DataRow
                Dim strPlace As String = String.Empty
                Dim drPlace As DataRow = dtResult.NewRow

                strPlace = lstPlaceIDs(inti)
                '日本国コードの変換
                subChangeToJapanesePlaceCd(strPlace)

                '対応する出荷場所名を取得
                If selLang.SelectedValue = "ja" Then
                    drTmp = dt_country.Select("country_cd='" & strPlace & "' AND language_cd='ja'")
                Else
                    drTmp = dt_country.Select("country_cd='" & strPlace & "' AND language_cd='en'")
                End If

                'タイ工場生産の場合「工場」また「Factory」を取り除くこと
                If strPlace.Equals("THF") Then
                    Dim strPlaceName As String = drTmp(0)("country_nm").ToString

                    If selLang.SelectedValue.Equals("ja") Then
                        strPlaceName = strPlaceName.Replace("工場", String.Empty)
                    Else
                        strPlaceName = strPlaceName.Replace("Factory", String.Empty)
                    End If
                    drPlace("PlaceID") = strPlace
                    drPlace("PlaceName") = "CKD " & strPlaceName.Trim

                Else
                    If drTmp.Length > 0 Then
                        drPlace("PlaceID") = strPlace
                        drPlace("PlaceName") = "CKD " & drTmp(0)("country_nm").ToString
                    Else
                        drPlace("PlaceID") = strPlace
                        drPlace("PlaceName") = strPlace
                    End If
                End If

                Dim drs() As DataRow = dtResult.Select("PlaceID='" & drPlace("PlaceID") & "' AND PlaceName='" & drPlace("PlaceName") & "'")
                If drs.Count <= 0 Then
                    dtResult.Rows.Add(drPlace)
                End If

            Next
        End If

        Return dtResult
    End Function

    ''' <summary>
    ''' 表示順番の調整
    ''' </summary>
    ''' <param name="dtPlace"></param>
    ''' <param name="lstCountriesByCountryCd">表示可能リスト(順番)</param>
    ''' <param name="lstPlaceIDs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSetOrder(ByVal dtPlace As DataTable, ByVal lstCountriesByCountryCd As ArrayList, _
                                 ByRef lstPlaceIDs As ArrayList, ByRef dtStrageEvaluation As DataTable) As DataTable
        Dim dtResult As New DataTable
        Dim lstNewPlaceIDs As New ArrayList
        Dim dtResultStrageEvaluation As New DataTable

        'テーブル構造をコピーする
        dtResult = dtPlace.Clone

        '表示可能リストの順番によりdtPlaceとlstPlaceIDsの順番を調整
        For Each strCountryCd As String In lstCountriesByCountryCd
            For Each strPlaceID In lstPlaceIDs
                Dim strTmp As String = strPlaceID
                subChangeToJapaneseCountryCd(strTmp)

                If strTmp.Equals(strCountryCd) Then
                    lstNewPlaceIDs.Add(strPlaceID)
                End If
            Next

            For Each dr As DataRow In dtPlace.Rows
                Dim strPlaceCdJP As String = dr.Item("PlaceID")
                Dim drNew As DataRow = dtResult.NewRow

                '日本コードへ変換
                subChangeToJapaneseCountryCd(strPlaceCdJP)

                If strCountryCd.Equals(strPlaceCdJP) Then
                    drNew("PlaceID") = dr("PlaceID")
                    drNew("PlaceName") = dr("PlaceName")
                    dtResult.Rows.Add(drNew)
                End If
            Next
        Next

        lstPlaceIDs = lstNewPlaceIDs

        '保管場所＆評価タイプの並び替え（優先順位：①評価タイプ海外　②保管場所GLC）
        'テーブル構造をコピーする
        dtResultStrageEvaluation = dtStrageEvaluation.Clone

        Dim intEvaluation As Integer = 0
        Dim intGLC As Integer = 0

        For Each dr As DataRow In dtStrageEvaluation.Rows
            Dim sortNo As Integer = 0
            Dim drNew As DataRow = dtResultStrageEvaluation.NewRow

            If dr.Item("StrageEvaluationID").ToString.EndsWith("JP1") Or _
                Right(dr.Item("StrageEvaluationID").ToString, 3) = "   " Then
                intEvaluation = 4
            Else
                intEvaluation = 2
            End If

            If dr.Item("StrageEvaluationID").ToString.StartsWith("A***") Then
                intGLC = 0
            Else
                intGLC = -1
            End If

            sortNo = intEvaluation + intGLC

            drNew("StrageEvaluationID") = sortNo
            drNew("StrageEvaluationName") = dr("StrageEvaluationName")
            dtResultStrageEvaluation.Rows.Add(drNew)
        Next

        Dim drarray() As DataRow
        drarray = dtResultStrageEvaluation.Select("", "StrageEvaluationID", DataViewRowState.CurrentRows)
        dtStrageEvaluation.Clear()

        For i = 0 To (drarray.Length - 1)
            Dim drNew As DataRow = dtStrageEvaluation.NewRow
            drNew("StrageEvaluationID") = drarray(i)("StrageEvaluationName").ToString
            drNew("StrageEvaluationName") = drarray(i)("StrageEvaluationName").ToString
            dtStrageEvaluation.Rows.Add(drNew)
        Next

        Return dtResult

    End Function

    ''' <summary>
    ''' 出荷場所選択ﾒｯｾｰｼﾞを表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subShowPlaceMessage(ByVal blnShowFlg As Boolean, ByVal strMessage As String, _
                                    ByVal strPlaceYes As String, ByVal strPlaceNo As String, _
                                    ByRef dt As DataTable, ByRef lstPlaceIDs As ArrayList, ByVal strFRLMsg As String, _
                                    ByVal strStockPlaceCdJa As String, ByVal strStockPlaceCdOs As String, ByVal strGLCMsg As String, _
                                    ByRef dtStrageEvaluation As DataTable)
        '引数にstrFRLMsgを追加  2017/02/22 追加
        'Dim drPlace As DataRow = dt.NewRow

        If blnShowFlg = True Then

            'Java引数にstrFRLMsgを追加  2017/02/22 追加
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), strPlaceYes, _
                                               "fncChangePlace('" & Me.ClientID & "_','" & strMessage & "','" & strPlaceYes & "','" & strPlaceNo & "','" & strFRLMsg & _
                                                "','" & strStockPlaceCdJa & "','" & strStockPlaceCdOs & "','" & strGLCMsg & "');", True)

            'drPlace("PlaceID") = strPlaceYes
            'drPlace("PlaceName") = strPlaceYes
            'If Not lstPlaceIDs.Contains(strPlaceYes) Then
            '    lstPlaceIDs.Add(strPlaceYes)
            '    dt.Rows.Add(drPlace)
            'End If

        End If

    End Sub

    ''' <summary>
    ''' 保管場所＆評価タイプを追加する
    ''' </summary>
    ''' <remarks></remarks>
    Private Function subSetStrageEvaluation(ByVal strStrageEvaluation As String, ByRef dtStrageEvaluation As DataTable)

        Dim drStrageEvaluation As DataRow = dtStrageEvaluation.NewRow
        drStrageEvaluation = dtStrageEvaluation.NewRow
        drStrageEvaluation("StrageEvaluationID") = strStrageEvaluation
        drStrageEvaluation("StrageEvaluationName") = strStrageEvaluation
        dtStrageEvaluation.Rows.Add(drStrageEvaluation)

        Return dtStrageEvaluation

    End Function

    ''' <summary>
    ''' 全ての出荷場所を取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetPlace()
        Dim lstPlace As New List(Of String)

        If cmbPlace.Enabled = True Then
            '選択できる場合
            For Each item As ListItem In cmbPlace.Items
                lstPlace.Add(item.Value)
            Next
        Else
            '選択できない場合
            lstPlace.Add(cmbPlace.SelectedValue)
        End If

        Session.Add("ShipPlaces", lstPlace)

    End Sub

    ''' <summary>
    ''' 価格マスタの出荷場所を追加するかどうかの判断ロジック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncUseSeriesCountryOrNot(ByRef strSeriesCountry As String) As Boolean
        Dim blnResult As Boolean = False
        Dim strCurrency As String = String.Empty

        strCurrency = objKtbnStrc.strcSelection.strCurrency

        If strCurrency Is Nothing OrElse strCurrency.Equals(String.Empty) Then
            strCurrency = "JPY"
        End If

        Using daPrice As New DS_TankaTableAdapters.kh_priceTableAdapter
            Dim dtPrice As New DS_Tanka.kh_priceDataTable

            dtPrice = daPrice.GetDataByKataban(objKtbnStrc.strcSelection.strFullKataban, strCurrency, Now)

            If dtPrice.Rows.Count > 0 Then
                'フル形番の場合は価格マスタの出荷場所を追加
                blnResult = False
            Else
                'フル形番以外の場合は
                Using daSeries As New DS_TankaTableAdapters.kh_series_katabanTableAdapter
                    Dim dtSeries As New DS_Tanka.kh_series_katabanDataTable

                    dtSeries = daSeries.GetDataBySeriesAndKey(objKtbnStrc.strcSelection.strSeriesKataban, objKtbnStrc.strcSelection.strKeyKataban, Now)

                    If dtSeries.Rows.Count > 0 Then
                        'シリーズの国コードが"JPN"以外の時はシリーズの国コードを使う
                        If Not dtSeries.Rows(0).Item("country_cd").ToString.Equals("JPN") Then
                            'RM1804035_生産国表示制御（中国ログイン且つＰＷＣの場合はシリーズの国コードをスルー）
                            If Me.objUserInfo.CountryCd = "PRC" And objKtbnStrc.strcSelection.strSeriesKataban = "PWC" Then
                                blnResult = True
                            Else
                                strSeriesCountry = dtSeries.Rows(0).Item("country_cd").ToString
                                blnResult = True
                            End If
                        Else
                            blnResult = False
                        End If
                    End If
                End Using
            End If
        End Using

        Return blnResult
    End Function

    ''' <summary>
    ''' 欧州対応メッセージ
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetEuropeMessage()

        Dim objKataban As New KHKataban

        Me.Label36.Visible = False                       '欧州市場では販売しておりません
        Me.Label37.Visible = False                       'CKDにお問い合わせください

        '欧州輸出不可設定
        If objUserInfo.BaseCd.Equals("07") Then

            '欧州市場では販売しておりません
            If objKataban.fncELKatabanCheck(objCon, objKtbnStrc.strcSelection.strFullKataban, "2") Then

                Me.Label36.Visible = True

            End If

            'CKDにお問い合わせください
            If objKataban.fncELKatabanCheck(objCon, objKtbnStrc.strcSelection.strFullKataban, "3") Then

                Me.Label37.Visible = True

            End If

        End If

    End Sub

    'RM1707001　中国対応メッセージ追加　2017/07/05
    ''' <summary>
    ''' 中国対応メッセージ
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetChinaMessage()

        Dim objKataban As New KHKataban

        Me.Label40.Visible = False                       '中国軍事懸念メッセージ

        '中国出荷不可設定
        If objUserInfo.CountryCd.Equals("PRC") Then

            '欧州市場では販売しておりません
            If objKataban.fncELKatabanCheck(objCon, objKtbnStrc.strcSelection.strFullKataban, "4") Then

                Me.Label40.Visible = True

            End If

        End If

    End Sub

    ''' <summary>
    ''' FRLメッセージ取得処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Function fncFRLMessage() As String

        Dim strResult As String = String.Empty

        '特定の型番の場合は追加メッセージを表示するよう変更  RM1702023  2017/02/22 追加

        If objKtbnStrc.strcSelection.strSeriesKataban Is Nothing Then
            '選択形番情報取得
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, objUserInfo.UserId, objLoginInfo.SessionId, 1)
        End If

        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban.Trim
        Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban.Trim

        Select Case strSeriesKata

            Case "C1000", "C1010", "C1020", "C1030", "C1040", "C1050", "C1060", "C2000", "C2010", "C2020", _
                 "C2030", "C2040", "C2050", "C2060", "C2500", "C2520", "C2530", "C2550", "C3000", "C3010", _
                 "C3020", "C3030", "C3040", "C3050", "C3060", "C3070", "C4000", "C4010", "C4020", "C4030", _
                 "C4040", "C4050", "C4060", "C4070", "C6500", "C8000", "C8010", "C8020", "C8030", "C8040", _
                 "C8050", "C8060", "C8070"

                '海外生産拠点との取引をしていないユーザ（日本、ベトナムなど）、かつ出荷場所を表示するユーザーのみ表示する

                If dt_Addinfo Is Nothing Then

                    '形番+仕様書Noセット 標準納期計算用
                    Dim strFullKatabanSiyouNo As String = lblSeriesKat.Text

                    dt_Addinfo = subAddInfoDispGet(objCon, Me.objUserInfo.UserId, _
                                                   Me.objLoginInfo.SessionId, selLang.SelectedValue, Me.objUserInfo.AddInformationLvl, _
                                                   strFullKatabanSiyouNo, objKtbnStrc)
                End If

                Select Case strKeyKata
                    Case "X"

                    Case Else
                        If dt_Addinfo.Select("strLevel='2'").Count > 0 Then
                            strResult = ClsCommon.fncGetMsg(selLang.SelectedValue, "W9200")
                        Else
                            strResult = ""
                        End If
                End Select

            Case Else
                'その他の場合は表示しない
                strResult = ""

        End Select

        Return strResult
    End Function

#End Region

End Class
