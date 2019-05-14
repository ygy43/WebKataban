Imports WebKataban.ClsCommon
Imports System.IO

Public Class WebUC_KatSep
    Inherits KHBase

#Region "プロパティー"
    'ビジネスロジック
    Private bllType As New TypeBLL
#End Region
    
#Region "イベント"
    Public Event BackToType()
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Me.lblSeparator.Text = String.Empty
        Me.lblPrice.Text = String.Empty
        Call Page_Load(Me, Nothing)
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
        txtKata.Style.Add("text-transform", "uppercase")

        Try
            GVDetail.Visible = False
            Call SetAllFontName(Me)
            Me.txtKata.Focus()
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' クリア
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub frmClear(Optional intMode As Integer = 0)
        If intMode = 0 Then Me.txtKata.Text = String.Empty
        lblKataName.Text = String.Empty
        Me.lblSeparator.Text = String.Empty
        Me.lblPrice.Text = String.Empty
        GVDetail.DataSource = New DataTable
        GVDetail.DataBind()
        GVTitle.DataSource = New DataTable
        GVTitle.DataBind()
        GVYouso.DataSource = New DataTable
        GVYouso.DataBind()
        Me.txtKata.Focus()
    End Sub

    ''' <summary>
    ''' 形番分解
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnKatSep_Click(sender As Object, e As EventArgs) Handles btnKatSep.Click
        If txtKata.Text.ToString.Trim.Length <= 0 Then
            Me.lblSeparator.Text = String.Empty
            Me.lblPrice.Text = String.Empty
            ' 確認メッセージ出力
            Dim sbScript As New StringBuilder
            Dim strMessage As String = "形番を入力してください。"
            sbScript.Append("alert('" & strMessage & "');")
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "alert", sbScript.ToString, True)
            Me.txtKata.Focus()
            Exit Sub
        Else
            Call frmClear(1) '画面クリア

            Dim strSeries As String = String.Empty
            Dim strKeyKata As String = String.Empty
            Dim strKataName As String = String.Empty
            Dim strSpecNo As String = String.Empty
            Dim strPriceNo As String = String.Empty
            Dim strItem1(24) As String
            Dim strItemName1(24) As String
            Dim strHyphen1(24) As String
            Dim strElement_div1(24) As String
            Dim strStructure_div(24) As String

            Dim dr As DataRow = Nothing
            Dim dtval As New DataTable
            Dim dc As New DataColumn("title_nm")
            dtval.Columns.Add(dc)
            dc = New DataColumn("colValue")
            dtval.Columns.Add(dc)

            Dim dtTitle As New DataTable
            dtTitle = dtval.Clone
            Dim dtYouso As New DataTable
            dtYouso = dtval.Clone
            dc = New DataColumn("colHyphen")
            dtYouso.Columns.Add(dc)

            If KHKatabanSeparator.GetSeparatorData(Me.txtKata.Text.ToString.Trim.ToUpper, strSeries, strKeyKata, _
                         strKataName, strSpecNo, strPriceNo, strItem1, strItemName1, strHyphen1, strStructure_div, strElement_div1) Then

                '韓国テストため一時的に”JPY”に設定する
                If objKtbnStrc.strcSelection.strCurrency Is Nothing Then
                    objKtbnStrc.strcSelection.strCurrency = "JPY"
                End If

                Me.lblKataName.Text = strKataName

                'タイトル
                dr = dtYouso.NewRow
                dr("title_nm") = "機種"
                dr("colValue") = strSeries
                dr("colHyphen") = IIf(strHyphen1(0) = "1", "－", "")
                dtYouso.Rows.Add(dr)
                For inti As Integer = 0 To strItem1.Length - 1
                    dr = dtYouso.NewRow
                    dr("title_nm") = strItemName1(inti).ToString
                    If strItem1(inti) Is Nothing Then
                        dr("colValue") = String.Empty
                    Else
                        dr("colValue") = strItem1(inti).ToString
                    End If
                    If inti < strItem1.Length - 1 Then
                        dr("colHyphen") = IIf(strHyphen1(inti) = "1", "－", "")
                    End If
                    dtYouso.Rows.Add(dr)
                Next
                Me.GVYouso.DataSource = dtYouso
                Me.GVYouso.DataBind()

                '引当シリーズ形番追加(機種)
                '通貨追加
                Call bllType.subInsertSelSrsKtbnMdl(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                    strSeries, strKeyKata, strKataName, objKtbnStrc.strcSelection.strCurrency)

                Dim strcCompData As New YousoBLL.CompData
                strcCompData.strSeriesKataban = strSeries
                strcCompData.strKeyKataban = strKeyKata

                YousoBLL.fncKatabanStrcSelect(objCon, strcCompData, "ja") '形番構成取得
                YousoBLL.subKtbnStrcEleSelect(objCon, strcCompData)                        '形番構成要素取得

                For intLoopCnt = 1 To strcCompData.strElementDiv.Length - 1
                    Call objKtbnStrc.subSelKtbnStrcIns(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                                       intLoopCnt, strcCompData.strElementDiv(intLoopCnt), _
                                                       strcCompData.strStructureDiv(intLoopCnt), _
                                                       strcCompData.strAdditionDiv(intLoopCnt), _
                                                       strcCompData.strHyphenDiv(intLoopCnt), _
                                                       strcCompData.strKtbnStrcNm(intLoopCnt), 0)
                Next
                '引当情報取得
                Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

                Dim strOp(strItem1.Length) As String
                strOp(0) = ""
                For inti As Integer = 0 To strItem1.Length - 1
                    If strItem1(inti) Is Nothing Then
                        strOp(inti + 1) = String.Empty
                    Else
                        strOp(inti + 1) = strItem1(inti)
                    End If
                Next
                Me.objKtbnStrc.strcSelection.strOpSymbol = strOp
            Else
                Me.lblSeparator.Text = "形番分解失敗（フル形番の可能性あり）"
                Me.GVDetail.Visible = False
            End If

            '分解失敗しても単価を計算する、特注形番の可能性がある
            Dim flgPrice As Boolean = False      'False:Full形番、True:分解

            Dim bolSpecInput As Boolean = False
            If Len(strSpecNo.Trim) <> 0 Then
                Select Case strSpecNo.Trim
                    Case "00"
                        'ページ遷移(ロッド先端形状オーダーメイド寸法入力画面)
                    Case "01", "02", "03", "04", "05", "06", "07", "08", "10", "11", _
                         "13", "14", "15", "16", "96"
                        bolSpecInput = True
                    Case "09"
                        If objKtbnStrc.strcSelection.strOpSymbol(6).ToString.Trim <> "" Then
                            bolSpecInput = True
                        End If
                    Case "12", "18", "19", "20", "21", "22", "23", "94", "95"
                        If KHKatabanSeparator.fncMixCheck(strSeries, objKtbnStrc.strcSelection.strOpSymbol) Then
                            bolSpecInput = True
                        End If
                    Case "17"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Then
                            bolSpecInput = True
                        End If
                    Case "51"
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                            bolSpecInput = True
                        End If
                    Case "52", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", _
                         "65", "66", "67", "68", "69", "70", "71", "72", "89", "90", "91", "92", "98"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            bolSpecInput = True
                        End If
                    Case "53", "73", "74", "75", "76", "77", "78", "79", "80", "81", _
                         "82", "83", "84", "85", "86", "87", "88", "93"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "80" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "81" Then
                            bolSpecInput = True
                        End If
                    Case "A1", "A2", "A9", "B1", "B2", "B3", "B4"
                        bolSpecInput = True
                End Select
            End If

            If bolSpecInput Then
                Me.lblPrice.Text = "マニホールド対象形番、価格を計算できません。"
                Me.GVDetail.Visible = False
            Else
                Dim objUnitPrice As New KHUnitPrice
                Dim strChangePlaceCd As String = String.Empty                         '出荷場所(変換)
                Dim objOption As New KHOptionCtl

                'フル形番の設定
                objKtbnStrc.strcSelection.strFullKataban = Me.txtKata.Text.ToString.Trim.ToUpper

                '価格の取得
                Call objUnitPrice.subPriceInfoSet_ForkatOut(objCon, objKtbnStrc, Me.objUserInfo.CountryCd, "")

                '原価積算No取得
                objKtbnStrc.strcSelection.strCostCalcNo = objOption.fncCostCalcNoGet(objKtbnStrc, objKtbnStrc.strcSelection.strKatabanCheckDiv)

                '①新しい形番ﾁｪｯｸ区分を反映する
                If KHKataban.subJapanChinaAmount(Me.txtKata.Text.ToString.Trim.ToUpper) Then objKtbnStrc.strcSelection.strKatabanCheckDiv = "1"

                'タイトル
                dr = dtTitle.NewRow
                dr("title_nm") = "機種"
                dr("colValue") = strSeries
                dtTitle.Rows.Add(dr)
                dr = dtTitle.NewRow
                dr("title_nm") = "キー"
                dr("colValue") = strKeyKata
                dtTitle.Rows.Add(dr)
                dr = dtTitle.NewRow
                dr("title_nm") = "仕様No"
                dr("colValue") = strSpecNo
                dtTitle.Rows.Add(dr)
                dr = dtTitle.NewRow
                dr("title_nm") = "価格No"
                dr("colValue") = strPriceNo
                dtTitle.Rows.Add(dr)
                dr = dtTitle.NewRow
                dr("title_nm") = "ﾁｪｯｸ区分"
                dr("colValue") = "Z" & objKtbnStrc.strcSelection.strKatabanCheckDiv
                dtTitle.Rows.Add(dr)
                dr = dtTitle.NewRow
                dr("title_nm") = "プラント"
                '変換必要があるかどうかの判断
                dr("colValue") = changeShipPlace(objKtbnStrc)
                dtTitle.Rows.Add(dr)

                '簡易オーダー
                dr = dtTitle.NewRow
                dr("title_nm") = "簡易オーダー"
                dr("colValue") = objKtbnStrc.strcSelection.strCostCalcNo
                dtTitle.Rows.Add(dr)

                Me.GVTitle.DataSource = dtTitle
                Me.GVTitle.DataBind()

                '価格
                Me.GVDetail.Visible = True
                dr = dtval.NewRow
                dr("title_nm") = "定価"
                dr("colValue") = CInt(objKtbnStrc.strcSelection.intListPrice)
                dtval.Rows.Add(dr)
                dr = dtval.NewRow
                dr("title_nm") = "登録店"
                dr("colValue") = CInt(objKtbnStrc.strcSelection.intRegPrice)
                dtval.Rows.Add(dr)
                dr = dtval.NewRow
                dr("title_nm") = "ＳＳ店"
                dr("colValue") = CInt(objKtbnStrc.strcSelection.intSsPrice)
                dtval.Rows.Add(dr)
                dr = dtval.NewRow
                dr("title_nm") = "ＢＳ店"
                dr("colValue") = CInt(objKtbnStrc.strcSelection.intBsPrice)
                dtval.Rows.Add(dr)
                dr = dtval.NewRow
                dr("title_nm") = "ＧＳ店"
                dr("colValue") = CInt(objKtbnStrc.strcSelection.intGsPrice)
                dtval.Rows.Add(dr)
                dr = dtval.NewRow
                dr("title_nm") = "ＰＳ店"
                dr("colValue") = CInt(objKtbnStrc.strcSelection.intPsPrice)

                dtval.Rows.Add(dr)
                Me.GVDetail.DataSource = dtval
                GVDetail.DataBind()
                Call SetAllFontName(Me)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 一括分解/一括分解（購入価格込み）
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnKatSepAll_Click(sender As Object, e As EventArgs) Handles btnKatSepAll.Click, btnKatSepAllWithNetPrice.Click

        Try
            If txtKatabanFilePath.Text.ToString.Trim.Length <= 0 Then
                Me.lblSeparator.Text = String.Empty
                Me.lblPrice.Text = String.Empty
                ' 確認メッセージ出力
                Dim sbScript As New StringBuilder
                Dim strMessage As String = "形番を入力してください。"
                sbScript.Append("alert('" & strMessage & "');")
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "alert", sbScript.ToString, True)
                Me.txtKata.Focus()
                Exit Sub
            Else
                ''ファイルをアップロード
                'Dim strFileName As String = Path.GetFileName(FileUpload1.FileName)
                'FileUpload1.SaveAs(Server.MapPath("~/TempFiles/" & strFileName))

                '画面クリア
                Call frmClear(1)

                '形番の読み込み
                Dim strKatabans As List(Of String) = IO.File.ReadAllLines(txtKatabanFilePath.Text.Trim).ToList

                '処理結果
                Dim strSepResults As New StringBuilder
                strSepResults.AppendLine("形番" & ControlChars.Tab & _
                                         "商品名" & ControlChars.Tab & _
                                         "LS" & ControlChars.Tab & _
                                         "RG" & ControlChars.Tab & _
                                         "SS" & ControlChars.Tab & _
                                         "BS" & ControlChars.Tab & _
                                         "GS" & ControlChars.Tab & _
                                         "PS" & ControlChars.Tab & _
                                         "ﾁｪｯｸ区分" & ControlChars.Tab & _
                                         "標準納期コード" & ControlChars.Tab & _
                                         "標準納期" & ControlChars.Tab & _
                                         "適用個数" & ControlChars.Tab & _
                                         "在庫区分" & ControlChars.Tab & _
                                         "プラント" & ControlChars.Tab & _
                                         "原価積算No." & ControlChars.Tab & _
                                         "購入価格" & ControlChars.Tab & _
                                         "簡易オーダー")
                For Each kataban In strKatabans
                    If kataban.Equals(String.Empty) Then
                        Continue For
                    End If
                    '初期化
                    strSepResults.Append(kataban & ControlChars.Tab)
                    objKtbnStrc = New KHKtbnStrc
                    Dim strBunkaiResult As String = ControlChars.Tab
                    Dim strSeries As String = String.Empty
                    Dim strKeyKata As String = String.Empty
                    Dim strKataName As String = String.Empty
                    Dim strSpecNo As String = String.Empty
                    Dim strPriceNo As String = String.Empty
                    Dim strItem1(24) As String
                    Dim strItemName1(24) As String
                    Dim strHyphen1(24) As String
                    Dim strElement_div1(24) As String
                    Dim strStructure_div(24) As String
                    Dim dr As DataRow = Nothing
                    Dim dtval As New DataTable
                    Dim dc As New DataColumn("title_nm")

                    dtval.Columns.Add(dc)
                    dc = New DataColumn("colValue")
                    dtval.Columns.Add(dc)

                    Dim dtTitle As New DataTable
                    dtTitle = dtval.Clone
                    Dim dtYouso As New DataTable
                    dtYouso = dtval.Clone
                    dc = New DataColumn("colHyphen")
                    dtYouso.Columns.Add(dc)
                    '形番分解
                    If KHKatabanSeparator.GetSeparatorData(kataban.Trim.ToUpper, strSeries, strKeyKata, _
                                 strKataName, strSpecNo, strPriceNo, strItem1, strItemName1, strHyphen1, strStructure_div, strElement_div1) Then

                        '韓国テストため一時的に”JPY”に設定する
                        If objKtbnStrc.strcSelection.strCurrency Is Nothing Then
                            objKtbnStrc.strcSelection.strCurrency = "JPY"
                        End If

                        Me.lblKataName.Text = strKataName

                        'タイトル
                        dr = dtYouso.NewRow
                        dr("title_nm") = "機種"
                        dr("colValue") = strSeries
                        dr("colHyphen") = IIf(strHyphen1(0) = "1", "－", "")
                        dtYouso.Rows.Add(dr)
                        For inti As Integer = 0 To strItem1.Length - 1
                            dr = dtYouso.NewRow
                            dr("title_nm") = strItemName1(inti).ToString
                            If strItem1(inti) Is Nothing Then
                                dr("colValue") = String.Empty
                            Else
                                dr("colValue") = strItem1(inti).ToString
                            End If
                            If inti < strItem1.Length - 1 Then
                                dr("colHyphen") = IIf(strHyphen1(inti) = "1", "－", "")
                            End If
                            dtYouso.Rows.Add(dr)
                        Next
                        Me.GVYouso.DataSource = dtYouso
                        Me.GVYouso.DataBind()

                        '引当シリーズ形番追加(機種)
                        '通貨の追加
                        Call bllType.subInsertSelSrsKtbnMdl(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                            strSeries, strKeyKata, strKataName, objKtbnStrc.strcSelection.strCurrency)

                        Dim strcCompData As New YousoBLL.CompData
                        strcCompData.strSeriesKataban = strSeries
                        strcCompData.strKeyKataban = strKeyKata

                        YousoBLL.fncKatabanStrcSelect(objCon, strcCompData, "ja") '形番構成取得
                        YousoBLL.subKtbnStrcEleSelect(objCon, strcCompData)                        '形番構成要素取得

                        For intLoopCnt = 1 To strcCompData.strElementDiv.Length - 1
                            Call objKtbnStrc.subSelKtbnStrcIns(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                                               intLoopCnt, strcCompData.strElementDiv(intLoopCnt), _
                                                               strcCompData.strStructureDiv(intLoopCnt), _
                                                               strcCompData.strAdditionDiv(intLoopCnt), _
                                                               strcCompData.strHyphenDiv(intLoopCnt), _
                                                               strcCompData.strKtbnStrcNm(intLoopCnt), 0)
                        Next
                        '引当情報取得
                        Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

                        Dim strOp(strItem1.Length) As String
                        strOp(0) = ""
                        For inti As Integer = 0 To strItem1.Length - 1
                            If strItem1(inti) Is Nothing Then
                                strOp(inti + 1) = String.Empty
                            Else
                                strOp(inti + 1) = strItem1(inti)
                            End If
                        Next
                        Me.objKtbnStrc.strcSelection.strOpSymbol = strOp
                    Else
                        strBunkaiResult &= "形番分解失敗" & ControlChars.Tab

                        'フル形番の場合は製品名を取得
                        objKtbnStrc.strcSelection.strGoodsNm = fncGetFullKatabanInfo(kataban)

                        'strSepResults.Append("形番分解失敗（フル形番の可能性あり）" & ControlChars.Tab)
                    End If

                    '分解失敗しても単価を計算する、特注形番の可能性がある
                    Dim flgPrice As Boolean = False      'False:Full形番、True:分解

                    Dim bolSpecInput As Boolean = False
                    If Len(strSpecNo.Trim) <> 0 Then
                        Select Case strSpecNo.Trim
                            Case "00"
                                'ページ遷移(ロッド先端形状オーダーメイド寸法入力画面)
                            Case "01", "02", "03", "04", "05", "06", "07", "08", "10", "11", _
                                 "13", "14", "15", "16", "96"
                                bolSpecInput = True
                            Case "09"
                                If objKtbnStrc.strcSelection.strOpSymbol(6).ToString.Trim <> "" Then
                                    bolSpecInput = True
                                End If
                            Case "12", "18", "19", "20", "21", "22", "23", "94", "95"
                                If KHKatabanSeparator.fncMixCheck(strSeries, objKtbnStrc.strcSelection.strOpSymbol) Then
                                    bolSpecInput = True
                                End If
                            Case "17"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Then
                                    bolSpecInput = True
                                End If
                            Case "51"
                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                                    bolSpecInput = True
                                End If
                            Case "52", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", _
                                 "65", "66", "67", "68", "69", "70", "71", "72", "89", "90", "91", "92"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                                    bolSpecInput = True
                                End If
                            Case "53", "73", "74", "75", "76", "77", "78", "79", "80", "81", _
                                 "82", "83", "84", "85", "86", "87", "88", "93"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "80" Or _
                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "81" Then
                                    bolSpecInput = True
                                End If
                            Case "A1", "A2", "A9", "B1", "B2", "B3", "B4"
                                bolSpecInput = True
                        End Select
                    End If

                    '結果表示
                    If bolSpecInput Then
                        strBunkaiResult &= ControlChars.Tab & ControlChars.Tab & ControlChars.Tab & ControlChars.Tab & _
                            ControlChars.Tab & ControlChars.Tab & ControlChars.Tab & ControlChars.Tab & ControlChars.Tab & _
                            "マニホールド対象形番、価格を計算できません。"
                        'strSepResults.AppendLine("マニホールド対象形番、価格を計算できません。")
                        strSepResults.AppendLine(strBunkaiResult)
                    Else
                        '適用個数、標準納期と在庫区分の取得
                        'Dim standardnoki As New WSKatahikirenkei.StandardNoki
                        'Dim KatabanWs As New WSKatahikirenkei.WSKatahikiRenkei

                        Dim standardnoki As New WcfKatahikiRenkei.StandardNouki
                        Dim KatabanWs As New WcfKatahikiRenkei.KatahikiRenkeiService

                        '価格の取得
                        Dim objUnitPrice As New KHUnitPrice
                        Dim objOption As New KHOptionCtl

                        'standardnoki = KatabanWs.GetStandardNoki(kataban, objKtbnStrc.strcSelection.strPlaceCd)
                        objKtbnStrc.strcSelection.strFullKataban = kataban

                        Call objUnitPrice.subPriceInfoSet_ForkatOut(objCon, objKtbnStrc, Me.objUserInfo.CountryCd, "")

                        '原価積算No取得
                        objKtbnStrc.strcSelection.strCostCalcNo = objOption.fncCostCalcNoGet(objKtbnStrc, objKtbnStrc.strcSelection.strKatabanCheckDiv)

                        '①新しい形番ﾁｪｯｸ区分を反映する
                        If KHKataban.subJapanChinaAmount(kataban) Then objKtbnStrc.strcSelection.strKatabanCheckDiv = "1"

                        '出荷場所の変換
                        objKtbnStrc.strcSelection.strPlaceCd = changeShipPlace(objKtbnStrc)

                        '出力内容の作成
                        strSepResults.AppendLine(objKtbnStrc.strcSelection.strGoodsNm & ControlChars.Tab & _
                                                 fncSetEmpty(objKtbnStrc.strcSelection.intListPrice) & ControlChars.Tab & _
                                                 fncSetEmpty(objKtbnStrc.strcSelection.intRegPrice) & ControlChars.Tab & _
                                                 fncSetEmpty(objKtbnStrc.strcSelection.intSsPrice) & ControlChars.Tab & _
                                                 fncSetEmpty(objKtbnStrc.strcSelection.intBsPrice) & ControlChars.Tab & _
                                                 fncSetEmpty(objKtbnStrc.strcSelection.intGsPrice) & ControlChars.Tab & _
                                                 fncSetEmpty(objKtbnStrc.strcSelection.intPsPrice) & ControlChars.Tab & _
                                                 "Z" & objKtbnStrc.strcSelection.strKatabanCheckDiv & ControlChars.Tab & _
                                                 standardnoki.StdDate & ControlChars.Tab & _
                                                 fncConvertNoki(standardnoki.StdDate.ToString) & ControlChars.Tab & _
                                                 standardnoki.Quantity & ControlChars.Tab & _
                                                 standardnoki.ZaikoFlg & ControlChars.Tab & _
                                                 objKtbnStrc.strcSelection.strPlaceCd & ControlChars.Tab & _
                                                 objKtbnStrc.strcSelection.strCostCalcNo & ControlChars.Tab & _
                                                 objKtbnStrc.strcSelection.strKatabanCheckDiv & ControlChars.Tab & _
                                                 standardnoki.StdDate & ControlChars.Tab & _
                                                 fncConvertNoki(standardnoki.StdDate.ToString) & ControlChars.Tab & _
                                                 standardnoki.Quantity & ControlChars.Tab & _
                                                 standardnoki.ZaikoFlg & ControlChars.Tab & _
                                                 objKtbnStrc.strcSelection.strPlaceCd & ControlChars.Tab & _
                                                 objKtbnStrc.strcSelection.strCostCalcNo & ControlChars.Tab & _
                                                 CalculateNetPrice(objKtbnStrc.strcSelection.intGsPrice,
                                                                   objUserInfo.CountryCd,
                                                                   kataban.Split("-").First) & ControlChars.Tab & _
                                                 strBunkaiResult)
                    End If

                    Dim strResult As String = String.Empty

                    If (Not objKtbnStrc.strcSelection.intGsPrice = 0D) AndAlso strSepResults.ToString.Contains("形番分解失敗") Then
                        'フル形番の場合は「形番分解失敗」メッセージを出力しない
                        strResult = strSepResults.ToString.Replace("形番分解失敗", String.Empty)
                    Else
                        strResult = strSepResults.ToString
                    End If
                    IO.File.WriteAllText(My.Settings.LogFolder & "一括形番分解結果.txt", strResult)
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 出荷場所の変換
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function changeShipPlace(ByVal objKtbnStrc As WebKataban.KHKtbnStrc) As String
        Dim strResult As String = String.Empty
        Dim strChangePlaceCd As String = String.Empty
        Dim strEvaluationType As String = String.Empty
        Dim strSearchDiv As String = String.Empty

        '変換必要があるかどうかの判断
        If (KHCountry.fncPlaceChangeInfo(objCon, objKtbnStrc.strcSelection.strFullKataban, _
                                                strChangePlaceCd, strEvaluationType, strSearchDiv)) Then
            strResult = strChangePlaceCd
        Else
            strResult = objKtbnStrc.strcSelection.strPlaceCd
        End If

        Return strResult
    End Function

    ''' <summary>
    ''' 戻る
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        RaiseEvent BackToType()
    End Sub

    ''' <summary>
    ''' クリア
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Call frmClear()
    End Sub

    ''' <summary>
    ''' 空白の変換
    ''' </summary>
    ''' <param name="strValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSetEmpty(ByVal strValue As String) As String

        Return CInt(IIf(strValue.Equals(String.Empty), 0, strValue)).ToString()

    End Function

    ''' <summary>
    ''' 標準納期の名称を取得
    ''' </summary>
    ''' <param name="strCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncConvertNoki(ByVal strCode As String) As String
        '結果
        Dim strResult As String = String.Empty

        Select Case strCode
            Case "90"
                '在庫対応
                strResult = "在庫対応"
            Case "91"
                '即日対応
                strResult = "即日対応"
            Case "92"
                'AM I/Pのみ即日対応
                strResult = "AM I/Pのみ即日対応"
            Case "97", "98"
                '納期工場へ問い合わせ
                strResult = "納期工場へ問い合わせ"
            Case "-1"
                '形番未入力エラー
                strResult = "形番未入力エラー"
            Case "-2"
                '形番分解エラー
                strResult = "形番分解エラー"
            Case "-3"
                '標準納期算出エラー
                strResult = "標準納期算出エラー"
            Case "-99"
                'その他エラー
                strResult = "その他エラー"
            Case Else
                'その他
                strResult = String.Format("{0}日間(実稼働日)", strCode)
        End Select

        Return strResult

    End Function

    ''' <summary>
    ''' フル形番の情報の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetFullKatabanInfo(ByVal strFullKataban As String) As String

        Dim strResult As String = String.Empty
        Dim dtFullKataban As New DS_KatSep.kh_fullKataban_infoDataTable

        'フル形番情報の取得
        Using da As New DS_KatSepTableAdapters.kh_fullKataban_infoTableAdapter
            da.FillByKataban(dtFullKataban, "ja", Now, "JPY", strFullKataban)
        End Using

        '製品名の設定
        If dtFullKataban.Rows.Count > 0 Then
            'コード文字列取得
            Dim strSystem As String = ClsCommon.fncGetMsg("ja", "I0040")
            Dim strParts As String = ClsCommon.fncGetMsg("ja", "I0030")
            Dim strFor As String = ClsCommon.fncGetMsg("ja", "I0050")

            'フル形番の場合
            With dtFullKataban.Rows(0)
                If CInt(.Item("kataban_check_div")) < 4 Then
                    If IsDBNull(.Item("model_nm")) _
                        OrElse .Item("model_nm").Equals(String.Empty) Then

                        If IsDBNull(.Item("parts_nm")) _
                        OrElse .Item("parts_nm").Equals(String.Empty) Then
                            strResult = "(" & strSystem & ")"
                        Else
                            strResult = .Item("parts_nm")
                        End If
                    Else
                        If IsDBNull(.Item("parts_nm")) _
                        OrElse .Item("parts_nm").Equals(String.Empty) Then
                            strResult = .Item("model_nm")
                        Else
                            strResult = .Item("model_nm") & "(" & .Item("parts_nm") & ")"
                        End If
                    End If
                Else
                    If IsDBNull(.Item("model_nm")) _
                        OrElse .Item("model_nm").Equals(String.Empty) Then

                        If IsDBNull(.Item("parts_nm")) _
                        OrElse .Item("parts_nm").Equals(String.Empty) Then
                            strResult = "(" & strSystem & ")"
                        Else
                            strResult = strParts & "(" & .Item("parts_nm") & ")"
                        End If
                    Else
                        If IsDBNull(.Item("parts_nm")) _
                        OrElse .Item("parts_nm").Equals(String.Empty) Then
                            strResult = strParts & "(" & .Item("model_nm") & ")"
                        Else
                            strResult = strParts & "(" & .Item("model_nm") & strFor & "(" & .Item("parts_nm") & "))"
                        End If
                    End If
                End If
            End With
        End If

        Return strResult

    End Function

    ''' <summary>
    ''' 購入価格の計算
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CalculateNetPrice(strGsPrice As String, strCountryCode As String, strSeries As String) As String
        Dim result As String = "0"
        Dim rateDal As New UnitPriceDAL
        Dim unitPrice As New KHUnitPrice

        Dim dtRate = rateDal.fncSelectRateFobprice(objConBase, strCountryCode, strSeries, "JPN")

        If dtRate.Rows.Count > 0 Then
            Dim strCurrency = dtRate.Rows(0)("currency_cd")    '変更通貨を取得する
            Dim strTypeFOB As String = dtRate.Rows(0)("TypeFOB").ToString.Trim
            Dim strPosFOB As String = dtRate.Rows(0)("PosFOB").ToString.Trim
            Dim decFOBRate As Decimal = dtRate.Rows(0)("fob_rate")

            Dim intGsPrice = CType(strGsPrice, Integer)

            '端数処理
            '購入定価 Fobprice = GS価格 * 掛率(fob_rate) * 為替レート 
            result = unitPrice.subFractionProc(intGsPrice * decFOBRate, strTypeFOB, strPosFOB).ToString
        End If

        Return result

    End Function

End Class
