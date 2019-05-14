Imports WebKataban.ClsCommon
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO

''' <summary>
''' メンテナンステスト関連
''' </summary>
''' <remarks></remarks>
Public Class WebUC_100Test
    Inherits KHBase

#Region "定数"

    Public Const strMaru As String = "○"

    Public Const strBatsu As String = "×"

    ''' <summary>
    ''' 形番分解結果
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum SeperateResult
        SUCCESS

        PRICE_ERROR

        MANIFOLD_ERROR

    End Enum

#End Region

#Region "イベント"

    Public Event BackToType()
    Public Event GoToType()

#End Region
    
    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Call ClearSession()
        Me.HidSelRowID.Value = String.Empty
        Me.HidSelKey.Value = String.Empty
        Page_Load(Me, Nothing)
    End Sub

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub
        If Me.objUserInfo.UserId Is Nothing Then Exit Sub
        Call SetAllFontName(Me)
        'Me.lblGroup.Text = "12-13"
    End Sub

    ''' <summary>
    ''' 100万件テスト（グループ毎に）
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btn100Test_Click(sender As Object, e As EventArgs) Handles btn100Test.Click
        Try
            For intloop As Integer = 1 To 31

                Dim ds_table As New DataSet
                'データを取得する（グループ毎に）
                Dim dt_100 As New DS_100Test.kh_TEST_NEWDataTable

                Using da As New DS_100TestTableAdapters.kh_TEST_NEWTableAdapter
                    da.FillByGroup(dt_100, intloop)

                    '形番分解用情報の取得
                    ds_table = fncLoadInfo(dt_100)

                End Using

                'グループごとの比較
                Call GetPriceData(ds_table, dt_100, intloop)
                ds_table = Nothing
                GC.Collect()
            Next
        Catch ex As Exception
            WriteLog(My.Settings.LogFolder & "YGY.txt", ex.Message & ex.StackTrace)
        End Try

    End Sub

    ''' <summary>
    ''' 閉じるボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        RaiseEvent BackToType()
    End Sub

    ''' <summary>
    ''' セッションクリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearSession()
        Me.Session.Remove("ManifoldKataban")
        Me.Session.Remove("ManifoldKatabanLoop")
        Me.Session.Remove("ManifoldSeriesKey")
        Me.Session.Remove("ManifoldItemKey")
        Me.Session.Remove("TestFlag")
        Me.Session.Remove("TestMode")
        Me.Session.Remove("ManifoldKataban")
        'TEST
        Me.Session.Remove("KtbnStrc")
        Me.Session.Remove("DS_Title")
        Me.Session.Remove("dt_Comb")
        Me.Session.Remove("KtbnStrc_Siyou")
    End Sub

    ''' <summary>
    ''' マニホールド逆展開(ISO)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnMFISO_Click(sender As Object, e As EventArgs) Handles btnMFISO.Click
        Dim dt_mfsiyou As New DS_100Test.MF_Siyou_ISODataTable
        Using da As New DS_100TestTableAdapters.MF_Siyou_ISOTableAdapter
            da.FillAll(dt_mfsiyou)
        End Using

        Dim intCount As Integer = dt_mfsiyou.Rows.Count
        Dim inti As Integer = My.MySettings.Default.ManifoldTestStart
        Session.Add("EventEndFlgISO", True)
        If intCount > 0 Then
            While inti < My.MySettings.Default.ManifoldTestEnd
                If Session("EventEndFlgISO").Equals(True) Then
                    ''TEST
                    'If Not dt_mfsiyou.Rows(inti)("仕様書№").ToString.Equals("#A277156") Then
                    '    inti += 1
                    '    Continue While
                    'End If

                    'If Not dt_mfsiyou.Rows(inti)("マニホールド形番").ToString.StartsWith("LMF0") Then
                    '    inti += 1
                    '    Continue While
                    'End If

                    Dim listKataban As New ManifoldKataban
                    Call ClearSession()
                    'マニホールドの価格計算
                    Session("EventEndFlg") = False
                    Me.Session("TestFlag") = Nothing    '
                    listKataban.KATABAN = dt_mfsiyou.Rows(inti)("マニホールド形番").ToString
                    listKataban.SIYOUSYO = dt_mfsiyou.Rows(inti)("仕様書№").ToString
                    Me.Session.Add("ManifoldKataban", listKataban)
                    Me.Session.Add("TestMode", 1)
                    RaiseEvent GoToType()
                    inti += 1
                    listKataban = Nothing
                End If
            End While
        End If
    End Sub

    ''' <summary>
    ''' マニホールド逆展開（通常）
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnMFSiyou_Click(sender As Object, e As EventArgs) Handles btnMFSiyou.Click

        If IsPostBack Then
            Call ClearSession()
            'CHANGED BY YGY 20140708 ↓↓↓↓↓↓
            Dim dt_mfsiyou As New DS_100Test.MF_SiyouDataTable
            Using da As New DS_100TestTableAdapters.MF_SiyouTableAdapter
                da.FillAll(dt_mfsiyou)
            End Using

            Dim intCount As Integer = dt_mfsiyou.Rows.Count
            Dim inti As Integer = My.MySettings.Default.ManifoldTestStart
            Dim intSum As Integer = 999999
            Session.Add("EventEndFlg", True)

            If intCount > 0 Then
                While inti <= My.MySettings.Default.ManifoldTestEnd
                    If Session("EventEndFlg") Then
                        Dim listKataban As New ManifoldKataban
                        'セッションのクリア
                        ClearSession()
                        'マニホールドの価格計算
                        Session("EventEndFlg") = False
                        Me.Session("TestFlag") = Nothing    '
                        listKataban.KATABAN = dt_mfsiyou.Rows(inti)("形番").ToString
                        listKataban.SIYOUSYO = dt_mfsiyou.Rows(inti)("仕様書№").ToString
                        listKataban.KATAPLACE = dt_mfsiyou.Rows(inti)("出荷場所").ToString
                        listKataban.KATACHECK = dt_mfsiyou.Rows(inti)("チェック区分").ToString()
                        listKataban.GSPRICE = dt_mfsiyou.Rows(inti)("ＧＳ店価格").ToString()
                        Me.Session.Add("ManifoldKataban", listKataban)
                        Me.Session.Add("TestMode", 0)
                        RaiseEvent GoToType()
                        inti += 1
                        listKataban = Nothing
                    End If
                End While
            End If
            'CHANGED BY YGY 20140708 ↑↑↑↑↑↑
        End If
    End Sub

    ''' <summary>
    ''' データバインド
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GVYouso_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVYouso.RowDataBound
        If e.Row.RowIndex < 0 Then Exit Sub
        Try
            Dim strName As String = Me.ClientID & "_"
            Dim intStartID As Integer = 0
            If e.Row.RowIndex = 0 Then
                intStartID = CInt(Strings.Right(e.Row.ClientID, 2))
            Else
                intStartID = CInt(Strings.Right(GVYouso.Rows(0).ClientID, 2))
            End If

            e.Row.Attributes.Add(CdCst.JavaScript.OnClick, "fncGridClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',0);")
            'e.Row.Attributes.Add(CdCst.JavaScript.OnDblClick, _
            '                     "MFHistorytest('" & btnMFTest.ClientID & "','" & _
            '                                         Me.HidSelKey.ClientID & "','" & _
            '                                         e.Row.Cells(2).Text & "," & e.Row.Cells(0).Text & "," & Me.txtUpdateUser.Text.Trim.ToUpper & "," & e.Row.Cells(3).Text & "');")
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 履歴取込テスト(MF_HISTORY)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnMFTest_Click(sender As Object, e As EventArgs) Handles btnMFTest.Click
        If Me.HidSelKey.Value.Length > 0 Then

            Dim listKataban As New ArrayList
            listKataban.Add(Me.HidSelKey.Value.ToString)

            Me.Session.Add("ManifoldKataban", listKataban)
            Me.Session.Add("TestMode", 2)

            RaiseEvent GoToType()
        End If
    End Sub

    ''' <summary>
    ''' 100万件テスト価格取得
    ''' </summary>
    ''' <param name="ds_table"></param>
    ''' <param name="dt100"></param>
    ''' <param name="intLoop"></param>
    ''' <remarks></remarks>
    Private Sub GetPriceData(ds_table As DataSet, dt100 As DataTable, intLoop As Integer)
        Dim strPath As String = My.Settings.LogFolder & "100Test_" & intLoop.ToString & ".txt"
        Dim strTitle As String = "形番" & ControlChars.Tab & _
                                 "チェック区分新" & ControlChars.Tab & "チェック区分旧" & ControlChars.Tab & _
                                 "出荷場所新" & ControlChars.Tab & "出荷場所旧" & ControlChars.Tab & _
                                 "LISTPRICE新" & ControlChars.Tab & "LISTPRICE旧" & ControlChars.Tab & _
                                 "REGISTPRICE新" & ControlChars.Tab & "REGISTPRICE旧" & ControlChars.Tab & _
                                 "SSPRICE新" & ControlChars.Tab & "SSPRICE旧" & ControlChars.Tab & _
                                 "BSPRICE新" & ControlChars.Tab & "BSPRICE旧" & ControlChars.Tab & _
                                 "GSPRICE新" & ControlChars.Tab & "GSPRICE旧" & ControlChars.Tab & _
                                 "PSPRICE新" & ControlChars.Tab & "PSPRICE旧" & ControlChars.Tab & _
                                 "簡易オーダー新" & ControlChars.Tab & "簡易オーダー旧" & ControlChars.NewLine
        'タイトルの出力
        File.AppendAllText(strPath, strTitle)

        '価格計算
        For Each dr100 As DataRow In dt100.Rows
            Dim strKataban As String = dr100.Item("KATABAN")
            Dim strSepResult As String = dr100("KATABAN").ToString & ControlChars.Tab

            '形番分解価格計算
            Select Case fncGetSeperateData(strKataban, ds_table)
                Case SeperateResult.SUCCESS
                    Dim strTmp As String = Compare100Test(dr100)

                    If strTmp.Equals(String.Empty) Then
                        strSepResult &= strMaru
                    Else
                        strSepResult &= strTmp
                    End If
                Case SeperateResult.PRICE_ERROR
                    strSepResult &= "分解失敗"
                Case SeperateResult.MANIFOLD_ERROR
                    strSepResult &= "マニホールド対象外"
            End Select

            File.AppendAllText(strPath, strSepResult & ControlChars.NewLine)
        Next

        RaiseEvent BackToType()
    End Sub

    ''' <summary>
    ''' 価格テスト
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnPriceTest_Click(sender As Object, e As EventArgs) Handles btnPriceTest.Click
        Try
            '比較データ
            Dim dtPriceTest As New DS_PriceTest.kh_price_testDataTable
            '処理結果
            Dim dtCompareResult As New DS_PriceTest.PriceTestResultDataTable
            '分解用情報
            Dim dsData As New DataSet

            '「kh_price_test」テーブルから価格データの取得
            Using da As New DS_PriceTestTableAdapters.kh_price_testTableAdapter
                da.Fill(dtPriceTest)
            End Using

            '形番分解用情報の取得
            dsData = fncLoadInfo(dtPriceTest)

            '形番ごとに分解して価格計算
            For Each drPriceTest As DS_PriceTest.kh_price_testRow In dtPriceTest
                Dim strKataban As String = drPriceTest.KATABAN

                'If Not strKataban.Equals("FH110-D-FP1") Then
                '    Continue For
                'End If

                If Not strKataban.Equals(String.Empty) Then

                    '形番分解価格計算
                    Select Case fncGetSeperateData(strKataban, dsData)
                        Case SeperateResult.SUCCESS

                            '分解成功
                            ComparePriceTest(drPriceTest, dtCompareResult)

                        Case SeperateResult.PRICE_ERROR
                            '価格取得失敗
                            Dim drComapreResult As DS_PriceTest.PriceTestResultRow

                            drComapreResult = dtCompareResult.NewPriceTestResultRow
                            drComapreResult.KATABAN = strKataban

                            If drPriceTest.SEPERATE_RESULT = SeperateResult.PRICE_ERROR Then
                                drComapreResult.COMPARE_RESULT = strMaru
                            Else
                                drComapreResult.SEPERATE_RESULT = "WEB版：" & SeperateResult.PRICE_ERROR & Space(4) & "NET版：" & drPriceTest.SEPERATE_RESULT
                                drComapreResult.COMPARE_RESULT = strBatsu
                            End If

                            dtCompareResult.Rows.Add(drComapreResult)
                        Case SeperateResult.MANIFOLD_ERROR
                            'マニホールド対象外
                            Dim drComapreResult As DS_PriceTest.PriceTestResultRow

                            drComapreResult = dtCompareResult.NewPriceTestResultRow
                            drComapreResult.KATABAN = strKataban

                            If drPriceTest.SEPERATE_RESULT = SeperateResult.MANIFOLD_ERROR Then
                                drComapreResult.COMPARE_RESULT = strMaru
                            Else
                                drComapreResult.SEPERATE_RESULT = "WEB版：" & SeperateResult.MANIFOLD_ERROR & Space(4) & "NET版：" & drPriceTest.SEPERATE_RESULT
                                drComapreResult.COMPARE_RESULT = strBatsu
                            End If

                            dtCompareResult.Rows.Add(drComapreResult)
                    End Select
                End If
            Next

            'ファイルの出力
            OutputPriceTestResult(dtCompareResult)

            'シリーズごとの件数を出力
            WriteCountPriceTest()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 仕様テスト
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnShiyouTest_Click(sender As Object, e As EventArgs) Handles btnShiyouTest.Click
        Dim inti As Integer = 0
        Dim dtShiyouTest As New DS_PriceTest.kh_shiyou_testDataTable
        '出力パス
        Dim strPath As String = My.Settings.LogFolder & "ShiyouTest_" & Now.ToString("yyyyMMdd") & ".txt"
        'タイトル
        Dim strTitle As String = "形番" & ControlChars.Tab & _
                                 "チェック区分" & ControlChars.Tab & _
                                 "出荷場所" & ControlChars.Tab & _
                                 "GS価格" & ControlChars.Tab & _
                                 "BS価格" & ControlChars.Tab & _
                                 "PS価格" & ControlChars.Tab & _
                                 "SS価格" & ControlChars.Tab & _
                                 "LS価格" & ControlChars.Tab & _
                                 "RG価格"

        '仕様テストデータの取得
        Using da As New DS_PriceTestTableAdapters.kh_shiyou_testTableAdapter
            da.Fill(dtShiyouTest)
        End Using

        '結果出力
        WriteLog(strPath, strTitle)

        '処理開始フラグ
        Session.Add("EventEndFlg", True)

        While inti < dtShiyouTest.Rows.Count
            Dim drShiyouTest As DataRow = dtShiyouTest.Rows(inti)

            'If Not drShiyouTest.Item("KATABAN").ToString.Equals("M4GB210R-CX-2-3") Then
            '    inti += 1
            '    Continue While
            'End If

            If Me.Session("EventEndFlg") Then
                Dim listKataban As New ArrayList

                'セッションクリア
                Call ClearSession()

                'マニホールドの価格計算用データの保存
                Me.Session("EventEndFlg") = False
                Me.Session("TestFlag") = Nothing
                Me.Session.Add("ManifoldKataban", drShiyouTest)
                Me.Session.Add("TestMode", 2)

                '機種選択画面へ遷移
                RaiseEvent GoToType()

                '次の履歴へ
                inti += 1
            End If
        End While

        'シリーズごとの件数を出力
        WriteCountShiyouTest()

    End Sub

#Region "メソッド"

    ''' <summary>
    ''' 形番分解用データの取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncLoadInfo(ByVal dtData As DataTable) As DataSet
        Dim dt_Option_All As New DS_100Test.GetOptionDataTable
        Dim dt_ElePattern_All As New DS_100Test.kh_ele_patternDataTable
        Dim dt_VolStd_All As New DS_100Test.GetVolDataTable
        Dim dt_Stroke_All As New DS_100Test.kh_strokeDataTable
        Dim dt_Series_All As New DS_KatSep.kh_series_katabanDataTable
        Dim dt_Hyphen_All As New DS_KatSep.kh_kataban_strcDataTable
        Dim dt_ItemName_All As New DS_KatSep.kh_ktbn_strc_nm_mstDataTable

        Dim dt_Option As New DS_100Test.GetOptionDataTable
        Dim dt_ElePattern As New DS_100Test.kh_ele_patternDataTable
        Dim dt_VolStd As New DS_100Test.GetVolDataTable
        Dim dt_Stroke As New DS_100Test.kh_strokeDataTable
        Dim dt_Series As New DS_KatSep.kh_series_katabanDataTable
        Dim dt_Hyphen As New DS_KatSep.kh_kataban_strcDataTable
        Dim dt_ItemName As New DS_KatSep.kh_ktbn_strc_nm_mstDataTable

        '取得結果
        Dim dsResult As New DataSet

        'データを取得する（グループ毎に）
        If dtData.Rows.Count > 0 Then 'データあれば

            Dim strKata As String = String.Empty
            Dim strLast As String = String.Empty

            For inti As Integer = 0 To dtData.Rows.Count - 1
                If strLast <> Strings.Left(dtData.Rows(inti)("KATABAN").ToString(), 1) Then
                    strKata = Strings.Left(dtData.Rows(inti)("KATABAN").ToString(), 1)
                    dt_Option = New DS_100Test.GetOptionDataTable
                    dt_ElePattern = New DS_100Test.kh_ele_patternDataTable
                    dt_VolStd = New DS_100Test.GetVolDataTable
                    dt_Stroke = New DS_100Test.kh_strokeDataTable

                    dt_Series = New DS_KatSep.kh_series_katabanDataTable
                    dt_Hyphen = New DS_KatSep.kh_kataban_strcDataTable
                    dt_ItemName = New DS_KatSep.kh_ktbn_strc_nm_mstDataTable

                    Using da_1 As New DS_KatSepTableAdapters.kh_series_katabanTableAdapter
                        da_1.FillBySeries(dt_Series, strKata & "%")
                        If dt_Series.Rows.Count > 0 Then dt_Series_All.Merge(dt_Series)
                    End Using

                    Using da_1 As New DS_KatSepTableAdapters.kh_ktbn_strc_nm_mstTableAdapter
                        da_1.FillBySeries(dt_ItemName, strKata & "%")
                        If dt_ItemName.Rows.Count > 0 Then dt_ItemName_All.Merge(dt_ItemName)
                    End Using

                    Using da_1 As New DS_KatSepTableAdapters.kh_kataban_strcTableAdapter
                        da_1.FillBySeries(dt_Hyphen, strKata & "%")
                        If dt_Hyphen.Rows.Count > 0 Then dt_Hyphen_All.Merge(dt_Hyphen)
                    End Using

                    Using da_1 As New DS_100TestTableAdapters.GetOptionTableAdapter
                        da_1.FillbySeries(dt_Option, Now, "en", "ja", strKata & "%")
                        If dt_Option.Rows.Count > 0 Then dt_Option_All.Merge(dt_Option)
                    End Using

                    Using da_1 As New DS_100TestTableAdapters.kh_ele_patternTableAdapter
                        da_1.FillbySeries(dt_ElePattern, Now, strKata & "%")
                        If dt_ElePattern.Rows.Count > 0 Then dt_ElePattern_All.Merge(dt_ElePattern)
                    End Using

                    Using da_1 As New DS_100TestTableAdapters.GetVolTableAdapter
                        da_1.FillbySeries(dt_VolStd, Now, strKata & "%")
                        If dt_VolStd.Rows.Count > 0 Then dt_VolStd_All.Merge(dt_VolStd)
                    End Using

                    Using da_1 As New DS_100TestTableAdapters.kh_strokeTableAdapter
                        da_1.FillbySeries(dt_Stroke, Now, strKata & "%")
                        If dt_Stroke.Rows.Count > 0 Then dt_Stroke_All.Merge(dt_Stroke)
                    End Using

                    strLast = Strings.Left(dtData.Rows(inti)("KATABAN").ToString(), 1)
                End If
            Next
        End If

        dt_Series_All.TableName = "dt_Series"
        dsResult.Tables.Add(dt_Series_All)
        dt_ItemName_All.TableName = "dt_ItemName"
        dsResult.Tables.Add(dt_ItemName_All)
        dt_Hyphen_All.TableName = "dt_Hyphen"
        dsResult.Tables.Add(dt_Hyphen_All)
        dt_Option_All.TableName = "dt_Option"
        dsResult.Tables.Add(dt_Option_All)
        dt_ElePattern_All.TableName = "dt_ElePattern"
        dsResult.Tables.Add(dt_ElePattern_All)
        dt_VolStd_All.TableName = "dt_VolStd"
        dsResult.Tables.Add(dt_VolStd_All)
        dt_Stroke_All.TableName = "dt_Stroke"
        dsResult.Tables.Add(dt_Stroke_All)

        Return dsResult
    End Function

    ''' <summary>
    ''' 指定した日付の履歴を利用してテスト
    ''' </summary>
    ''' <param name="strDate"></param>
    ''' <remarks></remarks>
    Private Sub subTestHistoryAll(ByVal strDate As String)
        Dim inti As Integer = 0
        Dim dt_History As New DS_History.MF_HistoryDataTable
        Dim dtTest As New DataTable

        '履歴データの取得
        Using da As New DS_HistoryTableAdapters.MF_HistoryTableAdapter
            da.FillByDate(dt_History, CDate(strDate), CDate(strDate).AddDays(1))
        End Using

        'テストデータの作成
        dtTest = fncGetTestData(dt_History)

        '処理開始フラグ
        Session.Add("EventEndFlg", True)

        While inti < dtTest.Rows.Count
            Dim drHistory As DataRow = dtTest.Rows(inti)

            If drHistory.Item("UpdateComputer").Equals("******    ") Then
                'NET版の履歴だけをテストする
                'If Not drHistory.Item("Kataban").Equals("M3GA180R-CX-T51RW1MHKASFZ1Z3D-12-4-P74") Or
                '    drHistory.Item("UpdateComputer").Equals("******    ") Then
                inti += 1
                Continue While
            Else
                If Me.Session("EventEndFlg") Then
                    Dim listKataban As New ArrayList

                    'セッションクリア
                    Call ClearSession()

                    'マニホールドの価格計算用データの保存
                    listKataban.Add(drHistory.Item("Kataban") & "," & CDate(drHistory.Item("UpdateDate")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "," & drHistory.Item("UpdateUser") & "," & drHistory.Item("GSPrice"))

                    Me.Session("EventEndFlg") = False
                    Me.Session("TestFlag") = Nothing
                    Me.Session.Add("ManifoldKataban", listKataban)
                    Me.Session.Add("TestMode", 2)

                    '機種選択画面へ遷移
                    RaiseEvent GoToType()

                    '次の履歴へ
                    inti += 1
                End If
            End If
        End While
        'End If
    End Sub

    ''' <summary>
    ''' テストデータテーブルの作成
    ''' </summary>
    ''' <param name="dtHistory"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetTestData(ByVal dtHistory As DataTable) As DataTable
        Dim dtResult As New DataTable

        'dtResult = dtHistory.Copy

        'Select Case rblDataType.SelectedValue
        '    Case "NET"
        '        dtResult.DefaultView.RowFilter = "UpdateComputer<>'******    '"
        '    Case "WEB"
        '        dtResult.DefaultView.RowFilter = "UpdateComputer='******    '"
        '    Case Else
        'End Select

        Return dtResult.DefaultView.ToTable
    End Function

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
    ''' 形番分解結果の作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetSeperateData(ByVal strKataban As String, ByVal dsData As DataSet) As SeperateResult
        Dim result As New SeperateResult

        objKtbnStrc = New KHKtbnStrc
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


        '形番分解
        If KHKatabanSeparator.GetSeparatorData(strKataban.Trim.ToUpper, _
                                               strSeries, _
                                               strKeyKata, _
                                               strKataName, _
                                               strSpecNo, _
                                               strPriceNo, _
                                               strItem1, _
                                               strItemName1, _
                                               strHyphen1, _
                                               strStructure_div, _
                                               strElement_div1, _
                                               dsData) Then

            '通貨の設定
            SetCurrency(strSeries, strKeyKata, strKataName)

            '引当形番構成の追加
            InsertKtbnStrc(strSeries, strKeyKata)

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
            '形番分解失敗
        End If

        '分解失敗しても単価を計算する、特注形番の可能性がある
        If IsManifold(strSeries, strSpecNo) Then
            'マニホールド対象形番、価格を計算できません。
            result = SeperateResult.MANIFOLD_ERROR
        Else
            Dim objUnitPrice As New KHUnitPrice
            Dim objOption As New KHOptionCtl

            objKtbnStrc.strcSelection.strFullKataban = strKataban

            '価格の取得
            Call objUnitPrice.subPriceInfoSet_ForkatOut(objCon, objKtbnStrc, Me.objUserInfo.CountryCd, "")

            '原価積算No取得
            objKtbnStrc.strcSelection.strCostCalcNo = objOption.fncCostCalcNoGet(objKtbnStrc, objKtbnStrc.strcSelection.strKatabanCheckDiv)

            '①新しい形番ﾁｪｯｸ区分を反映する
            If KHKataban.subJapanChinaAmount(strKataban) Then
                objKtbnStrc.strcSelection.strKatabanCheckDiv = "1"
            End If

            '出荷場所の変換
            objKtbnStrc.strcSelection.strPlaceCd = changeShipPlace(objKtbnStrc)

            '処理結果
            If objKtbnStrc.strcSelection.intGsPrice = 0 Then
                result = SeperateResult.PRICE_ERROR
            Else
                result = SeperateResult.SUCCESS
            End If

        End If

        Return result
    End Function

    ''' <summary>
    ''' 通貨の設定
    ''' </summary>
    ''' <param name="strSeries"></param>
    ''' <param name="strKeyKata"></param>
    ''' <param name="strKataName"></param>
    ''' <remarks></remarks>
    Private Sub SetCurrency(ByVal strSeries As String, ByVal strKeyKata As String, ByVal strKataName As String)
        Dim bllType As New TypeBLL

        '通貨がない場合はJPYに設定する
        If objKtbnStrc.strcSelection.strCurrency Is Nothing Then
            objKtbnStrc.strcSelection.strCurrency = "JPY"
        End If

        '引当シリーズ形番追加(機種)
        '通貨の追加
        Call bllType.subInsertSelSrsKtbnMdl(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
            strSeries, strKeyKata, strKataName, objKtbnStrc.strcSelection.strCurrency)
    End Sub

    ''' <summary>
    ''' 引当形番構成の追加
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InsertKtbnStrc(ByVal strSeries As String, ByVal strKeyKata As String)
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

    End Sub

    ''' <summary>
    ''' マニホールドかどうか
    ''' </summary>
    ''' <param name="strSeries"></param>
    ''' <param name="strSpecNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsManifold(ByVal strSeries As String, ByVal strSpecNo As String) As Boolean
        Dim blnResult As Boolean = False

        If Len(strSpecNo.Trim) <> 0 Then
            Select Case strSpecNo.Trim
                Case "00"
                    'ページ遷移(ロッド先端形状オーダーメイド寸法入力画面)
                Case "01", "02", "03", "04", "05", "06", "07", "08", "10", "11", _
                     "13", "14", "15", "16", "96"
                    blnResult = True
                Case "09"
                    If objKtbnStrc.strcSelection.strOpSymbol(6).ToString.Trim <> "" Then
                        blnResult = True
                    End If
                Case "12", "18", "19", "20", "21", "22", "23", "94", "95"
                    If KHKatabanSeparator.fncMixCheck(strSeries, objKtbnStrc.strcSelection.strOpSymbol) Then
                        blnResult = True
                    End If
                Case "17"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Then
                        blnResult = True
                    End If
                Case "51"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                        blnResult = True
                    End If
                Case "52", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", _
                     "65", "66", "67", "68", "69", "70", "71", "72", "89", "90", "91", "92", "98"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        blnResult = True
                    End If
                Case "53", "73", "74", "75", "76", "77", "78", "79", "80", "81", _
                     "82", "83", "84", "85", "86", "87", "88", "93"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "80" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "81" Then
                        blnResult = True
                    End If
                Case "A1", "A2", "A9", "B1", "B2", "B3", "B4"
                    blnResult = True
            End Select
        End If

        Return blnResult
    End Function

    ''' <summary>
    ''' 価格テスト処理結果の比較
    ''' </summary>
    ''' <param name="drPriceTest"></param>
    ''' <param name="dtCompareResult"></param>
    ''' <remarks></remarks>
    Private Sub ComparePriceTest(ByVal drPriceTest As DS_PriceTest.kh_price_testRow, _
                             ByRef dtCompareResult As DS_PriceTest.PriceTestResultDataTable)

        Dim drResult As DS_PriceTest.PriceTestResultRow
        Dim blnCompareResult As Boolean = True

        drResult = dtCompareResult.NewPriceTestResultRow
        drResult.KATABAN = drPriceTest.KATABAN

        With objKtbnStrc.strcSelection

            'ﾁｪｯｸ区分の比較
            If drPriceTest.IsCHECKKBNNull Then
                drPriceTest.CHECKKBN = String.Empty
            End If
            If .strKatabanCheckDiv.ToString = drPriceTest.CHECKKBN.ToString Then
                drResult.CHECKKBN = strMaru
            Else
                drResult.CHECKKBN = "WEB版：" & .strKatabanCheckDiv & Space(4) & "NET版：" & drPriceTest.CHECKKBN.ToString
                blnCompareResult = False
            End If

            '出荷場所の比較
            If drPriceTest.IsSHIPPLACENull Then
                drPriceTest.SHIPPLACE = String.Empty
            End If
            If .strPlaceCd.ToString = drPriceTest.SHIPPLACE.ToString Then
                drResult.SHIPPLACE = strMaru
            Else
                drResult.SHIPPLACE = "WEB版：" & .strPlaceCd & Space(4) & "NET版：" & drPriceTest.SHIPPLACE.ToString
                blnCompareResult = False
            End If

            'GS価格の比較
            If drPriceTest.IsGSPRICENull Then
                drPriceTest.GSPRICE = 0
            End If
            If .intGsPrice = drPriceTest.GSPRICE Then
                drResult.GSPRICE = strMaru
            Else
                drResult.GSPRICE = "WEB版：" & .intGsPrice & Space(4) & "NET版：" & drPriceTest.GSPRICE
                blnCompareResult = False
            End If

            'BS価格の比較
            If drPriceTest.IsBSPRICENull Then
                drPriceTest.BSPRICE = 0
            End If
            If .intBsPrice = drPriceTest.BSPRICE Then
                drResult.BSPRICE = strMaru
            Else
                drResult.BSPRICE = "WEB版：" & .intBsPrice & Space(4) & "NET版：" & drPriceTest.BSPRICE
                blnCompareResult = False
            End If

            'PS価格の比較
            If drPriceTest.IsPSPRICENull Then
                drPriceTest.PSPRICE = 0
            End If
            If .intPsPrice = drPriceTest.PSPRICE Then
                drResult.PSPRICE = strMaru
            Else
                drResult.PSPRICE = "WEB版：" & .intPsPrice & Space(4) & "NET版：" & drPriceTest.PSPRICE
                blnCompareResult = False
            End If

            'SS価格の比較
            If drPriceTest.IsSSPRICENull Then
                drPriceTest.SSPRICE = 0
            End If
            If .intSsPrice = drPriceTest.SSPRICE Then
                drResult.SSPRICE = strMaru
            Else
                drResult.SSPRICE = "WEB版：" & .intSsPrice & Space(4) & "NET版：" & drPriceTest.SSPRICE
                blnCompareResult = False
            End If

            'LS価格の比較
            If drPriceTest.IsPSPRICENull Then
                drPriceTest.PSPRICE = 0
            End If
            If .intListPrice = drPriceTest.LSPRICE Then
                drResult.LSPRICE = strMaru
            Else
                drResult.LSPRICE = "WEB版：" & .intListPrice & Space(4) & "NET版：" & drPriceTest.LSPRICE
                blnCompareResult = False
            End If

            'RG価格の比較
            If drPriceTest.IsRGPRICENull Then
                drPriceTest.RGPRICE = 0
            End If
            If .intRegPrice = drPriceTest.RGPRICE Then
                drResult.RGPRICE = strMaru
            Else
                drResult.RGPRICE = "WEB版：" & .intRegPrice & Space(4) & "NET版：" & drPriceTest.RGPRICE
                blnCompareResult = False
            End If

        End With

        If blnCompareResult Then
            drResult.COMPARE_RESULT = strMaru
        Else
            drResult.COMPARE_RESULT = strBatsu
        End If

        dtCompareResult.Rows.Add(drResult)
    End Sub

    ''' <summary>
    ''' 100万件テスト処理結果の比較
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Compare100Test(ByVal dr_100 As DataRow) As String
        Dim strSepResult As String = String.Empty

        With objKtbnStrc.strcSelection

            'チェック区分の比較
            If dr_100("CHECKKBN") <> .strKatabanCheckDiv Then
                strSepResult &= .strKatabanCheckDiv & ControlChars.Tab & _
                            dr_100("CHECKKBN").ToString & ControlChars.Tab
            End If

            '出荷場所の比較
            '変換必要があるかどうかの判断(P11,P21)
            .strPlaceCd = changeShipPlace(objKtbnStrc)

            If dr_100("SHIPPLACE") <> .strPlaceCd Then
                strSepResult &= .strPlaceCd & ControlChars.Tab & _
                            dr_100("SHIPPLACE").ToString & ControlChars.Tab
            End If

            'LISTPRICEの比較
            If dr_100("LISTPRICE") <> CInt(.intListPrice) Then
                strSepResult &= .intListPrice & ControlChars.Tab & _
                            dr_100("LISTPRICE").ToString & ControlChars.Tab
            End If

            'REGISTPRICEの比較
            If dr_100("REGISTPRICE") <> CInt(.intRegPrice) Then
                strSepResult &= .intRegPrice & ControlChars.Tab & _
                            dr_100("REGISTPRICE").ToString & ControlChars.Tab
            End If

            'SSPRICEの比較
            If dr_100("SSPRICE") <> CInt(.intSsPrice) Then
                strSepResult &= .intSsPrice & ControlChars.Tab & _
                            dr_100("SSPRICE").ToString & ControlChars.Tab
            End If

            'BSPRICEの比較
            If dr_100("BSPRICE") <> CInt(.intBsPrice) Then
                strSepResult &= .intBsPrice & ControlChars.Tab & _
                            dr_100("BSPRICE").ToString & ControlChars.Tab
            End If

            'GSPRICEの比較
            If dr_100("GSPRICE") <> CInt(.intGsPrice) Then
                strSepResult &= .intGsPrice & ControlChars.Tab & _
                            dr_100("GSPRICE").ToString & ControlChars.Tab
            End If

            'PSPRICEの比較
            If dr_100("PSPRICE") <> CInt(.intPsPrice) Then
                strSepResult &= .intPsPrice & ControlChars.Tab & _
                            dr_100("PSPRICE").ToString & ControlChars.Tab
            End If

            '簡易オーダーフラグの比較
            Dim strDBSimpleFlg As String = String.Empty
            'DBデータを変換
            strDBSimpleFlg = IIf(dr_100("SIMPLEORDERFLG") = "" OrElse dr_100("SIMPLEORDERFLG") = "0", "", "C5")

            If strDBSimpleFlg <> .strCostCalcNo Then
                strSepResult &= .strCostCalcNo & ControlChars.Tab & _
                            dr_100("SIMPLEORDERFLG").ToString & ControlChars.NewLine
            End If
        End With

        Return strSepResult
    End Function

    ''' <summary>
    ''' 結果の出力
    ''' </summary>
    ''' <param name="dtCompareResult"></param>
    ''' <remarks></remarks>
    Private Sub OutputPriceTestResult(ByVal dtCompareResult As DS_PriceTest.PriceTestResultDataTable)
        Dim strOutputPath As String = My.Settings.LogFolder & "PriceTest_" & Now.ToString("yyyyMMdd") & ".txt"
        Dim strResult As New StringBuilder

        '全体的な結果を出力
        Dim intCount As Integer = dtCompareResult.Rows.Count
        Dim intBatsu As Integer = dtCompareResult.Select("COMPARE_RESULT = '" & strBatsu & "'").Count
        Dim strDifference As String = String.Empty
        Dim strTitle As String = String.Empty

        strDifference = "差異" & intBatsu & "件/" & intCount & "件"
        strTitle = "形番" & ControlChars.Tab & "ﾁｪｯｸ区分" & ControlChars.Tab & "プラント" & ControlChars.Tab & _
                   "GS価格" & ControlChars.Tab & "BS価格" & ControlChars.Tab & "PS価格" & ControlChars.Tab & "SS価格" & ControlChars.Tab & "LS価格" & ControlChars.Tab & "RG価格"

        strResult.AppendLine(strDifference)
        strResult.AppendLine(strTitle)

        '結果が違うデータを出力
        For Each dr As DS_PriceTest.PriceTestResultRow In dtCompareResult
            If dr.COMPARE_RESULT.Equals(strBatsu) Then
                Dim strRowResult As String = String.Empty

                strRowResult = dr.Item("KATABAN") & ControlChars.Tab & _
                               dr.Item("CHECKKBN").ToString & ControlChars.Tab & _
                               dr.Item("SHIPPLACE").ToString & ControlChars.Tab & _
                               dr.Item("GSPRICE").ToString & ControlChars.Tab & _
                               dr.Item("BSPRICE").ToString & ControlChars.Tab & _
                               dr.Item("PSPRICE").ToString & ControlChars.Tab & _
                               dr.Item("SSPRICE").ToString & ControlChars.Tab & _
                               dr.Item("LSPRICE").ToString & ControlChars.Tab & _
                               dr.Item("RGPRICE").ToString

                strResult.AppendLine(strRowResult)
            End If
        Next

        '結果出力
        File.AppendAllText(strOutputPath, strResult.ToString)
    End Sub

    ''' <summary>
    ''' 処理したシリーズごとの件数を出力
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WriteCountPriceTest()
        Dim strOutputPath As String = My.Settings.LogFolder & "PriceTest_" & Now.ToString("yyyyMMdd") & ".txt"
        Dim strResult As New StringBuilder
        Dim strTitle As String = "実施対象形番" & ControlChars.NewLine & "第一ハイフン" & ControlChars.Tab & "件数"
        Dim dtCounter As New DS_PriceTest.KatabanCounterPriceTestDataTable

        'タイトルの出力
        strResult.AppendLine(strTitle)

        '件数の出力
        Using da As New DS_PriceTestTableAdapters.KatabanCounterPriceTestTableAdapter
            da.Fill(dtCounter)
        End Using

        For Each dr As DS_PriceTest.KatabanCounterPriceTestRow In dtCounter
            strResult.AppendLine(dr.第一ハイフン & ControlChars.Tab & dr.件数)
        Next

        '結果出力
        File.AppendAllText(strOutputPath, strResult.ToString)
    End Sub

    ''' <summary>
    ''' 処理したシリーズごとの件数を出力
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WriteCountShiyouTest()
        Dim strOutputPath As String = My.Settings.LogFolder & "ShiyouTest_" & Now.ToString("yyyyMMdd") & ".txt"
        Dim strResult As New StringBuilder
        Dim strTitle As String = "実施対象形番" & ControlChars.NewLine & "第一ハイフン" & ControlChars.Tab & "件数"
        Dim dtCounter As New DS_PriceTest.KatabanCounterShiyouTestDataTable

        'タイトルの出力
        strResult.AppendLine(strTitle)

        '件数の出力
        Using da As New DS_PriceTestTableAdapters.KatabanCounterShiyouTestTableAdapter
            da.Fill(dtCounter)
        End Using

        For Each dr As DS_PriceTest.KatabanCounterShiyouTestRow In dtCounter
            strResult.AppendLine(dr.第一ハイフン & ControlChars.Tab & dr.件数)
        Next

        '結果出力
        File.AppendAllText(strOutputPath, strResult.ToString)
    End Sub
#End Region
End Class