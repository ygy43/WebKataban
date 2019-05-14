Imports WebKataban.ClsCommon

Public Class M_KatabanDAL

    'Public Shared CommonDbService As CommonDbService.CommonDbServiceClient
    Public Shared dtData As New DataTable

    ''' -----------------------------------------------------------------------------------------
    ''' <summary>
    ''' 条件により形番マスタデータ取得(形引)
    ''' </summary>
    ''' <param name="userKatabanBeginWith">呼出形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' -----------------------------------------------------------------------------------------
    Public Shared Function LoadCommonDB_GetMKatabanByConditionsKatahiki(ByVal userKatabanBeginWith As String, _
                                                                       ByVal seriesSearch As KHSeriesSearch, ByRef dsData As DataSet) As DataSet
        LoadCommonDB_GetMKatabanByConditionsKatahiki = Nothing

        'Try
        '    Dim SalesGroupCodeList As New List(Of String)
        '    Dim soldToPartyCode As String
        '    Dim listData As New List(Of CommonDbService.M_Kataban)
        '    Dim count As Integer = 0

        '    '仕入品コード設定
        '    SalesGroupCodeList.Add("999")
        '    soldToPartyCode = "9999999999"

        '    '初期化
        '    LoadCommonDB_GetMKatabanByConditionsKatahiki = Nothing
        '    dtData.Clear()

        '    If dsData Is Nothing Then
        '        '仕入品検索の場合
        '        'テーブル追加
        '        LoadCommonDB_GetMKatabanByConditionsKatahiki = New DataSet
        '        dtData = LoadCommonDB_GetMKatabanByConditionsKatahiki.Tables.Add("KatabanTbl")

        '        '項目追加
        '        dtData.Columns.Add("sortKey", Type.GetType("System.String"))
        '        dtData.Columns.Add("SalesGroupCode", Type.GetType("System.String"))
        '        dtData.Columns.Add("SoldToPartyCode", Type.GetType("System.String"))
        '        dtData.Columns.Add("series_kataban", Type.GetType("System.String"))
        '        dtData.Columns.Add("key_kataban", Type.GetType("System.String"))
        '        dtData.Columns.Add("disp_kataban", Type.GetType("System.String"))
        '        dtData.Columns.Add("division", Type.GetType("System.String"))
        '        dtData.Columns.Add("disp_name", Type.GetType("System.String"))
        '        dtData.Columns.Add("currency_cd", Type.GetType("System.String"))
        '        dtData.Columns.Add("userkataban", Type.GetType("System.String"))

        '    Else
        '        '全て検索の場合
        '        '検索結果を追加する
        '        LoadCommonDB_GetMKatabanByConditionsKatahiki = dsData
        '        dtData = dsData.Tables(0)
        '    End If

        '    If soldToPartyCode.Length > 0 Or userKatabanBeginWith.Length > 0 Then
        '        '読み込み
        '        CommonDbService = New CommonDbService.CommonDbServiceClient
        '        count = CommonDbService.GetMKatabanCountByConditionsKatahiki(SalesGroupCodeList, soldToPartyCode, "JPY", userKatabanBeginWith)

        '        If count > 2000 Then
        '            '取得結果が2000件を超える場合はエラーを返す
        '            seriesSearch.strResultTypeCdValue = KHSeriesSearch.ResultType.MaxCountOver
        '        ElseIf count > 0 Then
        '            listData = CommonDbService.GetMKatabanByConditionsKatahiki(SalesGroupCodeList, soldToPartyCode, "JPY", userKatabanBeginWith)

        '            'Listデータをテーブルに変換
        '            listData.ForEach(AddressOf ConvertToDataTable)
        '        End If

        '        'いる？
        '        seriesSearch.strHeaderValue = "シリーズ形番/ＣＫＤ形番"

        '    End If

        '    Return LoadCommonDB_GetMKatabanByConditionsKatahiki

        'Catch ex As Exception
        '    WriteErrorLog("E001", ex)
        'Finally
        '    CommonDbService.Close()
        'End Try

    End Function

    ' ''' -----------------------------------------------------------------------------------------
    ' ''' <summary>
    ' ''' 形番マスタの取得
    ' ''' </summary>
    ' ''' <param name="salesGroupCode">販売グループコード</param>
    ' ''' <param name="soldToPartyCode">得意先コード</param>
    ' ''' <param name="userKataban">ユーザー形番</param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    ' ''' -----------------------------------------------------------------------------------------
    'Public Shared Function LoadCommonDB_GetMKataban(ByVal salesGroupCode As String, ByVal soldToPartyCode As String, ByVal userKataban As String) As CommonDbService.M_Kataban
    '    LoadCommonDB_GetMKataban = Nothing
    'Try
    '    '読み込み
    '    CommonDbService = New CommonDbService.CommonDbServiceClient
    '    LoadCommonDB_GetMKataban = CommonDbService.GetMKataban(salesGroupCode, soldToPartyCode, userKataban, "JPY")

    '    Return LoadCommonDB_GetMKataban

    'Catch ex As Exception
    '    WriteErrorLog("E001", ex)
    'Finally
    '    CommonDbService.Close()
    'End Try
    'End Function

    ''' -----------------------------------------------------------------------------------------
    ''' <summary>
    ''' 単価情報取得処理
    ''' </summary>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strKatabanCheckDiv">形番チェック区分</param>
    ''' <param name="strPlaceCd">プラント</param>
    ''' <param name="htPriceInfo">価格情報</param>
    ''' <param name="strMadeCountry">生産国コード</param>
    ''' <returns></returns>
    ''' <remarks>単価テーブルを読み込み単価情報を取得し返却する</remarks>
    ''' -----------------------------------------------------------------------------------------
    Public Shared Function fncSelectPrice(ByRef strKataban As String, ByRef strKatabanCheckDiv As String, ByRef strPlaceCd As String, _
                                  ByRef htPriceInfo As Hashtable, ByRef strMadeCountry As String, _
                                  ByRef strStorageLocation As String, ByRef strEvaluationType As String) As Boolean

        fncSelectPrice = False
        htPriceInfo = New Hashtable

        Try
            'Dim dt As CommonDbService.M_Kataban
            Dim salesGroupCode As String = "999"
            Dim soldToPartyCode As String = "9999999999"

            '読み込み
            ' CommonDbService = New CommonDbService.CommonDbServiceClient
            ' dt = CommonDbService.GetMKataban(salesGroupCode, soldToPartyCode, strKataban, "JPY")

            'If dt Is Nothing Then
            '単価が取得出来ない場合は0を設定する
            'strKatabanCheckDiv = String.Empty
            'strPlaceCd = String.Empty
            'htPriceInfo(CdCst.UnitPrice.ListPrice) = 0
            'htPriceInfo(CdCst.UnitPrice.RegPrice) = 0
            'htPriceInfo(CdCst.UnitPrice.SsPrice) = 0
            'htPriceInfo(CdCst.UnitPrice.BsPrice) = 0
            'htPriceInfo(CdCst.UnitPrice.GsPrice) = 0
            'htPriceInfo(CdCst.UnitPrice.PsPrice) = 0
            'strMadeCountry = String.Empty
            'strStorageLocation = String.Empty
            'strEvaluationType = String.Empty
            'fncSelectPrice = False

            'Else
            'strKatabanCheckDiv = Right(dt.CheckKubun, 1)
            'strPlaceCd = dt.DeliveryPlant
            'htPriceInfo(CdCst.UnitPrice.ListPrice) = dt.ListPrice
            'htPriceInfo(CdCst.UnitPrice.RegPrice) = dt.RegistPrice
            'htPriceInfo(CdCst.UnitPrice.SsPrice) = dt.SsPrice
            'htPriceInfo(CdCst.UnitPrice.BsPrice) = dt.BsPrice
            'htPriceInfo(CdCst.UnitPrice.GsPrice) = dt.GsPrice
            'htPriceInfo(CdCst.UnitPrice.PsPrice) = dt.PsPrice
            'strMadeCountry = "JPN"
            'strStorageLocation = IIf(IsNothing(dt.StorageLocation), String.Empty, dt.StorageLocation)
            'strEvaluationType = IIf(IsNothing(dt.EvaluationType), String.Empty, dt.EvaluationType)
            'fncSelectPrice = True
            ' End If

            Return fncSelectPrice

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'CommonDbService.Close()
        End Try

    End Function

    ''' -----------------------------------------------------------------------------------------
    ''' <summary>
    ''' ListからDataSetに変換
    ''' </summary>
    ''' -----------------------------------------------------------------------------------------
    'Public Shared Function ConvertToDataTable(ByVal list As CommonDbService.M_Kataban)
    '    Dim dr As DataRow

    '    dr = dtData.NewRow
    '    dr("sortKey") = list.UserKataban
    '    dr("series_kataban") = list.UserKataban
    '    dr("key_kataban") = String.Empty
    '    dr("disp_kataban") = list.Kataban
    '    dr("division") = "3"
    '    dr("disp_name") = list.Memo
    '    dr("currency_cd") = list.Currency
    '    dr("UserKataban") = list.UserKataban

    '    dtData.Rows.Add(dr)

    '    Return dtData

    'End Function

End Class
