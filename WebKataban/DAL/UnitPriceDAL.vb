Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class UnitPriceDAL

    ''' <summary>
    ''' 単価情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strKatabanCheckDiv">形番チェック区分</param>
    ''' <param name="strPlaceCd">出荷場所</param>
    ''' <param name="htPriceInfo">価格情報</param>
    ''' <param name="strCurrency"></param>
    ''' <param name="strMadeCountry"></param>
    ''' <returns></returns>
    ''' <remarks>単価テーブルを読み込み単価情報を取得し返却する</remarks>
    Public Function fncSelectPrice(objCon As SqlConnection, ByRef strKataban As String, _
                                  ByRef strKatabanCheckDiv As String, ByRef strPlaceCd As String, _
                                  ByRef htPriceInfo As Hashtable, _
                                  ByVal strCurrency As String, ByRef strMadeCountry As String) As Boolean
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        htPriceInfo = New Hashtable
        fncSelectPrice = False

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  kataban, ")
            sbSql.Append("         kataban_check_div, ")
            sbSql.Append("         place_cd, ")
            sbSql.Append("         ls_price, ")
            sbSql.Append("         rg_price, ")
            sbSql.Append("         ss_price, ")
            sbSql.Append("         bs_price, ")
            sbSql.Append("         gs_price, ")
            sbSql.Append("         ps_price, currency_cd, ")
            sbSql.Append("         country_cd ")
            sbSql.Append(" FROM    kh_price ")
            sbSql.Append(" WHERE   kataban             = @Kataban ")
            'ADD BY YGY 20150311    海外仕入れ追加のため    ↓↓↓↓↓↓
            sbSql.Append(" AND     currency_cd         = @CurrencyCd")
            'ADD BY YGY 20150311    海外仕入れ追加のため    ↑↑↑↑↑↑
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
                'ADD BY YGY 20150311    海外仕入れ追加のため
                .Parameters.Add("@CurrencyCd", SqlDbType.VarChar, 3).Value = strCurrency
            End With

            objRdr = objCmd.ExecuteReader

            If objRdr.HasRows = True Then
                objRdr.Read()
                strKatabanCheckDiv = objRdr.GetValue(objRdr.GetOrdinal("kataban_check_div"))
                strPlaceCd = objRdr.GetValue(objRdr.GetOrdinal("place_cd"))
                htPriceInfo(CdCst.UnitPrice.ListPrice) = objRdr.GetValue(objRdr.GetOrdinal("ls_price"))
                htPriceInfo(CdCst.UnitPrice.RegPrice) = objRdr.GetValue(objRdr.GetOrdinal("rg_price"))
                htPriceInfo(CdCst.UnitPrice.SsPrice) = objRdr.GetValue(objRdr.GetOrdinal("ss_price"))
                htPriceInfo(CdCst.UnitPrice.BsPrice) = objRdr.GetValue(objRdr.GetOrdinal("bs_price"))
                htPriceInfo(CdCst.UnitPrice.GsPrice) = objRdr.GetValue(objRdr.GetOrdinal("gs_price"))
                htPriceInfo(CdCst.UnitPrice.PsPrice) = objRdr.GetValue(objRdr.GetOrdinal("ps_price"))
                'strCurrency = objRdr.GetValue(objRdr.GetOrdinal("currency_cd"))
                strMadeCountry = objRdr.GetValue(objRdr.GetOrdinal("country_cd"))
                fncSelectPrice = True
            Else
                '単価が取得出来ない場合は0を設定する
                strKatabanCheckDiv = ""
                strPlaceCd = ""
                htPriceInfo(CdCst.UnitPrice.ListPrice) = 0
                htPriceInfo(CdCst.UnitPrice.RegPrice) = 0
                htPriceInfo(CdCst.UnitPrice.SsPrice) = 0
                htPriceInfo(CdCst.UnitPrice.BsPrice) = 0
                htPriceInfo(CdCst.UnitPrice.GsPrice) = 0
                htPriceInfo(CdCst.UnitPrice.PsPrice) = 0
                'strCurrency = String.Empty
                strMadeCountry = String.Empty
                fncSelectPrice = False
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 積上単価情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strKatabanCheckDiv">形番チェック区分</param>
    ''' <param name="strPlaceCd">出荷場所</param>
    ''' <param name="htPriceInfo">価格情報</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectAccumulatePrice(objCon As SqlConnection, ByVal strKataban As String, _
                                             ByRef strKatabanCheckDiv As String, ByRef strPlaceCd As String, _
                                             ByRef htPriceInfo As Hashtable, _
                                             ByVal strCurrency As String) As Boolean
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        htPriceInfo = New Hashtable
        fncSelectAccumulatePrice = False
        Try
            'SQL Query生成
            sbSql.Append(" SELECT  kataban, ")
            sbSql.Append("         kataban_check_div, ")
            sbSql.Append("         place_cd, ")
            sbSql.Append("         ls_price, ")
            sbSql.Append("         rg_price, ")
            sbSql.Append("         ss_price, ")
            sbSql.Append("         bs_price, ")
            sbSql.Append("         gs_price, ")
            sbSql.Append("         ps_price  ")
            sbSql.Append(" FROM    kh_accumulate_price ")
            sbSql.Append(" WHERE   kataban             = @Kataban ")
            'ADD BY YGY 20150311    海外仕入れ追加のため
            sbSql.Append(" AND     currency_cd         = @CurrencyCd ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
                'ADD BY YGY 20150311    海外仕入れ追加のため
                .Parameters.Add("@CurrencyCd", SqlDbType.Char, 3).Value = strCurrency
            End With

            objRdr = objCmd.ExecuteReader

            If objRdr.HasRows = True Then
                objRdr.Read()
                strKatabanCheckDiv = objRdr.GetValue(objRdr.GetOrdinal("kataban_check_div"))
                strPlaceCd = objRdr.GetValue(objRdr.GetOrdinal("place_cd"))
                htPriceInfo(CdCst.UnitPrice.ListPrice) = objRdr.GetValue(objRdr.GetOrdinal("ls_price"))
                htPriceInfo(CdCst.UnitPrice.RegPrice) = objRdr.GetValue(objRdr.GetOrdinal("rg_price"))
                htPriceInfo(CdCst.UnitPrice.SsPrice) = objRdr.GetValue(objRdr.GetOrdinal("ss_price"))
                htPriceInfo(CdCst.UnitPrice.BsPrice) = objRdr.GetValue(objRdr.GetOrdinal("bs_price"))
                htPriceInfo(CdCst.UnitPrice.GsPrice) = objRdr.GetValue(objRdr.GetOrdinal("gs_price"))
                htPriceInfo(CdCst.UnitPrice.PsPrice) = objRdr.GetValue(objRdr.GetOrdinal("ps_price"))

                fncSelectAccumulatePrice = True
            Else
                '単価が取得出来ない場合は0を設定する
                strKatabanCheckDiv = ""
                strPlaceCd = ""
                htPriceInfo(CdCst.UnitPrice.ListPrice) = 0
                htPriceInfo(CdCst.UnitPrice.RegPrice) = 0
                htPriceInfo(CdCst.UnitPrice.SsPrice) = 0
                htPriceInfo(CdCst.UnitPrice.BsPrice) = 0
                htPriceInfo(CdCst.UnitPrice.GsPrice) = 0
                htPriceInfo(CdCst.UnitPrice.PsPrice) = 0

                fncSelectAccumulatePrice = False
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Ｇねじ形番マスタ取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="htPriceInfo">価格情報</param>
    ''' <remarks>Ｇねじ形番マスタを読み込み単価情報を取得し返却する</remarks>
    Public Function subScrewKatabanMstSelect(objCon As SqlConnection, ByVal strKataban As String, _
                                        ByVal strCountryCd As String, ByVal strOfficeCd As String, _
                                        ByRef htPriceInfo As Hashtable) As Boolean

        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        htPriceInfo = New Hashtable
        subScrewKatabanMstSelect = False

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  ls_price, ")
            sbSql.Append("         rg_price, ")
            sbSql.Append("         ss_price, ")
            sbSql.Append("         bs_price, ")
            sbSql.Append("         gs_price, ")
            sbSql.Append("         ps_price ")
            sbSql.Append(" FROM    kh_screw_kataban_mst ")
            sbSql.Append(" WHERE   kataban = @Kataban ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban
            End With

            objRdr = objCmd.ExecuteReader

            If objRdr.HasRows = True Then
                objRdr.Read()

                If (strCountryCd <> CdCst.CountryCd.DefaultCountry) Or _
                      (strCountryCd = CdCst.CountryCd.DefaultCountry And strOfficeCd = CdCst.OfficeCd.Overseas) Then
                    ' If strCountryCd = CdCst.CountryCd.DefaultCountry Then
                    htPriceInfo(CdCst.UnitPrice.ListPrice) = objRdr.GetValue(objRdr.GetOrdinal("ls_price"))
                    htPriceInfo(CdCst.UnitPrice.RegPrice) = objRdr.GetValue(objRdr.GetOrdinal("rg_price"))
                    htPriceInfo(CdCst.UnitPrice.SsPrice) = objRdr.GetValue(objRdr.GetOrdinal("ss_price"))
                    htPriceInfo(CdCst.UnitPrice.BsPrice) = objRdr.GetValue(objRdr.GetOrdinal("bs_price"))
                    htPriceInfo(CdCst.UnitPrice.GsPrice) = objRdr.GetValue(objRdr.GetOrdinal("gs_price"))
                    htPriceInfo(CdCst.UnitPrice.PsPrice) = objRdr.GetValue(objRdr.GetOrdinal("ps_price"))
                Else
                    htPriceInfo(CdCst.UnitPrice.ListPrice) = 0
                    htPriceInfo(CdCst.UnitPrice.RegPrice) = 0
                    htPriceInfo(CdCst.UnitPrice.SsPrice) = 0
                    htPriceInfo(CdCst.UnitPrice.BsPrice) = 0
                    htPriceInfo(CdCst.UnitPrice.GsPrice) = 0
                    htPriceInfo(CdCst.UnitPrice.PsPrice) = 0
                End If
                subScrewKatabanMstSelect = True
            Else
                '単価が取得出来ない場合は0を設定する
                htPriceInfo(CdCst.UnitPrice.ListPrice) = 0
                htPriceInfo(CdCst.UnitPrice.RegPrice) = 0
                htPriceInfo(CdCst.UnitPrice.SsPrice) = 0
                htPriceInfo(CdCst.UnitPrice.BsPrice) = 0
                htPriceInfo(CdCst.UnitPrice.GsPrice) = 0
                htPriceInfo(CdCst.UnitPrice.PsPrice) = 0
                subScrewKatabanMstSelect = False
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 価格表示区分の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="intPriceDispLvl">価格表示レベル</param>
    ''' <param name="strLanguageCd">言語コード</param>
    ''' <param name="strPrice">単価リスト</param>
    ''' <remarks>価格表示レベルを元に価格リストを取得する</remarks>
    Public Function fncSelectDispLvl(objCon As SqlConnection, ByVal intPriceDispLvl As Integer, _
                                ByVal strLanguageCd As String, ByRef strPrice(,) As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  a.price_div, ")
            sbSql.Append("         a.price_lvl, ")
            sbSql.Append("         a.disp_seq_no, ")
            sbSql.Append("         ISNULL(c.price_nm, b.price_nm) AS price_nm ")
            sbSql.Append(" FROM    kh_price_div_mst  a ")
            sbSql.Append(" INNER JOIN  kh_price_div_nm_mst  b ")
            sbSql.Append(" ON      a.price_div = b.price_div ")
            sbSql.Append(" AND     b.language_cd = @DefaultLanguageCd ")
            sbSql.Append(" LEFT JOIN  kh_price_div_nm_mst  c ")
            sbSql.Append(" ON      a.price_div   = c.price_div ")
            sbSql.Append(" AND     c.language_cd = @LanguageCd ")
            sbSql.Append(" ORDER BY  a.price_lvl DESC ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@LanguageCd", SqlDbType.Char, 2).Value = strLanguageCd
                .Parameters.Add("@DefaultLanguageCd", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
            End With

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' 取引通貨マスタの取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectCurrMathAll(objConBase As SqlConnection) As DataTable
        Dim objCmd As New SqlCommand
        Dim objAdp As New SqlDataAdapter
        Dim sbSql As New Text.StringBuilder

        fncSelectCurrMathAll = New DataTable
        Try
            'DB接続
            objCmd = objConBase.CreateCommand
            sbSql.Append(" SELECT ")
            sbSql.Append("     math_Type, ")
            sbSql.Append("     math_Pos, ")
            sbSql.Append("     currency_cd ")
            sbSql.Append(" FROM ")
            sbSql.Append("     kh_currency_mst ")

            objCmd.CommandText = sbSql.ToString
            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncSelectCurrMathAll)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 取引通貨マスタより端数データを取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strCurrencyCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectCurrMath(objConBase As SqlConnection, ByVal strCurrencyCd As String) As DataTable
        Dim objCmd As New SqlCommand
        Dim objAdp As New SqlDataAdapter
        Dim sbSql As New Text.StringBuilder
        fncSelectCurrMath = New DataTable

        Try
            sbSql.Append(" SELECT ")
            sbSql.Append("     math_Type, ")
            sbSql.Append("     math_Pos ")
            sbSql.Append(" FROM ")
            sbSql.Append("     kh_currency_mst ")
            sbSql.Append(" WHERE currency_cd = @currency_cd ")
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@currency_cd", SqlDbType.Char, 3).Value = strCurrencyCd
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncSelectCurrMath)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 現地定価の為替レートを取得する
    ''' </summary>
    ''' <param name="strChangeCurr">ログイン国の通貨</param>
    ''' <param name="dblRate">為替レート</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectRateMstAprice(objConBase As SqlConnection, ByVal strChangeCurr As String, _
                                          ByVal strBaseCurr As String, ByRef dblRate As Double) As Boolean
        Dim objCmd As New SqlCommand
        Dim objAdp As New SqlDataAdapter
        Dim sbSql As New Text.StringBuilder
        Dim dt As New DataTable
        fncSelectRateMstAprice = False

        Try
            sbSql.Append(" SELECT ")
            sbSql.Append("     base_currency_cd, ")
            sbSql.Append("     change_currency_cd, ")
            sbSql.Append("     exchange_rate, ")
            sbSql.Append("     in_effective_date, ")
            sbSql.Append("     out_effective_date ")
            sbSql.Append(" FROM ")
            sbSql.Append("     kh_currency_exc_rate_mst ")
            sbSql.Append(" WHERE base_currency_cd = @BaseCurrencyCd ")
            sbSql.Append(" AND change_currency_cd = @ChangeCurrencyCd ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@BaseCurrencyCd", SqlDbType.Char, 3).Value = strBaseCurr
                .Parameters.Add("@ChangeCurrencyCd", SqlDbType.Char, 3).Value = strChangeCurr
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt)

            If dt.Rows.Count > 0 Then
                Dim strDateTime As DateTime = Now.Date

                dblRate = dt.Rows(0)("exchange_rate")
                '失効日対応
                If (dt.Rows(0)("in_effective_date") <= strDateTime) And _
                 (strDateTime < dt.Rows(0)("out_effective_date")) Then
                    fncSelectRateMstAprice = True
                Else
                    fncSelectRateMstAprice = False
                End If
            Else
                fncSelectRateMstAprice = False
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 現地定価掛け率の取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strSeries"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectRateAprice(objConBase As SqlConnection, ByVal strCountryCd As String, _
                                          ByVal strSeries As String) As DataTable
        Dim objCmd As New SqlCommand
        Dim objAdp As New SqlDataAdapter
        Dim sbSql As New Text.StringBuilder
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  ISNULL(list_price_rate1,0) AS list_price_rate1, ")
            sbSql.Append("         ISNULL(list_price_rate2,0) AS list_price_rate2, ")
            sbSql.Append("         ISNULL(math_TypeA,-1) AS TypeA, ")
            sbSql.Append("         ISNULL(math_PosA,1) AS PosA ")
            sbSql.Append(" FROM    kh_country_rate_localprice_mst as tblRate ")
            sbSql.Append(" INNER JOIN kh_country_mst as tblCun ")
            sbSql.Append(" ON tblRate.country_cd = tblCun.country_cd ")
            sbSql.Append(" WHERE   tblRate.country_cd          = @CountryCd ")
            sbSql.Append(" AND     tblRate.rate_search_key     = @RateSearchKey ")
            sbSql.Append(" AND     tblRate.in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     tblRate.out_effective_date  > @StandardDate ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@CountryCd", SqlDbType.Char, 3).Value = strCountryCd
                .Parameters.Add("@RateSearchKey", SqlDbType.VarChar, 60).Value = strSeries
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try
        Return dtResult
    End Function

    ''' <summary>
    ''' 購入価格の取引通貨マスタより端数データを取得する
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strCurrBase"></param>
    ''' <param name="strCurrChange"></param>
    ''' <param name="dblRate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetRateMst(objCon As SqlConnection, ByVal strCurrBase As String, _
                                   ByVal strCurrChange As String, ByRef dblRate As Double) As Boolean
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        fncGetRateMst = False
        Try
            sbSql.Append(" SELECT ")
            sbSql.Append("     base_currency_cd, ")
            sbSql.Append("     change_currency_cd, ")
            sbSql.Append("     exchange_rate, ")
            sbSql.Append("     in_effective_date, ")
            sbSql.Append("     out_effective_date ")
            sbSql.Append(" FROM ")
            sbSql.Append("     kh_currency_exc_rate_mst ")
            sbSql.Append(" WHERE base_currency_cd = @BcurrencyCd ")
            sbSql.Append(" AND change_currency_cd = @CcurrencyCd ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@BcurrencyCd", SqlDbType.Char, 3).Value = strCurrBase
                .Parameters.Add("@CcurrencyCd", SqlDbType.Char, 3).Value = strCurrChange
            End With

            objRdr = objCmd.ExecuteReader

            If objRdr.HasRows = True Then
                objRdr.Read()
                dblRate = objRdr.GetValue(objRdr.GetOrdinal("exchange_rate"))

                Dim strDateTime As DateTime = Now.Date
                '失効日対応
                If (objRdr.GetValue(objRdr.GetOrdinal("in_effective_date")) <= strDateTime) And _
                 (strDateTime < objRdr.GetValue(objRdr.GetOrdinal("out_effective_date"))) Then
                    fncGetRateMst = True
                Else
                    fncGetRateMst = False
                End If
            Else
                fncGetRateMst = False
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 購入価格掛け率の取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strSelCountry"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectRateFobprice(objConBase As SqlConnection, ByVal strCountryCd As String, _
                                          ByVal strSeries As String, ByVal strSelCountry As String) As DataTable
        Dim objCmd As New SqlCommand
        Dim objAdp As New SqlDataAdapter
        Dim sbSql As New Text.StringBuilder
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  ISNULL(fob_rate,0) AS fob_rate, ")
            sbSql.Append("         ISNULL(math_Type,-1) AS TypeFOB, ")
            sbSql.Append("         ISNULL(math_Pos,1) AS PosFOB,currency_cd, ")
            sbSql.Append("         authorization_no ")
            sbSql.Append(" FROM    kh_country_rate_netprice_mst as tblRate ")
            sbSql.Append(" INNER JOIN kh_currency_trade_mst as tblCun ")
            sbSql.Append(" ON tblRate.exp_country_cd = tblCun.exp_country_cd ")
            sbSql.Append(" AND tblRate.imp_country_cd = tblCun.imp_country_cd ")
            sbSql.Append(" WHERE   tblRate.exp_country_cd      = @MadeCountryCd ")
            sbSql.Append(" AND     tblRate.imp_country_cd      = @CountryCd ")
            sbSql.Append(" AND     tblRate.rate_search_key     = @RateSearchKey ")
            sbSql.Append(" AND     tblRate.in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     tblRate.out_effective_date  > @StandardDate ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@CountryCd", SqlDbType.Char, 3).Value = strCountryCd
                .Parameters.Add("@MadeCountryCd", SqlDbType.Char, 3).Value = strSelCountry
                .Parameters.Add("@RateSearchKey", SqlDbType.VarChar, 60).Value = strSeries
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dtResult)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try
        Return dtResult
    End Function
End Class
