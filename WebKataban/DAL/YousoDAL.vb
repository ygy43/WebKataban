Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class YousoDAL

    ''' <summary>
    ''' 形番構成取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strcCompData"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks>形番構成を読み込み該当するデータを変数に格納する</remarks>
    Public Function fncKatabanStrcSelect(ByVal objCon As SqlConnection, ByRef strcCompData As YousoBLL.CompData, _
                                                strLanguage As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  a.ktbn_strc_seq_no, ")
            sbSql.Append("         a.element_div, ")
            sbSql.Append("         a.structure_div, ")
            sbSql.Append("         a.addition_div, ")
            sbSql.Append("         a.hyphen_div, ")
            sbSql.Append("         b.ktbn_strc_nm as defaultNm, ")
            sbSql.Append("         c.ktbn_strc_nm ")
            sbSql.Append(" FROM    kh_kataban_strc a ")
            sbSql.Append(" INNER JOIN  kh_ktbn_strc_nm_mst b ")
            sbSql.Append(" ON      a.series_kataban      = b.series_kataban ")
            sbSql.Append(" AND     a.key_kataban         = b.key_kataban ")
            sbSql.Append(" AND     a.ktbn_strc_seq_no    = b.ktbn_strc_seq_no ")
            sbSql.Append(" AND     b.language_cd         = @DeaultLanguageCd ")
            sbSql.Append(" AND     b.in_effective_date  <= @PrmStandardDate ")
            sbSql.Append(" AND     b.out_effective_date  > @PrmStandardDate ")
            sbSql.Append(" LEFT  JOIN  kh_ktbn_strc_nm_mst c ")
            sbSql.Append(" ON      a.series_kataban      = c.series_kataban ")
            sbSql.Append(" AND     a.key_kataban         = c.key_kataban ")
            sbSql.Append(" AND     a.ktbn_strc_seq_no    = c.ktbn_strc_seq_no ")
            sbSql.Append(" AND     c.language_cd         = @LanguageCd ")
            sbSql.Append(" AND     c.in_effective_date  <= @PrmStandardDate ")
            sbSql.Append(" AND     c.out_effective_date  > @PrmStandardDate ")
            sbSql.Append(" WHERE   a.series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     a.key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     a.in_effective_date  <= @PrmStandardDate ")
            sbSql.Append(" AND     a.out_effective_date  > @PrmStandardDate ")
            sbSql.Append(" ORDER BY  a.ktbn_strc_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strcCompData.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strcCompData.strKeyKataban
                .Parameters.Add("@DeaultLanguageCd", SqlDbType.VarChar, 2).Value = CdCst.LanguageCd.DefaultLang
                .Parameters.Add("@LanguageCd", SqlDbType.VarChar, 2).Value = strLanguage
                .Parameters.Add("@PrmStandardDate", SqlDbType.DateTime).Value = Now()
            End With

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
    ''' 形番構成要素取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strcCompData"></param>
    ''' <returns></returns>
    ''' <remarks>形番構成要素を読み込み該当するデータを変数に格納する</remarks>
    Public Function subKtbnStrcEleSelect(ByVal objCon As SqlConnection, ByRef strcCompData As YousoBLL.CompData) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  ktbn_strc_seq_no, ")
            sbSql.Append("         option_symbol, ")
            sbSql.Append("         place_lvl ")
            sbSql.Append(" FROM    kh_kataban_strc_ele ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     in_effective_date  <= @PrmStandardDate ")
            sbSql.Append(" AND     out_effective_date  > @PrmStandardDate ")
            sbSql.Append(" ORDER BY  ktbn_strc_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strcCompData.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strcCompData.strKeyKataban
                .Parameters.Add("@PrmStandardDate", SqlDbType.DateTime).Value = Now()
            End With
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
    ''' 引当形番構成データのチェック
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks>引当形番構成を読み込み該当するデータがあるかチェックする</remarks>
    Public Function fncSelKtbnStrcCheck(ByVal objCon As SqlConnection, _
                                               strUserId As String, strSessionId As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  * ")
            sbSql.Append(" FROM    kh_sel_ktbn_strc ")
            sbSql.Append(" WHERE   user_id    = @PrmUserId ")
            sbSql.Append(" AND     session_id = @PrmSessionId ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@PrmUserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@PrmSessionId", SqlDbType.NVarChar, 88).Value = strSessionId
            End With
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
    ''' 出荷場所レベルの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <param name="strSymbol"></param>
    ''' <param name="intNo"></param>
    ''' <param name="intKtbnPlcaelvl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function subGetPlacelvl(ByVal objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                          ByVal strKeyKataban As String, ByVal strSymbol As String, _
                                          ByVal intNo As Integer, ByRef intKtbnPlcaelvl As Integer) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  place_lvl ")
            sbSql.Append(" FROM    kh_kataban_strc_ele ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     in_effective_date  <= @PrmStandardDate ")
            sbSql.Append(" AND     out_effective_date  > @PrmStandardDate ")
            sbSql.Append(" AND     ktbn_strc_seq_no  = @ktbn_strc_seq_no ")
            sbSql.Append(" AND     option_symbol  = @option_symbol ")
            sbSql.Append(" ORDER BY  in_effective_date DESC ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
                .Parameters.Add("@PrmStandardDate", SqlDbType.DateTime).Value = Now()
                .Parameters.Add("@option_symbol", SqlDbType.VarChar, 15).Value = strSymbol
                .Parameters.Add("@ktbn_strc_seq_no", SqlDbType.Int).Value = intNo
            End With

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
    ''' 要素パターン取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intFocusNo"></param>
    ''' <param name="intConSeqNoBr"></param>
    ''' <param name="strConOpSymbol"></param>
    ''' <returns></returns>
    ''' <remarks>引当形番構成を読み込み該当するデータがあるかチェックする</remarks>
    Public Function fncElePtnSelect(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc, _
                                           ByVal intFocusNo As Integer, ByRef intConSeqNoBr As ArrayList, _
                                           ByRef strConOpSymbol As ArrayList) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            '配列初期化
            'SQL Query生成
            sbSql.Append(" SELECT  condition_seq_no_br, ")
            sbSql.Append("         cond_option_symbol ")
            sbSql.Append(" FROM    kh_ele_pattern ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     ktbn_strc_seq_no    = @KtbnStrcSeqNo ")
            sbSql.Append(" AND     option_symbol       = @OptionSymbol ")
            sbSql.Append(" AND     in_effective_date  <= @PrmStandardDate ")
            sbSql.Append(" AND     out_effective_date  > @PrmStandardDate ")
            sbSql.Append(" ORDER BY  condition_seq_no_br ")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = objKtbnStrc.strcSelection.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = objKtbnStrc.strcSelection.strKeyKataban
                .Parameters.Add("@KtbnStrcSeqNo", SqlDbType.Int).Value = intFocusNo
                .Parameters.Add("@OptionSymbol", SqlDbType.VarChar, 15).Value = "#"
                .Parameters.Add("@PrmStandardDate", SqlDbType.DateTime).Value = Now()
            End With
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
    ''' '生産国データの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetPlacelvl(ByVal objCon As SqlConnection, ByVal strSeries As String, _
                                          ByVal strKey As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'DB接続
            objCmd = objCon.CreateCommand

            sbSql.Append("  SELECT  ktbn_strc_seq_no,option_symbol,disp_seq_no,place_lvl ")
            sbSql.Append("  FROM    kh_kataban_strc_ele ")
            sbSql.Append("  WHERE   series_kataban = '" & strSeries & "' ")
            sbSql.Append("  AND     key_kataban = '" & strKey & "' ")
            sbSql.Append("  AND     in_effective_date  <= '" & Now & "' ")
            sbSql.Append("  AND     out_effective_date  > '" & Now & "' ")

            objCmd.CommandText = sbSql.ToString
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
    ''' 'オプション外設定ファイルの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strSessionID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetOutofopPlacelvl(ByVal objCon As SqlConnection, ByVal strUserID As String, _
                                          ByVal strSessionID As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT      isnull(place_lvl,0) as place_lvl ")
            sbSql.Append(" FROM        kh_sel_outofop_order ")
            sbSql.Append(" WHERE       user_id    = @UserId ")
            sbSql.Append(" AND         session_id = @SessionId ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionID
            End With
            'DBオープン
            objAdp = New SqlDataAdapter(objCmd)
            Dim dt As New DataTable
            objAdp.Fill(dt)
            dtResult = dt
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
    ''' 全ての生産国名を取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetAllCountryLevel(ByVal objConBase As SqlConnection) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'DB接続
            objCmd = objConBase.CreateCommand

            'SQL Query生成
            sbSql.Append(" SELECT  place_div,disp_seq_no,place_lvl ")
            sbSql.Append(" FROM    kh_place_div_mst ")
            sbSql.Append(" ORDER BY  place_lvl DESC ")

            objCmd.CommandText = sbSql.ToString

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
    ''' ストローク国名を取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strKey"></param>
    ''' <param name="intBoreSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetStrokeCountry(ByVal objCon As SqlConnection, ByVal strSeries As String, _
                                               ByVal strKey As String, ByVal intBoreSize As Long) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try

            'DB接続
            objCmd = objCon.CreateCommand

            'SQL Query生成
            sbSql.Append(" SELECT  country_cd, min_stroke, max_stroke ")
            sbSql.Append(" FROM    kh_stroke  ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     bore_size           = @BoreSize ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" Group by  country_cd, min_stroke, max_stroke ")

            objCmd.CommandText = sbSql.ToString

            With objCmd.Parameters
                .Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeries
                .Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKey
                .Add("@BoreSize", SqlDbType.Int).Value = intBoreSize
                .Add("@StandardDate", SqlDbType.DateTime).Value = Now()
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
    ''' 標準ストロークを取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strKey"></param>
    ''' <param name="intBoreSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetStdStroke(ByVal objCon As SqlConnection, ByVal strSeries As String, _
                                    ByVal strKey As String, ByVal intBoreSize As Long, _
                                    ByVal intstroke As Long) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'DB接続
            objCmd = objCon.CreateCommand

            'SQL Query生成
            sbSql.Append(" SELECT  std_stroke ")
            sbSql.Append(" FROM    kh_std_stroke_mst  ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     bore_size           = @BoreSize ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     std_stroke  = @stroke ")

            objCmd.CommandText = sbSql.ToString
            With objCmd.Parameters
                .Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeries
                .Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKey
                .Add("@BoreSize", SqlDbType.Int).Value = intBoreSize
                .Add("@StandardDate", SqlDbType.DateTime).Value = Now()
                .Add("@stroke", SqlDbType.Int).Value = intstroke
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
    ''' 形番構成要素を取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <param name="strLanguageCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function subKataStrcEleSel(ByVal objCon As SqlConnection, _
                                             ByVal strSeriesKataban As String, ByVal strKeyKataban As String, _
                                             ByVal strLanguageCd As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'DB接続
            objCmd = objCon.CreateCommand

            'SQL Query生成
            sbSql.Append(" SELECT  b.element_div, ")
            sbSql.Append("         b.structure_div, ")
            sbSql.Append("         c.option_symbol, ")
            sbSql.Append("         d.option_nm as default_option_nm, ")
            sbSql.Append("         e.option_nm, ")
            sbSql.Append("         b.ktbn_strc_seq_no, ")
            sbSql.Append("         c.disp_seq_no ")
            sbSql.Append(" FROM    kh_series_kataban a ")
            sbSql.Append(" INNER JOIN  kh_kataban_strc b ")
            sbSql.Append(" ON      a.series_kataban         = b.series_kataban ")
            sbSql.Append(" AND     a.key_kataban            = b.key_kataban ")
            sbSql.Append(" AND     b.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     b.out_effective_date     > @StandardDate ")
            sbSql.Append(" INNER JOIN  kh_kataban_strc_ele c ")
            sbSql.Append(" ON      b.series_kataban         = c.series_kataban ")
            sbSql.Append(" AND     b.key_kataban            = c.key_kataban ")
            sbSql.Append(" AND     b.ktbn_strc_seq_no       = c.ktbn_strc_seq_no ")
            sbSql.Append(" AND     c.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     c.out_effective_date     > @StandardDate ")
            sbSql.Append(" INNER JOIN  kh_option_nm_mst d ")
            sbSql.Append(" ON      c.series_kataban         = d.series_kataban ")
            sbSql.Append(" AND     c.key_kataban            = d.key_kataban ")
            sbSql.Append(" AND     c.ktbn_strc_seq_no       = d.ktbn_strc_seq_no ")
            sbSql.Append(" AND     c.option_symbol          = d.option_symbol ")
            sbSql.Append(" AND     d.language_cd            = @DefaultLangCd ")
            sbSql.Append(" AND     d.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     d.out_effective_date     > @StandardDate ")
            sbSql.Append(" LEFT  JOIN  kh_option_nm_mst e ")
            sbSql.Append(" ON      c.series_kataban         = e.series_kataban ")
            sbSql.Append(" AND     c.key_kataban            = e.key_kataban ")
            sbSql.Append(" AND     c.ktbn_strc_seq_no       = e.ktbn_strc_seq_no ")
            sbSql.Append(" AND     c.option_symbol          = e.option_symbol ")
            sbSql.Append(" AND     e.language_cd            = @LanguageCd ")
            sbSql.Append(" AND     e.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     e.out_effective_date     > @StandardDate ")
            sbSql.Append(" WHERE   a.series_kataban         = @SeriesKataban ")
            sbSql.Append(" AND     a.key_kataban            = @KeyKataban ")
            sbSql.Append(" AND     a.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     a.out_effective_date     > @StandardDate ")
            sbSql.Append(" ORDER BY  c.disp_seq_no ")

            objCmd.CommandText = sbSql.ToString

            objCmd.Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
            objCmd.Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
            objCmd.Parameters.Add("@DefaultLangCd", SqlDbType.NVarChar, 150).Value = CdCst.LanguageCd.DefaultLang
            objCmd.Parameters.Add("@LanguageCd", SqlDbType.NVarChar, 150).Value = strLanguageCd
            objCmd.Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()

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
    ''' エレパタンの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function subElePatternSel(ByVal objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                    ByVal strKeyKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'DB接続
            objCmd = objCon.CreateCommand

            'SQL Query生成
            sbSql.Append(" SELECT  '1' as serach_seq_no, ")
            sbSql.Append("         option_symbol, ")
            sbSql.Append("         condition_cd, ")
            sbSql.Append("         condition_seq_no, ")
            sbSql.Append("         condition_seq_no_br, ")
            sbSql.Append("         cond_option_symbol, ")
            sbSql.Append("         ktbn_strc_seq_no ")
            sbSql.Append(" FROM    kh_ele_pattern ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     option_symbol       = '" & CdCst.ElePattern.Plural & "' ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT  '2' as serach_seq_no, ")
            sbSql.Append("         option_symbol, ")
            sbSql.Append("         condition_cd, ")
            sbSql.Append("         condition_seq_no, ")
            sbSql.Append("         condition_seq_no_br, ")
            sbSql.Append("         cond_option_symbol,ktbn_strc_seq_no ")
            sbSql.Append(" FROM    kh_ele_pattern ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     option_symbol       = '" & CdCst.ElePattern.All & "' ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT  '3' as serach_seq_no, ")
            sbSql.Append("         option_symbol, ")
            sbSql.Append("         condition_cd, ")
            sbSql.Append("         condition_seq_no, ")
            sbSql.Append("         condition_seq_no_br, ")
            sbSql.Append("         cond_option_symbol,ktbn_strc_seq_no ")
            sbSql.Append(" FROM    kh_ele_pattern ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     option_symbol  Not In ('" & CdCst.ElePattern.All & "','" & CdCst.ElePattern.Plural & "') ")
            sbSql.Append(" ORDER BY  serach_seq_no, option_symbol, condition_seq_no, condition_seq_no_br ")

            objCmd.CommandText = sbSql.ToString

            objCmd.Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
            objCmd.Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
            objCmd.Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()

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
    ''' 引当口径検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function subBoreSizeSelect(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  ktbn_strc_seq_no ")
            sbSql.Append(" FROM    kh_kataban_strc ")
            sbSql.Append(" WHERE   series_kataban  = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban     = @KeyKataban ")
            sbSql.Append(" AND     element_div = '" & CdCst.RodEndCstmOrder.EleBoreSize & "'")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = objKtbnStrc.strcSelection.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = objKtbnStrc.strcSelection.strKeyKataban
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
    ''' ストロークを取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intBoreSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetStroke(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByVal intBoreSize As Integer) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  min_stroke, ")
            sbSql.Append("         max_stroke, ")
            sbSql.Append("         stroke_unit ")
            sbSql.Append(" FROM    kh_stroke ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     bore_size           = @BoreSize")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     country_cd  = @countrycd ")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = objKtbnStrc.strcSelection.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = objKtbnStrc.strcSelection.strKeyKataban
                .Parameters.Add("@BoreSize", SqlDbType.Int).Value = intBoreSize
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
                .Parameters.Add("@countrycd", SqlDbType.VarChar, 3).Value = objKtbnStrc.strcSelection.strMadeCountry
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
