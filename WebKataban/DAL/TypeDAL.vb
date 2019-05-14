Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class TypeDAL

    ''' <summary>
    ''' 形番検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strRange"></param>
    ''' <param name="strMinFlg"></param>
    ''' <param name="strKataban"></param>
    ''' <param name="strCountrycd"></param>
    ''' <remarks></remarks>
    Public Function fncSelectBySeries(ByVal objCon As SqlConnection, ByVal strRange As String, _
                               ByVal strMinFlg As String, ByVal strKataban As String, _
                               ByVal strCountrycd As String, ByVal seriesSearch As KHSeriesSearch, _
                               ByVal lstWhereSeries As ArrayList) As DataSet
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand

        Try
            seriesSearch.dsKatabanValue = New DataSet
            seriesSearch.strHeaderValue = "形　番"

            If strRange IsNot Nothing Then
                sbSql.Append("SELECT ")
                sbSql.Append(strRange)
                sbSql.Append("       sortKey, ")
                sbSql.Append("       series_kataban, ")
                sbSql.Append("       key_kataban, ")
                sbSql.Append("       disp_kataban, ")
                sbSql.Append("       division, ")
                sbSql.Append("       disp_name, ")
                sbSql.Append("       price_no, ")
                sbSql.Append("       spec_no, ")
                sbSql.Append("       order_no, ")
                'ADD BY YGY 20150311    海外仕入れ追加のため
                sbSql.Append("       currency_cd ")
                sbSql.Append("FROM   ( ")
            End If
            sbSql.Append("SELECT ")
            sbSql.Append("       DISTINCT ")
            sbSql.Append("       KT.series_kataban + '_' + ")
            sbSql.Append("       KT.key_kataban  as sortKey, ")
            sbSql.Append("       KT.series_kataban, ")
            sbSql.Append("       KT.key_kataban, ")
            sbSql.Append("       ISNULL(NM.disp_kataban,DF.disp_kataban) AS disp_kataban, ")
            sbSql.Append("       '1'                                     AS division, ")
            sbSql.Append("       ISNULL(NM.series_nm,DF.series_nm)       AS disp_name, ")
            sbSql.Append("       KT.price_no, ")
            sbSql.Append("       KT.spec_no, ")
            sbSql.Append("       KT.order_no, ")
            sbSql.Append("       KT.currency_cd ")
            sbSql.Append(" FROM   kh_series_kataban  KT ")
            sbSql.Append("  INNER JOIN kh_series_nm_mst   DF ")
            sbSql.Append("    ON  KT.series_kataban      = DF.series_kataban ")
            sbSql.Append("    AND KT.key_kataban         = DF.key_kataban ")
            sbSql.Append("  LEFT OUTER JOIN kh_series_nm_mst   NM ")
            sbSql.Append("    ON  KT.series_kataban      = NM.series_kataban ")
            sbSql.Append("    AND KT.key_kataban         = NM.key_kataban ")
            sbSql.Append("    AND NM.language_cd         = @LangCd ")
            sbSql.Append("    AND NM.in_effective_date  <= @StandardDate ")
            sbSql.Append("    AND NM.out_effective_date  > @StandardDate ")
            sbSql.Append("  LEFT OUTER JOIN kh_country_group_mst   DI ")
            sbSql.Append("    ON  KT.country_group_cd      = DI.country_group_cd ")
            sbSql.Append("WHERE    KT.in_effective_date     <= @StandardDate ")
            sbSql.Append("AND    KT.out_effective_date     > @StandardDate ")
            sbSql.Append("AND    DF.language_cd            = @DefaultLangCd ")
            sbSql.Append("AND    DF.in_effective_date     <= @StandardDate ")
            sbSql.Append("AND    DF.out_effective_date     > @StandardDate ")
            sbSql.Append("AND    KT.series_kataban      LIKE @SrsKataban ")
            sbSql.Append("AND    ( KT.country_group_cd = 'ALL'  ")
            sbSql.Append("OR     DI.country_cd     = @country_cd ) ")

            If strMinFlg <> "0" Then
                Dim str() As String = seriesSearch.strMinKatabanValue.ToString.Split("_")
                If str.Length = 2 Then
                    If str(1).ToString.Trim.Length <= 0 Then
                        sbSql.Append(" AND ((KT.series_kataban > @MinSeries) OR (KT.series_kataban = @MinSeries AND KT.key_kataban <> '')) ")
                    Else
                        sbSql.Append(" AND  ((KT.series_kataban = @MinSeries AND KT.key_kataban > @MinKey) ")
                        sbSql.Append(" OR  (KT.series_kataban > @MinSeries)) ")
                    End If
                End If
            End If

            'CHANGED BY YGY 20150121  RM1501051  ↓↓↓↓↓↓
            '台湾のAX専門の特定取引代理店特殊対応
            If lstWhereSeries.Count > 0 Then
                Dim strWhereSeries As String = " AND (KT.series_kataban IN ("

                For inti As Integer = 0 To lstWhereSeries.Count - 1
                    If inti = 0 Then
                        strWhereSeries &= " '" & lstWhereSeries(inti) & "' "
                    Else
                        strWhereSeries &= " ,'" & lstWhereSeries(inti) & "' "
                    End If
                Next

                strWhereSeries &= "))"
                sbSql.Append(strWhereSeries)
            End If
            'CHANGED BY YGY 20150121  RM1501051  ↑↑↑↑↑↑

            If strRange IsNot Nothing Then
                sbSql.Append(" ) AS Kata ")
                sbSql.Append(" ORDER BY series_kataban, ")
                sbSql.Append("         order_no, ")
                sbSql.Append("         key_kataban ")

            Else
                sbSql.Append(" ORDER BY KT.series_kataban, ")
                sbSql.Append("         KT.order_no, ")
                sbSql.Append("         KT.key_kataban ")
            End If

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString
            objCmd.Parameters.Add("@LangCd", SqlDbType.VarChar, 2).Value = seriesSearch.strLangCdValue
            objCmd.Parameters.Add("@DefaultLangCd", SqlDbType.VarChar, 2).Value = CdCst.LanguageCd.DefaultLang
            objCmd.Parameters.Add("@SrsKataban", SqlDbType.VarChar, 10).Value = seriesSearch.strSrsKataValue & "%"
            If strMinFlg <> "0" Then
                Dim str() As String = seriesSearch.strMinKatabanValue.ToString.Split("_")
                If str.Length = 2 Then
                    objCmd.Parameters.Add("@MinSeries", SqlDbType.VarChar, 60).Value = str(0).ToString
                    If str(1).ToString.Trim.Length > 0 Then  'キー形番あれば
                        objCmd.Parameters.Add("@MinKey", SqlDbType.VarChar, 1).Value = str(1).ToString
                    End If
                End If
            End If
            objCmd.Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            objCmd.Parameters.Add("@country_cd", SqlDbType.Char, 3).Value = strCountrycd

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(seriesSearch.dsKatabanValue, "KatabanTbl")

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return seriesSearch.dsKatabanValue
    End Function

    ''' <summary>
    ''' ＣＫＤ形番検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strRange"></param>
    ''' <param name="strMinFlg"></param>
    ''' <param name="strCountrycd"></param>
    ''' <remarks></remarks>
    Public Function fncSelectByFullKataban(ByVal objCon As SqlConnection, ByVal strRange As String, _
                              ByVal strMinFlg As String, ByVal strCountrycd As String, _
                              ByVal seriesSearch As KHSeriesSearch, ByVal lstWhereSeries As ArrayList) As DataSet
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand

        Try
            seriesSearch.dsKatabanValue = New DataSet
            seriesSearch.strHeaderValue = "ＣＫＤ形番"

            sbSql.Append("SELECT ")
            If strRange IsNot Nothing Then
                sbSql.Append(strRange)
            End If
            sbSql.Append("        PRC.kataban + ")
            sbSql.Append("        CONVERT(VARCHAR, PRC.in_effective_date, 120)  as sortKey, ")
            sbSql.Append("        ' '            AS series_kataban, ")
            sbSql.Append("        ' '            AS key_kataban, ")
            sbSql.Append("        PRC.kataban    AS disp_kataban, ")
            sbSql.Append("        '2'            AS division, ")
            sbSql.Append("        PRC.kataban_check_div, ")
            sbSql.Append("        PNM.parts_nm, ")
            sbSql.Append("        PNM.model_nm, ")
            sbSql.Append("        ''    AS disp_name, ")
            sbSql.Append("        ''    AS price_no, ")
            sbSql.Append("        ''    AS spec_no, ")
            'ADD BY YGY 20150311    海外仕入れ追加のため
            sbSql.Append("        PRC.currency_cd ")

            sbSql.Append("FROM    kh_price  PRC ")
            sbSql.Append("  LEFT OUTER JOIN kh_parts_nm_mst   PNM ")
            sbSql.Append("    ON  PRC.kataban              = PNM.kataban ")
            sbSql.Append("    AND PNM.language_cd          = @LangCd ")
            sbSql.Append("    AND PNM.in_effective_date   <= @StandardDate ")
            sbSql.Append("    AND PNM.out_effective_date   > @StandardDate ")
            'ADD BY YGY 20150311    海外仕入れ追加のため
            sbSql.Append("    AND PRC.currency_cd          = PNM.currency_cd ")

            sbSql.Append("  LEFT OUTER JOIN kh_country_group_mst   DI ")
            sbSql.Append("    ON  PRC.country_group_cd      = DI.country_group_cd ")
            sbSql.Append("WHERE  PRC.kataban            LIKE @SrsKataban ")
            sbSql.Append("AND    PRC.in_effective_date    <= @StandardDate ")
            sbSql.Append("AND    PRC.out_effective_date    > @StandardDate ")
            sbSql.Append("AND    ( PRC.country_group_cd = 'ALL'  ")
            sbSql.Append("OR     DI.country_cd     = @country_cd ) ")
            If strMinFlg <> "0" Then
                sbSql.Append("AND  PRC.kataban + ")
                sbSql.Append("     CONVERT(VARCHAR, PRC.in_effective_date, 120) > @MinKataban ")
            End If

            'CHANGED BY YGY 20150121  RM1501051  ↓↓↓↓↓↓
            '台湾のAX専門の特定取引代理店特殊対応
            'フル形番で検索する時にAXで始まる全てのフル形番を表示
            If lstWhereSeries.Count > 0 Then
                Dim strWhereSeries As String = " AND ((PRC.kataban LIKE 'AX%') OR ("

                For inti As Integer = 0 To lstWhereSeries.Count - 1
                    If inti = 0 Then
                        strWhereSeries &= " PRC.kataban LIKE '" & lstWhereSeries(inti) & "%' "
                    Else
                        strWhereSeries &= " OR PRC.kataban LIKE '" & lstWhereSeries(inti) & "%' "
                    End If
                Next

                strWhereSeries &= "))"
                sbSql.Append(strWhereSeries)
            End If
            'CHANGED BY YGY 20150121  RM1501051  ↑↑↑↑↑↑

            sbSql.Append("ORDER BY PRC.kataban ")

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString
            objCmd.Parameters.Add("@SrsKataban", SqlDbType.VarChar, 60).Value = seriesSearch.strSrsKataValue & "%"
            objCmd.Parameters.Add("@LangCd", SqlDbType.VarChar, 2).Value = seriesSearch.strLangCdValue
            If strMinFlg <> "0" Then
                objCmd.Parameters.Add("@MinKataban", SqlDbType.VarChar, 61).Value = seriesSearch.strMinKatabanValue
            End If
            objCmd.Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            objCmd.Parameters.Add("@country_cd", SqlDbType.Char, 3).Value = strCountrycd

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(seriesSearch.dsKatabanValue, "KatabanTbl")

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return seriesSearch.dsKatabanValue
    End Function

    ''' <summary>
    '''  全て検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strRange"></param>
    ''' <param name="strMinFlg"></param>
    ''' <param name="strKataban"></param>
    ''' <param name="strCountrycd"></param>
    ''' <remarks></remarks>
    Public Function fncSelectByAll(ByVal objCon As SqlConnection, ByVal strRange As String, _
                             ByVal strMinFlg As String, ByVal strKataban As String, _
                             ByVal strCountrycd As String, ByVal seriesSearch As KHSeriesSearch, _
                             ByVal lstWhereSeries As ArrayList) As DataSet
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand

        Try
            seriesSearch.dsKatabanValue = New DataSet
            seriesSearch.strHeaderValue = "シリーズ形番/ＣＫＤ形番"

            sbSql.Append("SELECT ")
            If strRange IsNot Nothing Then
                sbSql.Append(strRange)
            End If

            sbSql.Append("       sortKey, ")
            sbSql.Append("       series_kataban, ")
            sbSql.Append("       key_kataban, ")
            sbSql.Append("       disp_kataban, ")
            sbSql.Append("       division, ")
            sbSql.Append("       kataban_check_div, ")
            sbSql.Append("       parts_nm, ")
            sbSql.Append("       model_nm, ")
            sbSql.Append("       disp_name, ")
            sbSql.Append("       price_no, ")
            sbSql.Append("       spec_no, ")
            sbSql.Append("       currency_cd, ")
            sbSql.Append("       country_cd ")
            sbSql.Append("FROM   ( ")
            sbSql.Append("  SELECT ")
            sbSql.Append("         DISTINCT ")
            sbSql.Append("         '1_' + ")
            sbSql.Append("         KT.series_kataban + '_' + ")
            sbSql.Append("         KT.key_kataban  as sortKey, ")
            sbSql.Append("         KT.series_kataban, ")
            sbSql.Append("         KT.key_kataban, ")
            sbSql.Append("         ISNULL(NM.disp_kataban,DF.disp_kataban) AS disp_kataban, ")
            sbSql.Append("         '1'                                     AS division, ")
            sbSql.Append("         ''                                      AS kataban_check_div, ")
            sbSql.Append("         ''                                      AS parts_nm, ")
            sbSql.Append("         ''                                      AS model_nm, ")
            sbSql.Append("         ISNULL(NM.series_nm,DF.series_nm)       AS disp_name, ")
            sbSql.Append("         KT.price_no, ")
            sbSql.Append("         KT.spec_no, ")
            sbSql.Append("         KT.currency_cd, ")
            sbSql.Append("         KT.country_cd ")
            sbSql.Append("  FROM   kh_series_kataban  KT ")
            sbSql.Append("    INNER JOIN kh_series_nm_mst   DF ")
            sbSql.Append("      ON  KT.series_kataban       = DF.series_kataban ")
            sbSql.Append("      AND KT.key_kataban          = DF.key_kataban ")
            sbSql.Append("    LEFT OUTER JOIN kh_series_nm_mst   NM ")
            sbSql.Append("      ON  KT.series_kataban       = NM.series_kataban ")
            sbSql.Append("      AND KT.key_kataban          = NM.key_kataban ")
            sbSql.Append("      AND NM.language_cd          = @LangCd ")
            sbSql.Append("      AND NM.in_effective_date   <= @StandardDate ")
            sbSql.Append("      AND NM.out_effective_date   > @StandardDate ")
            sbSql.Append("  LEFT OUTER JOIN kh_country_group_mst   DI ")
            sbSql.Append("    ON  KT.country_group_cd      = DI.country_group_cd ")
            sbSql.Append("  WHERE    KT.in_effective_date    <= @StandardDate ")
            sbSql.Append("  AND    KT.out_effective_date    > @StandardDate ")
            sbSql.Append("  AND    DF.language_cd           = @DefaultLangCd ")
            sbSql.Append("  AND    DF.in_effective_date    <= @StandardDate ")
            sbSql.Append("  AND    DF.out_effective_date    > @StandardDate ")
            sbSql.Append("AND      KT.series_kataban      LIKE @SrsKataban ")
            sbSql.Append("AND     (KT.country_group_cd = 'ALL'  ")
            sbSql.Append("OR       DI.country_cd     = @country_cd ) ")

            'CHANGED BY YGY 20150121  RM1501051  ↓↓↓↓↓↓
            '台湾のAX専門の特定取引代理店特殊対応
            If lstWhereSeries.Count > 0 Then
                Dim strWhereSeries As String = " AND (KT.series_kataban IN ("

                For inti As Integer = 0 To lstWhereSeries.Count - 1
                    If inti = 0 Then
                        strWhereSeries &= " '" & lstWhereSeries(inti) & "' "
                    Else
                        strWhereSeries &= " ,'" & lstWhereSeries(inti) & "' "
                    End If
                Next

                strWhereSeries &= "))"
                sbSql.Append(strWhereSeries)
            End If
            'CHANGED BY YGY 20150121  RM1501051  ↑↑↑↑↑↑

            sbSql.Append("UNION ALL ")
            sbSql.Append("  SELECT '2_' + ")
            sbSql.Append("         PRC.kataban + '_' + ")
            sbSql.Append("         CONVERT(VARCHAR, PRC.in_effective_date, 120)  AS sortKey, ")
            sbSql.Append("         ' '            AS series_kataban, ")
            sbSql.Append("         ' '            AS key_kataban, ")
            sbSql.Append("         PRC.kataban    AS disp_kataban, ")
            sbSql.Append("         '2'            AS division, ")
            sbSql.Append("         PRC.kataban_check_div, ")
            sbSql.Append("         PNM.parts_nm, ")
            sbSql.Append("         PNM.model_nm, ")
            sbSql.Append("         ''    AS disp_name, ")
            sbSql.Append("         ''      AS price_no, ")
            sbSql.Append("         ''      AS spec_no, ")
            sbSql.Append("         PRC.currency_cd AS currency_cd, ")
            sbSql.Append("         PRC.country_cd AS country_cd ")
            sbSql.Append("  FROM   kh_price  PRC ")
            sbSql.Append("    LEFT OUTER JOIN kh_parts_nm_mst   PNM ")
            sbSql.Append("      ON  PRC.kataban              = PNM.kataban ")
            sbSql.Append("      AND PNM.language_cd          = @LangCd ")
            sbSql.Append("      AND PNM.in_effective_date   <= @StandardDate ")
            sbSql.Append("      AND PNM.out_effective_date   > @StandardDate ")

            'ADD BY YGY 20150311    海外仕入れ追加のため
            sbSql.Append("    AND PRC.currency_cd          = PNM.currency_cd ")

            sbSql.Append("  LEFT OUTER JOIN kh_country_group_mst   DI ")
            sbSql.Append("    ON  PRC.country_group_cd      = DI.country_group_cd ")
            sbSql.Append("  WHERE  PRC.kataban            LIKE @SrsKataban ")
            sbSql.Append("  AND    PRC.in_effective_date    <= @StandardDate ")
            sbSql.Append("  AND    PRC.out_effective_date    > @StandardDate ")
            sbSql.Append("AND    ( PRC.country_group_cd = 'ALL'  ")
            sbSql.Append("OR     DI.country_cd     = @country_cd ) ")

            'CHANGED BY YGY 20150121  RM1501051  ↓↓↓↓↓↓
            '台湾のAX専門の特定取引代理店特殊対応
            If lstWhereSeries.Count > 0 Then
                Dim strWhereSeries As String = " AND ((PRC.kataban LIKE 'AX%') OR ("

                For inti As Integer = 0 To lstWhereSeries.Count - 1
                    If inti = 0 Then
                        strWhereSeries &= " PRC.kataban LIKE '" & lstWhereSeries(inti) & "%' "
                    Else
                        strWhereSeries &= " OR PRC.kataban LIKE '" & lstWhereSeries(inti) & "%' "
                    End If
                Next

                strWhereSeries &= "))"
                sbSql.Append(strWhereSeries)
            End If
            'CHANGED BY YGY 20150121  RM1501051  ↑↑↑↑↑↑

            sbSql.Append(") AS Kata ")

            If strMinFlg <> "0" Then
                Dim str() As String = seriesSearch.strMinKatabanValue.ToString.Split("_")
                If str.Length = 3 Then
                    If str(0).ToString.Trim = "2" Then
                        sbSql.Append("WHERE  sortKey > @MinKataban ")
                    ElseIf str(2).ToString.Trim.Length <= 0 Then
                        sbSql.Append("WHERE  series_kataban > @MinSeries OR series_kataban =' ' ")
                    Else
                        sbSql.Append("WHERE  ((series_kataban = @MinSeries AND key_kataban > @MinKey)")
                        sbSql.Append("OR  (series_kataban > @MinSeries)) OR  series_kataban =' ' ")
                    End If
                End If
            End If
            sbSql.Append("  ORDER BY 1,2,3 ")

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString
            objCmd.Parameters.Add("@SrsKataban", SqlDbType.VarChar, 60).Value = seriesSearch.strSrsKataValue & "%"
            objCmd.Parameters.Add("@LangCd", SqlDbType.VarChar, 2).Value = seriesSearch.strLangCdValue
            objCmd.Parameters.Add("@DefaultLangCd", SqlDbType.VarChar, 2).Value = CdCst.LanguageCd.DefaultLang
            If strMinFlg <> "0" Then
                Dim str() As String = seriesSearch.strMinKatabanValue.ToString.Split("_")
                If str.Length = 3 Then
                    If str(0).ToString.Trim = "2" Then
                        objCmd.Parameters.Add("@MinKataban", SqlDbType.VarChar, 62).Value = seriesSearch.strMinKatabanValue
                    ElseIf str(2).ToString.Trim.Length > 0 Then
                        objCmd.Parameters.Add("@MinSeries", SqlDbType.VarChar, 60).Value = str(1).ToString
                        objCmd.Parameters.Add("@MinKey", SqlDbType.VarChar, 1).Value = str(2).ToString
                    Else
                        objCmd.Parameters.Add("@MinSeries", SqlDbType.VarChar, 60).Value = str(1).ToString
                    End If
                End If
            End If
            objCmd.Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            objCmd.Parameters.Add("@country_cd", SqlDbType.Char, 3).Value = strCountrycd

            objAdp = New SqlDataAdapter()
            objAdp.SelectCommand = objCmd
            objAdp.Fill(seriesSearch.dsKatabanValue, "KatabanTbl")

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return seriesSearch.dsKatabanValue
    End Function

    ''' <summary>
    ''' 引当シリーズ形番追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strSrsKataban">シリーズ形番</param>
    ''' <param name="strKeyKataban">キー形番</param>
    ''' <param name="strGoodsNm">商品名</param>
    ''' <remarks></remarks>
    Public Sub subInsertSelSrsKtbnMdl(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                      ByVal strSessionId As String, ByVal strSrsKataban As String, _
                                      ByVal strKeyKataban As String, ByVal strGoodsNm As String, _
                                      ByVal strCurrencyCd As String)
        Dim objCmd As SqlCommand = Nothing
        Try
            objCmd = objCon.CreateCommand

            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHSelSrsKtbnMdlIns
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
                .Parameters.Add("@series_kataban", SqlDbType.VarChar, 10).Value = strSrsKataban
                .Parameters.Add("@key_kataban", SqlDbType.VarChar, 2).Value = strKeyKataban
                .Parameters.Add("@goods_nm", SqlDbType.NVarChar, 150).Value = strGoodsNm
                'ADD BY YGY 20150311    海外仕入れ追加のため
                .Parameters.Add("@currency_cd", SqlDbType.VarChar, 3).Value = strCurrencyCd
            End With
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 引当シリーズ形番追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strFullKataban">フル形番</param>
    ''' <param name="strGoodsNm">商品名</param>
    ''' <remarks></remarks>
    Public Sub subInsertSelSrsKtbnFull(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                    ByVal strSessionId As String, ByVal strFullKataban As String, _
                                    ByVal strGoodsNm As String, ByVal strCurrencyCd As String)
        Dim objCmd As SqlCommand = Nothing

        Try
            objCmd = objCon.CreateCommand
            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHSelSrsKtbnFullIns
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
                .Parameters.Add("@full_kataban", SqlDbType.VarChar, 60).Value = strFullKataban
                .Parameters.Add("@goods_nm", SqlDbType.NVarChar, 150).Value = strGoodsNm
                'ADD BY YGY 20150311    海外仕入れ追加のため
                .Parameters.Add("@currency_cd", SqlDbType.VarChar, 3).Value = strCurrencyCd
            End With
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 引当シリーズ形番追加処理（仕入品）
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strFullKataban">フル形番</param>
    ''' <param name="strGoodsNm">商品名</param>
    ''' <remarks></remarks>
    Public Sub subInsertSelSrsKtbnShiire(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                    ByVal strSessionId As String, ByVal strFullKataban As String, _
                                    ByVal strGoodsNm As String, ByVal strCurrencyCd As String)
        Dim objCmd As SqlCommand = Nothing

        Try
            objCmd = objCon.CreateCommand
            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHSelSrsKtbnShiireIns
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
                .Parameters.Add("@full_kataban", SqlDbType.VarChar, 60).Value = strFullKataban
                .Parameters.Add("@goods_nm", SqlDbType.NVarChar, 150).Value = strGoodsNm
                .Parameters.Add("@currency_cd", SqlDbType.VarChar, 3).Value = strCurrencyCd
            End With
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 引当シリーズ形番削除処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <remarks></remarks>
    Public Sub subDeleteSelKtbnInfo(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                          ByVal strSessionId As String)
        Dim objCmd As New SqlCommand
        Try
            objCmd = objCon.CreateCommand

            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHSelKtbnInfoDel
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
            End With
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try
    End Sub
End Class
