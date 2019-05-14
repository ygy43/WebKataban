Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class MasterDAL

    ''' <summary>
    ''' ユーザー情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strDate"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_UserMstList(objCon As SqlConnection, strUserID As String, _
                                              strDate As String, strLanguage As String, _
                                              intStartIndex As Integer, intEndIndex As Integer, _
                                              likeflg As Boolean) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append(" SELECT * FROM ( ")
            'ページング
            sbSql.Append(" SELECT ROW_NUMBER() OVER (ORDER BY a.user_id, a.seq_no) AS rownum, ")
            sbSql.Append("     a.user_id, ")
            sbSql.Append("     a.seq_no, ")
            sbSql.Append("     convert(VARCHAR,a.in_effective_date,111) as in_effective_date, ")
            sbSql.Append("     convert(VARCHAR,a.out_effective_date,111) as out_effective_date, ")
            sbSql.Append("     a.user_nm, ")
            sbSql.Append("     a.country_cd, ")
            sbSql.Append("     a.office_cd, ")
            sbSql.Append("     a.person_cd, ")
            sbSql.Append("     a.mail_address, ")
            sbSql.Append("     a.password, ")
            sbSql.Append("     convert(VARCHAR,a.password_exp_date,111) as password_exp_date, ")
            sbSql.Append("     a.user_class, ")
            sbSql.Append("     a.price_disp_lvl, ")
            sbSql.Append("     a.add_information_lvl, ")
            sbSql.Append("     a.use_function_lvl, ")
            sbSql.Append("     a.register_person, ")
            sbSql.Append("     a.register_datetime, ")
            sbSql.Append("     a.current_person, ")
            sbSql.Append("     a.current_datetime, ")
            sbSql.Append("     c.user_class_nm as user_class_nm_def, ")
            sbSql.Append("     d.user_class_nm as user_class_nm_sel ")
            sbSql.Append(" FROM ")
            sbSql.Append("     kh_user_mst a ")
            sbSql.Append("     INNER JOIN kh_user_cls_mst    b ")
            sbSql.Append("     ON a.user_class = b.user_class ")
            sbSql.Append("     INNER JOIN kh_user_cls_nm_mst c ")
            sbSql.Append("     ON c.user_class = b.user_class ")
            sbSql.Append("     AND c.language_cd = @DefaultLang ")
            sbSql.Append("     LEFT JOIN kh_user_cls_nm_mst d ")
            sbSql.Append("     ON  d.user_class = b.user_class ")
            sbSql.Append("     AND d.language_cd = @SelectLang ")
            sbSql.Append(" WHERE a.user_id like @UserId ")

            If Not strDate.Trim = String.Empty Then
                sbSql.Append(" AND   a.in_effective_date <= @StdDate ")
                sbSql.Append(" AND   a.out_effective_date > @StdDate ")
            End If

            'ページング
            sbSql.Append(" ) AS T ")
            sbSql.AppendLine(" WHERE T.rownum >= @StartIndex ")
            sbSql.AppendLine(" AND T.rownum <= @EndIndex ")

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString
            objCmd.Parameters.Add("@DefaultLang", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
            objCmd.Parameters.Add("@SelectLang", SqlDbType.Char, 2).Value = strLanguage
            If likeflg Then
                objCmd.Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID.Trim & "%"
            Else
                objCmd.Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID.Trim
            End If

            If Not strDate.Trim = String.Empty Then
                objCmd.Parameters.Add("@StdDate", SqlDbType.VarChar, 10).Value = strDate.Trim
            End If

            'ページング
            objCmd.Parameters.Add("@StartIndex", SqlDbType.Int).Value = intStartIndex
            objCmd.Parameters.Add("@EndIndex", SqlDbType.Int).Value = intEndIndex

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)
        Catch ex As Exception
            fncSQL_UserMstList = Nothing
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
    ''' ユーザー情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strDate"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_UserMstCount(objCon As SqlConnection, strUserID As String, _
                                              strDate As String, strLanguage As String) As Integer
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable
        Dim intResult As Integer = 0

        Try
            sbSql.Append(" SELECT COUNT(*) ")
            sbSql.Append(" FROM ")
            sbSql.Append("     kh_user_mst a ")
            sbSql.Append("     INNER JOIN kh_user_cls_mst    b ")
            sbSql.Append("     ON a.user_class = b.user_class ")
            sbSql.Append("     INNER JOIN kh_user_cls_nm_mst c ")
            sbSql.Append("     ON c.user_class = b.user_class ")
            sbSql.Append("     AND c.language_cd = @DefaultLang ")
            sbSql.Append("     LEFT JOIN kh_user_cls_nm_mst d ")
            sbSql.Append("     ON  d.user_class = b.user_class ")
            sbSql.Append("     AND d.language_cd = @SelectLang ")
            sbSql.Append(" WHERE a.user_id like @UserId ")

            If Not strDate.Trim = String.Empty Then
                sbSql.Append(" AND   a.in_effective_date <= @StdDate ")
                sbSql.Append(" AND   a.out_effective_date > @StdDate ")
            End If
            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString
            objCmd.Parameters.Add("@DefaultLang", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
            objCmd.Parameters.Add("@SelectLang", SqlDbType.Char, 2).Value = strLanguage
            objCmd.Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID.Trim & "%"
            If Not strDate.Trim = String.Empty Then
                objCmd.Parameters.Add("@StdDate", SqlDbType.VarChar, 10).Value = strDate.Trim
            End If

            intResult = objCmd.ExecuteScalar

        Catch ex As Exception
            fncSQL_UserMstCount = 0
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return intResult
    End Function

    ''' <summary>
    ''' 国別生産品情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_CountryItemMstList(objCon As SqlConnection, _
                                                     strCountryCd As String, strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append("     SELECT ")
            sbSql.Append("         kataban, ")
            sbSql.Append("         country_cd, ")
            sbSql.Append("         seq_no, ")
            sbSql.Append("         convert(VARCHAR,in_effective_date,111) as in_effective_date, ")
            sbSql.Append("         convert(VARCHAR,out_effective_date,111) as out_effective_date, ")
            sbSql.Append("         register_person, ")
            sbSql.Append("         register_datetime, ")
            sbSql.Append("         current_person, ")
            sbSql.Append("         current_datetime ")
            sbSql.Append("     FROM ")
            sbSql.Append("         kh_country_item_mst ")
            If Not strKataban.Trim = "" Then sbSql.Append(" WHERE kataban like @Kataban ")
            sbSql.Append(" AND   country_cd = @CountryCd ")
            sbSql.Append("     ORDER BY ")
            sbSql.Append("         kataban, ")
            sbSql.Append("         country_cd, ")
            sbSql.Append("         seq_no ")

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString
            If Not strKataban.Trim = "" Then
                objCmd.Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban.Trim & "%"
            End If
            objCmd.Parameters.Add("@CountryCd", SqlDbType.Char, 3).Value = strCountryCd

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)
        Catch ex As Exception
            fncSQL_CountryItemMstList = Nothing
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
    ''' 現地定価
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_RateMstList_L(objCon As SqlConnection, _
                                                strCountryCd As String, strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append("     SELECT ")
            sbSql.Append("         country_cd, ")
            sbSql.Append("         rate_search_key, ")
            sbSql.Append("         seq_no, ")
            sbSql.Append("         list_price_rate1, ")
            sbSql.Append("         list_price_rate2, ")
            sbSql.Append("         convert(VARCHAR,in_effective_date,111) as in_effective_date, ")
            sbSql.Append("         convert(VARCHAR,out_effective_date,111) as out_effective_date ")
            sbSql.Append("     FROM ")
            sbSql.Append("         kh_country_rate_localprice_mst ")
            'sbSql.Append(" Left Join    kh_country_mst on ")
            'sbSql.Append(" kh_country_rate_localprice_mst.country_cd = kh_country_mst.country_cd ")
            If Not strCountryCd.Trim = "" Then
                sbSql.Append(" WHERE   kh_country_rate_localprice_mst.country_cd = @country_cd ")
            End If
            If Not strKataban.Trim = "" And strCountryCd.Trim = "" Then
                sbSql.Append(" WHERE rate_search_key like @Kataban ")
            ElseIf Not strKataban.Trim = "" And Not strCountryCd.Trim = "" Then
                sbSql.Append(" AND rate_search_key like @Kataban ")
            End If
            sbSql.Append("     ORDER BY ")
            sbSql.Append("         rate_search_key, ")
            sbSql.Append("         kh_country_rate_localprice_mst.country_cd, ")
            sbSql.Append("         seq_no ")

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString
            If Not strKataban.Trim = "" Then
                objCmd.Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban.Trim & "%"
            End If
            If Not strCountryCd.Trim = "" Then
                objCmd.Parameters.Add("@country_cd", SqlDbType.Char, 3).Value = strCountryCd
            End If

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)
        Catch ex As Exception
            fncSQL_RateMstList_L = Nothing
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
    ''' 購入価格
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strMadeCountryCd"></param>
    ''' <param name="strSaleCountryCd"></param>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_RateMstList_N(objCon As SqlConnection, strMadeCountryCd As String, _
                                                strSaleCountryCd As String, strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append("     SELECT ")
            sbSql.Append("         exp_country_cd,imp_country_cd, ")
            sbSql.Append("         rate_search_key, ")
            sbSql.Append("         seq_no, ")
            sbSql.Append("         fob_rate, ")
            sbSql.Append("         convert(VARCHAR,in_effective_date,111) as in_effective_date, ")
            sbSql.Append("         convert(VARCHAR,out_effective_date,111) as out_effective_date ")
            sbSql.Append("     FROM ")
            sbSql.Append("         kh_country_rate_netprice_mst ")

            If Not strMadeCountryCd.Trim = "" Then sbSql.Append(" WHERE   exp_country_cd = @Made_country ")
            If strMadeCountryCd.Trim.Length > 0 Then
                If Not strSaleCountryCd.Trim = "" Then sbSql.Append(" AND   imp_country_cd = @Sale_country ")
            Else
                If Not strSaleCountryCd.Trim = "" Then sbSql.Append(" WHERE   imp_country_cd = @Sale_country ")
            End If

            If strMadeCountryCd.Trim.Length > 0 Or strSaleCountryCd.Trim.Length > 0 Then
                If Not strKataban.Trim = "" Then sbSql.Append(" AND   rate_search_key LIKE @Kataban ")
            ElseIf strMadeCountryCd.Trim.Length <= 0 AndAlso strSaleCountryCd.Trim.Length <= 0 Then
                If Not strKataban.Trim = "" Then sbSql.Append(" WHERE   rate_search_key LIKE @Kataban ")
            End If
            sbSql.Append("     ORDER BY ")
            sbSql.Append("         rate_search_key, ")
            sbSql.Append("         exp_country_cd, ")
            sbSql.Append("         seq_no ")

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString
            If Not strKataban.Trim = "" Then
                objCmd.Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban.Trim & "%"
            End If
            If Not strMadeCountryCd.Trim = "" Then
                objCmd.Parameters.Add("@Made_country", SqlDbType.Char, 3).Value = strMadeCountryCd
            End If
            If Not strSaleCountryCd.Trim = "" Then
                objCmd.Parameters.Add("@Sale_country", SqlDbType.Char, 3).Value = strSaleCountryCd
            End If

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dtResult)
        Catch ex As Exception
            fncSQL_RateMstList_N = Nothing
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
    ''' 国マスタリストの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_CountryMst(objCon As SqlConnection) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append(" SELECT  ' ' AS country_cd, ' ' AS base_cd ")
            sbSql.Append(" FROM    kh_country_mst ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT  country_cd, base_cd ")
            sbSql.Append(" FROM    kh_country_mst ")
            sbSql.Append(" ORDER BY  country_cd ")

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString

            objAdp = New SqlDataAdapter(objCmd)
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
    ''' 国マスタの取得(Sort_no順)
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetAllCountryMst(objCon As SqlConnection) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append(" SELECT  country_cd, ")
            sbSql.Append("         country_nm_omi_ja, ")
            sbSql.Append("         country_nm_omi_en, ")
            sbSql.Append("         currency_cd ")
            sbSql.Append(" FROM    kh_country_mst ")
            sbSql.Append(" ORDER BY  country_cd ")

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString

            objAdp = New SqlDataAdapter(objCmd)
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
    ''' 営業所情報の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_OfficeMst(objCon As SqlConnection) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append(" SELECT  ' ' AS office_cd ")
            sbSql.Append(" FROM    kh_office_mst ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT  office_cd ")
            sbSql.Append(" FROM    kh_office_mst ")
            sbSql.Append(" ORDER BY  office_cd ")

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString

            objAdp = New SqlDataAdapter(objCmd)
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
    ''' ユーザクラスの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_UserClassMst(objCon As SqlConnection, strLanguage As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  ' ' AS user_class, ")
            sbSql.Append("         ' ' AS user_class_nm ")
            sbSql.Append(" FROM    kh_user_cls_nm_mst ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT b.user_class, ")
            sbSql.Append("        a.user_class_nm ")
            sbSql.Append(" from kh_user_cls_nm_mst a, ")
            sbSql.Append("      kh_user_cls_mst b ")
            sbSql.Append(" where a.user_class = b.user_class ")
            If strLanguage = CdCst.LanguageCd.Japanese Then
                sbSql.Append(" and a.language_cd = '" & CdCst.LanguageCd.Japanese & "' ")
            Else
                sbSql.Append(" and a.language_cd = '" & CdCst.LanguageCd.DefaultLang & "' ")
            End If

            objCmd = objCon.CreateCommand
            objCmd.CommandText = sbSql.ToString

            objAdp = New SqlDataAdapter(objCmd)
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
    ''' ドロップダウンリスト作成
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_CountryCodeList(objCon As SqlConnection, strLanguage As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try

            'SQL Query生成
            If strLanguage = CdCst.LanguageCd.Japanese Then
                sbSql.Append(" SELECT  ' ' AS country_cd, ")
                sbSql.Append("         ' ' AS country_nm ")
                sbSql.Append(" FROM    kh_country_mst ")
                sbSql.Append(" UNION ")
                sbSql.Append(" SELECT  country_cd, ")
                sbSql.Append("         country_nm_omi_ja AS country_nm ")
                sbSql.Append(" FROM    kh_country_mst ")
                sbSql.Append(" ORDER BY  country_cd ")
            Else
                sbSql.Append(" SELECT  ' ' AS country_cd, ")
                sbSql.Append("         ' ' AS country_nm ")
                sbSql.Append(" FROM    kh_country_mst ")
                sbSql.Append(" UNION ")
                sbSql.Append(" SELECT  country_cd, ")
                sbSql.Append("         country_nm_omi_en AS country_nm ")
                sbSql.Append(" FROM    kh_country_mst ")
                sbSql.Append(" ORDER BY  country_cd ")
            End If

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            objCmd.CommandType = CommandType.Text
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
    ''' 国コードによりベースコードの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strCountryCod"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetBaseCdByCountryCd(objCon As SqlConnection, strCountryCod As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            sbSql.Append(" SELECT  base_cd ")
            sbSql.Append(" FROM    kh_country_mst ")
            sbSql.Append(" WHERE   country_cd = '" & strCountryCod & "' ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            objCmd.CommandType = CommandType.Text
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
