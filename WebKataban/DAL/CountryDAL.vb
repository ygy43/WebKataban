Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class CountryDAL
    ''' <summary>
    ''' 国別生産品の対象国判定
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncCountryTradeGet(objConBase As SqlConnection, ByVal strCountryCd As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  disp_country_cd ")
            sbSql.Append(" FROM    kh_country_item_trade_mst ")
            sbSql.Append(" WHERE   country_cd          = @country_cd ")
            sbSql.Append(" ORDER BY country_cd, seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@country_cd", SqlDbType.Char, 3).Value = strCountryCd
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
    ''' フル形番の国コードの取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strFullKataban">形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncCountryItmMstChkP(objConBase As SqlConnection, ByVal strFullKataban As String, ByVal strCountry As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  kataban, country_cd ")
            sbSql.Append(" FROM    kh_country_item_mst ")
            sbSql.Append(" WHERE   kataban             = @kataban")
            sbSql.Append(" AND     country_cd          = @country_cd")
            sbSql.Append(" AND     in_effective_date  <= @time")
            sbSql.Append(" AND     out_effective_date  > @time")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@kataban", SqlDbType.VarChar, 60).Value = strFullKataban
                .Parameters.Add("@country_cd", SqlDbType.Char, 3).Value = strCountry
                .Parameters.Add("@time", SqlDbType.DateTime).Value = Now
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
    ''' 生産国名の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="intPlacelvl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetPlacelvlName(objCon As SqlConnection, ByVal intPlacelvl As Long) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  place_div,disp_seq_no,place_lvl ")
            sbSql.Append(" FROM    kh_place_div_mst ")
            sbSql.Append(" WHERE  place_lvl <= " & intPlacelvl)
            sbSql.Append(" ORDER BY  place_lvl DESC ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
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
    ''' 形番のすべての国コードを取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncCountryKeyGet(objCon As SqlConnection, ByVal strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            '機種形番取得
            Dim strMdlKataban As String = KHKataban.fncMdlKtbnGet(strKataban)

            'SQL Query生成
            sbSql.Append(" SELECT  country_cd,in_effective_date,out_effective_date ")
            sbSql.Append(" FROM    kh_country_item_key_mst ")
            sbSql.Append(" WHERE   item_search_key          = @item_search_key ")
            sbSql.Append(" AND     use_flag = 1 ")
            sbSql.Append(" ORDER BY seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@item_search_key", SqlDbType.Char, 30).Value = strMdlKataban
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
    ''' 出荷場所の表示名の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetCountryName(objCon As SqlConnection) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append("  SELECT  country_nm,country_cd,language_cd ")
            sbSql.Append("  FROM    kh_country_nm_mst ")
            sbSql.Append("  WHERE   language_cd = 'en' OR language_cd ='ja' ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

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
    ''' 出荷場所の表示名の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetAllCountryName(objCon As SqlConnection) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append("  SELECT  country_nm,country_cd,language_cd ")
            sbSql.Append("  FROM    kh_country_nm_mst ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

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
    ''' 出荷場所変更情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncPlaceChangeInfo(objCon As SqlConnection, ByVal strKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  search_div, ")
            sbSql.Append("         place_cd, ")
            sbSql.Append("         evaluation_type, ")
            sbSql.Append("         search_div ")
            sbSql.Append(" FROM    kh_place_change ")
            sbSql.Append(" WHERE   kataban             = @Kataban ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     search_div          = '1' ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT  search_div, ")
            sbSql.Append("         place_cd, ")
            sbSql.Append("         evaluation_type, ")
            sbSql.Append("         search_div ")
            sbSql.Append(" FROM    kh_place_change ")
            sbSql.Append(" WHERE   @Kataban      LIKE kataban ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     search_div          = '2' ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT  search_div, ")
            sbSql.Append("         place_cd, ")
            sbSql.Append("         evaluation_type, ")
            sbSql.Append("         search_div ")
            sbSql.Append(" FROM    kh_place_change ")
            sbSql.Append(" WHERE   kataban             = @Kataban ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     search_div          = '4' ")

            '  sbSql.Append(" ORDER BY  Search_Div ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)


            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Kataban", SqlDbType.VarChar, 60).Value = strKataban
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
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
    ''' 出荷場所変更情報取得処理（GLC在庫品）
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncStockPlaceInfo(objCon As SqlConnection, ByVal strKataban As String, ByVal strPlaceCd As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append("  Select stock_place_cd,evaluation_type,storage_Location,message_type ")
            sbSql.Append("  FROM kh_stock ")
            sbSql.Append(" WHERE (kataban = @FullKata) AND storage_Location = 'G000'")
            sbSql.Append(" AND (in_effective_date <= @StandardDate) AND (out_effective_date > @StandardDate) ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@FullKata", SqlDbType.VarChar, 60).Value = strKataban
                .Parameters.Add("@StockPlace", SqlDbType.VarChar, 60).Value = strPlaceCd
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
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
End Class
