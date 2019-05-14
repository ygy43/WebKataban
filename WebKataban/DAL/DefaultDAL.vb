Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class DefaultDAL

    ''' <summary>
    ''' メニュー情報取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strLangCd"></param>
    ''' <param name="strAuthorityCd"></param>
    ''' <param name="strMenuId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncMenuMstSelect(ByVal objConBase As SqlConnection, ByVal strLangCd As String, _
                                            ByVal strAuthorityCd As String, _
                                            Optional ByVal strMenuId As String = Nothing) As DataTable
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        fncMenuMstSelect = New DataTable
        Try
            'SQL Query生成
            sbSql.Append(" SELECT  a.menu_id, ")
            sbSql.Append("         a.use_function_lvl, ")
            sbSql.Append("         ISNULL(c.menu_nm, b.menu_nm) AS menu_nm ")
            sbSql.Append(" FROM    kh_menu_mst  a ")
            sbSql.Append(" INNER JOIN  kh_menu_nm_mst  b ")
            sbSql.Append(" ON      a.menu_id      = b.menu_id ")
            sbSql.Append(" AND     b.language_cd  = @DefLanguageCd ")
            sbSql.Append(" LEFT  JOIN  kh_menu_nm_mst  c ")
            sbSql.Append(" ON      a.menu_id      = c.menu_id ")
            sbSql.Append(" AND     c.language_cd  = @LanguageCd ")
            If strMenuId Is Nothing Then
                sbSql.Append(" WHERE   p_menu_id IS NULL ")
            Else
                sbSql.Append(" WHERE   p_menu_id  = @PMenuId ")
            End If
            sbSql.Append(" ORDER BY  a.disp_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@DefLanguageCd", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
                .Parameters.Add("@LanguageCd", SqlDbType.Char, 2).Value = strLangCd
                If Not strMenuId Is Nothing Then
                    .Parameters.Add("@PMenuId", SqlDbType.VarChar, 30).Value = strMenuId
                End If
            End With
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncMenuMstSelect)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Function

    ''' <summary>
    ''' ログイン日時によりログインデータの削除
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strDate"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteFromLoginByLoginDate(ByVal objConBase As SqlConnection, ByVal strDate As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt_del As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" DELETE FROM KH_LOGIN ")
            sbSql.Append(" WHERE login_datetime <= '" & strDate & "'")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt_del)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 登録日時によりデータの削除
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strDate"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteFromSelAccPrcStrcByRegDate(ByVal objConBase As SqlConnection, ByVal strDate As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt_del As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" DELETE FROM kh_sel_acc_prc_strc ")
            sbSql.Append(" WHERE register_datetime <= '" & strDate & "'")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt_del)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 登録日時によりデータの削除
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strDate"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteFromSelKtbnStrcByRegDate(ByVal objConBase As SqlConnection, ByVal strDate As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt_del As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" DELETE FROM kh_sel_ktbn_strc ")
            sbSql.Append(" WHERE register_datetime <= '" & strDate & "'")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt_del)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 登録日時によりデータの削除
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strDate"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteFromSelOutofopOrderByRegDate(ByVal objConBase As SqlConnection, ByVal strDate As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt_del As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" DELETE FROM kh_sel_outofop_order ")
            sbSql.Append(" WHERE register_datetime <= '" & strDate & "'")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt_del)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 登録日時によりデータの削除
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strDate"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteFromSelRodEndOrderByRegDate(ByVal objConBase As SqlConnection, ByVal strDate As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt_del As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" DELETE FROM kh_sel_rod_end_order ")
            sbSql.Append(" WHERE register_datetime <= '" & strDate & "'")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt_del)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 登録日時によりデータの削除
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strDate"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteFromSelSpecByRegDate(ByVal objConBase As SqlConnection, ByVal strDate As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt_del As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" DELETE FROM kh_sel_spec ")
            sbSql.Append(" WHERE register_datetime <= '" & strDate & "'")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt_del)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 登録日時によりデータの削除
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strDate"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteFromSelSpecStrcByRegDate(ByVal objConBase As SqlConnection, ByVal strDate As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt_del As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" DELETE FROM kh_sel_spec_strc ")
            sbSql.Append(" WHERE register_datetime <= '" & strDate & "'")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt_del)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 登録日時によりデータの削除
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strDate"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteFromSelSrsKtbnByRegDate(ByVal objConBase As SqlConnection, ByVal strDate As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt_del As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" DELETE FROM kh_sel_srs_ktbn ")
            sbSql.Append(" WHERE register_datetime <= '" & strDate & "'")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt_del)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 言語選択欄の内容を取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectLanguageList(ByVal objConBase As SqlConnection, ByVal strLang As String) As DataTable
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing

        fncSelectLanguageList = New DataTable
        Try
            objCmd = objConBase.CreateCommand
            sbSql.Append("SELECT DISTINCT ")
            sbSql.Append("       ISNULL(NM.disp_language_cd, DF.disp_language_cd) AS language_cd, ")
            sbSql.Append("       ISNULL(NM.language_nm, DF.language_nm) AS language_nm ")
            sbSql.Append("FROM (SELECT b.disp_language_cd,b.language_nm ")
            sbSql.Append("      FROM sales.kh_language_mst AS a,sales.kh_language_nm_mst AS b ")
            sbSql.Append("      WHERE a.language_cd = b.language_cd ")
            sbSql.Append("      AND   b.language_cd = @DefaultLangCd) AS DF, ")
            sbSql.Append("      sales.kh_language_mst AS LA ")
            sbSql.Append("      LEFT OUTER JOIN ")
            sbSql.Append("           sales.kh_language_nm_mst AS NM ")
            sbSql.Append("      ON LA.language_cd = NM.disp_language_cd ")
            sbSql.Append("      AND NM.language_cd = @LangCd ")
            sbSql.Append("ORDER BY  language_cd ")

            objCmd.CommandText = sbSql.ToString
            If Len(Trim(strLang)) = 0 Then
                objCmd.Parameters.Add("@LangCd", SqlDbType.VarChar, 3).Value = CdCst.LanguageCd.DefaultLang
            Else
                objCmd.Parameters.Add("@LangCd", SqlDbType.VarChar, 3).Value = strLang
            End If
            objCmd.Parameters.Add("@DefaultLangCd", SqlDbType.VarChar, 3).Value = CdCst.LanguageCd.DefaultLang

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncSelectLanguageList)
        Catch ex As Exception
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Function

End Class
