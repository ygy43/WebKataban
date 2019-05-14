Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class MenuDAL

    ''' <summary>
    ''' 通知情報の取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strSelLang">言語</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectInformation(ByVal objConBase As SqlConnection, ByVal strSelLang As String) As DataTable
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        fncSelectInformation = New DataTable

        Try
            sbSql.Append(" SELECT  language_cd, ")
            sbSql.Append(" 		message, ")
            sbSql.Append("         seq_no ")
            sbSql.Append(" FROM    kh_Information ")
            sbSql.Append(" WHERE   language_cd         = @LanguageCd ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT  language_cd, ")
            sbSql.Append(" 		message, ")
            sbSql.Append("         seq_no ")
            sbSql.Append(" FROM    kh_Information ")
            sbSql.Append(" WHERE   language_cd         = @DefLanguageCd ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" order by language_cd,seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objConBase)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@LanguageCd", SqlDbType.Char, 2).Value = strSelLang
                .Parameters.Add("@DefLanguageCd", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            End With
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncSelectInformation)

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
End Class
