Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports WebKataban.CdCst

Public Class TankaISODAL

    ''' <summary>
    ''' 小数点区分取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncDecPointDivSelect(ByVal objCon As SqlConnection, strUserID As String) As String
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        fncDecPointDivSelect = String.Empty
        Try
            'SQL Query生成
            sbSql.Append(" SELECT  a.edit_div ")
            sbSql.Append(" FROM    kh_country_mst  a ")
            sbSql.Append(" INNER JOIN kh_user_mst  b ")
            sbSql.Append(" ON      a.country_cd = b.country_cd ")
            sbSql.Append(" WHERE   b.user_id = @UserId ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID
            End With
            objRdr = objCmd.ExecuteReader
            objRdr.Read()
            fncDecPointDivSelect = objRdr.GetValue(objRdr.GetOrdinal("edit_div"))
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
    ''' 選択情報の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strSession"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSQL_GetCompData(objCon As SqlConnection, strUserID As String, strSession As String) As DataTable
        Dim sbSql As New StringBuilder
        Dim objAdp As SqlDataAdapter = Nothing
        Dim objCmd As SqlCommand
        Dim dsSpecInfo As New DataSet
        fncSQL_GetCompData = New DataTable

        Try
            sbSql.Append(" SELECT ")
            sbSql.Append("     a.option_kataban, ")
            sbSql.Append("     a.quantity, ")
            sbSql.Append("     a.spec_strc_seq_no, ")
            sbSql.Append("     ISNULL(b.kataban , '') AS kataban, ")
            sbSql.Append("     ISNULL(b.kataban_check_div , '') AS kataban_check_div, ")
            sbSql.Append("     ISNULL(b.place_cd , '') AS place_cd, ")
            sbSql.Append("     ISNULL(b.ls_price , 0) AS ls_price, ")
            sbSql.Append("     ISNULL(b.rg_price , 0) AS rg_price, ")
            sbSql.Append("     ISNULL(b.ss_price , 0) AS ss_price, ")
            sbSql.Append("     ISNULL(b.bs_price , 0) AS bs_price, ")
            sbSql.Append("     ISNULL(b.gs_price , 0) AS gs_price, ")
            sbSql.Append("     ISNULL(b.ps_price , 0) AS ps_price, ")
            sbSql.Append("     ISNULL(b.amount , 0) AS amount ")
            sbSql.Append(" FROM ")
            sbSql.Append("           sales.kh_sel_spec_strc a ")
            sbSql.Append(" LEFT JOIN sales.kh_sel_acc_prc_strc b ")
            sbSql.Append(" ON    a.user_id          = b.user_id ")
            sbSql.Append(" AND   a.session_id       = b.session_id ")
            sbSql.Append(" AND   a.spec_strc_seq_no = b.disp_seq_no ")
            sbSql.Append(" WHERE a.user_id          = @UserId ")
            sbSql.Append(" AND   a.session_id       = @SessionId ")
            sbSql.Append(" AND   a.option_kataban  <> '' ")
            sbSql.Append(" AND   a.quantity         > 0 ")
            sbSql.Append(" ORDER BY a.spec_strc_seq_no ")

            objCmd = objCon.CreateCommand

            With objCmd
                .CommandText = sbSql.ToString
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 176).Value = strSession
            End With

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.SelectCommand = objCmd
            objAdp.Fill(dsSpecInfo, "SpecInfo")

            If Not dsSpecInfo Is Nothing Then
                fncSQL_GetCompData = dsSpecInfo.Tables("SpecInfo")
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
            fncSQL_GetCompData = Nothing
        Finally
            'DBオブジェクト破棄
            objCmd = Nothing
            sbSql = Nothing
            objAdp = Nothing
        End Try
    End Function
End Class
