Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class OutOfOptionDAL

    ''' <summary>
    ''' 引当口径検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncBoreSizeSelect(ByVal objCon As SqlConnection, ByVal strSeriesKataban As String, ByVal strKeyKataban As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL作成
            sbSql.Append(" SELECT  ktbn_strc_seq_no ")
            sbSql.Append(" FROM    kh_kataban_strc ")
            sbSql.Append(" WHERE   series_kataban  = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban     = @KeyKataban ")
            sbSql.Append(" AND     element_div = '" & CdCst.RodEndCstmOrder.EleBoreSize & "'")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
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
    ''' 画面表示情報検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strLang"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <param name="strBoreSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncOutofOpDataSelect(ByVal objCon As SqlConnection, ByVal strLang As String, _
                                          ByVal strSeriesKataban As String, ByVal strKeyKataban As String, _
                                          ByVal strBoreSize As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  A.pattern_key as pattern_key")
            sbSql.Append(" ,       A.pattern_seq_no as pattern_seq_no ")
            sbSql.Append(" ,       isnull(A.disp_value,'') as disp_value ")
            sbSql.Append(" ,       isnull(B.place_lvl,'1') as place_lvl ")
            sbSql.Append(" FROM    sales.kh_outofop_datainfo as A ")

            sbSql.Append(" LEFT OUTER JOIN sales.kh_outofop_detailplace As B")
            sbSql.Append("      ON B.series_kataban = A.series_kataban")
            sbSql.Append("     AND B.key_kataban = A.key_kataban")
            sbSql.Append("     AND B.bore_size = A.bore_size")
            sbSql.Append("     AND B.pattern_key = A.pattern_key")
            sbSql.Append("     AND B.pattern_seq_no = A.pattern_seq_no")
            sbSql.Append("     AND A.language_cd = @Languagecd ")

            sbSql.Append(" WHERE   A.language_cd     = @Languagecd ")
            sbSql.Append(" AND     A.series_kataban  = @SeriesKataban ")
            sbSql.Append(" AND     (A.key_kataban     = '%' ")
            sbSql.Append(" OR       A.key_kataban     = @KeyKataban)")
            sbSql.Append(" AND     (A.bore_size       = 999")
            sbSql.Append(" OR       A.bore_size       = @Boresize)")
            sbSql.Append(" ORDER BY  A.pattern_key ")
            sbSql.Append(" ,         A.pattern_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Languagecd", SqlDbType.Char, 2).Value = strLang
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
                .Parameters.Add("@Boresize", SqlDbType.Int).Value = strBoreSize
            End With
            '実行
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
    ''' 引当オプション外特注取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks>引当オプション外特注情報を取得し、メンバ変数にセットする</remarks>
    Public Function fncSelOutOfOpSelect(ByVal objCon As SqlConnection, ByVal strUserID As String, _
                                        ByVal strSessionID As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT      isnull(port_cushion,0) as port_cushion ")
            sbSql.Append("            ,isnull(port_cushion_place,'') as port_cushion_place ")
            sbSql.Append("            ,isnull(port,0) as port ")
            sbSql.Append("            ,isnull(port_size,0) as port_size ")
            sbSql.Append("            ,isnull(mounting,0) as mounting ")
            sbSql.Append("            ,isnull(trunnion,'') as trunnion ")
            sbSql.Append("            ,isnull(clevis,0) as clevis ")
            sbSql.Append("            ,isnull(tierod_radio,0) as tierod_radio ")
            sbSql.Append("            ,isnull(tierod_default,0) as tierod_default ")
            sbSql.Append("            ,isnull(tierod_custom,0) as tierod_custom ")
            sbSql.Append("            ,isnull(sus,0) as sus ")
            sbSql.Append("            ,isnull(jm,0) as jm ")
            sbSql.Append("            ,isnull(fluororub,0) as fluororub ")
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
    ''' 引当オプション外特注テーブル削除処理
    ''' </summary>
    ''' <param name="objCon">DB接続オブジェクト</param>
    ''' <returns></returns>
    ''' <remarks>引当オプション外特注テーブルからデータを削除する</remarks>
    Public Function fncSPSelOutOpDel(ByVal objCon As SqlConnection, ByVal strUserID As String, _
                                        ByVal strSessionID As String) As Boolean
        Dim objCmd As SqlCommand = Nothing
        fncSPSelOutOpDel = False

        Try
            objCmd = objCon.CreateCommand
            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHSelOutOfOpDel
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionID
            End With
            '実行
            objCmd.ExecuteNonQuery()
            fncSPSelOutOpDel = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then objCmd.Dispose()
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' フル形番生成
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strSessionID"></param>
    ''' <param name="strBoreSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncFullKtbnCreate(ByVal objCon As SqlConnection, ByVal strUserID As String, _
                                        ByVal strSessionID As String, ByVal strBoreSize As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT      isnull(a.port_cushion,0) as port_cushion ")
            sbSql.Append("            ,isnull(a.port_cushion_place,'') as port_cushion_place ")
            sbSql.Append("            ,isnull(a.port,0) as port ")
            sbSql.Append("            ,substring(tblPortKey.disp_value,1,charindex(':',tblPortKey.disp_value)-1) as K_port ")
            sbSql.Append("            ,isnull(a.port_size,0) as port_size ")
            sbSql.Append("            ,substring(tblPortSize.disp_value,1,charindex(':',tblPortSize.disp_value)-1) as K_portSize ")
            sbSql.Append("            ,isnull(a.mounting,0) as mounting ")
            sbSql.Append("            ,substring(tblMounting.disp_value,1,charindex(':',tblMounting.disp_value)-1) as K_mounting ")
            sbSql.Append("            ,isnull(a.trunnion,'') as trunnion ")
            sbSql.Append("            ,isnull(a.clevis,0) as clevis ")
            sbSql.Append("            ,substring(tblClevis.disp_value,1,charindex(':',tblClevis.disp_value)-1) as K_Clevis ")
            sbSql.Append("            ,isnull(a.tierod_radio,0) as tierod_radio ")
            sbSql.Append("            ,isnull(a.tierod_default,0) as tierod_default ")
            sbSql.Append("            ,isnull(a.tierod_custom,0) as tierod_custom ")
            sbSql.Append("            ,isnull(a.sus,0) as sus ")
            sbSql.Append("            ,isnull(a.jm,0) as jm ")
            sbSql.Append("            ,isnull(a.fluororub,0) as fluororub ")
            sbSql.Append(" FROM        kh_sel_outofop_order a ")
            sbSql.Append(" INNER JOIN  kh_sel_srs_ktbn b ")
            sbSql.Append(" ON          a.user_id               = b.user_id ")
            sbSql.Append(" AND         a.session_id            = b.session_id ")
            sbSql.Append(" LEFT JOIN  kh_outofop_datainfo tblPortKey ")
            sbSql.Append(" ON          b.series_kataban          =  tblPortKey.series_kataban ")
            sbSql.Append(" AND         b.key_kataban           like tblPortKey.key_kataban ")
            sbSql.Append(" AND         (tblPortKey.bore_size     =  @BoreSize ")
            sbSql.Append(" OR           tblPortKey.bore_size     =  '999') ")
            sbSql.Append(" AND         tblPortKey.pattern_key    =  1 ")
            sbSql.Append(" AND         a.port                    =  tblPortKey.pattern_seq_no ")
            sbSql.Append(" LEFT JOIN  kh_outofop_datainfo tblPortSize ")
            sbSql.Append(" ON          b.series_kataban          =  tblPortSize.series_kataban ")
            sbSql.Append(" AND         b.key_kataban           like tblPortSize.key_kataban ")
            sbSql.Append(" AND         (tblPortSize.bore_size     =  @BoreSize ")
            sbSql.Append(" OR           tblPortSize.bore_size     =  '999') ")
            sbSql.Append(" AND         tblPortSize.pattern_key    =  2 ")
            sbSql.Append(" AND         a.port_size               =  tblPortSize.pattern_seq_no ")
            sbSql.Append(" LEFT JOIN  kh_outofop_datainfo tblMounting ")
            sbSql.Append(" ON          b.series_kataban          =  tblMounting.series_kataban ")
            sbSql.Append(" AND         b.key_kataban           like tblMounting.key_kataban ")
            sbSql.Append(" AND         (tblMounting.bore_size     =  @BoreSize ")
            sbSql.Append(" OR           tblMounting.bore_size     =  '999') ")
            sbSql.Append(" AND         tblMounting.pattern_key    =  3 ")
            sbSql.Append(" AND         a.mounting                =  tblMounting.pattern_seq_no ")
            sbSql.Append(" LEFT JOIN  kh_outofop_datainfo tblClevis ")
            sbSql.Append(" ON          b.series_kataban          =  tblClevis.series_kataban ")
            sbSql.Append(" AND         b.key_kataban           like tblClevis.key_kataban ")
            sbSql.Append(" AND         (tblClevis.bore_size     =  @BoreSize ")
            sbSql.Append(" OR           tblClevis.bore_size     =  '999') ")
            sbSql.Append(" AND         tblClevis.pattern_key    =  5 ")
            sbSql.Append(" AND         a.clevis                  =  tblClevis.pattern_seq_no ")
            sbSql.Append(" WHERE       a.user_id    = @UserId ")
            sbSql.Append(" AND         a.session_id = @SessionId ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@BoreSize", SqlDbType.Int).Value = strBoreSize
                .Parameters.Add("@UserID", SqlDbType.VarChar, 10).Value = strUserID
                .Parameters.Add("@SessionID", SqlDbType.VarChar, 88).Value = strSessionID
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
    ''' 画面表示情報検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <param name="strSelectvalue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncOutofOpDataChack(ByVal objCon As SqlConnection, _
                                          ByVal strSeriesKataban As String, ByVal strKeyKataban As String, _
                                          ByVal strSelectvalue As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  place_lvl as place_lvl")
            sbSql.Append(" FROM    sales.kh_outofop_detailplace")

            sbSql.Append(" WHERE   series_kataban  = @SeriesKataban ")
            sbSql.Append(" AND     (key_kataban     = '%' ")
            sbSql.Append(" OR       key_kataban     = @KeyKataban)")
            sbSql.Append(" AND      pattern_key     = '99' ")
            sbSql.Append(" AND      pattern_seq_no  = @Selectvalue ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
                .Parameters.Add("@Selectvalue", SqlDbType.Int).Value = strSelectvalue
            End With
            '実行
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

End Class
