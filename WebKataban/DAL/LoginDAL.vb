Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class LoginDAL

    ''' <summary>
    ''' ログインユーザー情報の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectUserInfo(objConBase As SqlConnection, _
                                ByVal strUserid As String, ByVal strPasswd As String, _
                                ByVal strSelLang As String, Optional ByVal bolOverWrite As Boolean = False) As DataTable
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        fncSelectUserInfo = New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  a.user_id, ")
            sbSql.Append("         b.base_cd, ")
            sbSql.Append("         a.country_cd, ")
            sbSql.Append("         a.office_cd, ")
            sbSql.Append("         a.person_cd, ")
            sbSql.Append("         a.mail_address, ")
            sbSql.Append("         b.language_cd, ")
            sbSql.Append("         b.currency_cd, ")
            sbSql.Append("         b.edit_div, ")
            sbSql.Append("         a.user_class, ")
            sbSql.Append("         a.price_disp_lvl, ")
            sbSql.Append("         a.add_information_lvl, ")
            sbSql.Append("         a.use_function_lvl, ")
            'sbSql.Append("         DATEDIFF(MINUTE, c.login_datetime,CURRENT_TIMESTAMP) AS datediff, ")
            sbSql.Append("         a.current_datetime ")
            sbSql.Append(" FROM    kh_user_mst  a ")
            sbSql.Append(" INNER JOIN kh_country_mst  b ")
            sbSql.Append(" ON      a.country_cd = b.country_cd ")
            'sbSql.Append(" LEFT JOIN  kh_login  c ")
            'sbSql.Append(" ON      a.user_id                     = c.user_id ")
            sbSql.Append(" WHERE   CAST(a.user_id AS VARBINARY)  = CAST(@UserId AS VARBINARY) ")
            sbSql.Append(" AND     CAST(a.password AS VARBINARY) = CAST(@PassWord AS VARBINARY) ")
            sbSql.Append(" AND     a.in_effective_date          <= @StandardDate ")
            sbSql.Append(" AND     a.out_effective_date          > @StandardDate ")

            'DB接続文字列の取得
            'objCon.ConnectionString = My.Settings.connkhBase
            objCmd = New SqlCommand(sbSql.ToString, objConBase)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserid
                .Parameters.Add("@PassWord", SqlDbType.VarChar, 10).Value = strPasswd
                .Parameters.Add("@StandardDate", SqlDbType.DateTime, 8).Value = Now()
            End With

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncSelectUserInfo)

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
    ''' ログイン情報追加
    ''' </summary>
    ''' <param name="strUserid">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncInsertLoginInfo(ByVal strUserid As String, _
                                        ByVal strSessionId As String) As Integer
        Dim objCmd As SqlCommand
        Dim objConBase As New SqlConnection
        objConBase = New SqlClient.SqlConnection(My.Settings.connkhBase)
        objConBase.Open()

        Try
            objCmd = New SqlCommand(CdCst.DB.SPL.KHLoginRec, objConBase)

            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserid
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
            End With

            '実行
            objCmd.ExecuteNonQuery()
            fncInsertLoginInfo = 0
        Catch ex As Exception
            'システムエラー
            fncInsertLoginInfo = 9
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
            If Not objConBase Is Nothing Then If Not objConBase.State = ConnectionState.Closed Then objConBase.Close()
        End Try

    End Function

    ''' <summary>
    ''' ログイン情報取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectLoginInfo(ByVal objCon As SqlConnection, ByVal strSessionId As String) As DataTable
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        fncSelectLoginInfo = New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  user_id ")
            sbSql.Append(" FROM    kh_login ")
            sbSql.Append(" WHERE   session_id = @SessionId ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
            End With

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncSelectLoginInfo)

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
    ''' パスワードの取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strUserid">ユーザーＩＤ</param>
    ''' <returns></returns>
    ''' <remarks>パスワードを取得する</remarks>
    Public Function fncSelectPassword(ByVal objConBase As SqlConnection, ByVal strUserId As String) As String
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        Dim dt As New DataTable
        'デフォルト設定
        fncSelectPassword = String.Empty

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  password ")
            sbSql.Append(" FROM    kh_user_mst ")
            sbSql.Append(" WHERE   CAST(user_id AS VARBINARY)  = CAST(@UserId AS VARBINARY) ")

            objCmd = New SqlCommand(sbSql.ToString, objConBase)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
            End With

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dt)

            If dt.Rows.Count > 0 Then
                fncSelectPassword = dt.Rows(0)("password")
            End If

        Catch ex As Exception
            fncSelectPassword = String.Empty
            WriteErrorLog("E001", ex)
        Finally
            sbSql = Nothing
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' パスワードの更新
    ''' </summary>
    ''' <param name="strUserId"></param>
    ''' <param name="strNewPasswd"></param>
    ''' <param name="strCurrentDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncUpdatePassword(ByVal strUserId As String, ByVal strNewPasswd As String, ByVal strCurrentDate As String) As Boolean
        Dim objCmd As SqlCommand
        Dim objConBase As New SqlConnection
        Dim blnResult As Boolean = False

        objConBase = New SqlClient.SqlConnection(My.Settings.connkhBase)
        objConBase.Open()

        Try
            objCmd = New SqlCommand(CdCst.DB.SPL.KHLoginRec, objConBase)

            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHUserPasswdChg

                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@Passwd", SqlDbType.VarChar, 10).Value = strNewPasswd
                .Parameters.Add("@CurrentDatetime", SqlDbType.DateTime).Value = strCurrentDate
            End With

            '実行
            If objCmd.ExecuteNonQuery() > 0 Then
                blnResult = True
            Else
                blnResult = False
            End If

        Catch ex As Exception
            'システムエラー
            blnResult = False
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
            If Not objConBase Is Nothing Then If Not objConBase.State = ConnectionState.Closed Then objConBase.Close()
        End Try

        Return blnResult
    End Function

    ''' <summary>
    ''' ユーザ情報の削除(DBBase)
    ''' </summary>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncDeleteUserInfoFromBase(ByVal objConBase As SqlConnection, ByVal strUserId As String, ByVal strSessionId As String) As Boolean
        Dim objCmd As SqlCommand
        Dim blnResult As Boolean = False

        Try
            objCmd = objConBase.CreateCommand

            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHLogoutRec
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
            End With

            '実行
            If objCmd.ExecuteNonQuery() > 0 Then
                blnResult = True
            Else
                blnResult = False
            End If

        Catch ex As Exception
            'システムエラー
            blnResult = False
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
            'If Not objConBase Is Nothing Then If Not objConBase.State = ConnectionState.Closed Then objConBase.Close()
        End Try

        Return blnResult
    End Function

    ''' <summary>
    ''' ユーザ情報の削除(DBWeb)
    ''' </summary>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncDeleteUserInfoFromWeb(ByVal objCon As SqlConnection, ByVal strUserId As String, ByVal strSessionId As String) As Boolean
        Dim objCmd As SqlCommand
        Dim blnResult As Boolean = False

        Try
            objCmd = objCon.CreateCommand

            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHSelKtbnInfoDel
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionId
            End With

            '実行
            If objCmd.ExecuteNonQuery() > 0 Then
                blnResult = True
            Else
                blnResult = False
            End If

        Catch ex As Exception
            'システムエラー
            blnResult = False
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
            'If Not objConBase Is Nothing Then If Not objConBase.State = ConnectionState.Closed Then objConBase.Close()
        End Try

        Return blnResult
    End Function
End Class
