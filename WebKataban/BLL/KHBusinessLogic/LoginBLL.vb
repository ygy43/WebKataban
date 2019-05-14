Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class LoginBLL

    Private Property dllLogin As New LoginDAL
    ''' <summary>
    ''' ログイン認証
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strUserid">ユーザーＩＤ</param>
    ''' <param name="strPasswd">パスワード</param>
    ''' <param name="strSelLang">選択言語コード</param>
    ''' <param name="strDateFlg"></param>
    ''' <returns>0：認証成功  1：認証失敗  2：ログイン情報あり　3：40日以上警告</returns>
    ''' <remarks>
    ''' 引数で渡されたユーザＩＤ、パスワードにてユーザマスタ（kh_user_mst）より認証を行う
    ''' 認証した場合は、ログインマスタ（kh_login）及びセッションにユーザ情報をセットする
    ''' 認証時、ログインマスタ（kh_login）にセッション時間以前の情報が存在している場合、
    ''' ログインせずに通知する。
    ''' </remarks>
    Public Function fncUserChk(objConBase As SqlConnection, _
                               ByVal strUserid As String, ByVal strPasswd As String, _
                               ByVal strSelLang As String, ByVal strDateFlg As String) As Integer

        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim objUserInfo As KHSessionInfo.UserInfo
        Dim objLoginInfo As KHSessionInfo.LoginInfo
        Dim result As DateTime
        Dim dt As New DataTable

        'デフォルト設定
        fncUserChk = 0

        Try
            'データの取得
            dt = dllLogin.fncSelectUserInfo(objConBase, strUserid, strPasswd, strSelLang)

            If dt.Rows.Count > 0 Then
                '認証成功
                If strDateFlg = 1 Then
                    If Not IsDBNull(dt.Rows(0)("current_datetime")) Then
                        result = DateAdd("d", 40, dt.Rows(0)("current_datetime"))
                        If result < Now() Then
                            'ログインパスワードが前回更新時から40日以上経過したユーザーには警告メッセージの表示をする。
                            fncUserChk = 3
                        End If
                    End If
                End If
            Else
                '認証失敗
                fncUserChk = 1
            End If

            '認証成功時はセッション情報作成
            Select Case fncUserChk
                Case 0
                    'ユーザー情報設定
                    With objUserInfo
                        .UserId = dt.Rows(0)("user_id")
                        .BaseCd = dt.Rows(0)("base_cd")
                        .CountryCd = dt.Rows(0)("country_cd")

                        .OfficeCd = IIf(IsDBNull(dt.Rows(0)("office_cd")), "", dt.Rows(0)("office_cd"))
                        .PersonCd = IIf(IsDBNull(dt.Rows(0)("person_cd")), "", dt.Rows(0)("person_cd"))
                        .MailAddress = IIf(IsDBNull(dt.Rows(0)("mail_address")), "", dt.Rows(0)("mail_address"))

                        .LanguageCd = dt.Rows(0)("language_cd")
                        .CurrencyCd = dt.Rows(0)("currency_cd")
                        .EditDiv = dt.Rows(0)("edit_div")

                        .UserClass = dt.Rows(0)("user_class")
                        .PriceDispLvl = dt.Rows(0)("price_disp_lvl")
                        .AddInformationLvl = dt.Rows(0)("add_information_lvl")
                        .UseFunctionLvl = dt.Rows(0)("use_function_lvl")
                        .TnkDispCnt = 1
                    End With

                    'セッション情報に確保
                    httpCon.Session(CdCst.SessionInfo.Key.UserInfo) = objUserInfo

                    'ログイン情報設定
                    With objLoginInfo
                        .SessionId = httpCon.Session.SessionID
                        If strSelLang.Trim = "" Then
                            .SelectLang = dt.Rows(0)("language_cd")
                        Else
                            .SelectLang = strSelLang
                        End If
                    End With

                    'セッション情報に確保
                    httpCon.Session(CdCst.SessionInfo.Key.LoginInfo) = objLoginInfo

                    'ログイン情報追加
                    fncUserChk = dllLogin.fncInsertLoginInfo(strUserid, httpCon.Session.SessionID)
                Case 3
                    'ユーザー情報設定
                    With objUserInfo
                        .UserId = dt.Rows(0)("user_id")
                        .BaseCd = dt.Rows(0)("base_cd")
                        .CountryCd = dt.Rows(0)("country_cd")

                        .OfficeCd = IIf(IsDBNull(dt.Rows(0)("office_cd")), "", dt.Rows(0)("office_cd"))
                        .PersonCd = IIf(IsDBNull(dt.Rows(0)("person_cd")), "", dt.Rows(0)("person_cd"))
                        .MailAddress = IIf(IsDBNull(dt.Rows(0)("mail_address")), "", dt.Rows(0)("mail_address"))

                        .LanguageCd = dt.Rows(0)("language_cd")
                        .CurrencyCd = dt.Rows(0)("currency_cd")
                        .EditDiv = dt.Rows(0)("edit_div")

                        .UserClass = dt.Rows(0)("user_class")
                        .PriceDispLvl = dt.Rows(0)("price_disp_lvl")
                        .AddInformationLvl = dt.Rows(0)("add_information_lvl")
                        .UseFunctionLvl = dt.Rows(0)("use_function_lvl")
                        .TnkDispCnt = 1
                    End With

                    'セッション情報に確保
                    httpCon.Session(CdCst.SessionInfo.Key.UserInfo) = objUserInfo

                    'ログイン情報設定
                    With objLoginInfo
                        .SessionId = httpCon.Session.SessionID
                        If strSelLang.Trim = "" Then
                            .SelectLang = dt.Rows(0)("language_cd")
                        Else
                            .SelectLang = strSelLang
                        End If
                    End With

                    'セッション情報に確保
                    httpCon.Session(CdCst.SessionInfo.Key.LoginInfo) = objLoginInfo

                    'ログイン情報追加
                    fncUserChk = dllLogin.fncInsertLoginInfo(strUserid, httpCon.Session.SessionID)
                    fncUserChk = 3
            End Select
        Catch ex As Exception
            'システムエラー
            fncUserChk = 9
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' ログイン情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectLoginInfo(ByVal objCon As SqlConnection, ByVal strSessionId As String) As Boolean
        Dim dt As New DataTable
        Dim blnResult As Boolean = False

        Try
            'ログイン情報の取得
            dt = dllLogin.fncSelectLoginInfo(objCon, strSessionId)
            If dt.Rows.Count > 0 Then
                blnResult = True
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return blnResult
    End Function

    ''' <summary>
    ''' パスワードの取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectPassword(ByVal objConBase As SqlConnection, ByVal strUserId As String) As String
        Dim strResult As String = String.Empty

        Try
            dllLogin.fncSelectPassword(objConBase, strUserId)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return strResult
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
        Dim blnResult As Boolean = False
        Try
            blnResult = dllLogin.fncUpdatePassword(strUserId, strNewPasswd, strCurrentDate)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
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
        Dim blnResult As Boolean = False
        Try
            blnResult = dllLogin.fncDeleteUserInfoFromBase(objConBase, strUserId, strSessionId)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
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
        Dim blnResult As Boolean = False
        Try
            blnResult = dllLogin.fncDeleteUserInfoFromWeb(objCon, strUserId, strSessionId)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return blnResult
    End Function
End Class
