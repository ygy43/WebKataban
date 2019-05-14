Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class KHUser

    Private objUserInfo As KHSessionInfo.UserInfo                   'ユーザー情報
    Private objLoginInfo As KHSessionInfo.LoginInfo                 'ログイン情報
    Private bllLogin As New LoginBLL                                'ビジネスロジック

#Region " Property "

    '**********************************************************************************************
    '*【プロパティ】UserInfo
    '*  ユーザー情報を取得する
    '**********************************************************************************************
    Public ReadOnly Property UserInfo() As KHSessionInfo.UserInfo
        Get
            Return Me.objUserInfo
        End Get
    End Property

    '**********************************************************************************************
    '*【プロパティ】LoginInfo
    '*  ログイン情報を取得する
    '**********************************************************************************************
    Public ReadOnly Property LoginInfo() As KHSessionInfo.LoginInfo
        Get
            Return Me.objLoginInfo
        End Get
    End Property

#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Try
            Me.objUserInfo = Nothing
            Me.objLoginInfo = Nothing
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' セッションユーザ情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' セッション上のユーザ情報を取得する。
    ''' DB上のログイン情報と一致しない場合、セッション情報をクリアする。
    ''' 取得に失敗した際はログイン画面へ遷移する。
    ''' </remarks>
    Public Function subGetSession(objCon As SqlConnection) As Boolean
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        subGetSession = False
        Try
            'セッション情報取得  取得できなければログイン画面に遷移
            If httpCon.Session Is Nothing Then Exit Function

            'ユーザー情報取得    取得できなければログイン画面に遷移
            If httpCon.Session.Item(CdCst.SessionInfo.Key.UserInfo) Is Nothing Then
                Exit Function
            Else
                Me.objUserInfo = httpCon.Session.Item(CdCst.SessionInfo.Key.UserInfo) 'ユーザー情報をセット
            End If

            'ログイン情報取得    取得できなければログイン画面に遷移
            If httpCon.Session.Item(CdCst.SessionInfo.Key.LoginInfo) Is Nothing Then
                Exit Function
            Else
                Me.objLoginInfo = httpCon.Session.Item(CdCst.SessionInfo.Key.LoginInfo) 'ログイン情報をセット
            End If

            'DB情報と同期していない為、セッションをクリアしてログイン画面へ遷移
            If bllLogin.fncSelectLoginInfo(objCon, Me.LoginInfo.SessionId) = False Then
                httpCon.Session.Clear()
                Exit Function
            End If

            subGetSession = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' ログアウト処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objConBase"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <remarks>
    ''' 引数で渡されたユーザＩＤにてログインマスタ（kh_login）の情報を削除する。
    ''' 現プロセスのセッション情報を削除する。
    ''' </remarks>
    Public Sub subUserLogout(ByVal objCon As SqlConnection, ByVal objConBase As SqlConnection, _
                             ByVal strUserId As String, ByVal strSessionId As String)

        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current

        Try
            'ログイン情報を削除
            bllLogin.fncDeleteUserInfoFromBase(objConBase, strUserId, strSessionId)

            '引当形番情報削除を削除
            bllLogin.fncDeleteUserInfoFromWeb(objCon, strUserId, strSessionId)

            'セッション情報削除
            httpCon.Session.Clear()

            'キャッシュ削除
            httpCon.Cache.Remove(CdCst.CacheInfo.Key.MenuClass)
            httpCon.Cache.Remove(CdCst.CacheInfo.Key.MenuContent)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' ログイン処理
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strUserId"> ユーザーＩＤ</param>
    ''' <param name="strPasswd">パスワード</param>
    ''' <param name="strLangCd">言語コード</param>
    ''' <param name="strDateFlg"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 引数で渡されたユーザＩＤ、パスワードにてユーザマスタ（kh_user_mst）より
    ''' 認証をコールする。
    ''' </remarks>
    Public Function fncUserLogin(objConBase As SqlConnection, ByVal strUserId As String, ByVal strPasswd As String, _
                                 ByVal strLangCd As String, ByVal strDateFlg As String) As Integer
        fncUserLogin = -1
        Try
            'ログイン認証
            fncUserLogin = bllLogin.fncUserChk(objConBase, strUserId, strPasswd, strLangCd, strDateFlg)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' パスワード変更処理
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strPasswd">パスワード</param>
    ''' <param name="strNewPasswd">新パスワード</param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strDateFlg"></param>
    ''' <param name="strCurrentDate"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 引数で渡されたユーザＩＤ、パスワードにてユーザマスタ（kh_user_mst）より認証をコールする。
    ''' 認証が成功した場合は、ユーザＩＤ、パスワード、新パスワードにてユーザマスタ（kh_user_mst）に対し、
    ''' パスワードの更新を行う。
    ''' </remarks>
    Public Function fncUserPasswdChg(objConBase As SqlConnection, ByVal strUserId As String, _
                                     ByVal strPasswd As String, ByVal strNewPasswd As String, _
                                     ByVal strCountryCd As String, ByVal strDateFlg As String, _
                                     ByVal strCurrentDate As String) As Integer
        fncUserPasswdChg = 9

        Try
            'ユーザ認証
            Select Case bllLogin.fncUserChk(objConBase, strUserId, strPasswd, strCountryCd, strDateFlg)
                Case 0
                    If bllLogin.fncUpdatePassword(strUserId, strNewPasswd, strCurrentDate) Then
                        '更新成功
                        Return 0
                    Else
                        '更新失敗
                        Return 9
                    End If
                Case 1
                    '認証失敗
                    Return 1
                Case 2
                    'ログインパスワードが前回更新時から40日以上経過したユーザーには警告メッセージの表示をする。
                    Return 2
                Case 9
                    'システムエラー
                    Return 9
            End Select

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' パスワードの取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strUserid">ユーザーＩＤ</param>
    ''' <returns></returns>
    ''' <remarks>パスワードを取得する</remarks>
    Public Function fncGetPassword(objConBase As SqlConnection, ByVal strUserId As String) As String
        fncGetPassword = String.Empty
        Try
            fncGetPassword = bllLogin.fncSelectPassword(objConBase, strUserId)
        Catch ex As Exception
            fncGetPassword = String.Empty
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 利用機能情報取得
    ''' </summary>
    ''' <param name="intUseFunctionLvl">利用機能レベル</param>
    ''' <returns></returns>
    ''' <remarks>利用機能レベルをもとに利用機能情報を取得する</remarks>
    Public Function fncUseFunctionInfoGet(ByVal intUseFunctionLvl As Integer) As Integer()

        Dim intUserFunctionInfo() As Integer
        Dim intWkUseFunctionLvl As Integer = intUseFunctionLvl
        Dim intLoopCnt As Integer
        Dim intLevel(8) As Integer

        'ReDim intLevel(8)
        intLevel(0) = 0
        intLevel(1) = 1
        intLevel(2) = 2
        intLevel(3) = 4
        intLevel(4) = 8
        intLevel(5) = 16
        intLevel(6) = 32
        intLevel(7) = 64
        intLevel(8) = 128

        ReDim intUserFunctionInfo(0)
        fncUseFunctionInfoGet = Nothing
        Try
            For intLoopCnt = intLevel.Length - 1 To 1 Step -1
                Select Case intLevel(intLoopCnt)
                    Case strcUserFunctionLvl.CountryItemMstMnt
                        If intWkUseFunctionLvl >= strcUserFunctionLvl.CountryItemMstMnt Then
                            ReDim Preserve intUserFunctionInfo(UBound(intUserFunctionInfo) + 1)
                            intUserFunctionInfo(UBound(intUserFunctionInfo)) = strcUserFunctionLvl.CountryItemMstMnt
                            intWkUseFunctionLvl = intWkUseFunctionLvl - strcUserFunctionLvl.CountryItemMstMnt
                        End If
                    Case strcUserFunctionLvl.CurrencyExcRateMstMnt
                        If intWkUseFunctionLvl >= strcUserFunctionLvl.CurrencyExcRateMstMnt Then
                            ReDim Preserve intUserFunctionInfo(UBound(intUserFunctionInfo) + 1)
                            intUserFunctionInfo(UBound(intUserFunctionInfo)) = strcUserFunctionLvl.CurrencyExcRateMstMnt
                            intWkUseFunctionLvl = intWkUseFunctionLvl - strcUserFunctionLvl.CurrencyExcRateMstMnt
                        End If
                    Case strcUserFunctionLvl.InfoMstMnt
                        If intWkUseFunctionLvl >= strcUserFunctionLvl.InfoMstMnt Then
                            ReDim Preserve intUserFunctionInfo(UBound(intUserFunctionInfo) + 1)
                            intUserFunctionInfo(UBound(intUserFunctionInfo)) = strcUserFunctionLvl.InfoMstMnt
                            intWkUseFunctionLvl = intWkUseFunctionLvl - strcUserFunctionLvl.InfoMstMnt
                        End If
                    Case strcUserFunctionLvl.UserMstMnt
                        If intWkUseFunctionLvl >= strcUserFunctionLvl.UserMstMnt Then
                            ReDim Preserve intUserFunctionInfo(UBound(intUserFunctionInfo) + 1)
                            intUserFunctionInfo(UBound(intUserFunctionInfo)) = strcUserFunctionLvl.UserMstMnt
                            intWkUseFunctionLvl = intWkUseFunctionLvl - strcUserFunctionLvl.UserMstMnt
                        End If
                    Case strcUserFunctionLvl.RateMstMnt
                        If intWkUseFunctionLvl >= strcUserFunctionLvl.RateMstMnt Then
                            ReDim Preserve intUserFunctionInfo(UBound(intUserFunctionInfo) + 1)
                            intUserFunctionInfo(UBound(intUserFunctionInfo)) = strcUserFunctionLvl.RateMstMnt
                            intWkUseFunctionLvl = intWkUseFunctionLvl - strcUserFunctionLvl.RateMstMnt
                        End If
                    Case strcUserFunctionLvl.SapIF
                        If intWkUseFunctionLvl >= strcUserFunctionLvl.SapIF Then
                            ReDim Preserve intUserFunctionInfo(UBound(intUserFunctionInfo) + 1)
                            intUserFunctionInfo(UBound(intUserFunctionInfo)) = strcUserFunctionLvl.SapIF
                            intWkUseFunctionLvl = intWkUseFunctionLvl - strcUserFunctionLvl.SapIF
                        End If
                    Case strcUserFunctionLvl.AcosIF
                        If intWkUseFunctionLvl >= strcUserFunctionLvl.AcosIF Then
                            ReDim Preserve intUserFunctionInfo(UBound(intUserFunctionInfo) + 1)
                            intUserFunctionInfo(UBound(intUserFunctionInfo)) = strcUserFunctionLvl.AcosIF
                            intWkUseFunctionLvl = intWkUseFunctionLvl - strcUserFunctionLvl.AcosIF
                        End If
                End Select
            Next

            fncUseFunctionInfoGet = intUserFunctionInfo

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function
End Class
