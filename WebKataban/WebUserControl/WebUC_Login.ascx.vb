Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class WebUC_Login
    Inherits System.Web.UI.UserControl

#Region "プロパティ"
    'DB接続
    Public objCon As New SqlConnection
    Public objConBase As New SqlConnection

    Private strUserID As String = String.Empty          'ユーザーID
    Private strPassword As String = String.Empty        'パスワード
    Private strParaE As String = String.Empty           '
    Private Const WebforGS As String = "WebforGS"       'WebforGS連携用
    Private bolRes As Boolean = False                   ' 
    Private selLang As DropDownList = Nothing           '言語選択欄

    Private Structure ErrCode                       'エラーコード
        Private strDummy As String
        Public Const I0060 As String = "I0060"
        Public Const E0020 As String = "E0020"
        Public Const E0030 As String = "E0030"
        Public Const E9999 As String = "E9999"
    End Structure

    'イベント
    Public Event Goto_ErrPage(strErr As String)
    Public Event GoToMenuPage()
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
    End Sub

    ''' <summary>
    ''' ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Me.Visible Then Exit Sub

        '言語欄の取得
        selLang = Me.Parent.Parent.Parent.Parent.FindControl("ContentTitle").FindControl("selLang")
        Me.txtUserID.Focus()

        '初期化
        If Not LoginInit() Then
            Exit Sub
        End If

        'Me.txtUserID.Text = "SYSPRC"
        'Me.txtPasswd.Text = "system"
        'Call Button1_Click(Nothing, Nothing)
    End Sub

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function LoginInit() As Boolean
        LoginInit = False

        Dim strMacAddress As String = String.Empty
        Dim strSerialNo As String = String.Empty
        Dim strDummyNo As String = "99999"
        Dim bolLoginCheck As Boolean = False
        Dim bolEncrypted As Boolean = False
        Dim bolUrlEncoded As Boolean = False
        Dim bolEdiReturn As Boolean = False
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim objUserInfo As New KHSessionInfo.UserInfo

        Try
            bolLoginCheck = True
            bolEncrypted = True
            bolUrlEncoded = True
            bolEdiReturn = True

            'Hidden項目をクリア
            hiddenCurrentDatetime.Value = String.Empty
            hiddenOver.Value = String.Empty
            hiddenPasswd.Value = String.Empty
            hiddenNewPasswd.Value = String.Empty

            'Config取得
            bolLoginCheck = My.Settings.LoginCheck
            bolEncrypted = My.Settings.Encrypted
            bolUrlEncoded = My.Settings.UrlEncoded
            bolEdiReturn = My.Settings.EdiReturn

            If bolLoginCheck Then   '認証を行う場合
                'Webサービスインスタンス作成
                Dim wsWebloginSystem As New weblogin.LoginCheck
                '構成ファイルからURL指定
                wsWebloginSystem.Url = My.Settings.WebKataban_weblogin_LoginCheck

                'パラメータの取得
                strUserID = Request.QueryString("A")
                strPassword = Request.QueryString("B")
                strMacAddress = Request.QueryString("C")
                strSerialNo = Request.QueryString("D")

                '国内代理店用に追加
                If Not Request.QueryString("E") Is Nothing Then
                    strParaE = Request.QueryString("E")
                End If

                If bolUrlEncoded Then   'URLエンコードされている場合、URLデコードする
                    strUserID = System.Web.HttpUtility.UrlDecode(strUserID)
                    strPassword = System.Web.HttpUtility.UrlDecode(strPassword)

                    strMacAddress = System.Web.HttpUtility.UrlDecode(strMacAddress)
                    strSerialNo = System.Web.HttpUtility.UrlDecode(strSerialNo)
                    strParaE = System.Web.HttpUtility.UrlDecode(strParaE)
                End If

                '暗号化されている場合、復号する
                If bolEncrypted Then
                    strUserID = wsWebloginSystem.Decode(strUserID)
                    strPassword = wsWebloginSystem.Decode(strPassword)
                    strMacAddress = wsWebloginSystem.Decode(strMacAddress)
                    strSerialNo = wsWebloginSystem.Decode(strSerialNo)
                    strParaE = wsWebloginSystem.Decode(strParaE)
                End If

                'パラメータが無い場合エラー画面を表示する
                If strUserID = String.Empty Or strPassword = String.Empty Or strMacAddress = String.Empty Then
                    Throw New ArgumentException
                    Exit Function
                End If

                'シリアルNoが無い場合「99999」(ダミーNo)をセットする
                If strSerialNo = String.Empty Then
                    'Webサービス呼び出し
                    bolRes = wsWebloginSystem.CheckMacaddress(strUserID, strPassword, strMacAddress)
                Else
                    'Webサービス呼び出し
                    bolRes = wsWebloginSystem.CheckAuthentication(strUserID, strPassword, strMacAddress, strSerialNo)
                End If
            Else
                '国内営業所用に追加
                If bolEdiReturn And Not String.IsNullOrEmpty(Request.QueryString("E")) Then
                    'Webサービスインスタンス作成
                    Dim wsEDISystem As New weblogin.LoginCheck

                    strParaE = Request.QueryString("E")

                    '暗号化されているので復号する
                    strParaE = wsEDISystem.Decode(strParaE.ToString.Trim)

                    'ユーザIDとパスワードの設定
                    Dim strArray As String()
                    strArray = strParaE.Split("_")
                    For intLoopCnt As Integer = 0 To strArray.Length - 1
                        Select Case intLoopCnt
                            Case 0
                                strUserID = strArray(intLoopCnt)
                            Case 1
                                strPassword = strArray(intLoopCnt)
                            Case Else
                        End Select
                    Next
                ElseIf Not String.IsNullOrEmpty(Request.QueryString("role")) Then
                    '匿名ユーザー
                    If Request.QueryString("role").ToString.Trim.Equals("custom") Then
                        strUserID = My.Settings.AnonymousUserName
                        strPassword = My.Settings.AnonymousPassword
                        If String.IsNullOrEmpty(Request.QueryString("lang")) Then
                            selLang.SelectedValue = CdCst.LanguageCd.DefaultLang
                        Else
                            If Request.QueryString("lang").Trim() <> CdCst.LanguageCd.DefaultLang AndAlso 
                               Request.QueryString("lang").Trim() <> CdCst.LanguageCd.Japanese AndAlso 
                               Request.QueryString("lang").Trim() <> CdCst.LanguageCd.Korean AndAlso 
                               Request.QueryString("lang").Trim() <> CdCst.LanguageCd.SimplifiedChinese AndAlso 
                               Request.QueryString("lang").Trim() <> CdCst.LanguageCd.TraditionalChinese Then
                                Throw New Exception("URLをご確認ください。")
                            End If
                            selLang.SelectedValue = Request.QueryString("lang").Trim()
                        End If

                        'ログイン処理
                        Me.txtUserID.Text = strUserID
                        Me.txtPasswd.Text = strPassword
                        Call Button1_Click(Nothing, Nothing)
                    Else
                        'role<>customの時、エラー画面へ遷移
                        Throw New Exception("URLをご確認ください。")
                    End If
                End If

                bolRes = True
            End If

            If bolRes Then
                '画面ラベル設定
                Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHLogin, selLang.SelectedValue, Me)

                If My.Settings.EdiReturn Then
                    '代理店EDIシステムから起動する場合
                    subJuchuEDIProc()
                End If

                'TextBoxにEnterキーが押された場合Submitしない
                subSetInitScript()
            Else
                objUserInfo.UserId = strUserID
                'セッション情報に確保
                httpCon.Session(CdCst.SessionInfo.Key.UserInfo) = objUserInfo
                Session.Add(CdCst.SessionInfo.Key.ErrorCode, "TIMEOUT")
                'エラー画面に遷移する
                RaiseEvent Goto_ErrPage("")
                Exit Function
            End If
            LoginInit = True
        Catch ex As Exception
            RaiseEvent Goto_ErrPage(ex.Message) 'エラー画面へ
        End Try
    End Function

    ''' <summary>
    ''' アップデートログインボタン押下時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim clsUser As New KHUser
        Dim objLoginInfo As KHSessionInfo.LoginInfo
        Dim strOverFlg As String
        Dim strUserId As String
        Dim strPasswd As String
        Dim intReturn As Integer
        Dim strLangCd As String = ""
        Dim strNewPasswd As String
        Dim strDateFlg As String = 0
        Dim strCurrentDate As DateTime

        Try
            If Len(Trim(Me.txtNewPasswd.Text)) <> 0 Then
                '選択言語設定
                strLangCd = selLang.SelectedValue

                '入力チェック
                If fncValidate("1", strLangCd) = False Then
                    Exit Try
                End If

                '入力情報取得
                strOverFlg = Me.hiddenOver.Value
                strUserId = Me.txtUserID.Text
                strPasswd = Me.txtPasswd.Text
                strNewPasswd = Me.txtNewPasswd.Text
                strCurrentDate = Now()

                'ログイン処理
                Session.Clear()
                If strOverFlg = "0" Then
                    intReturn = clsUser.fncUserPasswdChg(objConBase, strUserId, strPasswd, strNewPasswd, strLangCd, strDateFlg, strCurrentDate)
                Else
                    intReturn = clsUser.fncUserPasswdChg(objConBase, strUserId, strPasswd, strNewPasswd, strLangCd, strDateFlg, strCurrentDate)
                End If

                If httpCon.Session(CdCst.SessionInfo.Key.LoginInfo) IsNot Nothing Then
                    objLoginInfo = httpCon.Session(CdCst.SessionInfo.Key.LoginInfo)
                    strLangCd = objLoginInfo.SelectLang
                End If
                If Len(Trim(strLangCd)) = 0 Then
                    strLangCd = CdCst.LanguageCd.DefaultLang
                End If

                Select Case intReturn
                    Case 0
                        'メニュー画面へ遷移
                        RaiseEvent GoToMenuPage()
                    Case 1
                        'ログイン失敗
                        AlertMessage(strLangCd, "W0030") 'ログイン情報が間違っています。
                        ''エラー画面に遷移
                        'Session.Add(CdCst.SessionInfo.Key.ErrorCode, CdCst.SessionInfo.Key.LoginError)
                        'RaiseEvent Goto_ErrPage(String.Empty)
                    Case 2
                        'パスワードを隠しエリアに保持しておく
                        Me.hiddenPasswd.Value = Me.txtPasswd.Text
                        Me.hiddenNewPasswd.Value = Me.txtNewPasswd.Text
                        Me.hiddenCurrentDatetime.Value = Now()
                        ConfirmMessage(strLangCd, "I0010", "1") '既に同じユーザがログインしています。
                End Select
            Else
                'メニュー画面へ遷移
                RaiseEvent GoToMenuPage()
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
            'エラー画面に遷移する
            RaiseEvent Goto_ErrPage(ex.Message)
        Finally
            clsUser = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 入力項目のチェック
    ''' </summary>
    ''' <param name="strDv"></param>
    ''' <param name="strLang"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncValidate(ByVal strDv As String, ByVal strLang As String) As Boolean
        Dim strNewPasswd As String
        Dim strNewPasswdRe As String
        fncValidate = False
        Try
            If Len(Trim(strLang)) = 0 Then strLang = CdCst.LanguageCd.DefaultLang

            If Len(Trim(Me.txtUserID.Text)) = 0 Then
                AlertMessage(strLang, "W0010") 'User IDを入力してください。
                Return False
            End If

            If Len(Trim(Me.hiddenPasswd.Value)) > 0 Then Me.txtPasswd.Text = Me.hiddenPasswd.Value
            If Len(Trim(Me.txtPasswd.Text)) = 0 Then
                'CHANGED BY YGY 20141029
                AlertMessage(selLang.SelectedValue, "W0020")
                'WriteErrorLog("W0020", selLang.SelectedValue) 'Passwordを入力してください。
                Return False
            End If

            If strDv = "1" Then
                If Len(Trim(Me.hiddenNewPasswd.Value)) > 0 Then
                    Me.txtNewPasswd.Text = Me.hiddenNewPasswd.Value
                    Me.txtNewPasswdRe.Text = Me.hiddenNewPasswd.Value
                Else
                    If Len(Trim(Me.txtNewPasswd.Text)) = 0 Then
                        AlertMessage(strLang, "W0040") 'New Passwordを入力してください。
                        Return False
                    Else
                        'パスワードの文字数チェックを追加
                        If Len(Trim(Me.txtNewPasswd.Text)) < 5 Or Len(Trim(Me.txtNewPasswd.Text)) > 10 Then
                            AlertMessage(strLang, "W8810") 'New Passwordの文字数が不正です。
                            Return False
                        End If
                        strNewPasswd = Trim(Me.txtNewPasswd.Text)
                    End If
                    If Len(Trim(Me.txtNewPasswdRe.Text)) = 0 Then
                        AlertMessage(strLang, "W0040") 'New Passwordを入力してください。
                        Return False
                    Else
                        'パスワードの文字数チェックを追加
                        If Len(Trim(Me.txtNewPasswdRe.Text)) < 5 Or Len(Trim(Me.txtNewPasswdRe.Text)) > 10 Then
                            AlertMessage(strLang, "W8810") 'New Passwordの文字数が不正です。
                            Return False
                        End If
                        strNewPasswdRe = Trim(Me.txtNewPasswdRe.Text)
                    End If
                    If strNewPasswd <> strNewPasswdRe Then
                        AlertMessage(strLang, "W0050") 'New Passwordが異なっています。
                        Return False
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            RaiseEvent Goto_ErrPage(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' ログインボタン押下時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim clsUser As New KHUser
        Dim objLoginInfo As KHSessionInfo.LoginInfo
        Dim strOverFlg As String
        Dim strUserId As String
        Dim strPasswd As String
        Dim intReturn As Integer
        Dim strLangCd As String = ""
        Dim strDateFlg As String = 1

        Try
            '選択言語設定
            strLangCd = selLang.SelectedValue
            If strLangCd = "" Then strLangCd = CdCst.LanguageCd.DefaultLang

            '入力チェック
            If Me.fncValidate("0", strLangCd) = False Then
                'DELETE BY YGY 20141029
                'fncInpCheckの中に既にメッセージ表示した
                'WriteErrorLog("W0030", strLangCd) 'ログイン情報が間違っています。
                Exit Try
            End If

            '入力情報取得
            strOverFlg = Me.hiddenOver.Value
            strUserId = Me.txtUserID.Text
            strPasswd = Me.txtPasswd.Text
            Me.hiddenNewPasswd.Value = Me.txtPasswd.Text

            '代理店EDI　WebforGSから起動する場合以外はセッションをクリア
            If httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) Is Nothing Then
                Session.Clear()
            End If

            'ログイン処理
            If strOverFlg = "0" Then
                intReturn = clsUser.fncUserLogin(objConBase, strUserId, strPasswd, strLangCd, strDateFlg)
            Else
                intReturn = clsUser.fncUserLogin(objConBase, strUserId, strPasswd, strLangCd, strDateFlg)
            End If

            If httpCon.Session(CdCst.SessionInfo.Key.LoginInfo) IsNot Nothing Then
                objLoginInfo = httpCon.Session(CdCst.SessionInfo.Key.LoginInfo)
                strLangCd = objLoginInfo.SelectLang
            End If
            If selLang.SelectedValue.ToString.Length > 0 Then
                strLangCd = selLang.SelectedValue
            Else
                Dim objUserInfo As KHSessionInfo.UserInfo = httpCon.Session(CdCst.SessionInfo.Key.UserInfo)
                If Not objUserInfo.LanguageCd Is Nothing Then strLangCd = objUserInfo.LanguageCd
            End If
            If Len(Trim(strLangCd)) = 0 Then strLangCd = CdCst.LanguageCd.DefaultLang

            Dim objSysCtrl As New KHSystem
            '稼動状況確認
            Dim objUserInfo1 As KHSessionInfo.UserInfo = httpCon.Session(CdCst.SessionInfo.Key.UserInfo)
            If Not objUserInfo1.BaseCd Is Nothing Then
                Select Case objSysCtrl.fncOpeStateChk(objConBase, objUserInfo1.BaseCd)
                    Case CdCst.OpeState.Operating
                    Case CdCst.OpeState.Stopping
                        'メンテナンス中
                        Me.Session(CdCst.SessionInfo.Key.ErrorCode) = ErrCode.E0020
                        Throw New System.ApplicationException(ErrCode.E0020)
                    Case CdCst.OpeState.Trouble
                        'トラブル中
                        Me.Session(CdCst.SessionInfo.Key.ErrorCode) = ErrCode.E0030
                        Throw New System.ApplicationException(ErrCode.E0030)
                        Call Page_Load(Me, Nothing)
                    Case Else
                        Me.Session(CdCst.SessionInfo.Key.ErrorCode) = ErrCode.E9999
                        Throw New System.ApplicationException(ErrCode.E9999)
                        Call Page_Load(Me, Nothing)
                End Select
            End If

            Select Case intReturn
                Case 0
                    'メニュー画面へ遷移
                    clsUser = Nothing
                    RaiseEvent GoToMenuPage()
                Case 1
                    If Not httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) Is Nothing AndAlso _
                       Not httpCon.Session(CdCst.SessionInfo.Key.EdiInfo).KeyInfo.Equals(String.Empty) Then
                        'EDI連動の場合
                        Session.Add(CdCst.SessionInfo.Key.ErrorCode, CdCst.SessionInfo.Key.LoginError)
                        RaiseEvent Goto_ErrPage(String.Empty)
                    Else
                        'ログイン失敗
                        AlertMessage(strLangCd, "W0030") 'ログイン情報が間違っています。
                    End If
                Case 2
                    'パスワードを隠しエリアに保持しておく
                    Me.hiddenPasswd.Value = Me.txtPasswd.Text
                    ConfirmMessage(strLangCd, "I0010", "0") '既に同じユーザがログインしています。
                Case 3
                    'パスワードを隠しエリアに保持しておく
                    Me.hiddenPasswd.Value = Me.txtPasswd.Text
                    Me.hiddenNewPasswd.Value = Me.txtNewPasswd.Text

                    ''Manifoldtest
                    'RaiseEvent GoToMenuPage()

                    'EDI
                    If (Not httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) Is Nothing AndAlso _
                        Not httpCon.Session(CdCst.SessionInfo.Key.EdiInfo).KeyInfo.Equals(String.Empty)) Then
                        '代理店EDIシステムから起動する場合
                        RaiseEvent GoToMenuPage()
                    ElseIf strParaE.Equals(WebforGS) Then
                        'WebforGSで起動する場合
                        RaiseEvent GoToMenuPage()
                    ElseIf strParaE.Contains("EDI_") Then
                        'P3システムから起動する場合
                        Dim strUID As String = txtUserID.Text
                        Dim strUPD As String = txtPasswd.Text

                        strUPD &= "EDI"
                        If strUID = strUPD Then
                            RaiseEvent GoToMenuPage()
                        End If
                    Else
                        Call ConfirmMessage(strLangCd, "W5250", "1") 'パスワードの更新時期です。新しいパスワードに更新して下さい。
                    End If
            End Select

            '選択した機種情報を削除
            If Me.Session("KisyuInfo") IsNot Nothing Then
                Me.Session.Remove("KisyuInfo")
            End If
        Catch ex As Exception
            'エラー画面に遷移する
            clsUser = Nothing
            RaiseEvent Goto_ErrPage(ex.Message)
        Finally
            clsUser = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 受注EDI・WebforGS連携の場合は自動でログイン認証を行い、形番引当画面へ遷移させる
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subJuchuEDIProc()
        Dim objEdiInfo As KHSessionInfo.EdiInfo
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current

        Try
            If strParaE Is Nothing OrElse strParaE.Equals(String.Empty) Then
                Exit Sub
            End If
            objEdiInfo.KeyInfo = String.Empty

            If strParaE.Equals(WebforGS) OrElse strParaE.Length <= 20 Then
            Else
                '受注EDI連携のとき
                'セッション情報に確保
                Session.Clear()
                objEdiInfo.KeyInfo = strParaE
                httpCon.Session(CdCst.SessionInfo.Key.EdiInfo) = objEdiInfo
            End If

            'ログイン処理
            Me.txtUserID.Text = strUserID
            Me.txtPasswd.Text = strPassword
            Call Button1_Click(Nothing, Nothing)

        Catch ex As Exception
            'エラー画面に遷移する
            RaiseEvent Goto_ErrPage(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' エラーメッセージを表示する
    ''' </summary>
    ''' <param name="strMessageID"></param>
    ''' <remarks></remarks>
    Private Sub AlertMessage(ByVal strLang As String, ByVal strMessageID As String)
        'メッセージ内容の取得
        Dim strMessage As String = String.Empty
        Try
            strMessage = ClsCommon.fncGetMsg(strLang, strMessageID)

            If Session("TestMode") Is Nothing OrElse Session("TestMode").Equals(2) Then
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "Alert", "alert('" & strMessage & "');", True)
            Else
                'マニホールドテスト専用
                If Not Session("ManifoldKataban") Is Nothing Then
                    Dim listKataban As New ManifoldKataban(Session("ManifoldKataban"))
                    IO.File.AppendAllText(My.Settings.LogFolder & "Error.txt", strMessage & ControlChars.Tab & listKataban.ToString & ControlChars.NewLine)
                    'エラーも処理完了の一種
                    Session("EventEndFlg") = True
                    GC.Collect()
                End If
            End If
        Catch ex As Exception
            RaiseEvent Goto_ErrPage(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 確認メッセージを表示する
    ''' </summary>
    ''' <param name="strMessageID">メッセージコード</param>
    ''' <param name="strDiv">呼び出すボタン区分</param>
    ''' <remarks></remarks>
    Private Sub ConfirmMessage(ByVal strLang As String, ByVal strMessageID As String, ByVal strDiv As String)
        'メッセージ
        Dim strMessage As String = String.Empty
        Dim sbScript As New StringBuilder
        Dim strControlFullname As String = "ctl00_ContentDetail_WebUC_Login_"
        Try
            'メッセージを取得
            strMessage = ClsCommon.fncGetMsg(strLang, strMessageID)
            'Javascriptコードの作成
            sbScript.Append(" if(confirm('" & strMessage & "')){" & vbCrLf)
            sbScript.Append("    document.getElementById('" & strControlFullname & "hiddenOver').value = '1';" & vbCrLf)
            If strDiv = "0" Then
                sbScript.Append("    document.getElementById('" & strControlFullname & "Button1').click();" & vbCrLf)
            Else
                sbScript.Append("    document.getElementById('" & strControlFullname & "Button3').click();" & vbCrLf)
            End If
            sbScript.Append(" }else{" & vbCrLf)
            sbScript.Append("    //保持していたパスワードを破棄する" & vbCrLf)
            sbScript.Append("    document.getElementById('" & strControlFullname & "hiddenPasswd').value = '';" & vbCrLf)
            sbScript.Append("    document.getElementById('" & strControlFullname & "hiddenNewPasswd').value = '';" & vbCrLf)
            sbScript.Append(" }" & vbCrLf)

            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "Confirm", sbScript.ToString, True)
        Catch ex As Exception
            RaiseEvent Goto_ErrPage(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' EnterKeyの無効化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInitScript()
        Dim strKeyDown As String
        Dim strFocus As String
        Dim strBlur As String

        Try
            'テキストエリアのEnterKey無効化
            strKeyDown = "if (event.keyCode == 13){return false;}else{return true;}"
            Me.txtUserID.Attributes.Add(CdCst.JavaScript.OnKeyDown, strKeyDown)
            Me.txtPasswd.Attributes.Add(CdCst.JavaScript.OnKeyDown, strKeyDown)
            Me.txtNewPasswd.Attributes.Add(CdCst.JavaScript.OnKeyDown, strKeyDown)
            Me.txtNewPasswdRe.Attributes.Add(CdCst.JavaScript.OnKeyDown, strKeyDown)
            'OnFocus
            strFocus = "this.select(); this.style.background = '#FFCC33';"
            Me.txtUserID.Attributes.Add(CdCst.JavaScript.OnFocus, strFocus)
            Me.txtPasswd.Attributes.Add(CdCst.JavaScript.OnFocus, strFocus)
            Me.txtNewPasswd.Attributes.Add(CdCst.JavaScript.OnFocus, strFocus)
            Me.txtNewPasswdRe.Attributes.Add(CdCst.JavaScript.OnFocus, strFocus)
            'OnBlur
            strBlur = "this.style.background = '#FFFFC0';"
            Me.txtUserID.Attributes.Add(CdCst.JavaScript.OnBlur, strBlur)
            Me.txtPasswd.Attributes.Add(CdCst.JavaScript.OnBlur, strBlur)
            Me.txtNewPasswd.Attributes.Add(CdCst.JavaScript.OnBlur, strBlur)
            Me.txtNewPasswdRe.Attributes.Add(CdCst.JavaScript.OnBlur, strBlur)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' セッションのクリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearSession()
        If Not Session.Item("KisyuInfo") Is Nothing Then
            Session.Remove("KisyuInfo")
        End If
    End Sub
End Class