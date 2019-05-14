Imports System.Data.SqlClient

Public Class WebUC_Error
    Inherits KHBase

#Region "プロパティ"
    Public strMessage As String = String.Empty
    Public Event Goto_Login()
#End Region

    ''' <summary>
    ''' ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub

        Call SetAllFontName(Me)

        'タイトルの設定
        If selLang.SelectedValue = CdCst.LanguageCd.Japanese Then
            If objLoginInfo.SelectLang Is Nothing Then objLoginInfo.SelectLang = CdCst.LanguageCd.Japanese
            Me.lblTitle.Text = CdCst.Message.Title.Japanese
        Else
            If objLoginInfo.SelectLang Is Nothing Then objLoginInfo.SelectLang = CdCst.LanguageCd.DefaultLang
            Me.lblTitle.Text = CdCst.Message.Title.English
        End If

        'エラーメッセージの設定
        If Session(CdCst.SessionInfo.Key.ErrorCode) Is Nothing Then
            'SYSTEM ERROR
            If selLang.SelectedValue = CdCst.LanguageCd.Japanese Then
                strMessage = CdCst.Message.SystemError.Japanese
            Else
                strMessage = CdCst.Message.SystemError.English
            End If
            strMessage = HidErrMsg.Value
        ElseIf Session(CdCst.SessionInfo.Key.ErrorCode).Equals(CdCst.SessionInfo.Key.LoginError) Then
            'LOGIN ERROR
            strMessage = CdCst.Message.AuthenticationErr2.Japanese
        ElseIf Session(CdCst.SessionInfo.Key.ErrorCode).Equals("TIMEOUT") Then
            'TIMEOUT
            If selLang.SelectedValue = CdCst.LanguageCd.Japanese Then
                strMessage = CdCst.Message.AuthenticationErr.Japanese
            Else
                strMessage = CdCst.Message.AuthenticationErr.English
            End If
        Else
            strMessage = ClsCommon.fncGetMsg(objLoginInfo.SelectLang, Session(CdCst.SessionInfo.Key.ErrorCode))
        End If

        'エラーメッセージの登録
        If objKtbnStrc IsNot Nothing AndAlso _
            objKtbnStrc.strcSelection IsNot Nothing AndAlso _
            objKtbnStrc.strcSelection.strFullKataban IsNot Nothing Then
            Call Me.subErrorContentInsert(objConBase, objUserInfo.UserId, strMessage & objKtbnStrc.strcSelection.strFullKataban)
        Else
            Call Me.subErrorContentInsert(objConBase, objUserInfo.UserId, strMessage)
        End If


        Me.lblMessage.Text = strMessage
    End Sub

    ''' <summary>
    ''' エラー情報追加処理
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strException">エラー内容</param>
    ''' <remarks>エラー情報をエラー内容テーブルに追加する</remarks>
    Private Sub subErrorContentInsert(objConBase As SqlConnection, ByVal strUserId As String, _
                                      ByVal strException As String)
        Dim objCmd As SqlCommand
        Dim strErrMessage As String
        Try
            'メッセージ内容設定
            'strErrMessage = objException.Message.ToString & CdCst.Sign.Colon
            'strErrMessage = strErrMessage & objException.StackTrace.ToString
            strErrMessage = strException

            'DB接続文字列の取得
            objCmd = New SqlCommand(CdCst.DB.SPL.KHErrorContentIns, objConBase)
            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = IIf(strUserId Is Nothing, "Missing", strUserId)
                .Parameters.Add("@ErrorContent", SqlDbType.VarChar, 2000).Value = strErrMessage
            End With
            '実行
            objCmd.ExecuteNonQuery()
        Catch ex As Exception
        Finally
            objCmd = Nothing
        End Try

    End Sub

    ''' <summary>
    ''' ログインボタンを押す
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        '再ログイン
        RaiseEvent Goto_Login()
    End Sub
End Class