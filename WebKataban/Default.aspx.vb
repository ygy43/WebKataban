Imports System.Web.HttpUtility

Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim strPath As String                                           'メイン画面パス
        Dim strFullUrl As String                                        'URL
        Dim strAppPath As String                                        'システムパス
        Dim strJS As String                                             'Javascript
        Dim strUserID As String                                         'ユーザー名
        Dim strPassword As String                                       'パスワード
        Dim strMacAddress As String                                     'MAC
        Dim strSerialNo As String                                       'シリーズNo.
        Dim strEdiKey As String                                         '受注EDI連携

        'パラメータの取得
        strUserID = UrlEncode(Request.QueryString("A"))
        strPassword = UrlEncode(Request.QueryString("B"))
        strMacAddress = UrlEncode(Request.QueryString("C"))
        strSerialNo = UrlEncode(Request.QueryString("D"))
        strEdiKey = UrlEncode(Request.QueryString("E"))

        'URL取得
        strAppPath = System.Web.HttpContext.Current.Request.ApplicationPath.TrimEnd("/") & "/"
        strPath = strAppPath & "Main.aspx"

        '受注EDI連携
        strFullUrl = strPath & "?a=" & strUserID & "&b=" & strPassword & "&c=" & strMacAddress & "&d=" & strSerialNo & "&e=" & strEdiKey

        'JavaScript生成
        strJS = ""
        strJS = strJS & "window.open('" & strFullUrl & "'"
        strJS = strJS & "," & "'Kataban'"
        strJS = strJS & "," & "'');"
        strJS = strJS & "window.opener = true;" & vbCrLf
        strJS = strJS & "(window.open('', '_self').opener = window).close(); " & vbCrLf

        'JavaScript登録
        Me.ClientScript.RegisterStartupScript(Me.GetType(), "openWindow", strJS, True)

        'クライアントキャッシュしない
        Me.Response.Cache.SetCacheability(HttpCacheability.NoCache)

    End Sub

End Class