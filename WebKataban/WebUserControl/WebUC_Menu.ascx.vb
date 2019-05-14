Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class WebUC_Menu
    Inherits KHBase

#Region "プロパティ"

    Public Event GoToErrorPage(strErrMsg As String)

    Private Const intMenuClassInit As Integer = 2           'メニュー(分類)ボタン順序初期値
    Private Const intMenuContentInit As Integer = 3         'メニュー(内容)ボタン順序初期値
    Private bllMenu As New MenuBLL                          'ビジネスロジック
#End Region

    ''' <summary>
    ''' 外部からの呼出
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Me.OnLoad(Nothing)
    End Sub

    ''' <summary>
    ''' ロード処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Me.Visible = False Then Exit Sub
        If Me.objUserInfo.UserId Is Nothing Then Exit Sub
        If Not FormIDCheck() Then Exit Sub

        Try
            Dim blnHikiate As Boolean = False

            '画面ラベル設定
            Call SetFontName(ListMsg, selLang.SelectedValue)
            Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHMenu, selLang.SelectedValue, Me)

            '画面初期設定
            Call Me.subSetInformation()

            If Not Me.IsPostBack Then
                Dim strTest As String = String.Empty
            Else
                '選択言語をセッション情報に設定する
                If Me.Session(CdCst.SessionInfo.Key.LoginInfo) IsNot Nothing Then
                    Me.objLoginInfo.SelectLang = selLang.SelectedValue
                    Me.Session(CdCst.SessionInfo.Key.LoginInfo) = Me.objLoginInfo
                End If
            End If

            '受注EDI セッションが有効のときの処理
            If Me.Session(CdCst.SessionInfo.Key.EdiInfo) IsNot Nothing Then
                blnHikiate = True
            End If
            Select Case Me.objUserInfo.UserClass
                Case CdCst.UserClass.DmAgentRs, CdCst.UserClass.DmAgentSs, CdCst.UserClass.DmAgentBs, CdCst.UserClass.DmAgentGs, CdCst.UserClass.DmAgentPs
                    blnHikiate = True
            End Select

            '形番引当画面を起動するためにEDIから起動するフラグを保存
            If blnHikiate Then
                'ADD BY YGY 20140630
                If Me.Session(CdCst.SessionInfo.Key.HikiateFlg) Is Nothing Then
                    Me.Session(CdCst.SessionInfo.Key.HikiateFlg) = blnHikiate
                Else
                    If Me.Session(CdCst.SessionInfo.Key.HikiateFlg) = True Then
                        Me.Session(CdCst.SessionInfo.Key.HikiateFlg) = False
                    End If
                End If
            End If

            'Call SetAllFontName(Me)
        Catch ex As Exception
            AlertMessage(ex)
            'エラー画面に遷移する
            RaiseEvent GoToErrorPage(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 画面初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInformation()
        Try
            'Information設定
            'If Me.objUserInfo.UserClass >= CdCst.UserClass.DmSalesOffice Then
            '    Me.ListMsg.DataSource = bllMenu.fncSelectInformation(objConBase, Me.selLang.SelectedValue)
            'Else
            '    Me.ListMsg.DataSource = New ArrayList
            'End If

            If Me.objUserInfo.UserClass >= CdCst.UserClass.DmSalesOffice Then
                Dim strMsgs As New ArrayList
                strMsgs = bllMenu.fncSelectInformation(objConBase, Me.selLang.SelectedValue)

                For Each strMsg In strMsgs
                    Me.ListMsg.Text &= strMsg & ControlChars.NewLine
                Next

            Else
                Me.ListMsg.Text = String.Empty
            End If
            Me.ListMsg.DataBind()
        Catch ex As Exception
            AlertMessage(ex)
            'エラー画面に遷移する
            RaiseEvent GoToErrorPage(ex.Message)
        End Try
    End Sub

End Class
