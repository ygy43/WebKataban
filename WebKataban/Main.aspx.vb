Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports CKDStandard.ManifoldExport.Common
Imports CKDStandard.ManifoldExport.Data
Imports WebKataban.DS_MasterTableAdapters

Public Class _Main
    Inherits Page

#Region "プロパティ"

    Private HT_Menu As ArrayList = Nothing
    'ログイン情報
    Private objLoginInfo As KHSessionInfo.LoginInfo
    'ユーザー情報
    Private objUserInfo As KHSessionInfo.UserInfo
    'KHDBへの接続
    Private objCon As New SqlConnection
    'KHBASEへの接続
    Private objConBase As New SqlConnection
    '形番情報
    Private ReadOnly objKtbnStrc As New KHKtbnStrc
    Private ReadOnly clsUserInf As New KHUser
    'ビジネスロジック
    Private ReadOnly bllDefault As New DefaultBLL
    '時計
    Public watch As New Stopwatch

#End Region

#Region "初期化"

    ''' <summary>
    '''     初期化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
        'ボタン自動サブミットを無効にする
        subSetNoSubmit()
        Call HideAllButton()
        objConBase = New SqlConnection(My.Settings.connkhBase)
        objCon = New SqlConnection(My.Settings.connkhdb)
        objConBase.Open()
        objCon.Open()
        Call SetObjCon()
        Call fncMakeLanguageList(selLang.SelectedValue)
        ScriptManager1.RegisterPostBackControl(Me.btnDownload)
        ScriptManager1.AsyncPostBackTimeout = 360000
        'Footer
        Me.pageFooter.Visible = My.Settings.IsShowFooter
    End Sub

    ''' <summary>
    '''     各画面のDB接続の初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetObjCon() 
        Dim strName() As String = CdCst.strPageIDs
        If Not Me.Controls(0).FindControl("ContentDetail") Is Nothing Then
            For inti = 0 To strName.Length - 1
                If Not Me.Controls(0).FindControl("ContentDetail").FindControl(strName(inti)) Is Nothing Then
                    CType(Me.Controls(0).FindControl("ContentDetail").FindControl(strName(inti)), Object).objCon =
                        objCon
                    CType(Me.Controls(0).FindControl("ContentDetail").FindControl(strName(inti)), Object).objConBase =
                        objConBase
                End If
            Next
        End If
        strName = Nothing
    End Sub

    ''' <summary>
    '''     画面ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            If Not GetUserSession() Then
                Page.Response.Redirect(Page.Request.Url.ToString(), True)
                Exit Sub
            End If

            If Not WebUC_Login.Visible Then
                If WebUC_Youso.Visible = False Then
                    Me.Session.Remove("KtbnStrc")
                End If
                If WebUC_Siyou.Visible = False Then
                    Me.Session.Remove("KtbnStrc_Siyou")
                End If
            End If

            '二重実行防ぐ
            If Not Me.Session("FormID") Is Nothing Then Me.Session.Remove("FormID")
            If Me.HidRunForm.Value.Length > 0 Then
                Select Case Me.HidRunForm.Value
                    Case "1"
                        Me.Session("FormID") = "WebUC_Menu"
                    Case "3", "4"
                        Me.Session("FormID") = "WebUC_Type"
                    Case "5"
                        Me.Session("FormID") = "WebUC_Youso"
                    Case "6"
                        Me.Session("FormID") = "WebUC_Siyou"
                End Select
                Me.HidRunForm.Value = String.Empty
            End If

            '画面ラベル設定
            Dim strLang As String = CdCst.LanguageCd.DefaultLang
            If selLang.SelectedValue <> "" Then strLang = selLang.SelectedValue

            If Not WebUC_Youso.Visible And Not WebUC_Tanka.Visible Then
                Call _
                    KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHMenu_Head, strLang,
                                           Me.Controls(0).FindControl("ContentTitle"))
            End If

        Catch ex As Exception
            Call ShowErrPage(ex.Message)
        End Try
    End Sub

    ''' <summary>
    '''     セッション情報の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetUserSession() As Boolean
        GetUserSession = False
        Try
            If Not WebUC_Login.Visible Then
                If Not clsUserInf.subGetSession(objConBase) Then 'セッション情報取得
                    '失敗したら、エラー画面へ
                    'ShowErrPage("セッション情報取得エラー、再登録してください。")
                    Exit Function
                Else
                    'ユーザー情報取得
                    Me.objUserInfo = clsUserInf.UserInfo
                    'ログイン情報
                    Me.objLoginInfo = clsUserInf.LoginInfo
                    '各画面のユーザ情報の初期化
                    Call SetUserInfo()
                End If

                If Me.objUserInfo.UserId Is Nothing Or Me.objLoginInfo.SessionId Is Nothing Or
                   selLang.SelectedValue Is Nothing Then
                    'セッションTimeOut、エラー画面へ
                    'ShowErrPage("セッションTimeOut、再登録してください。")
                    Exit Function
                End If
            End If
            GetUserSession = True
        Catch ex As Exception
            Call ShowErrPage(ex.Message)
        End Try
    End Function

    ''' <summary>
    '''     各画面のユーザ情報の初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetUserInfo()
        Dim strName() As String = CdCst.strPageIDs
        If Not Me.Controls(0).FindControl("ContentDetail") Is Nothing Then
            For inti = 0 To strName.Length - 1
                If strName(inti) = "WebUC_Login" Then Continue For
                If Not Me.Controls(0).FindControl("ContentDetail").FindControl(strName(inti)) Is Nothing Then
                    CType(Me.Controls(0).FindControl("ContentDetail").FindControl(strName(inti)), Object).objLoginInfo =
                        objLoginInfo
                    CType(Me.Controls(0).FindControl("ContentDetail").FindControl(strName(inti)), Object).objUserInfo =
                        objUserInfo
                End If
            Next
        End If
        strName = Nothing
    End Sub

    ''' <summary>
    '''     ボタンの自動サブミットを無効する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetNoSubmit()
        Me.Button1.UseSubmitBehavior = False
        Me.Button2.UseSubmitBehavior = False
        Me.Button4.UseSubmitBehavior = False
        Me.Button5.UseSubmitBehavior = False
        Me.Button6.UseSubmitBehavior = False
        '匿名ユーザー機種選択ボタン
        Me.Button7.UseSubmitBehavior = False
        CType(Me.WebUC_Error.FindControl("btnClose"), Button).UseSubmitBehavior = False
    End Sub

#End Region

#Region "画面遷移"

    ''' <summary>
    '''     メニューボタンクリックイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btn_Click(sender As Object, e As EventArgs) Handles Button2.Click, Button1.Click,
                                                                    Button3.Click, Button10.Click, Button11.Click,
                                                                    Button12.Click, Button13.Click, Button14.Click,
                                                                    Button15.Click,
                                                                    Button4.Click, Button5.Click, Button6.Click,
                                                                    Button7.Click
        Try
            If Not GetUserSession() Then
                Page.Response.Redirect(Page.Request.Url.ToString(), True)
                Exit Sub
            End If

            If CInt(sender.ID.ToString.Replace("Button", "")) < 6 AndAlso
               CInt(sender.ID.ToString.Replace("Button", "")) <> 2 Then
                Me.Session.Remove("KtbnStrc_Siyou")
                Me.Session.Remove("dt_Comb")
                Me.Session.Remove("DS_Title")
            End If

            ClearHiddenField("Youso")

            Me.Session.Remove("ManifoldKataban")
            Me.Session.Remove("ManifoldKatabanLoop")
            Me.Session.Remove("ManifoldSeriesKey")
            Me.Session.Remove("ManifoldItemKey")
            Me.Session.Remove("TestFlag")
            Me.Session.Remove("TestMode")

            '単価画面のリセット
            Dim labels = New List(Of String) From {"Label6", "Label7", "Label9", "Label10", "Label11", "Label27"}
            SetTankaLabelVisible(labels)

            Select Case CInt(sender.ID.ToString.Replace("Button", ""))
                Case 1 'メニュー
                    Call WebUC_Login_GoToMenuPage()
                    Call HideAllButton()
                    ShowButton(2)        'Logoff
                    ShowButton(3)        '形番引当
                    Me.selLang.Enabled = True
                    Call Me.subButtonSet()
                Case 2 'LogOff
                    Call CloseMe()
                Case 3 '形番引当
                    pnlMaster.Visible = False
                    Call Show_Type()
                    Me.selLang.Enabled = False
                Case 4 '機種選択
                    selLang.Enabled = True
                    Me.Session.Remove("SeriesSelectPageType")
                    Me.Session.Add("SeriesSelectPageType",SeriesSelectPageType.Search)
                    Call Show_Type()
                Case 5 '形番引当
                    Call Back_Youso()
                Case 6 '仕様入力
                    Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)
                    Dim objOption As New KHOptionCtl
                    Dim intMode As Integer = YousoBLL.GetNextFormMode(objKtbnStrc, objOption)
                    Select Case intMode '0：単価見積、1：仕様書画面、2：ロッド先端形状オーダーメイド寸法入力画面
                        Case 1 '仕様入力画面
                            Call Back_Siyou()
                        Case 2 'ロッド先端形状オーダーメイド寸法入力画面+                
                            Call Show_RodEndOrder(objKtbnStrc)
                            WebUC_RodEndOrder.txtRodEndSize.Text = objKtbnStrc.strcSelection.strRodEndOption
                    End Select
                Case 7 '匿名ユーザー機種選択
                    selLang.Enabled = True
                    Me.Session.Remove("SeriesSelectPageType")
                    Me.Session.Add("SeriesSelectPageType",SeriesSelectPageType.List)
                    Call Show_TypeAnonymous()
                Case 10, 11, 12, 13, 14, 15 'マスタメンテ
                    '10 国別生産品マスタメンテナンス
                    '11 為替率マスタメンテナンス
                    '12 情報マスタメンテナンス
                    '13 掛率マスタメンテナンス
                    '14 ユーザーマスタメンテナンス
                    '15 マスタメンテナンス
                    Call HideAllWebUC()
                    Me.WebUC_Master.Visible = True
                    WebUC_Master.HidMode.Value = CInt(sender.ID.ToString.Replace("Button", ""))
                    WebUC_Master.frmInit()
                    Call HideAllButton()

                    'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
                    '    ShowButton(1)        'メニュー
                    'End If

                    ShowButton(2)        'Logoff
                    Me.selLang.Enabled = False
            End Select
        Catch ex As Exception
            Call ShowErrPage(ex.Message)
        End Try
    End Sub

    ''' <summary>
    '''     ログイン画面→メニュー画面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Login_GoToMenuPage() Handles WebUC_Login.GoToMenuPage
        If selLang.SelectedValue = "" Then 'ログイン画面の言語欄非選択する時に、登録ユーザーの言語を反映する
            Try
                'セッション情報取得
                Call clsUserInf.subGetSession(objConBase)
                'ユーザー情報取得
                Me.objUserInfo = clsUserInf.UserInfo
                'ログイン情報
                Me.objLoginInfo = clsUserInf.LoginInfo

                WebUC_Menu.objLoginInfo = objLoginInfo
                WebUC_Menu.objUserInfo = objUserInfo

                Call fncMakeLanguageList(Me.objUserInfo.LanguageCd)
                Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHMenu_Head, Me.objUserInfo.LanguageCd,
                                            Me.Controls(0).FindControl("ContentTitle"))
            Catch ex As Exception
                Call ShowErrPage(ex.Message)
            End Try
        ElseIf Me.objUserInfo.UserId Is Nothing Then
            'セッション情報取得
            Call clsUserInf.subGetSession(objConBase)
            'ユーザー情報取得
            Me.objUserInfo = clsUserInf.UserInfo
            'ログイン情報
            Me.objLoginInfo = clsUserInf.LoginInfo

            WebUC_Menu.objLoginInfo = objLoginInfo
            WebUC_Menu.objUserInfo = objUserInfo

            Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHMenu_Head, selLang.SelectedValue,
                                        Me.Controls(0).FindControl("ContentTitle"))
        End If
        Call HideAllWebUC()
        Me.WebUC_Menu.Visible = True
        Me.WebUC_Menu.frmInit() '画面初期化
        ShowButton(2) 'Logoff
        ShowButton(3) '形番引当
        'メニューを設置する（形番引当とマスタ）
        'Button設定
        Call Me.subButtonSet()

        'ADD BY YGY 20140630
        'EDIシステムから起動する場合機種選択画面へ
        If Not Me.Session(CdCst.SessionInfo.Key.HikiateFlg) Is Nothing Then
            If Not Me.Session(CdCst.SessionInfo.Key.EdiInfo) Is Nothing OrElse
               Me.Session(CdCst.SessionInfo.Key.HikiateFlg) = True Then
                Call btn_Click(Me.Button3, Nothing)
            End If
        End If

        '匿名ユーザーの場合匿名ユーザー機種選択画面へ
        If Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
            Call fncMakeLanguageList(selLang.SelectedValue)
            Call btn_Click(Me.Button7, Nothing)
        End If
    End Sub

    ''' <summary>
    '''     要素画面のOKボタン
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Protected Sub WebUC_Youso_BtnOKGo(intMode As Integer) Handles WebUC_Youso.BtnOKGo
        If Not Me.Session("ManifoldItemKey") Is Nothing Then Me.Session("TestFlag") = Nothing 'マニホールドテスト専用
        Select Case intMode
            Case 0 '単価見積画面
                Call Show_Tanka()
                ShowButton(5) '形番引当
            Case 1 '仕様書画面
                Call Show_Siyou()
            Case 2 'ロッド先端形状オーダーメイド寸法入力画面
                Call Show_RodEndOrder()
        End Select
    End Sub

    ''' <summary>
    '''     要素画面→OutOfOption
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Youso_GotoOutOfOption() Handles WebUC_Youso.GotoOutOfOption
        Call Show_OutOfOption()
    End Sub

    ''' <summary>
    '''     要素画面→RodEnd
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Youso_GotoRodEnd() Handles WebUC_Youso.GotoRodEnd
        Call Show_RodEnd()
    End Sub

    ''' <summary>
    '''     要素画面→RodEnd
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Youso_GotoMotor() Handles WebUC_Youso.GotoMotor
        Call Show_Motor()
    End Sub

    ''' <summary>
    '''     要素画面→Stopper
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Youso_GotoStopper() Handles WebUC_Youso.GotoStopper
        Call Show_Stopper()
    End Sub

    ''' <summary>
    '''     機種選択→要素画面
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub WebUC_Type_GotoYouso() Handles WebUC_Type.GotoYouso
        Call Show_Youso()
    End Sub

    ''' <summary>
    '''     機種選択→単価画面
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub WebUC_Type_GotoTanka() Handles WebUC_Type.GotoTanka
        Call Show_Tanka()
    End Sub

    ''' <summary>
    '''     匿名ユーザー機種選択→要素選択画面
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub WebUC_TypeAnonymous_GotoYouso() Handles WebUC_TypeAnonymous.GotoYouso
        sellang.Enabled = False
        Call Show_Youso()
    End Sub

    ''' <summary>
    '''     先端特注→要素画面、オプション外→要素画面、Stopper→要素画面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_RodEnd_GotoYouso() Handles WebUC_RodEnd.BacktoYouso, WebUC_OutOfOption.BacktoYouso,
                                                 WebUC_Stopper.BacktoYouso, WebUC_Motor.BacktoYouso
        Call Back_Youso()
    End Sub

    ''' <summary>
    '''     エラー画面→登録画面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Error_Goto_Login() Handles WebUC_Error.Goto_Login
        Call HideAllWebUC()
        Me.WebUC_Login.Visible = True
        Me.selLang.Enabled = True
        Call CloseMe(False, False)
        Call HideAllButton()
        'CHANGED BY YGY 20141106
        'ShowButton(1)        'メニュー
        'ShowButton(2)        'Logoff
        WebUC_Login.Page_Load(WebUC_Login, Nothing)
    End Sub

    ''' <summary>
    '''     単価画面→価格積上げ画面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Tanka_GotoCopyPrice() Handles WebUC_Tanka.GotoCopyPrice
        Call Show_CopyPrice(1)
    End Sub

    ''' <summary>
    '''     ISO単価画面→価格積上げ画面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_ISOTanka_GotoCopyPrice() Handles WebUC_ISOTanka.GotoCopyPrice
        Call Show_CopyPrice(2)
    End Sub

    ''' <summary>
    '''     単価画面→価格積上げ画面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Tanka_GotoPriceDetail() Handles WebUC_Tanka.GotoPriceDetail
        Call Show_PriceDetail(1)
    End Sub

    ''' <summary>
    '''     単価画面→価格積上げ画面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_ISOTanka_GotoPriceDetail() Handles WebUC_ISOTanka.GotoPriceDetail
        Call Show_PriceDetail(2)
    End Sub

    ''' <summary>
    '''     価格積上げ画面→単価画面
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub WebUC_PriceCopy_BackToTanka(intMode As Integer) _
        Handles WebUC_PriceCopy.BackToTanka, WebUC_PriceDetail.BackToTanka
        Call HideAllWebUC()
        '二重実行防ぐ
        Select Case intMode
            Case 1
                Me.WebUC_Tanka.Visible = True
                Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
                WebUC_Tanka.objKtbnStrc = objKtbnStrc
                WebUC_Tanka.selLang = Me.selLang
                WebUC_Tanka.Page_Load(WebUC_Tanka, Nothing)
            Case 2
                Me.WebUC_ISOTanka.Visible = True
                Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)
                WebUC_ISOTanka.objKtbnStrc = objKtbnStrc
                WebUC_ISOTanka.selLang = Me.selLang
                WebUC_ISOTanka.Page_Load(WebUC_ISOTanka, Nothing)
        End Select
    End Sub

    ' ''' <summary>
    ' ''' 価格詳細画面から単価画面に戻る
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Private Sub WebUC_PriceDetail_BackToTanka() Handles WebUC_PriceDetail.BackToTanka
    '    Call HideAllWebUC()

    '    Me.WebUC_Tanka.Visible = True
    '    Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
    '    WebUC_Tanka.objKtbnStrc = objKtbnStrc
    '    WebUC_Tanka.selLang = Me.selLang
    '    WebUC_Tanka.Page_Load(WebUC_Tanka, Nothing)
    'End Sub

    ''' <summary>
    '''     ロッド先端形状オーダーメイド寸法入力画面→単価画面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_RodEndOrder_GotoTanka() Handles WebUC_RodEndOrder.GotoTanka
        ShowButton(5) '形番引当
        ShowButton(6) '仕様入力
        Call Show_Tanka()
    End Sub

    ''' <summary>
    '''     ISO単価画面へ
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Siyou_GotoISOTanka() Handles WebUC_Siyou.GotoISOTanka
        ShowButton(5) '形番引当
        ShowButton(6) '仕様入力
        Call Show_ISOTanka()
    End Sub

    ''' <summary>
    '''     単価画面へ
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Siyou_GotoTanka() Handles WebUC_Siyou.GotoTanka
        ShowButton(5) '形番引当
        ShowButton(6) '仕様入力
        Call Show_Tanka(1)
    End Sub

#End Region

#Region "画面表示設定"

    ''' <summary>
    '''     機種選択画面を表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Show_Type()

        Call HideAllWebUC()
        Me.WebUC_Type.Visible = True
        WebUC_Type.frmInit() '画面初期化
        Call HideAllButton()

        'ShowButton(4)        '機種選択（検索）
        ShowButton(7)        '機種選択（一覧）
        ShowButton(2)        'Logoff
    End Sub

    ''' <summary>
    '''     匿名ユーザー機種選択画面を表示
    ''' </summary>
    Private Sub Show_TypeAnonymous()
        '全画面非表示
        Call HideAllWebUC()
        Me.WebUC_TypeAnonymous.Visible = True
        me.WebUC_TypeAnonymous.selLang = me.selLang
        WebUC_TypeAnonymous.FrmInit() '画面初期化
        Call HideAllButton()

        ShowButton(4)        '機種選択（検索）
        'ShowButton(7)        '機種選択（一覧）
        ShowButton(2)        'Logoff
    End Sub

    ''' <summary>
    '''     要素画面に戻る
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Back_Youso()
        Call HideAllWebUC()
        Me.WebUC_Youso.Visible = True
        Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
        WebUC_Youso.objKtbnStrc = objKtbnStrc
        Me.Session("KtbnStrc") = objKtbnStrc
        WebUC_Youso.selLang = Me.selLang
        WebUC_Youso.Page_Load(WebUC_Youso, Nothing)
        Call HideAllButton()

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If
    End Sub

    ''' <summary>
    '''     要素画面を表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Show_Youso()
        'Process Time Test
        watch.Start()

        Call HideAllButton()

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If

        Call HideAllWebUC()
        Me.WebUC_Youso.Visible = True
        Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
        WebUC_Youso.objKtbnStrc = objKtbnStrc
        'WebUC_Youso.HidMaxSelCount.Value = String.Empty
        WebUC_Youso.selLang = Me.selLang
        WebUC_Youso.frmInit() '画面初期化
        If Not Me.Session("ManifoldItemKey") Is Nothing Then 'マニホールドテスト専用
            WebUC_Youso.Page_Load(WebUC_Youso, Nothing)
        End If

        'Process Time Test
        watch.Stop()

        If Me.objUserInfo.UserId = "IDH303" Then
            Debug.WriteLine(My.Settings.LogFolder & "負荷テスト\新.txt", "要素画面：" & watch.Elapsed.ToString & ControlChars.Tab)
        End If
        watch.Reset()
    End Sub

    ''' <summary>
    '''     単価画面を表示する
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub Show_Tanka(Optional intMode As Integer = 0)
        'Process Time Test
        watch.Start()

        Me.Session("decDinRailLength") = WebUC_Siyou.objKtbnStrc.strcSelection.decDinRailLength '2018/03/08追加

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If

        Call HideAllWebUC()
        Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, intMode)
        Me.WebUC_Tanka.Visible = True
        WebUC_Tanka.objKtbnStrc = objKtbnStrc
        WebUC_Tanka.frmInit() '画面初期化

        'Process Time Test
        watch.Stop()
        If Me.objUserInfo.UserId = "IDH303" Then
            Debug.WriteLine(My.Settings.LogFolder & "負荷テスト\新.txt",
                            "単価画面：" & watch.Elapsed.ToString & ControlChars.NewLine)
        End If
        watch.Reset()
    End Sub

    ''' <summary>
    '''     ISO単価画面を表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Show_ISOTanka()

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff
        Call HideAllWebUC()
        Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)
        Me.WebUC_ISOTanka.Visible = True
        WebUC_ISOTanka.objKtbnStrc = objKtbnStrc
        WebUC_ISOTanka.frmInit() '画面初期化
    End Sub

    ''' <summary>
    '''     先端特注画面を表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Show_RodEnd()
        Call HideAllWebUC()
        Me.WebUC_RodEnd.Visible = True
        WebUC_RodEnd.frmInit() '画面初期化
        Call HideAllButton()

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If

        ShowButton(5)       '形番引当
    End Sub

    ''' <summary>
    '''     オプション外画面を表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Show_OutOfOption()
        Call HideAllWebUC()
        Me.WebUC_OutOfOption.Visible = True
        WebUC_OutOfOption.frmInit() '画面初期化
        Call HideAllButton()

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If

        ShowButton(5)       '形番引当
    End Sub

    ''' <summary>
    '''     Stopper画面を表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Show_Stopper()
        Call HideAllWebUC()
        Me.WebUC_Stopper.Visible = True
        WebUC_Stopper.frmInit() '画面初期化
        Call HideAllButton()

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If

        ShowButton(5)       '形番引当
    End Sub

    ''' <summary>
    '''     取付モータ画面を表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Show_Motor()
        Call HideAllWebUC()
        Me.WebUC_Motor.Visible = True
        Me.WebUC_Motor.frmInit() '画面初期化
        Call HideAllButton()

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If

        ShowButton(5)       '形番引当
    End Sub

    ''' <summary>
    '''     価格積上げ表示画面を表示する
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub Show_CopyPrice(intMode As Integer)
        Call HideAllWebUC()

        If intMode = 1 Then
            '一般の場合
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
        Else

            'ISOの場合
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)
        End If

        Me.WebUC_PriceCopy.Visible = True
        WebUC_PriceCopy.objKtbnStrc = objKtbnStrc
        WebUC_PriceCopy.selLang = Me.selLang
        WebUC_PriceCopy.frmInit(intMode) '画面初期化
    End Sub

    ''' <summary>
    '''     価格詳細画面を表示する
    ''' </summary>
    ''' <param name="intMode">
    '''     1:一般
    '''     2:ISO
    ''' </param>
    ''' <remarks></remarks>
    Private Sub Show_PriceDetail(intMode As Integer)
        Call HideAllWebUC()

        If intMode = 1 Then
            '一般の場合
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
        Else
            'ISOの場合
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)
        End If

        Me.WebUC_PriceDetail.Visible = True
        WebUC_PriceDetail.objKtbnStrc = objKtbnStrc
        WebUC_PriceDetail.selLang = Me.selLang
        WebUC_PriceDetail.frmInit(intMode) '画面初期化(ISO/一般)
    End Sub

    ''' <summary>
    '''     ロッド先端形状オーダーメイド寸法入力画面を表示する
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <remarks></remarks>
    Private Sub Show_RodEndOrder(Optional objKtbnStrc As KHKtbnStrc = Nothing)
        Call HideAllWebUC()
        If objKtbnStrc Is Nothing Then
            objKtbnStrc = New KHKtbnStrc
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
        End If
        Me.WebUC_RodEndOrder.Visible = True
        WebUC_RodEndOrder.objKtbnStrc = objKtbnStrc
        WebUC_RodEndOrder.frmInit() '画面初期化
        Call HideAllButton()

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If

        ShowButton(5)       '形番引当
    End Sub

    ''' <summary>
    '''     仕様画面の表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Show_Siyou()
        'Process Time Test
        watch.Start()

        Call HideAllWebUC()
        Me.WebUC_Siyou.Visible = True
        WebUC_Siyou.frmInit() '画面初期化
        Call HideAllButton()

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If

        ShowButton(5)       '形番引当

        'Process Time Test
        watch.Stop()
        If Me.objUserInfo.UserId = "IDH303" Then
            Debug.WriteLine(My.Settings.LogFolder & "負荷テスト\新.txt", "仕様画面：" & watch.Elapsed.ToString & ControlChars.Tab)
        End If
        watch.Reset()
    End Sub

    ''' <summary>
    '''     要素画面に戻る
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Back_Siyou()
        Call HideAllWebUC()
        Me.WebUC_Siyou.Visible = True
        WebUC_Siyou.objKtbnStrc = objKtbnStrc
        Me.Session("KtbnStrc") = objKtbnStrc
        WebUC_Siyou.selLang = Me.selLang
        WebUC_Siyou.frmBack()
        Call HideAllButton()

        'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
        '    ShowButton(1)        'メニュー
        'End If

        ShowButton(2)        'Logoff

        If Me.Session("SeriesSelectPageType")=SeriesSelectPageType.List Then
            ShowButton(7)        '機種選択一覧
        Else
            ShowButton(4)        '機種選択検索
        End If

        ShowButton(5)       '形番引当
    End Sub

#End Region

    ''' <summary>
    '''     LogOff
    ''' </summary>
    ''' <param name="bolCloseFlag"></param>
    ''' <remarks></remarks>
    Private Sub CloseMe(Optional bolCloseFlag As Boolean = True, Optional bolDeleteFlag As Boolean = True)
        Try
            Try
                'セッション情報すべてクリア            'セッションが切れていない場合
                If Not Session(CdCst.SessionInfo.Key.UserInfo) Is Nothing Then
                    'ADD BY YGY 20150619
                    'エラー修正のため一時的に保存
                    If bolDeleteFlag Then
                        'セッション情報取得(ユーザー情報)
                        objUserInfo = Session(CdCst.SessionInfo.Key.UserInfo)
                        'セッション情報取得(ログイン情報)
                        objLoginInfo = Session(CdCst.SessionInfo.Key.LoginInfo)
                        'ログイン情報＆セッション情報削除
                        Call clsUserInf.subUserLogout(objCon, objConBase, objUserInfo.UserId, objLoginInfo.SessionId)
                    End If
                End If

                '一ヶ月前の臨時テーブル情報をクリアする（異常終了する時、残る可能性がある）
                Call bllDefault.subDelErrHistory(objCon, objConBase)

                Me.Session.Clear()
            Catch ex As Exception
            End Try
            '画面を閉じる
            If bolCloseFlag Then
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "close",
                                                    "window.open(""about:blank"",""_self"").close();", True)
            End If
        Catch ex As Exception
            Call ShowErrPage(ex.Message) 'エラー画面に遷移する
        End Try
    End Sub

    ''' <summary>
    '''     EDI送信成功したら、終了
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_EDIReturn() Handles WebUC_Tanka.EDIReturn
        Call CloseMe()
    End Sub

    ''' <summary>
    '''     EDI送信成功したら、終了(ISO)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_ISO_EDIReturn() Handles WebUC_ISOTanka.EDIReturn
        Call CloseMe()
    End Sub

#Region "ファイル出力"

    ''' <summary>
    '''     仕様ファイル出力
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strSiyou"></param>
    ''' <remarks></remarks>
    Private Sub WebUC_SiyouFileOutput(objKtbnStrc As KHKtbnStrc, Optional ByVal strSiyou As String = "") _
        Handles WebUC_Tanka.SiyouFileOutput, WebUC_ISOTanka.SiyouFileOutput

        Try
            '仕様書Excelダウンロード
            If strSiyou.Equals(String.Empty) Then '自動テストする時にダウンロードしない

                '仕様書出力
                If fncCreateManifold(objKtbnStrc) Then

                    Me.Session("strDownloadMode") = "3"
                    'ダウンロード
                    Dim sbScript As New StringBuilder
                    sbScript.Append("fncDownload('" & Me.btnDownload.ClientID & "');")
                    ScriptManager.RegisterStartupScript(Me.UpdatePanelPage, Page.GetType(), "downloadfile",
                                                        sbScript.ToString, True)

                Else

                    Call ClsCommon.WriteErrorLog("E9999", selLang.SelectedValue)

                End If

            End If

        Catch ex As Exception

            Call ShowErrPage(ex.Message) 'エラー画面に遷移する

        End Try
    End Sub

    ''' <summary>
    '''     ＪＳＯＮファイル出力（年末はＰＯＳＴ送信に変更）
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <remarks></remarks>
    Private Sub WebUC_JSONFileOutput(objKtbnStrc As KHKtbnStrc, strName As String) Handles WebUC_Tanka.JSONFileOutput

        Try
            Dim strFileDir As String = My.Settings.FileOutputDir
            Dim strFilePath As String = strFileDir & Me.objUserInfo.UserId & CdCst.File.JsonExtension
            Dim bolDownload As Boolean
            
            'ディレクトリ存在確認
            If Directory.Exists(strFileDir) = False Then
                '存在しない場合は作成する
                Directory.CreateDirectory(strFileDir)
            End If

            'ファイル存在確認
            If IO.File.Exists(strFilePath) Then
                'bolDownload = OverwriteConfirm(strName, "Button12")
                bolDownload = True
            Else
                'ファイルが存在しないとき
                bolDownload = True
            End If

            If bolDownload = True Then

                '出力情報の作成
                Dim strFileData As String = subMakeJSONData(objKtbnStrc)

                If strFileData.Length > 0 Then
                    IO.File.WriteAllText(strFilePath, strFileData, Encoding.UTF8)
                End If

                Me.Session("DownloadFlg") = String.Empty
                Me.Session("strDownloadMode") = "4"
                Me.Session("strJsonFileName") = objKtbnStrc.strcSelection.strFullKataban & CdCst.File.JsonExtension

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "downloadfile", "fncDownload('" & Me.btnDownload.ClientID & "');", True)
            End If

        Catch ex As Exception
            Me.Session("DownloadFlg") = String.Empty
            Call ShowErrPage(ex.Message) 'エラー画面に遷移する
        End Try
    End Sub

    Private Sub SendJsonToCad()

        Using client As New WebClient
            client.Headers(HttpRequestHeader.ContentType) = "application/json"
            Dim json As String = subMakeJSONData(objKtbnStrc)
            Dim response = client.UploadString("https://service.web2cad.co.jp/maker/ckd_assy/call_assy.php/?language=ja", json)
        End Using

    End Sub

    ''' <summary>
    '''     ファイル出力
    ''' </summary>
    ''' <param name="strpath"></param>
    ''' <param name="strXlUserFile"></param>
    ''' <remarks></remarks>
    Private Sub FileOutput(strpath As String, strXlUserFile As String)
        'ダウンロード
        Try
            Response.Clear()
            Response.ContentEncoding = New UTF8Encoding
            Response.ContentType = "application/octet-stream"
            Response.AppendHeader("Content-Disposition", "attachment;filename=" + strXlUserFile)
            Response.Flush()
            Response.WriteFile(strpath)
            Try
                Response.End() '異常エラーになる
            Catch ex As Exception
            End Try
        Catch ex As Exception
            'Call AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    '''     ダウンロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDownload_Click(sender As Object, e As EventArgs) Handles btnDownload.Click
        Dim strUser As String = Me.objUserInfo.UserId
        Dim intMode As String = Me.Session("strDownloadMode")

        Select Case intMode
            Case "1"
                Dim strSBOFolder As String = My.Settings.FileOutputFolder
                Dim strSBOFilePath As String = String.Empty
                Dim strDownLoadFileName As String = My.Settings.DownLoadFileName
                'SBOインターフェースファイルパス設定
                strSBOFilePath = strSBOFolder & strUser & CdCst.File.TextExtension
                'ダウンロード
                FileOutput(strSBOFilePath, strDownLoadFileName)
            Case "2"
                Dim strFileDir As String = My.Settings.FileOutputDir
                Dim strFilePath As String = String.Empty
                Dim strFileName As String = My.Settings.FileOutputName
                strFilePath = strFileDir & strUser & CdCst.File.CsvExtension
                FileOutput(strFilePath, strFileName)
            Case "3"
                'Dim strXlDir As String = My.Settings.ExcelDir
                'Dim strXlUserDir As String = My.Settings.ExcelUserDir
                'Dim strXlUserFile As String = My.Settings.ExcelUserFile
                Dim strpath As String = HttpContext.Current.Server.MapPath("TempFiles") & "\" & strUser & "_" &
                                        CdCst.strExcelTmpFileName
                FileOutput(strpath, CdCst.strExcelTmpFileName)
            Case "4" 'JSONファイル
                Dim strFileDir As String = My.Settings.FileOutputDir
                Dim strFilePath As String = String.Empty
                Dim strFileName As String = Me.Session("strJsonFileName")
                strFilePath = strFileDir & strUser & CdCst.File.JsonExtension
                FileOutput(strFilePath, strFileName)
        End Select
        Me.Session("strDownloadMode") = String.Empty
    End Sub

    ''' <summary>
    '''     I/Fファイル出力
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strName"></param>
    ''' <param name="strNewPlace"></param>
    ''' <remarks></remarks>
    Private Sub WebUC_IFFileOutput(objKtbnStrc As KHKtbnStrc, strName As String, strNewPlace As String) _
        Handles WebUC_Tanka.IFFileOutput, WebUC_ISOTanka.IFFileOutput
        Dim strSBOFolder As String = My.Settings.FileOutputFolder
        Dim strSBOFilePath As String = String.Empty
        Dim strDownLoadFileName As String = My.Settings.DownLoadFileName
        Dim bolDownload As Boolean

        Try
            bolDownload = False

            'SBOインターフェースファイルパス設定
            strSBOFilePath = strSBOFolder & Me.objUserInfo.UserId & CdCst.File.TextExtension

            'ディレクトリ存在確認
            If Directory.Exists(strSBOFolder) = False Then
                '存在しない場合は作成する
                Directory.CreateDirectory(strSBOFolder)
            End If

            ' ファイル存在確認
            If IO.File.Exists(strSBOFilePath) = False Then
                'SBOインターフェースファイル出力フラグをonにする
                bolDownload = True
            Else
                bolDownload = OverwriteConfirm(strName, "Button3")
            End If

            If bolDownload = True Then
                Me.Session("DownloadFlg") = String.Empty
                'SBOインターフェースファイル出力
                '出荷場所の中国生産品対応
                Dim bolApp = CType(Me.AppFlg.Value, Boolean)

                'SBO出力内容の作成
                Dim strOutputText As String = fncCreateIFOutput(objKtbnStrc, strNewPlace)

                'ファイルOpen
                If strOutputText.Length > 0 Then
                    If bolApp Then
                        IO.File.AppendAllText(strSBOFilePath, strOutputText, Encoding.GetEncoding("Shift-Jis"))
                    Else
                        IO.File.WriteAllText(strSBOFilePath, strOutputText, Encoding.GetEncoding("Shift-Jis"))
                    End If
                    Me.Session("strDownloadMode") = "1"
                    Dim sbScript As New StringBuilder
                    sbScript.Append("fncDownload('" & Me.btnDownload.ClientID & "');")
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "downloadfile", sbScript.ToString, True)
                End If
            End If
        Catch ex As Exception
            Me.Session("DownloadFlg") = String.Empty
            Call ShowErrPage(ex.Message) 'エラー画面に遷移する
        End Try
    End Sub

    ''' <summary>
    '''     I/F出力データの作成
    ''' </summary>
    ''' <param name="strNewPlace"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreateIFOutput(objKtbnStrc As KHKtbnStrc,
                                       strNewPlace As String) As String

        Dim strResult As String = String.Empty

        'FOB対応
        Dim strFobPrice As String = 0
        Dim strCountryCd As String = String.Empty
        Dim strSessionIDFob As String = String.Empty

        'セッション名の取得
        Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
            Case "05", "06"
                'ISOの場合
                strSessionIDFob = "strPriceListFobISO"
            Case Else
                strSessionIDFob = "strPriceListFob"
        End Select

        'セッション情報の取得
        If Not (Session(strSessionIDFob) Is Nothing) Then
            strFobPrice = Session(strSessionIDFob)
            Session.Remove(strSessionIDFob)
        Else
            strFobPrice = 0
        End If
        If Not (Session("strCountryCod") Is Nothing) Then
            strCountryCd = Session("strCountryCod")
            Session.Remove("strCountryCod")
        Else
            strCountryCd = String.Empty
        End If

        'SBO出力内容の作成
        strResult = KHSBOInterface.fncSBOInterfaceGet(objCon, objKtbnStrc, strFobPrice, strCountryCd,
                                                      Me.objUserInfo.OfficeCd, Me.objUserInfo.UserId,
                                                      Me.objLoginInfo.SessionId, strNewPlace)

        Return strResult
    End Function

#Region "仕様書出力関連"

    ''' <summary>
    '''     マニホールド仕様書の作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreateManifold(objKtbnStrc As KHKtbnStrc) As Boolean

        Dim blnResult = False
        Dim objExcel As New KHExcelCtl

        Dim strSpecData As String = fncCreateIFOutput(objKtbnStrc, String.Empty) _
        'SpecData

        Dim dtItem As New DS_Master.kh_item_mstDataTable                                  'Itemマスタデータ
        Dim lstData As List(Of ManifoldBaseData)                     '仕様データ

        Dim strDBLanguage As String = String.Empty
        Dim strResourceLanguage As String = String.Empty

        subGetLanguage(selLang.SelectedValue, strDBLanguage, strResourceLanguage)

        'Itemデータ
        dtItem = fncGetItemData(strDBLanguage, objKtbnStrc.strcSelection.strSpecNo)

        '仕様情報
        lstData = fncGetBaseData(objKtbnStrc.strcSelection.strFullKataban, objKtbnStrc.strcSelection.strSpecNo,
                                 strSpecData)

        '仕様書Excel作成出力
        blnResult = objExcel.fncExportManifold(strResourceLanguage, Define.SystemID.WEBKATAHIKI, lstData, dtItem,
                                               Me.objUserInfo.UserId)

        Return blnResult
    End Function

    ''' <summary>
    '''     ItemMasterデータの取得
    ''' </summary>
    ''' <param name="strLanguage">言語</param>
    ''' <param name="strSpecNo">SpecNo.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetItemData(strLanguage As String, strSpecNo As String) As DataTable
        'Itemデータ
        Dim dtItem As New DS_Master.kh_item_mstDataTable

        Using da As New kh_item_mstTableAdapter

            'M_ITEMデータの取得
            dtItem = da.GetDataByLanguageSpecNo(strLanguage, strSpecNo)

        End Using

        Return dtItem
    End Function

    ''' <summary>
    '''     形引システムデータの作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetBaseData(strKataban As String,
                                    strSpecNo As String,
                                    strSpecData As String) As List(Of ManifoldBaseData)

        Dim result As New List(Of ManifoldBaseData)

        Dim item As New ManifoldBaseData(strKataban, strSpecNo, strSpecData)

        result.Add(item)

        Return result
    End Function

    ''' <summary>
    '''     言語の取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subGetLanguage(strLanguage As String,
                               ByRef strDBLanguage As String,
                               ByRef strResourceLanguage As String)

        Select Case strLanguage

            Case Define.DBLanguageCode.SIMPLECHINESE

                strResourceLanguage = Define.ResourceLanguageCode.SIMPLECHINESE
                strDBLanguage = Define.DBLanguageCode.SIMPLECHINESE

            Case Define.DBLanguageCode.TRADITIONALCHINESE

                strResourceLanguage = Define.ResourceLanguageCode.TRADITIONALCHINESE
                strDBLanguage = Define.DBLanguageCode.TRADITIONALCHINESE

            Case Define.DBLanguageCode.ENGLISH

                strResourceLanguage = Define.ResourceLanguageCode.ENGLISH
                strDBLanguage = Define.DBLanguageCode.ENGLISH

            Case Define.DBLanguageCode.JAPANESE

                strResourceLanguage = Define.ResourceLanguageCode.JAPANESE
                strDBLanguage = Define.DBLanguageCode.JAPANESE

            Case Define.DBLanguageCode.KOREAN

                strResourceLanguage = Define.ResourceLanguageCode.KOREAN
                strDBLanguage = Define.DBLanguageCode.KOREAN

        End Select
    End Sub

#End Region

    ''' <summary>
    '''     ファイル出力
    ''' </summary>
    ''' <param name="objKtbnStrc">全ての情報</param>
    ''' <param name="strName">画面ID</param>
    ''' <param name="strOrder">掛率、単価、数量、金額、消費税と合計の情報</param>
    ''' <param name="strPriceList">価格リストの情報</param>
    ''' <param name="intMode">出力タイプ（通常・ISO）</param>
    ''' <remarks></remarks>
    Private Sub WebUC_FileOutput(objKtbnStrc As KHKtbnStrc,
                                 strName As String,
                                 strOrder As String,
                                 strPriceList As String,
                                 intMode As Integer) Handles WebUC_Tanka.FileOutput, WebUC_ISOTanka.FileOutput

        Dim strFileDir As String = My.Settings.FileOutputDir
        Dim strFilePath As String = String.Empty
        Dim strFileName As String = My.Settings.FileOutputName
        Dim bolDownload = False

        Try
            'ファイルパス設定
            strFilePath = strFileDir & Me.objUserInfo.UserId & CdCst.File.CsvExtension

            'ディレクトリ存在確認
            If Directory.Exists(strFileDir) = False Then
                '存在しない場合は作成する
                Directory.CreateDirectory(strFileDir)
            End If

            'ファイル存在確認
            If IO.File.Exists(strFilePath) Then
                bolDownload = OverwriteConfirm(strName, "Button6")
            Else
                'ファイルが存在しないとき
                bolDownload = True
            End If

            If bolDownload = True Then

                Me.Session("DownloadFlg") = String.Empty

                '価格情報をファイルに出力
                Dim bolApp = CType(Me.AppFlg.Value, Boolean)

                '出力情報の作成
                Dim strFileData As String = subMakeFileData(objKtbnStrc, strOrder, strPriceList, intMode, bolApp)

                'ファイルOpen
                If strFileData.Length > 0 Then
                    If bolApp Then
                        IO.File.AppendAllText(strFilePath, strFileData, Encoding.UTF8)
                    Else
                        IO.File.WriteAllText(strFilePath, strFileData, Encoding.UTF8)
                    End If
                    Me.Session("strDownloadMode") = "2"
                    Dim sbScript As New StringBuilder
                    sbScript.Append("fncDownload('" & Me.btnDownload.ClientID & "');")
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "downloadfile", sbScript.ToString, True)
                End If
            End If
        Catch ex As Exception
            Me.Session("DownloadFlg") = String.Empty
            Call ShowErrPage(ex.Message) 'エラー画面に遷移する
        End Try
    End Sub

    ''' <summary>
    '''     ファイル出力データを作成する
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strOrder"></param>
    ''' <param name="intMode"></param>
    ''' <param name="blnAppFlg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function subMakeFileData(objKtbnStrc As KHKtbnStrc,
                                     strOrder As String,
                                     strPriceList As String,
                                     intMode As Integer,
                                     Optional ByVal blnAppFlg As Boolean = False) As String
        Dim strArray() As String
        Dim sbFileData As New StringBuilder
        subMakeFileData = String.Empty

        Try
            '注文情報の取得（掛率～合計）
            strOrder = strOrder.Replace(",", "")

            '注文情報（掛率～合計）を配列に入れる
            strArray = Split(strOrder, CdCst.Sign.Delimiter.Pipe)

            '価格表示レベルの取得
            Dim objUnitPrice As New KHUnitPrice
            Dim strPriceDispLvl() As Boolean = objUnitPrice.fncPriceDispLvlInfoGet(Me.objUserInfo.PriceDispLvl)

            With sbFileData
                '価格リスト項目の出力
                Dim lstPriceInfo As New List(Of KeyValuePair(Of String, String))

                'ラベル表示名の取得
                Dim dt_Title As DataTable = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHFileOutput,
                                                                           selLang.SelectedValue)

                '価格リスト情報の分解
                'ISOの場合は1行目によりタイトルを設定
                If strPriceList.Contains("_") Then
                    lstPriceInfo = fncGetPriceList(strPriceList.Split("_")(0), strPriceDispLvl,
                                                   objKtbnStrc.strcSelection.strMadeCountry, dt_Title)
                Else
                    lstPriceInfo = fncGetPriceList(strPriceList, strPriceDispLvl,
                                                   objKtbnStrc.strcSelection.strMadeCountry, dt_Title)
                End If

                If Not blnAppFlg Then '上書きの場合はヘッダー情報を作成する
                    Dim strTitles As String = String.Empty
                    'ヘッダー情報の作成
                    strTitles = fncFileOutputTitle(lstPriceInfo, dt_Title, objKtbnStrc.strcSelection.strMadeCountry)
                    .AppendLine(strTitles)
                End If

                Select Case intMode
                    Case FileOutputType.Normal
                        '普通の形番
                        .AppendLine(fncCreateOutputNormal(objKtbnStrc, strArray, lstPriceInfo))

                    Case FileOutputType.ISO
                        Dim arrylistPriceList As New ArrayList
                        '全ての価格リスト
                        Dim arrPriceList() As String = strPriceList.Split("_")

                        For Each price As String In arrPriceList
                            '価格リスト
                            Dim lstPrice As New List(Of KeyValuePair(Of String, String))

                            lstPrice = fncGetPriceList(price, strPriceDispLvl, objKtbnStrc.strcSelection.strMadeCountry,
                                                       dt_Title)

                            arrylistPriceList.Add(lstPrice)
                        Next

                        'ISOの形番
                        .AppendLine(fncCreateOutputISO(objKtbnStrc, strOrder, arrylistPriceList))

                End Select
            End With
            subMakeFileData = sbFileData.ToString
        Catch ex As Exception
            Call ShowErrPage(ex.Message) 'エラー画面に遷移する
        End Try
    End Function

    'Private Structure JSONDataInfo
    '    Public AttributeSymbol As String                    '属性記号
    '    Public OptionKataban As String                      'オプション形番
    '    Public PositionInfo As String                       '設置位置
    '    Public Quantity As String                           '使用数
    '    Public OrderNo As String                            '受注No.
    'End Structure
    'Private Shared strcJSONDataInfo(,) As JSONDataInfo

    ''' <summary>
    '''     ＪＳＯＮファイル出力データを作成する
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function subMakeJSONData(objKtbnStrc As KHKtbnStrc) As String

        Dim strTab(4) As String
        Dim sbFileData As New StringBuilder
        Dim intLoopCnt As Integer
        Dim intLoopCnt2 As Integer
        Dim intLoopMax_Option As Integer
        Dim intLoopMax_Station As Integer
        strTab(1) = CdCst.Sign.Delimiter.Tab
        strTab(2) = strTab(1) & CdCst.Sign.Delimiter.Tab
        strTab(3) = strTab(2) & CdCst.Sign.Delimiter.Tab
        strTab(4) = strTab(3) & CdCst.Sign.Delimiter.Tab
        subMakeJSONData = String.Empty


        Try

            intLoopMax_Option = CdCst.Siyou_04.Spacer4   '電磁弁～スペーサ最終行までの最大（2018/9時点では、SpecNo.04のみ対応）

            '属性記号D1～D4の合計分ステーション数
            For intLoopCnt = 1 To intLoopMax_Option
                Select Case objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).ToString
                    Case "D1", "D2", "D3", "D4"
                        intLoopMax_Station += CInt(objKtbnStrc.strcSelection.intQuantity(intLoopCnt))
                End Select
            Next

            With sbFileData

                .AppendLine("{")
                .AppendLine(strTab(1) & ClsCommon.fncAddQuote("ASSY") & ": {")
                .AppendLine(
                    strTab(2) & ClsCommon.fncAddQuote("ASSY-ON") & ": " &
                    ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.strFullKataban) & CdCst.Sign.Delimiter.Comma)
                .AppendLine(strTab(2) & ClsCommon.fncAddQuote("MANIFOLD-BASE") & ": {")
                If objKtbnStrc.strcSelection.decDinRailLength = 0 Then
                    '.AppendLine(strTab(3) & ClsCommon.fncAddQuote("BASE-RAIL-ON") & ": " & ClsCommon.fncAddQuote("BAA"))
                Else
                    .AppendLine(
                        strTab(3) & ClsCommon.fncAddQuote("BASE-RAIL-ON") & ": " & ClsCommon.fncAddQuote("BAA") &
                        CdCst.Sign.Delimiter.Comma)
                    .AppendLine(
                        strTab(3) & ClsCommon.fncAddQuote("BASE-RAIL-LEN") & ": " &
                        ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.decDinRailLength.ToString("#.#")))
                End If
                .AppendLine(strTab(2) & "}" & CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 1 To intLoopMax_Station
                    Dim strValveKataban As String = String.Empty
                    Dim strSpacerKataban As String = String.Empty
                    Dim strAport As String = String.Empty
                    Dim strBport As String = String.Empty

                    .AppendLine(strTab(2) & ClsCommon.fncAddQuote("STATION-" & intLoopCnt.ToString("00")) & ": {")

                    For intLoopCnt2 = 1 To intLoopMax_Option
                        If Mid(objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt2).ToString, intLoopCnt, 1) = "1" _
                            Then
                            If objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt2).Trim <> "" Then
                                strAport = objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt2).Trim
                                strBport = objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt2).Trim
                            End If
                            Select Case objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt2).ToString
                                Case "D1", "D2"
                                    'VALVE
                                    strValveKataban = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt2).Trim
                                Case "GB", "S7", "S3", "S2"
                                    'SPACER
                                    strSpacerKataban = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt2).Trim
                                Case "D3", "D4"
                                    'MASPLATE
                                    .AppendLine(
                                        strTab(3) & ClsCommon.fncAddQuote("MASKING-PLATE-ON") & ": " &
                                        ClsCommon.fncAddQuote(
                                            objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt2).Trim) &
                                        CdCst.Sign.Delimiter.Comma)
                                    .AppendLine(strTab(3) & ClsCommon.fncAddQuote("OPTIONS-ON") & ": [")
                                    .AppendLine(
                                        strTab(4) &
                                        ClsCommon.fncAddQuote(
                                            objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt2).Trim))
                                    .AppendLine(strTab(3) & "]")
                                Case Else
                            End Select
                        End If
                    Next
                    'VALVE
                    If strValveKataban <> "" Then
                        .AppendLine(strTab(3) & ClsCommon.fncAddQuote("VALVE") & ": {")
                        .AppendLine(
                            strTab(4) & ClsCommon.fncAddQuote("VALVE-ON") & ": " &
                            ClsCommon.fncAddQuote(strValveKataban))
                        .AppendLine(strTab(3) & "}" & CdCst.Sign.Delimiter.Comma)
                        'SPACER
                        If strSpacerKataban <> "" Then
                            .AppendLine(
                                strTab(3) & ClsCommon.fncAddQuote("SPACER-ON") & ": " &
                                ClsCommon.fncAddQuote(strSpacerKataban) & CdCst.Sign.Delimiter.Comma)
                            .AppendLine(strTab(3) & ClsCommon.fncAddQuote("OPTIONS-ON") & ": [")
                            .AppendLine(strTab(4) & ClsCommon.fncAddQuote(strSpacerKataban) & CdCst.Sign.Delimiter.Comma)
                            .AppendLine(strTab(4) & ClsCommon.fncAddQuote(strValveKataban))
                        Else
                            .AppendLine(strTab(3) & ClsCommon.fncAddQuote("OPTIONS-ON") & ": [")
                            .AppendLine(strTab(4) & ClsCommon.fncAddQuote(strValveKataban))
                        End If
                        'CX
                        If strAport <> "" Then
                            .AppendLine(strTab(3) & "]" & CdCst.Sign.Delimiter.Comma)
                            .AppendLine(strTab(3) & ClsCommon.fncAddQuote("ST-A-PORT") & ": {")
                            .AppendLine(
                                strTab(4) & ClsCommon.fncAddQuote("ST-PIPING-SIZE") & ": " &
                                ClsCommon.fncAddQuote(strAport))
                            .AppendLine(strTab(3) & "}" & CdCst.Sign.Delimiter.Comma)
                            .AppendLine(strTab(3) & ClsCommon.fncAddQuote("ST-B-PORT") & ": {")
                            .AppendLine(
                                strTab(4) & ClsCommon.fncAddQuote("ST-PIPING-SIZE") & ": " &
                                ClsCommon.fncAddQuote(strBport))
                            .AppendLine(strTab(3) & "}")
                        Else
                            .AppendLine(strTab(3) & "]")
                        End If
                    End If
                    If intLoopCnt = intLoopMax_Station Then
                        .AppendLine(strTab(2) & "}")
                    Else
                        .AppendLine(strTab(2) & "}" & CdCst.Sign.Delimiter.Comma)
                    End If
                Next
                .AppendLine(strTab(1) & "}")
                .Append("}")

            End With
            subMakeJSONData = sbFileData.ToString
        Catch ex As Exception
            Call ShowErrPage(ex.Message) 'エラー画面に遷移する
        End Try
    End Function

    ''' <summary>
    '''     上書き確認
    ''' </summary>
    ''' <param name="strName"></param>
    ''' <param name="strBtnName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function OverwriteConfirm(strName As String, strBtnName As String) As Boolean
        OverwriteConfirm = False
        Dim sbScript As New StringBuilder
        Try
            If Me.Session("DownloadFlg") Is Nothing OrElse Me.Session("DownloadFlg") <> "1" Then
                Me.Session("DownloadFlg") = "1"
                'ダウンロードする指示が何もないとき(1回目)
                Dim strMessage As String = ClsCommon.fncGetMsg(selLang.SelectedValue, "I0150")
                sbScript.Append("fncOverwriteConfirm('" & strMessage & "','" & strName & "_','" & strBtnName & "');")
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "download", sbScript.ToString, True)
            Else
                'ダウンロードする指示をユーザが出したとき(2回目)
                OverwriteConfirm = True
            End If
        Catch ex As Exception
            Call ShowErrPage(ex.Message) 'エラー画面に遷移する
        End Try
    End Function

#End Region

    ''' <summary>
    '''     ボタン全部非表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub HideAllButton()
        If Not Me.Controls(0).FindControl("ContentTitle") Is Nothing Then
            For inti = 1 To 20
                If Not Me.Controls(0).FindControl("ContentTitle").FindControl("Button" & inti) Is Nothing Then
                    Me.Controls(0).FindControl("ContentTitle").FindControl("Button" & inti).Visible = False
                End If
                If Not Me.Controls(0).FindControl("ContentTitle").FindControl("lbl" & inti) Is Nothing Then
                    Me.Controls(0).FindControl("ContentTitle").FindControl("lbl" & inti).Visible = False
                End If

            Next
        End If
    End Sub

    ''' <summary>
    '''     ボタン表示
    ''' </summary>
    ''' <param name="intID"></param>
    ''' <remarks></remarks>
    Private Sub ShowButton(intID As Integer)
        If Not Me.Controls(0).FindControl("ContentTitle") Is Nothing Then
            If Not Me.Controls(0).FindControl("ContentTitle").FindControl("Button" & intID) Is Nothing Then
                Me.Controls(0).FindControl("ContentTitle").FindControl("Button" & intID).Visible = True
            End If
            If Not Me.Controls(0).FindControl("ContentTitle").FindControl("lbl" & intID) Is Nothing Then
                Me.Controls(0).FindControl("ContentTitle").FindControl("lbl" & intID).Visible = True
            End If
        End If
    End Sub

    ''' <summary>
    '''     ボタン非表示
    ''' </summary>
    ''' <param name="intID"></param>
    ''' <remarks></remarks>
    Private Sub HideButton(intID As Integer)
        If Not Me.Controls(0).FindControl("ContentTitle") Is Nothing Then
            If Not Me.Controls(0).FindControl("ContentTitle").FindControl("Button" & intID) Is Nothing Then
                Me.Controls(0).FindControl("ContentTitle").FindControl("Button" & intID).Visible = False
            End If
            If Not Me.Controls(0).FindControl("ContentTitle").FindControl("lbl" & intID) Is Nothing Then
                Me.Controls(0).FindControl("ContentTitle").FindControl("lbl" & intID).Visible = False
            End If
        End If
    End Sub

    ''' <summary>
    '''     子画面全部非表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub HideAllWebUC()
        Dim strName() As String = CdCst.strPageIDs
        If Not Me.Controls(0).FindControl("ContentDetail") Is Nothing Then
            For inti = 0 To strName.Length - 1
                If Not Me.Controls(0).FindControl("ContentDetail").FindControl(strName(inti)) Is Nothing Then
                    Me.Controls(0).FindControl("ContentDetail").FindControl(strName(inti)).Visible = False
                End If
            Next
        End If
        strName = Nothing
    End Sub

    ''' <summary>
    '''     言語欄変更イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub selLang_SelectedIndexChanged(sender As Object, e As EventArgs) Handles selLang.SelectedIndexChanged
        If Not selLang.SelectedValue Is Nothing AndAlso selLang.SelectedValue.Length > 0 Then
            Call fncMakeLanguageList(selLang.SelectedValue)
            if me.WebUC_TypeAnonymous.Visible = true
                me.WebUC_TypeAnonymous.FrmInit()
            End If
        End If
    End Sub

    ''' <summary>
    '''     ドロップダウンリスト作成
    ''' </summary>
    ''' <param name="strLang"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncMakeLanguageList(strLang As String) As Integer
        Dim dtLang As New DataTable
        Dim intReturn = 0

        Try
            '言語リストの取得
            dtLang = bllDefault.fncSelectLanguageList(objConBase, strLang)

            Me.selLang.Items.Clear()
            Me.selLang.DataSource = dtLang
            Dim dr As DataRow = dtLang.NewRow
            dtLang.Rows.InsertAt(dr, 0)
            Me.selLang.DataBind()

            If Len(Trim(strLang)) > 0 Then
                Me.selLang.SelectedValue = strLang
                Dim strMsg As String = ClsCommon.fncGetMsg(strLang, "I0060")
                Me.Button2.Attributes.Clear()
                'ログオフボタン設定
                Me.Button2.Attributes.Add(CdCst.JavaScript.OnClick, ClsCommon.strConfirm(strMsg))
            End If
        Catch ex As Exception
            intReturn = 9
        End Try
        fncMakeLanguageList = intReturn
    End Function

    ''' <summary>
    '''     ボタンセット
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subButtonSet()
        Dim objDataTbl As DataTable = Nothing
        Dim intLoopCnt As Integer
        Dim intLoopCnt1 As Integer
        Dim intUseFunctionInfo() As Integer         '利用機能情報

        Try
            pnlMaster.Visible = False
            '利用機能情報取得
            intUseFunctionInfo = clsUserInf.fncUseFunctionInfoGet(Me.objUserInfo.UseFunctionLvl)
            If intUseFunctionInfo.Length > 0 Then pnlMaster.Visible = True
            'メニュー(内容)を取得する
            Dim ShowList As New ArrayList

            objDataTbl = bllDefault.fncMenuMstSelect(objConBase, selLang.SelectedValue, Me.objUserInfo.UserClass,
                                                     "MasterMaintenance")

            For intLoopCnt = 0 To objDataTbl.Rows.Count - 1
                If objDataTbl.Rows(intLoopCnt).Item("use_function_lvl").ToString = "0" Then
                    '利用機能レベルが0の場合は全ユーザーが利用可能
                    HideButton(intLoopCnt + 10)
                Else
                    'メニューの利用機能レベルを取得する
                    For intLoopCnt1 = 0 To intUseFunctionInfo.Length - 1
                        If _
                            intUseFunctionInfo(intLoopCnt1).ToString =
                            objDataTbl.Rows(intLoopCnt).Item("use_function_lvl").ToString Then
                            ShowButton(intLoopCnt + 10)
                            Exit For
                        End If
                    Next
                End If
            Next
            'ADD BY YGY 20141029
            'フォカス


            Button3.Focus()
        Catch ex As Exception
            Call ShowErrPage(ex.Message)
        Finally
            objDataTbl = Nothing
        End Try
    End Sub

#Region "自動テスト関連"

    ''' <summary>
    '''     組合せ出力
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Type_GotoKatOut() Handles WebUC_Type.GotoKatOut
        Call HideAllWebUC()
        Me.WebUC_KatOut.Visible = True
        WebUC_KatOut.selLang = Me.selLang
        Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
        WebUC_KatOut.objKtbnStrc = objKtbnStrc
        Call WebUC_KatOut.frmInit()
    End Sub

    ''' <summary>
    '''     形番分解
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Type_GotoKatsepchk() Handles WebUC_Type.GotoKatsepchk
        Call HideAllWebUC()
        Me.WebUC_KatSep.Visible = True
        WebUC_KatSep.selLang = Me.selLang
        Call WebUC_KatSep.frmInit()
    End Sub

    ''' <summary>
    '''     機種選択画面に戻る
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_KatOut_BackToType() _
        Handles WebUC_KatOut.BackToType, WebUC_KatSep.BackToType, WebUC_100Test.BackToType
        pnlMaster.Visible = False
        Call Show_Type()
        Me.selLang.Enabled = False
    End Sub

    ''' <summary>
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_Type_Goto100test() Handles WebUC_Type.Goto100test
        Call HideAllWebUC()
        Me.WebUC_100Test.Visible = True
        WebUC_100Test.selLang = Me.selLang
        Call WebUC_100Test.frmInit()
    End Sub

    ''' <summary>
    '''     機種画面に戻る
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WebUC_100Test_GoToType() Handles WebUC_100Test.GoToType
        Select Case Me.Session("TestMode").ToString
            Case "2"
                Me.Session.Add("ManifoldKatabanLoop", 0)  '開始行、履歴テストの場合、一件しかない
            Case Else
                Me.Session.Add("ManifoldKatabanLoop", 0)
        End Select
        Call Bunkai()
    End Sub

    ''' <summary>
    '''     形番分解
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Bunkai()
        Dim strSeries As String = String.Empty
        Dim strKeyKata As String = String.Empty
        Dim strKataName As String = String.Empty
        Dim strSpecNo As String = String.Empty
        Dim strPriceNo As String = String.Empty
        Dim strItem1(24) As String
        Dim strItemName1(24) As String
        Dim strHyphen1(24) As String
        Dim strElement_div1(24) As String
        Dim strStructure_div(24) As String

        'CHANGED BY YGY 20140708 ↓↓↓↓↓↓
        Dim strKataban As String = String.Empty
        If Me.Session("TestMode").Equals(2) Then
            Dim drShiyouTest As DS_PriceTest.kh_shiyou_testRow = Me.Session("ManifoldKataban")
            strKataban = drShiyouTest.KATABAN
        Else
            Dim listKataban As ManifoldKataban = Me.Session("ManifoldKataban")
            If listKataban Is Nothing Then Exit Sub
            strKataban = listKataban.KATABAN.ToString
        End If

        'CHANGED BY YGY 20140708 ↑↑↑↑↑↑

        Dim strTxtX As String = String.Empty
        If strKataban.StartsWith("GAMD0") Then
            Dim strK() As String = strKataban.Split("-")
            If strK(strK.Length - 1).StartsWith("X") Then
                strTxtX = Right(strK(strK.Length - 1), strK(strK.Length - 1).Length - 1)
                strKataban = Left(strKataban, strKataban.Length - strTxtX.Length - 2)
            End If
        End If

        If strKataban.StartsWith("MN4KB180A-") Or strKataban.StartsWith("MN4KB280-") Then
            Dim strKataKey() As String = strKataban.Split("-")
            If strKataKey(strKataKey.Length - 1) <> "ST" Then
                strKataban = Left(strKataban, strKataban.Length - strKataKey(strKataKey.Length - 1).Length - 1)
            Else
                strKataban = strKataban.Replace("-ST", "")
                strKataban = Left(strKataban, strKataban.Length - strKataKey(strKataKey.Length - 1).Length - 1)
                strKataban &= "-ST"
            End If
        End If

        If strKataban Like ("N*P51*") Then
            Dim strKataKey() As String = strKataban.Split("-")
            If Not strKataKey(strKataKey.Length - 1).Contains("V") Then
                If strKataKey(strKataKey.Length - 1) <> "ST" Then
                    strKataban = Left(strKataban, strKataban.Length - strKataKey(strKataKey.Length - 1).Length - 1)
                Else
                    strKataban = strKataban.Replace("-ST", "")
                    strKataban = Left(strKataban, strKataban.Length - strKataKey(strKataKey.Length - 1).Length - 1)
                    strKataban &= "-ST"
                End If
            End If
        End If

        Dim strPath As String = String.Empty
        Select Case Me.Session("TestMode").ToString
            Case "1"
                strPath = My.Settings.LogFolder & Now.ToString("yyyyMMdd") & "_ManifoldTest_ISO.txt"
            Case "0"
                strPath = My.Settings.LogFolder & Now.ToString("yyyyMMdd") & "_ManifoldTest.txt"
            Case Else
                strPath = My.Settings.LogFolder & "ShiyouTest_" & Now.ToString("yyyyMMdd") & ".txt"
        End Select

        '形番分解
        If KHKatabanSeparator.GetSeparatorData(strKataban.Trim.ToUpper, strSeries, strKeyKata, strKataName,
                                               strSpecNo, strPriceNo, strItem1, strItemName1, strHyphen1,
                                               strStructure_div, strElement_div1) Then
            pnlMaster.Visible = False

            Me.Session.Add("ManifoldSeriesKey", strSeries & "," & strKeyKata & "," & strTxtX)
            Me.Session.Add("ManifoldItemKey", strItem1)

            'Call HideAllButton()

            'If Not Me.objUserInfo.UserId.Equals(My.Settings.AnonymousUserName) Then
            '    ShowButton(1)        'メニュー
            'End If

            ShowButton(2)        'Logoff
            Call HideAllWebUC()
            Me.WebUC_Type.Visible = True
            Me.WebUC_Type.txtKataban.Text = strSeries
            WebUC_Type.frmInit() '画面初期化
            'WebUC_Type.Dispose()
            Me.selLang.Enabled = False

            Session("EventEndFlg") = True
        Else
            If Me.Session("TestMode").Equals(2) Then
                '仕様テストの場合
                Dim drShiyouTest As DS_PriceTest.kh_shiyou_testRow = Me.Session("ManifoldKataban")
                Dim strSeperateResult = drShiyouTest.SEPERATE_RESULT

                If Not strSeperateResult.Equals("0") Then
                    'NET版もエラーの場合はOK
                    ClsCommon.WriteLog(strPath, strKataban & ControlChars.Tab & "○")
                Else
                    'NET版は分解できる場合はNG
                    ClsCommon.WriteLog(strPath, strKataban & ControlChars.Tab & "形番分解エラー")
                End If
            Else
                '価格テストの場合
                ClsCommon.WriteLog(strPath, "形番分解エラー：" & "→" & strKataban)
            End If

            Session("EventEndFlg") = True
            GC.Collect()

        End If
    End Sub

#End Region

    ''' <summary>
    '''     画面のHiddenFiledをクリア
    ''' </summary>
    ''' <param name="strPageName"></param>
    ''' <remarks></remarks>
    Public Sub ClearHiddenField(strPageName As String)
        Select Case strPageName
            Case "Youso"
                'Me.WebUC_Youso.HidDblClick.Value = String.Empty
                'Me.WebUC_Youso.HidGotID.Value = String.Empty
                'Me.WebUC_Youso.HidGVStartID.Value = String.Empty
                'Me.WebUC_Youso.HidLostID.Value = String.Empty
                'Me.WebUC_Youso.HidMaxSelCount.Value = String.Empty


                'Me.WebUC_Youso.HidMultiple.Value = String.Empty
                'Me.WebUC_Youso.HidSelAll.Value = String.Empty
                'Me.WebUC_Youso.HidDblClick.Value = String.Empty


                'Me.WebUC_Youso.HidSelRowID.Value = String.Empty
        End Select
    End Sub

    ''' <summary>
    '''     エラー処理
    ''' </summary>
    ''' <param name="strErrMsg"></param>
    ''' <remarks></remarks>
    Private Sub UC_Goto_ErrPage(strErrMsg As String) Handles WebUC_ISOTanka.Goto_ErrPage,
                                                             WebUC_Menu.Goto_ErrPage, WebUC_Type.Goto_ErrPage,
                                                             WebUC_Youso.Goto_ErrPage,
                                                             WebUC_Tanka.Goto_ErrPage, WebUC_RodEndOrder.Goto_ErrPage,
                                                             WebUC_RodEnd.Goto_ErrPage,
                                                             WebUC_OutOfOption.Goto_ErrPage, WebUC_Stopper.Goto_ErrPage,
                                                             WebUC_PriceCopy.Goto_ErrPage,
                                                             WebUC_Siyou.Goto_ErrPage, WebUC_ISOTanka.Goto_ErrPage,
                                                             WebUC_Login.Goto_ErrPage, WebUC_Motor.Goto_ErrPage
        ShowErrPage(strErrMsg)
    End Sub

    ''' <summary>
    '''     エラーページの表示
    ''' </summary>
    ''' <param name="strMsg"></param>
    ''' <remarks></remarks>
    Private Sub ShowErrPage(strMsg As String)
        Call HideAllWebUC()
        Call HideAllButton()
        WebUC_Error.HidErrMsg.Value = strMsg
        Me.WebUC_Error.Visible = True
        Me.HidRunForm.Value = "1"
        WebUC_Error.selLang = Me.selLang
        Call WebUC_Error.Page_Load(WebUC_Error, Nothing)
    End Sub

    ''' <summary>
    '''     ログインページの表示
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShowLogin()
        Call HideAllWebUC()
        Call HideAllButton()
        Me.WebUC_Login.Visible = True
        Me.selLang.Enabled = True

        Call WebUC_Login.Page_Load(Nothing, Nothing)
    End Sub

    ''' <summary>
    '''     単価画面のリセット
    ''' </summary>
    ''' <param name="strLabelName"></param>
    ''' <remarks></remarks>
    Private Sub SetTankaLabelVisible(strLabelName As List(Of String))
        For Each labelName In strLabelName
            If WebUC_Tanka.FindControl(labelName) IsNot Nothing Then
                WebUC_Tanka.FindControl(labelName).Visible = False
            End If
        Next
    End Sub

#Region "ファイル出力関連"

    ''' <summary>
    '''     ファイル出力のタイトルを出力
    ''' </summary>
    ''' <param name="dt_Title">表示ラベル</param>
    ''' <param name="strMadeCountry">生産国</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncFileOutputTitle(lstPriceInfo As List(Of KeyValuePair(Of String, String)),
                                        dt_Title As DataTable,
                                        strMadeCountry As String) As String
        Dim strResult As String = String.Empty
        Dim strResult2 As String = String.Empty

        '価格リストの項目名
        Dim strPriceListColumns As String = String.Empty

        '画面に選択した情報の項目を出力
        strResult = fncSetSelectInfoTitle(dt_Title, strResult2)


        '価格リストタイトルの出力
        strPriceListColumns = fncSetPriceListTitle(lstPriceInfo, dt_Title, strMadeCountry)

        strResult = String.Format(strResult, strResult2, strPriceListColumns)

        Return strResult
    End Function

    ''' <summary>
    '''     画面に選択した情報の項目を出力
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSetSelectInfoTitle(dt_Title As DataTable, ByRef strResult2 As String) As String
        Dim strResult As String = String.Empty
        Dim flgShowShipPlace = False

        '価格リスト以外の項目の出力
        For Each dr As DataRow In dt_Title.Rows

            If CInt(dr.Item("label_seq")) = FileOutputColumns.ListPrice Then
                '価格リスト項目の位置
                strResult &= "{0}{1}" & CdCst.Sign.Comma

            ElseIf CInt(dr.Item("label_seq")) > FileOutputColumns.ListPrice AndAlso
                   CInt(dr.Item("label_seq")) <= FileOutputColumns.FobPrice Then
                '価格リストの場合はスキップ
            Else
                Select Case dr.Item("label_seq")
                    Case FileOutputColumns.CheckKBN

                        'チェック区分
                        If fncShowCheckKbn(Me.objUserInfo.AddInformationLvl) Then

                            strResult &= ClsCommon.fncAddQuote(dr.Item("label_content")) & CdCst.Sign.Comma
                        End If

                    Case FileOutputColumns.ShipPlace

                        '出荷場所　→プラント
                        If fncShowShipPlace(Me.objUserInfo.AddInformationLvl) Then
                            flgShowShipPlace = True
                            strResult &= ClsCommon.fncAddQuote(dr.Item("label_content")) & CdCst.Sign.Comma
                        End If

                    Case FileOutputColumns.Tax, FileOutputColumns.Total

                        '海外代理店は金額と消費税を非表示にする
                        If Me.objUserInfo.UserClass = CdCst.UserClass.OsAgentCs OrElse
                           Me.objUserInfo.UserClass = CdCst.UserClass.OsAgentLs OrElse
                           Me.objUserInfo.CountryCd <> "JPN" Then
                        Else
                            strResult &= ClsCommon.fncAddQuote(dr.Item("label_content")) & CdCst.Sign.Comma
                        End If

                    Case FileOutputColumns.StorageLocation, FileOutputColumns.EvaluationType
                        '保管場所、評価タイプ
                        If flgShowShipPlace Then
                            strResult2 &= ClsCommon.fncAddQuote(dr.Item("label_content")) & CdCst.Sign.Comma
                        End If

                    Case Else

                        If dt_Title.Rows.IndexOf(dr).Equals(dt_Title.Rows.Count - 1) Then
                            '最後の場合はカンマを出力しない
                            strResult &= ClsCommon.fncAddQuote(dr.Item("label_content"))
                        Else
                            strResult &= ClsCommon.fncAddQuote(dr.Item("label_content")) & CdCst.Sign.Comma
                        End If

                End Select
            End If
        Next

        Return strResult
    End Function

    ''' <summary>
    '''     価格リストタイトルの出力
    ''' </summary>
    ''' <param name="dt_Title">表示ラベル</param>
    ''' <param name="strMadeCountry">生産国</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSetPriceListTitle(lstPriceInfo As List(Of KeyValuePair(Of String, String)),
                                          dt_Title As DataTable,
                                          strMadeCountry As String) As String

        Dim strResult As String = String.Empty


        '価格リストタイトルの作成
        For introw = 0 To lstPriceInfo.Count - 1
            Dim pricePair As KeyValuePair(Of String, String) = lstPriceInfo(introw)
            Dim columnKbn As String = String.Empty
            Dim columnName As String = String.Empty

            columnKbn = pricePair.Key.Split(":")(0)
            columnName = pricePair.Key.Split(":")(1)

            If CInt(columnKbn).Equals(FileOutputColumns.APrice) OrElse
               CInt(columnKbn).Equals(FileOutputColumns.FobPrice) Then
                '購入価格と現地定価の場合は2列を出力
                If lstPriceInfo.IndexOf(pricePair) = lstPriceInfo.Count - 1 Then
                    '最後の場合
                    strResult &= ClsCommon.fncAddQuote(columnName) & CdCst.Sign.Delimiter.Comma &
                                 ClsCommon.fncAddQuote(columnName)
                Else
                    strResult &= ClsCommon.fncAddQuote(columnName) & CdCst.Sign.Delimiter.Comma &
                                 ClsCommon.fncAddQuote(columnName) & CdCst.Sign.Delimiter.Comma
                End If
            Else
                If lstPriceInfo.IndexOf(pricePair) = lstPriceInfo.Count - 1 Then
                    '最後の場合はカンマ出力しない
                    strResult &= ClsCommon.fncAddQuote(columnName)
                Else
                    strResult &= ClsCommon.fncAddQuote(columnName) & CdCst.Sign.Delimiter.Comma
                End If
            End If
        Next

        Return strResult
    End Function

    ''' <summary>
    '''     価格リスト情報の分解
    ''' </summary>
    ''' <param name="strPriceList">画面に選択した価格情報</param>
    ''' <param name="strPriceDispLvl">価格表示権限</param>
    ''' <param name="dt_Title">表示ラベル</param>
    ''' <param name="strMadeCountry">生産国</param>
    ''' <returns></returns>
    ''' <remarks>価格区分：価格名称：価格</remarks>
    Private Function fncGetPriceList(strPriceList As String,
                                     strPriceDispLvl() As Boolean,
                                     strMadeCountry As String,
                                     dt_Title As DataTable
                                     ) As List(Of KeyValuePair(Of String, String))

        Dim result As New List(Of KeyValuePair(Of String, String))
        Dim lstPrice As New List(Of String)

        '生産国が海外の場合は表示権限により仮データを作成
        Select Case strMadeCountry
            Case "P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52",
                "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"
            Case Else
                'PS
                If strPriceDispLvl(2) Then
                    Dim strKey As String = fncCreateFakeKey(FileOutputColumns.PsPrice, dt_Title)
                    Dim strPrice As String = String.Empty
                    result.Add(New KeyValuePair(Of String, String)(strKey, strPrice))
                End If
                'GS
                If strPriceDispLvl(3) Then
                    Dim strKey As String = fncCreateFakeKey(FileOutputColumns.GsPrice, dt_Title)
                    Dim strPrice As String = String.Empty
                    result.Add(New KeyValuePair(Of String, String)(strKey, strPrice))
                End If
                'BS
                If strPriceDispLvl(4) Then
                    Dim strKey As String = fncCreateFakeKey(FileOutputColumns.BsPrice, dt_Title)
                    Dim strPrice As String = String.Empty
                    result.Add(New KeyValuePair(Of String, String)(strKey, strPrice))
                End If
                'SS
                If strPriceDispLvl(5) Then
                    Dim strKey As String = fncCreateFakeKey(FileOutputColumns.SsPrice, dt_Title)
                    Dim strPrice As String = String.Empty
                    result.Add(New KeyValuePair(Of String, String)(strKey, strPrice))
                End If
                '登録店
                If strPriceDispLvl(6) Then
                    Dim strKey As String = fncCreateFakeKey(FileOutputColumns.RegPrice, dt_Title)
                    Dim strPrice As String = String.Empty
                    result.Add(New KeyValuePair(Of String, String)(strKey, strPrice))
                End If
                '定価
                If strPriceDispLvl(7) Then
                    Dim strKey As String = fncCreateFakeKey(FileOutputColumns.ListPrice, dt_Title)
                    Dim strPrice As String = String.Empty
                    result.Add(New KeyValuePair(Of String, String)(strKey, strPrice))
                End If
        End Select

        '画面データの出力
        lstPrice = strPriceList.Split("|").ToList

        For Each pricePair In lstPrice
            If Not pricePair.Equals(String.Empty) Then
                Dim strKey As String = String.Empty
                Dim strPrice As String = String.Empty
                Dim columnKbn As String = String.Empty

                '価格
                strPrice = pricePair.Split(":").Last

                strKey = pricePair.Replace(strPrice, String.Empty)

                '価格区分
                columnKbn = pricePair.Split(":").First

                If columnKbn.Equals(String.Empty) Then
                    Continue For
                End If

                If CInt(columnKbn).Equals(FileOutputColumns.APrice) OrElse
                   CInt(columnKbn).Equals(FileOutputColumns.FobPrice) Then
                Else
                    '現地定価と購入価格以外の場合は通貨を非表示にする
                    strPrice = strPrice.Split(Space(1)).First
                End If

                result.Add(New KeyValuePair(Of String, String)(strKey, strPrice))

            End If
        Next

        Return result
    End Function

    ''' <summary>
    '''     仮キーの作成
    ''' </summary>
    ''' <param name="priceKbn"></param>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreateFakeKey(priceKbn As FileOutputColumns,
                                      dt As DataTable) As String
        Dim strName As String = dt.Select("label_seq = '" & priceKbn & "'")(0).Item("label_content")
        Dim strKey As String = priceKbn & ":" & strName

        Return strKey
    End Function

    ''' <summary>
    '''     形番出力データの作成（普通）
    ''' </summary>
    ''' <param name="objKtbnStrc">全ての情報</param>
    ''' <param name="strArray">掛率,単価,数量,金額,消費税</param>
    ''' <param name="strPriceList">価格リスト</param>
    ''' <param name="strPriceDispLvl">各価格の表示フラグ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreateOutputNormal(objKtbnStrc As KHKtbnStrc,
                                           strArray() As String,
                                           lstPriceInfo As List(Of KeyValuePair(Of String, String))) As String
        Dim strResult As New StringBuilder

        With strResult
            '製品名
            If objKtbnStrc.strcSelection.strDivision = "3" Then
                '仕入品
                Select Case selLang.SelectedValue.Trim
                    Case "ja", String.Empty
                        .Append(CdCst.GoogsName_Shiire.ja & CdCst.Sign.Delimiter.Comma)
                    Case "en"
                        .Append(CdCst.GoogsName_Shiire.en & CdCst.Sign.Delimiter.Comma)
                    Case "ko"
                        .Append(CdCst.GoogsName_Shiire.ko & CdCst.Sign.Delimiter.Comma)
                    Case "tw"
                        .Append(CdCst.GoogsName_Shiire.tw & CdCst.Sign.Delimiter.Comma)
                    Case "zh"
                        .Append(CdCst.GoogsName_Shiire.zh & CdCst.Sign.Delimiter.Comma)
                End Select
            Else
                .Append(ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.strGoodsNm) & CdCst.Sign.Delimiter.Comma)
            End If

            '形番
            .Append(ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.strFullKataban) & CdCst.Sign.Delimiter.Comma)

            'チェック区分・出荷場所
            If fncShowCheckKbn(Me.objUserInfo.AddInformationLvl) Then
                'チェック区分
                .Append(
                    ClsCommon.fncAddQuote("Z" & objKtbnStrc.strcSelection.strKatabanCheckDiv) &
                    CdCst.Sign.Delimiter.Comma)
            End If

            If fncShowShipPlace(Me.objUserInfo.AddInformationLvl) Then
                '出荷場所
                .Append(ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.strPlaceCd) & CdCst.Sign.Delimiter.Comma)
                ''保管場所
                .Append(ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.strStorageLocation) & CdCst.Sign.Delimiter.Comma)
                '評価タイプ
                .Append(ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.strEvaluationType) & CdCst.Sign.Delimiter.Comma)
            End If

            '価格リストの追加
            .Append(fncCreatePriceList(lstPriceInfo))

            '掛率
            .Append(ClsCommon.fncAddQuote(ClsCommon.fncIsInputed(strArray(0), "0.000")) & CdCst.Sign.Delimiter.Comma)

            '掛率,単価,数量,金額,消費税
            For inti = 1 To strArray.Length - 1
                If inti = 4 OrElse inti = 5 Then
                    '海外代理店は金額と消費税を非表示にする
                    If Me.objUserInfo.UserClass = CdCst.UserClass.OsAgentCs OrElse
                       Me.objUserInfo.UserClass = CdCst.UserClass.OsAgentLs OrElse
                       Me.objUserInfo.CountryCd <> "JPN" Then
                    Else
                        .Append(
                            ClsCommon.fncAddQuote(ClsCommon.fncIsInputed(strArray(inti), "0")) &
                            CdCst.Sign.Delimiter.Comma)
                    End If
                Else
                    .Append(
                        ClsCommon.fncAddQuote(ClsCommon.fncIsInputed(strArray(inti), "0")) & CdCst.Sign.Delimiter.Comma)
                End If
            Next

            '更新日
            .Append(ClsCommon.fncAddQuote(DateTime.Now))
        End With

        Return strResult.ToString
    End Function

    ''' <summary>
    '''     価格リストの作成
    ''' </summary>
    ''' <param name="lstPrice"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreatePriceList(lstPrice As List(Of KeyValuePair(Of String, String))) As String
        Dim strResult As New StringBuilder

        With strResult
            For Each pricePair In lstPrice
                Dim columnKbn As String = String.Empty
                Dim columnName As String = String.Empty

                columnKbn = pricePair.Key.Split(":")(0)
                columnName = pricePair.Key.Split(":")(1)

                If CInt(columnKbn).Equals(FileOutputColumns.APrice) OrElse
                   CInt(columnKbn).Equals(FileOutputColumns.FobPrice) Then
                    '現地定価と購入価格の場合は通貨と価格を分ける
                    If pricePair.Value.Contains(Space(1)) Then
                        'スペースで区切する場合
                        Dim strPrice As String = String.Empty
                        Dim strCurrency As String = String.Empty

                        strPrice = pricePair.Value.Split(Space(1)).First
                        strCurrency = pricePair.Value.Split(Space(1)).Last

                        .Append(
                            ClsCommon.fncAddQuote(strCurrency) & CdCst.Sign.Delimiter.Comma &
                            ClsCommon.fncAddQuote(strPrice) & CdCst.Sign.Delimiter.Comma)
                    ElseIf pricePair.Value.Contains("(") Then
                        '括弧で区切する場合
                        Dim strPrice As String = String.Empty
                        Dim strCurrency As String = String.Empty

                        strPrice = pricePair.Value.Split("(").First
                        strCurrency = pricePair.Value.Split("(").Last.Replace(")", String.Empty)

                        .Append(
                            ClsCommon.fncAddQuote(strCurrency) & CdCst.Sign.Delimiter.Comma &
                            ClsCommon.fncAddQuote(strPrice) & CdCst.Sign.Delimiter.Comma)
                    Else
                        .Append(
                            ClsCommon.fncAddQuote(pricePair.Value) & CdCst.Sign.Delimiter.Comma &
                            CdCst.Sign.Delimiter.Comma)
                    End If
                Else
                    Dim strPrice As String = pricePair.Value

                    If pricePair.Value.Contains("(") Then
                        strPrice = strPrice.Split("(")(0)
                    End If

                    .Append(ClsCommon.fncAddQuote(strPrice) & CdCst.Sign.Delimiter.Comma)
                End If
            Next
        End With

        Return strResult.ToString
    End Function

    ''' <summary>
    '''     形番出力データの作成（ISO）
    ''' </summary>
    ''' <param name="objKtbnStrc">全ての情報</param>
    ''' <param name="strOrder">掛率,単価,数量,金額,消費税</param>
    ''' <param name="arrylistPriceList">価格リスト</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreateOutputISO(objKtbnStrc As KHKtbnStrc,
                                        strOrder As String,
                                        arrylistPriceList As ArrayList) As String
        Dim strResult As New StringBuilder
        '引当形番情報取得
        objKtbnStrc = New KHKtbnStrc
        Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, 1)

        'オプション名称データ取得()
        Dim strOpNm As New ArrayList
        strOpNm = fncGetISOOptionNames()

        With strResult

            Dim intLoop = 0
            '掛率,単価,数量,金額,消費税
            Dim strPrice() As String = strOrder.Split("_")

            For j = 1 To objKtbnStrc.strcSelection.strOpKataban.Length - 1
                If objKtbnStrc.strcSelection.strOpKataban(j).ToString.Trim.Length > 0 Then
                    If objKtbnStrc.strcSelection.intQuantity(j) > 0 Then

                        '価格リスト
                        Dim lstPriceInfo As New List(Of KeyValuePair(Of String, String))

                        '表示された形番だけを出力する
                        If strPrice.Length > intLoop Then
                            Dim strOptionNm As String = String.Empty

                            'ISOオプション名称の取得
                            strOptionNm = fncGetISOOptionName(strOpNm, j, objKtbnStrc)

                            '製品名
                            .Append(ClsCommon.fncAddQuote(strOptionNm) & CdCst.Sign.Delimiter.Comma)

                            '形番
                            .Append(
                                ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.strOpKataban(j).ToString) &
                                CdCst.Sign.Delimiter.Comma)

                            'チェック区分・出荷場所
                            If fncShowCheckKbn(Me.objUserInfo.AddInformationLvl) Then
                                'チェック区分
                                .Append(
                                    ClsCommon.fncAddQuote("Z" & objKtbnStrc.strcSelection.strOpKatabanCheckDiv(j)) &
                                    CdCst.Sign.Delimiter.Comma)
                            End If

                            If fncShowShipPlace(Me.objUserInfo.AddInformationLvl) Then
                                '出荷場所
                                .Append(
                                    ClsCommon.fncAddQuote(objKtbnStrc.strcSelection.strOpPlaceCd(j)) &
                                    CdCst.Sign.Delimiter.Comma)
                                '保管場所
                                .Append(ClsCommon.fncAddQuote(String.Empty) & CdCst.Sign.Delimiter.Comma)
                                '評価タイプ
                                .Append(ClsCommon.fncAddQuote(String.Empty) & CdCst.Sign.Delimiter.Comma)
                            End If

                            Dim strArray As String()

                            '価格リストの追加
                            lstPriceInfo = arrylistPriceList(intLoop)
                            .Append(fncCreatePriceList(lstPriceInfo))

                            '掛率,単価,数量,金額,消費税
                            strArray = strPrice(intLoop).ToString.Split("|")
                            intLoop += 1
                            If Not strArray Is Nothing AndAlso strArray.Length > 0 Then
                                .Append(
                                    ClsCommon.fncAddQuote(ClsCommon.fncIsInputed(strArray(0), "0.000")) &
                                    CdCst.Sign.Delimiter.Comma)    '掛率
                                For inti = 1 To strArray.Length - 1

                                    If inti = 4 OrElse inti = 5 Then
                                        '海外代理店は金額と消費税を非表示にする
                                        If Me.objUserInfo.UserClass = CdCst.UserClass.OsAgentCs OrElse
                                           Me.objUserInfo.UserClass = CdCst.UserClass.OsAgentLs OrElse
                                           Me.objUserInfo.CountryCd <> "JPN" Then
                                        Else
                                            .Append(
                                                ClsCommon.fncAddQuote(ClsCommon.fncIsInputed(strArray(inti), "0")) &
                                                CdCst.Sign.Delimiter.Comma)
                                        End If
                                    Else
                                        .Append(
                                            ClsCommon.fncAddQuote(ClsCommon.fncIsInputed(strArray(inti), "0")) &
                                            CdCst.Sign.Delimiter.Comma)
                                    End If
                                Next
                            End If
                            .AppendLine(ClsCommon.fncAddQuote(DateTime.Now))
                        End If

                    End If
                End If
            Next

        End With
        Return strResult.ToString
    End Function

    ''' <summary>
    '''     ISOオプション名称の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetISOOptionNames() As ArrayList

        Dim strOpNm As New ArrayList
        Dim dt_Title As DataTable = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHISOTanka, selLang.SelectedValue)

        For inti = 0 To dt_Title.Rows.Count - 1
            If dt_Title.Rows(inti)("label_div").ToString = "L" Then
                strOpNm.Add(dt_Title.Rows(inti)("label_content").ToString)
            End If
        Next

        Return strOpNm
    End Function

    ''' <summary>
    '''     ISOオプション名の取得
    ''' </summary>
    ''' <param name="strOpNm"></param>
    ''' <param name="inti"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetISOOptionName(strOpNm As ArrayList,
                                         inti As Integer,
                                         objKtbnStrc As KHKtbnStrc) As String
        Dim strOptionNm As String = String.Empty

        '明細情報の作成
        Select Case objKtbnStrc.strcSelection.strSeriesKataban
            Case "CMF", "GMF"
                Select Case inti
                    Case 1
                        strOptionNm = strOpNm(0)  'ベース
                    Case 2, 3, 4, 5, 6, 7
                        strOptionNm = strOpNm(1)  '電磁弁形式
                    Case 13, 14
                        strOptionNm = strOpNm(2)  '給気スペーサ
                    Case 15, 16
                        strOptionNm = strOpNm(3)  '排気スペーサ
                    Case 17, 18
                        strOptionNm = strOpNm(4)  'パイロットチェック弁
                    Case 19, 20, 21, 22
                        strOptionNm = strOpNm(5)  'スペーサ形減圧弁
                    Case 23, 24
                        strOptionNm = strOpNm(6)  '流露遮蔽板
                End Select
            Case "LMF0"
                Select Case inti
                    Case 1
                        strOptionNm = strOpNm(0)
                    Case 2, 3, 4, 5, 6, 7
                        strOptionNm = strOpNm(1)
                    Case 13, 14
                        strOptionNm = strOpNm(2)
                    Case 15, 16
                        strOptionNm = strOpNm(3)
                    Case 17
                        strOptionNm = strOpNm(4)
                    Case 18, 19
                        strOptionNm = strOpNm(6)
                End Select
        End Select

        Return strOptionNm
    End Function

    ''' <summary>
    '''     チェック区分を表示かどうか
    ''' </summary>
    ''' <param name="intAddinfoLvl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncShowCheckKbn(intAddinfoLvl As Integer) As Boolean

        Dim blnResult = False
        Dim strKey = "1024,512,256,128,64,32,16,8,4,2,1"
        Dim strLevel() As String = strKey.Split(",")

        For inti = 0 To strLevel.Length - 1
            If intAddinfoLvl >= CInt(strLevel(inti)) Then

                Select Case CInt(strLevel(inti))
                    Case 128 '中国輸出不可
                        '表示のみ
                    Case 64 'EL品情報
                    Case 32 '販売数量単位
                    Case 16 '標準納期
                    Case 8 '担当者情報
                    Case 4 '在庫情報
                    Case 2 '出荷場所
                    Case 1 '形番チェック区分
                        blnResult = True
                End Select

                intAddinfoLvl -= CInt(strLevel(inti))

            End If
        Next

        Return blnResult
    End Function

    ''' <summary>
    '''     出荷場所を表示かどうか
    ''' </summary>
    ''' <param name="intAddinfoLvl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncShowShipPlace(intAddinfoLvl As Integer) As Boolean

        Dim blnResult = False
        Dim strKey = "1024,512,256,128,64,32,16,8,4,2,1"
        Dim strLevel() As String = strKey.Split(",")

        For inti = 0 To strLevel.Length - 1
            If intAddinfoLvl >= CInt(strLevel(inti)) Then

                Select Case CInt(strLevel(inti))
                    Case 128 '中国輸出不可
                        '表示のみ
                    Case 64 'EL品情報
                    Case 32 '販売数量単位
                    Case 16 '標準納期
                    Case 8 '担当者情報
                    Case 4 '在庫情報
                    Case 2 '出荷場所
                        blnResult = True
                    Case 1 '形番チェック区分
                End Select

                intAddinfoLvl -= CInt(strLevel(inti))

            End If
        Next

        Return blnResult
    End Function

#End Region
End Class