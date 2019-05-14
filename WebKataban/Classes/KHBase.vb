Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class KHBase
    Inherits System.Web.UI.UserControl

#Region "プロパティ"
    Protected strFormId As String                         'フォームＩＤ
    Public objLoginInfo As New KHSessionInfo.LoginInfo 'ログイン情報
    Public objUserInfo As New KHSessionInfo.UserInfo   'ユーザー情報
    Public selLang As DropDownList = Nothing           '言語欄

    Protected strParent As String = "ctl00_ContentDetail_"
    Protected myFontSize As System.Web.UI.WebControls.FontUnit = System.Web.UI.WebControls.FontUnit.Medium
    Protected DefaultColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 192)
    Protected pnlBackColor As System.Drawing.Color = System.Drawing.Color.FromArgb(202, 255, 202)
    Protected txtBackColor As System.Drawing.Color = System.Drawing.Color.FromArgb(173, 205, 207)

    Public objKtbnStrc As New KHKtbnStrc
    Public objCon As New SqlConnection
    Public objConBase As New SqlConnection

    Public Event Goto_ErrPage(strErrMsg As String)
#End Region

    ''' <summary>
    ''' エラーが出る場合
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Page_Error(sender As Object, e As System.EventArgs) Handles Me.Error
        Dim bllType As New TypeBLL
        '引当シリーズ形番削除処理
        bllType.subDeleteSelKtbnInfo(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

        'データベース接続を閉じる
        If Not objCon Is Nothing Then If Not objCon.State = ConnectionState.Closed Then objCon.Close()
        objCon = Nothing
        If Not objConBase Is Nothing Then If Not objConBase.State = ConnectionState.Closed Then objConBase.Close()
        objConBase = Nothing
    End Sub

    ''' <summary>
    ''' エラー表示
    ''' </summary>
    ''' <param name="strMessageID"></param>
    ''' <param name="strMsgValue"></param>
    ''' <remarks></remarks>
    Public Sub AlertMessage(strMessageID As String, Optional ByVal strMsgValue As String = "")
        'メッセージ内容の取得
        Dim strMessage As String = String.Empty
        Try
            If strMessageID = "E001" Then strMessage = "[1]"
            If strMessage.Length <= 0 Then
                strMessage = ClsCommon.fncGetMsg(selLang.SelectedValue, strMessageID)
            End If

            '埋め込み文字を変換する
            Dim strMsgArray() As String = Nothing
            Dim strReplaceValue As String

            strMsgArray = Split(strMsgValue, ",")

            If strMsgArray IsNot Nothing Then
                For i As Integer = 0 To strMsgArray.Length - 1
                    strReplaceValue = "[" + (i + 1).ToString + "]"
                    strMessage = strMessage.Replace(strReplaceValue, strMsgArray(i))
                Next
            End If

            If Session("TestMode") Is Nothing Then
                'エラーメッセージの出力
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "Alert", "alert('" & strMessage & "');", True)

            ElseIf Session("TestMode").Equals(2) Then

                '仕様テストの結果出力
                WriteShiyouTestResult(strMessage)

            Else
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
    ''' 仕様テストの結果を出力
    ''' </summary>
    ''' <param name="strMessage"></param>
    ''' <remarks></remarks>
    Private Sub WriteShiyouTestResult(ByVal strMessage As String)
        'エラーの時にダイアログを自動的に閉じる
        If Session("ManifoldKataban") IsNot Nothing Then
            'NET版処理結果
            Dim drShiyouTest As DS_PriceTest.kh_shiyou_testRow = Me.Session("ManifoldKataban")

            '形番
            Dim strKataban As String = drShiyouTest.KATABAN

            'NET版分解結果
            Dim strSeperateResult As String = drShiyouTest.SEPERATE_RESULT

            'ダイアログを閉じるJavascript
            Dim strAlert As String = "var newWindow = window.showModelessDialog('javascript:alert('" & strMessage & "');window.close();','','status:no;resizable:no;help:no;dialogHeight:30px;dialogWidth:40px;'); setTimeout('newWindow.close();',1000);"

            Dim strPath As String = My.Settings.LogFolder & "Shiyoutest_" & Now.ToString("yyyyMMdd") & ".txt"

            If Not strSeperateResult.Equals("0") Then
                'NET版も失敗の場合はOK
                WriteLog(strPath, strKataban & ControlChars.Tab & "○")
            Else
                'NET版は成功の場合はNG
                WriteLog(strPath, strKataban & ControlChars.Tab & strMessage)
            End If

            'ダイアログを閉じる
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "Alert", strAlert, True)
        End If
    End Sub

    ''' <summary>
    ''' エラー表示
    ''' </summary>
    ''' <param name="myex"></param>
    ''' <remarks></remarks>
    Protected Sub AlertMessage(myex As System.Exception)
        'メッセージ内容の取得
        Try
            Dim strMessage As String = String.Empty
            If Not myex Is Nothing Then
                strMessage &= ControlChars.NewLine
                strMessage &= "関数名：" & myex.TargetSite.Name
                subWriteLog(" -------------------エラー 開始------------------- ")
                subWriteLog(myex.Message)
                subWriteLog(myex.StackTrace)
                subWriteLog(" -------------------エラー 終了------------------- ")
            End If

            If Session("TestMode") Is Nothing Then
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "Alert", "alert('" & strMessage & ")');", True)
            Else
                'ManifoldTest
                'システムエラーの時にダイアログを自動的に閉じる
                IO.File.AppendAllText(My.Settings.LogFolder & "SystemError_" & Now.ToString("yyyyMMdd") & ".txt", strMessage & ControlChars.NewLine)
                Dim strAlert As String = "var newWindow = window.showModelessDialog('javascript:alert('" & strMessage & "');window.close();','','status:no;resizable:no;help:no;dialogHeight:30px;dialogWidth:40px;'); setTimeout('newWindow.close();',1000);"
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "Alert", strAlert, True)
            End If

            RaiseEvent Goto_ErrPage(myex.Message)
        Catch ex As Exception
            RaiseEvent Goto_ErrPage(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ログの出力
    ''' </summary>
    ''' <param name="strLog"></param>
    ''' <remarks></remarks>
    Protected Sub AppendDebugLog(strLog As String)
        'System.IO.File.AppendAllText("D:\Log\Debug.txt", Now & " " & Now.Millisecond.ToString.PadLeft(3, "0") & " " & _
        '                             strLog & ControlChars.NewLine)
    End Sub

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Me.Visible = False Then Exit Sub
        Try
            If Me.ID.StartsWith("ISODetail") Then Exit Sub
            selLang = Me.Parent.Parent.Parent.Parent.FindControl("ContentTitle").FindControl("selLang")
            Select Case sender.ID
                Case "WebUC_Youso"
            End Select
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 画面IDのチェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function FormIDCheck() As Boolean
        FormIDCheck = True
        If Not Me.Session("FormID") Is Nothing AndAlso Me.Session("FormID").ToString.Length > 0 Then
            If Me.Session("FormID").ToString <> Me.ID Then
                FormIDCheck = False
            Else
                Me.Session.Remove("FormID")
            End If
        End If
    End Function

    ''' <summary>
    ''' フォントの設定
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Protected Sub SetAllFontName(obj As Object)
        Try
            'If obj.Controls.Count > 0 Then
            '    Select Case obj.GetType.Name.ToUpper
            '        Case "DROPDOWNLIST", "GRIDVIEW", "LISTBOX"
            '            obj.Font.Name = GetFontName(selLang.SelectedValue)
            '        Case Else
            '            For inti As Integer = 0 To obj.Controls.Count - 1
            '                SetAllFontName(obj.Controls(inti))
            '            Next
            '    End Select
            'Else
            '    Select Case obj.GetType.Name.ToUpper
            '        Case "LABEL", "CTLCHARTEXT", "CTLNUMTEXT", "TEXTBOX", "RADIOBUTTON", "DROPDOWNLIST", "GRIDVIEW", "LISTBOX"
            '            If obj.ID <> "txtKataban" Then
            '                obj.Font.Name = GetFontName(selLang.SelectedValue)
            '            End If
            '    End Select
            'End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' テキスト位置
    ''' </summary>
    ''' <remarks></remarks>
    Protected Enum TextAlign
        Left = 0
        Right = 1
        Center = 2
    End Enum

    ''' <summary>
    ''' テキスト位置の設定
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="TextAlign"></param>
    ''' <remarks></remarks>
    Protected Sub SetFontAlign(obj As Object, TextAlign As TextAlign)
        Select Case TextAlign
            Case KHBase.TextAlign.Left
                obj.Style.Add("text-align", "left")
            Case KHBase.TextAlign.Right
                obj.Style.Add("text-align", "right")
            Case KHBase.TextAlign.Center
                obj.Style.Add("text-align", "center")
        End Select
    End Sub
End Class