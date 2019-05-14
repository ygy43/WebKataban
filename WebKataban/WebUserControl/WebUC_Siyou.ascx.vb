Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports System.Drawing
Imports WebKataban.KHCodeConstants

Public Class WebUC_Siyou
    Inherits KHBase

#Region "プロパティ"
    Public Event GotoTanka()
    Public Event GotoISOTanka()

    Private dt_Comb As New ArrayList
    Private DS_Title As New DataSet
    Private Const strMaru As String = "●"
    Private bllSiyou As New SiyouBLL
#End Region

    ''' <summary>
    ''' 画面初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()

        Me.Session("TestFlag") = Nothing
        Me.Session("DS_Title") = Nothing
        Me.Session("dt_Comb") = Nothing
        If Not Me.Session("ManifoldKatabanLoop") Is Nothing Then
            Me.Session("KtbnStrc_Siyou") = Nothing
        End If
        HidColMerge.Value = String.Empty
        HidManifoldMode.Value = 0
        HidColCount.Value = 0
        HidClick.Value = String.Empty
        HidStdNum.Value = String.Empty
        HidOther.Value = String.Empty
        HidSelect.Value = String.Empty
        HidUse.Value = String.Empty

        HidSetCX.Value = String.Empty
        HidStartID.Value = String.Empty
        HidSimpleOther.Value = String.Empty
        HidCXA.Value = String.Empty
        HidCXB.Value = String.Empty
        HidTube.Value = String.Empty
        HidRailChangeFlg.Value = String.Empty
        DS_Title = New DataSet
        dt_Comb = New ArrayList
        Me.btnOK.UseSubmitBehavior = False
        Me.OnLoad(Nothing)
    End Sub

    ''' <summary>
    ''' 単価画面から戻るとき
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmBack()
        HidPostBack.Value = "1"
        Me.OnLoad(Nothing)
    End Sub

    ''' <summary>
    ''' 情報をロードする
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub
        If Me.objUserInfo.UserId Is Nothing Then Exit Sub
        If Not FormIDCheck() Then Exit Sub

        Try
            '画面の作成
            Call CreatePage()

            'CX情報の設定
            Call SetCXChoicesInfo()

            'タグ銘板の設定
            Call SetTagMeiban()

            'マニホールドテスト時実行
            Call ManifoldTest_Siyou()

        Catch ex As Exception
            AlertMessage(ex)
            HidClick.Value = String.Empty
        End Try
    End Sub

    ''' <summary>
    ''' フォントとタイトルの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetFontAndTitle()
        Me.GridViewDetail.Font.Name = GetFontName(selLang.SelectedValue)
        Me.GridViewTitle.Font.Name = GetFontName(selLang.SelectedValue)
        Me.lblSeriesNm.Font.Name = GetFontName(selLang.SelectedValue)
        Me.lblSeriesKat.Font.Name = GetFontName(selLang.SelectedValue)
        Me.btnOK.Font.Name = GetFontName(selLang.SelectedValue)
        lblSeriesKat.Text = objKtbnStrc.strcSelection.strFullKataban
        lblSeriesNm.Text = objKtbnStrc.strcSelection.strGoodsNm
    End Sub

    ''' <summary>
    ''' CX選択情報の設定
    ''' 修正が必要
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetCXChoicesInfo()

        '操作区分
        If HidManifoldMode.Value = "3" Then
            HidSetCX.Value = objKtbnStrc.strcSelection.strOpSymbol(6).ToString
        ElseIf HidManifoldMode.Value = "4" Then
            HidSetCX.Value = objKtbnStrc.strcSelection.strOpSymbol(2).ToString
        Else
            HidSetCX.Value = String.Empty
        End If

        'OKボタンを押す時にCXAとCXBの設定
        Dim strCXA() As String = HidCXA.Value.Split(",")
        Dim strCXB() As String = HidCXB.Value.Split(",")
        Dim strSelect() As String = HidSelect.Value.Split(",")

        For inti As Integer = 0 To strCXA.Count - 1

            If Not strCXA(inti).Equals(String.Empty) Then
                Dim strkata As String = String.Empty
                Dim intRow As Integer
                Dim CXList As New ArrayList
                Dim drp As DropDownList

                strkata = strSelect(inti)
                intRow = inti
                CXList = bllSiyou.GetCXList(HidManifoldMode.Value, objKtbnStrc, strkata, intRow)

                For intj As Integer = 0 To 1
                    drp = Me.GridViewDetail.Rows(intRow).Cells(intj).Controls(0)
                    drp.ViewStateMode = UI.ViewStateMode.Enabled
                    drp.DataSource = CXList
                    drp.DataBind()
                Next
            End If
        Next
    End Sub

    ''' <summary>
    ''' タグ銘板の使用数を設置
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetTagMeiban()
        Dim str_itemdiv() As String = Me.HidOther.Value.ToString.Split(",")

        For inti As Integer = 0 To str_itemdiv.Length - 1
            If str_itemdiv(inti) = "99" Then
                Dim strKataSel() As String = Me.HidSelect.Value.ToString.Split(",")
                Dim dt_data As DataTable = Nothing
                If Not DS_Title.Tables("data") Is Nothing Then dt_data = DS_Title.Tables("data")
                If dt_data Is Nothing Then Exit For

                If Not strKataSel Is Nothing AndAlso strKataSel.Length > inti AndAlso strKataSel(inti).ToString.Length > 0 Then
                    dt_data.Rows(inti)("Col0") = "1"
                Else
                    dt_data.Rows(inti)("Col0") = String.Empty
                End If
                Dim cel As System.Web.UI.WebControls.TableCell = Nothing
                If Me.GridViewDetail.Rows.Count <= inti Then Exit For
                cel = Me.GridViewDetail.Rows(inti).Cells(0)
                If cel Is Nothing OrElse cel.Controls.Count <= 0 Then Exit For

                Dim txt As New TextBox
                Select Case HidManifoldMode.Value
                    Case "3", "4"
                        txt = cel.Controls(2)
                    Case Else
                        txt = cel.Controls(0)
                End Select

                txt.Text = dt_data.Rows(inti)("Col0")
                Me.Session("DS_Title") = DS_Title
                Exit For
            End If
        Next
    End Sub

    ''' <summary>
    ''' 画面の初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreatePage()
        Try
            '位置構成データを取得する
            Dim dt_position As New DataTable

            'Session情報の取得
            If Me.Session("KtbnStrc_Siyou") Is Nothing Then
                Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
                Me.Session.Add("KtbnStrc_Siyou", objKtbnStrc)
            Else
                objKtbnStrc = Me.Session("KtbnStrc_Siyou")
            End If

            'フォントとタイトルのセット
            Call SetFontAndTitle()

            '画面の作成
            If Not HidPostBack.Value.Equals("1") Then

                '初期化すること
                HidStdNum.Value = String.Empty
                HidOther.Value = String.Empty

                '@@@@仕様に関係ない    ↓↓↓↓↓↓
                '画面ラベルの設定()
                Dim subSetLbl As New DataTable
                subSetLbl = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHSiyou, selLang.SelectedValue)
                subSetLbl.TableName = "LabelName"
                DS_Title.Tables.Add(subSetLbl)

                '@@@@仕様関連    ↓↓↓↓↓↓
                '画面構造を取得(ラベル表示名と表示名に対応する行数)
                dt_position = bllSiyou.LoadPositionData(objCon, selLang.SelectedValue, objKtbnStrc.strcSelection.strSpecNo)

                '形番選択肢の取得
                If dt_position.Rows.Count < 0 Then Exit Sub
                dt_Comb = New ArrayList
                If Not bllSiyou.GetSelKata(objCon, objKtbnStrc, dt_Comb) Then Exit Sub

                '特殊選択肢の追加
                Call SetComb_Change()

                'マニホールド画面とデータテーブルの作成
                CreateMode(dt_position)

                'DataGridをバインドする
                Call BindGridView()

                '形番の初期値の設定
                If HidManifoldMode.Value <> "0" AndAlso Not DS_Title.Tables("data") Is Nothing Then
                    For inti As Integer = 1 To DS_Title.Tables("title").Rows.Count
                        If CType(dt_Comb(inti), ArrayList).Count = 1 Then
                            DS_Title.Tables("title").Rows(inti - 1)("ColKata") = CType(dt_Comb(inti), ArrayList)(0)
                        End If
                    Next
                    'レール長さの設定
                    Call SetRail()
                End If

                '選択した仕様情報をDataTableに登録
                SaveInfoToDatatable(DS_Title.Tables("title"), DS_Title.Tables("data"))

                'DataGridをバインドする
                Call BindGridView()

                '画面選択不可範囲の設定
                Call SetNoSelect()

                Me.Session.Add("dt_Comb", dt_Comb)
                Me.Session.Add("DS_Title", DS_Title)
            Else
                HidPostBack.Value = "0"
                'DataGridをバインドする
                dt_Comb = Me.Session.Item("dt_Comb")
                DS_Title = Me.Session.Item("DS_Title")
                Call BindGridView()
                '画面選択不可範囲の設定
                Call SetNoSelect()
                'DropDownListの設定
                Call SetDropDownList()
                Call SetTextBox()
            End If

            If Me.HidStartID.Value = String.Empty Then
                Me.HidStartID.Value = Me.GridViewTitle.Rows(0).ClientID & "," & Me.GridViewDetail.Rows(0).ClientID
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' レール長さの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetRail()
        'レール長さの設定
        Dim intRail As Integer = -1
        intRail = GetRailTubeIndex(0) '「取付ﾚｰﾙ長さ」の行番号
        '
        If dt_Comb.Count > intRail + 1 AndAlso CType(dt_Comb(intRail + 1), ArrayList).Count > 0 Then
            Call SetRailData(intRail)
        End If
    End Sub

    ''' <summary>
    ''' ﾚｰﾙ長さの設定
    ''' </summary>
    ''' <param name="intRail"></param>
    ''' <remarks></remarks>
    Private Sub SetRailData(intRail As Integer)
        If intRail >= 0 Then
            Dim dblSelValue As Double = 0D

            'レール長さの計算
            dblSelValue = GetRailLength(DS_Title, intRail)

            HidStdNum.Value = intRail & "," & HidStdNum.Value

            DS_Title.Tables("data").Rows(intRail)("Col0") = dblSelValue
            DS_Title.Tables("title").Rows(intRail)("ColKata") = dblSelValue

            Dim celcmb As System.Web.UI.WebControls.TableCell
            Dim celtxt As System.Web.UI.WebControls.TableCell
            celcmb = Me.GridViewTitle.Rows(intRail).Cells(Me.GridViewTitle.Rows(intRail).Cells.Count - 1)

            If HidManifoldMode.Value.Equals("3") OrElse HidManifoldMode.Value.Equals("4") Then
                celtxt = Me.GridViewDetail.Rows(intRail).Cells(2)
            Else
                celtxt = Me.GridViewDetail.Rows(intRail).Cells(0)
            End If

            If celcmb.Controls.Count > 0 AndAlso celtxt.Controls.Count > 0 Then
                Dim drp As DropDownList = celcmb.Controls(0)
                Dim txt As TextBox = celtxt.Controls(0)
                Dim blnHasScript As Boolean = False
                'レール長さ初期値の設定
                drp.SelectedValue = dblSelValue
                txt.Text = dblSelValue

                'ADD BY YGY 20140807    ↓↓↓↓↓↓
                'レール長さ更新ボタン押す時に新しい長さを画面に反映
                If Me.Session("RailUpdate") IsNot Nothing Then
                    Me.Session.Remove("RailUpdate")
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "SetData", "SetSiyouData('" & drp.ClientID & "','" & txt.ClientID & "','" & dblSelValue & "');", True)
                End If
                'ADD BY YGY 20140807    ↑↑↑↑↑↑

                If objKtbnStrc.strcSelection.decDinRailLength > 0 Then
                    If Me.GridViewDetail.Rows(intRail).Cells.Count > 1 Then
                        DS_Title.Tables("data").Rows(intRail)("Col1") = " L1=" & objKtbnStrc.strcSelection.decDinRailLength
                        celtxt = Me.GridViewDetail.Rows(intRail).Cells(1)
                        celtxt.Text = " L1=" & objKtbnStrc.strcSelection.decDinRailLength
                    End If
                End If
            End If
        End If
    End Sub

    'RM1803032_一部SpecNo連数変更
    ''' <summary>
    ''' dt_position    '設置位置情報
    ''' </summary>
    ''' <param name="dt_position"></param>
    ''' <remarks></remarks>
    Private Sub CreateMode(ByRef dt_position As DataTable)
        Try
            Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban.ToString
            Dim strValue() As String = objKtbnStrc.strcSelection.strOpSymbol
            Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban

            '画面にバインドするdt_detailとdt_titleの作成    ↓↓↓↓↓↓
            Select Case objKtbnStrc.strcSelection.strSpecNo.ToString.Trim
                Case "64", "66", "68", "70", "72", "S", "T", "U"    'RM1805001_4Rシリーズ追加
                    Select Case objKtbnStrc.strcSelection.strSpecNo.ToString
                        Case "70", "72"  'M4F6E
                            If strValue(4).ToString.Length > 0 Then
                                For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                    If dt_position.Rows(inti)("df_label_content") <> "Masking plate" AndAlso _
                                        dt_position.Rows(inti)("item_div") = "2" Then
                                        Dim str As String = dt_position.Rows(inti)("label_content").ToString
                                        dt_position.Rows(inti)("label_content") = str.Split("-")(0) & "-" & strValue(4)
                                    End If
                                Next
                            End If
                    End Select
                    If strValue(1).ToString.StartsWith("8") Then
                        HidColCount.Value = 25
                        Call CreatSiyou_DT(dt_position)
                    End If
                Case "52", "60", "61", "62", "63", "65", "67", "69", "71"
                    If strKeyKata = "M" Then
                        'マスターバルブマニホールド
                        If (strSeriesKata = "M4F2" Or strSeriesKata = "M4F3") And (strValue(8) = "C" Or strValue(8) = "I") Then
                            For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                If dt_position.Rows(inti)("label_content") = "MP" Then
                                    dt_position.Rows.RemoveAt(inti)
                                    dt_position.AcceptChanges()
                                    Exit For
                                End If
                            Next
                        End If
                        Select Case strSeriesKata
                            Case "M4F0", "M4F1"
                                For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                    If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                        dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "1-" & strValue(3).ToString
                                    ElseIf dt_position.Rows(inti)("label_content").ToString.StartsWith("A4F") Then
                                        dt_position.Rows(inti)("label_content") = dt_position.Rows(inti)("label_content").ToString & "-" & strValue(3).ToString
                                    End If
                                Next
                            Case "M4F2", "M4F3"
                                Select Case strValue(8)
                                    Case "C"
                                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                            If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                                dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "8-" & strValue(3).ToString
                                            ElseIf dt_position.Rows(inti)("label_content").ToString.StartsWith("A4F") Then
                                                dt_position.Rows(inti)("label_content") = dt_position.Rows(inti)("label_content").ToString & "-" & strValue(3).ToString
                                            End If
                                        Next
                                    Case "I"
                                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                            If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                                dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "8-" & strValue(3).ToString
                                            ElseIf dt_position.Rows(inti)("label_content").ToString.StartsWith("A4F") Then
                                                dt_position.Rows(inti)("label_content") = dt_position.Rows(inti)("label_content").ToString & "-" & strValue(3).ToString
                                            End If
                                        Next
                                    Case Else
                                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                            If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                                dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "1-" & strValue(3).ToString
                                            ElseIf dt_position.Rows(inti)("label_content").ToString.StartsWith("A4F") Then
                                                dt_position.Rows(inti)("label_content") = dt_position.Rows(inti)("label_content").ToString & "-" & strValue(3).ToString
                                            End If
                                        Next
                                End Select

                            Case "M4F4", "M4F5"
                                For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                    If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                        dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "8-00"
                                    End If
                                Next
                            Case "M4F6"
                                For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                    If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                        dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "8-D00"
                                    End If
                                Next
                            Case "M4F7"
                                For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                    If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                        dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "8-E00"
                                    End If
                                Next
                        End Select
                    Else
                        'マニホールド
                        If (strSeriesKata = "M4F2" Or strSeriesKata = "M4F3") And (strValue(8) = "C" Or strValue(8) = "I") Then
                            For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                If dt_position.Rows(inti)("label_content") = "MP" Then
                                    dt_position.Rows.RemoveAt(inti)
                                    dt_position.AcceptChanges()
                                    Exit For
                                End If
                            Next
                        End If
                        Select Case strSeriesKata
                            Case "M4F0", "M4F1"
                                For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                    If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                        dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "0-" & strValue(3).ToString
                                    ElseIf dt_position.Rows(inti)("label_content").ToString.StartsWith("A4F") Then
                                        dt_position.Rows(inti)("label_content") = dt_position.Rows(inti)("label_content").ToString & "-" & strValue(3).ToString
                                    End If
                                Next
                            Case "M4F2", "M4F3"
                                Select Case strValue(8)
                                    Case "C"
                                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                            If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                                dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "9-" & strValue(3).ToString
                                            End If
                                        Next
                                    Case "I"
                                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                            If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                                dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "9-" & strValue(3).ToString
                                            End If
                                        Next
                                    Case Else
                                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                            If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                                dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "0-" & strValue(3).ToString
                                            End If
                                        Next
                                End Select
                            Case "M4F4", "M4F5"

                                For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                    If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                        dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "9-00"
                                    End If
                                Next
                        End Select
                    End If
                    dt_position.AcceptChanges()
                    If strValue(1).ToString.StartsWith("8") Then
                        HidColCount.Value = 25
                        Call CreatSiyou_DT(dt_position)
                    End If
                Case "A4", "A5", "A6", "A7", "A8"
                    If (strSeriesKata = "M4F3") And (strValue(8) = "C" Or strValue(8) = "I") Then
                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                            If dt_position.Rows(inti)("label_content") = "MP" Then
                                dt_position.Rows.RemoveAt(inti)
                                dt_position.AcceptChanges()
                                Exit For
                            End If
                        Next
                    End If
                    Select Case strSeriesKata
                        Case "M4F3"
                            'RM1312006
                            For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                    dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "0EX"
                                End If
                            Next

                        Case "M4F4", "M4F5", "M4F6", "M4F7"
                            'RM1312006
                            For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                If dt_position.Rows(inti)("label_content").ToString.StartsWith("4F") Then
                                    dt_position.Rows(inti)("label_content") = Strings.Left(dt_position.Rows(inti)("label_content"), 4) & "9EX"
                                End If
                            Next
                    End Select
                    dt_position.AcceptChanges()
                    If strValue(1).ToString.StartsWith("8") Then
                        HidColCount.Value = 25
                        Call CreatSiyou_DT(dt_position)
                    End If
                Case "51"
                    If strValue(4) = "6" Or (strValue(5) = "M0" Or strValue(5) = "M1" Or strValue(5) = "M4") Then
                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                            If dt_position.Rows(inti)("label_content") = "M4" Then
                                dt_position.Rows.RemoveAt(6)
                                dt_position.Rows.RemoveAt(inti)
                                dt_position.AcceptChanges()
                                Exit For
                            End If
                        Next
                    End If
                    If strValue(3).ToString.StartsWith("8") Then
                        HidColCount.Value = 25
                        Call CreatSiyou_DT(dt_position)
                    End If
                Case "A1", "A2"
                    If strValue(1) <> "8" Then
                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                            'CHANGED BY YGY 20141027    M3QRA110-M5-D2-2-3
                            'If dt_position.Rows(inti)("df_label_content") = "Masking plate" Then
                            If dt_position.Rows(inti)("label_content") = "MP" Then
                                dt_position.Rows.RemoveAt(inti)
                                dt_position.AcceptChanges()
                                Exit For
                            End If
                        Next
                    End If

                    If strValue(1) = "2" Then
                        Select Case strSeriesKata
                            Case "M3QRA1", "M3QRB1"
                                For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                    If dt_position.Rows(inti)("label_content") = "3QRA119" Then
                                        dt_position.Rows(inti)("label_content") = "3QRA129"
                                    ElseIf dt_position.Rows(inti)("label_content") = "3QRB119" Then
                                        dt_position.Rows(inti)("label_content") = "3QRB129"
                                    ElseIf dt_position.Rows(inti)("label_content") = "S1" Then
                                        dt_position.Rows(inti)("label_content") = "S2"
                                    End If
                                Next
                        End Select
                    End If

                    HidColCount.Value = 25
                    Call CreatSiyou_DT(dt_position)
                Case "A9", "B1"
                    If strValue(1) <> "8" Then
                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                            If dt_position.Rows(inti)("label_content") = "MP" Then
                                dt_position.Rows.RemoveAt(inti)
                                dt_position.AcceptChanges()
                                Exit For
                            End If
                        Next
                    End If
                    If strValue(1) <> "8" Then
                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                            If dt_position.Rows(inti)("label_content") = "S2" Then
                                dt_position.Rows.RemoveAt(inti)
                                dt_position.AcceptChanges()
                                Exit For
                            End If
                        Next
                    End If

                    If strValue(1) = "2" Then
                        Select Case strSeriesKata
                            Case "MV3QRA1", "MV3QRB1"
                                For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                    If dt_position.Rows(inti)("label_content") = "3QRA119" Then
                                        If strValue(8) = "V1" Then
                                            'dt_position.Rows(inti)("label_content") = "3QRB129+ｾﾝｻ"
                                            'CHANGED BY YGY 20141027
                                            '多言語の対応
                                            If selLang.SelectedValue.Equals("ja") Then
                                                dt_position.Rows(inti)("label_content") = "3QRA129+ｾﾝｻ"
                                            Else
                                                dt_position.Rows(inti)("label_content") = "3QRA129+Senser"
                                            End If
                                        Else
                                            dt_position.Rows(inti)("label_content") = "3QRA129"
                                        End If
                                    ElseIf dt_position.Rows(inti)("label_content") = "3QRB119" Then
                                        If strValue(8) = "V1" Then
                                            'dt_position.Rows(inti)("label_content") = "3QRB129+ｾﾝｻ"
                                            'CHANGED BY YGY 20141027
                                            '多言語の対応
                                            If selLang.SelectedValue.Equals("ja") Then
                                                dt_position.Rows(inti)("label_content") = "3QRB129+ｾﾝｻ"
                                            Else
                                                dt_position.Rows(inti)("label_content") = "3QRB129+Senser"
                                            End If
                                        Else
                                            dt_position.Rows(inti)("label_content") = "3QRB129"
                                        End If
                                    ElseIf dt_position.Rows(inti)("label_content") = "S1" Then
                                        dt_position.Rows(inti)("label_content") = "S2"
                                    End If
                                Next
                        End Select
                    Else
                        If strValue(1) = "1" Then
                            Select Case strSeriesKata
                                Case "MV3QRA1", "MV3QRB1"
                                    For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                        If dt_position.Rows(inti)("label_content") = "3QRA119" Then
                                            If strValue(8) = "V1" Then
                                                'dt_position.Rows(inti)("label_content") = "3QRA119+ｾﾝｻ"
                                                'CHANGED BY YGY 20141027
                                                '多言語の対応
                                                If selLang.SelectedValue.Equals("ja") Then
                                                    dt_position.Rows(inti)("label_content") = "3QRA119+ｾﾝｻ"
                                                Else
                                                    dt_position.Rows(inti)("label_content") = "3QRA119+Senser"
                                                End If
                                                Exit For
                                            Else
                                                dt_position.Rows(inti)("label_content") = "3QRA119"
                                                Exit For
                                            End If
                                        ElseIf dt_position.Rows(inti)("label_content") = "3QRB119" Then
                                            If strValue(8) = "V1" Then
                                                'dt_position.Rows(inti)("label_content") = "3QRB119+ｾﾝｻ"
                                                'CHANGED BY YGY 20141027
                                                '多言語の対応
                                                If selLang.SelectedValue.Equals("ja") Then
                                                    dt_position.Rows(inti)("label_content") = "3QRB119+ｾﾝｻ"
                                                Else
                                                    dt_position.Rows(inti)("label_content") = "3QRB119+Senser"
                                                End If
                                                Exit For
                                            Else
                                                dt_position.Rows(inti)("label_content") = "3QRB119"
                                                Exit For
                                            End If

                                        End If
                                    Next
                            End Select
                        End If
                    End If
                    If strValue(1) = "8" Then
                        If strValue(8) = "V1" Then
                            For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                                If dt_position.Rows(inti)("label_content") = "3QRA119" Then
                                    'dt_position.Rows(inti)("label_content") = "3QRA119+ｾﾝｻ"
                                    'CHANGED BY YGY 20141027
                                    '多言語の対応
                                    If selLang.SelectedValue.Equals("ja") Then
                                        dt_position.Rows(inti)("label_content") = "3QRA119+ｾﾝｻ"
                                    Else
                                        dt_position.Rows(inti)("label_content") = "3QRA119+Senser"
                                    End If
                                ElseIf dt_position.Rows(inti)("label_content") = "3QRA129" Then
                                    'dt_position.Rows(inti)("label_content") = "3QRA129+ｾﾝｻ"
                                    'CHANGED BY YGY 20141027
                                    '多言語の対応
                                    If selLang.SelectedValue.Equals("ja") Then
                                        dt_position.Rows(inti)("label_content") = "3QRA129+ｾﾝｻ"
                                    Else
                                        dt_position.Rows(inti)("label_content") = "3QRA129+Senser"
                                    End If
                                End If
                            Next
                        End If

                    End If
                    HidColCount.Value = 25
                    Call CreatSiyou_DT(dt_position)
                Case "B2", "B3", "B4"
                    If strValue(1) <> "8" Then
                        For inti As Integer = dt_position.Rows.Count - 1 To 0 Step -1
                            If dt_position.Rows(inti)("label_content") = "MP" Then
                                dt_position.Rows.RemoveAt(inti)
                                dt_position.AcceptChanges()
                                Exit For
                            End If
                        Next
                    End If
                    HidColCount.Value = 25
                    Call CreatSiyou_DT(dt_position)
                Case "54", "55", "56", "57", "58", "59", "91", "92"
                    If strValue(1) = "8" Then
                        HidColCount.Value = 25
                        Call CreatSiyou_DT(dt_position)
                    End If
                Case "53", "73", "74", "75", "76", "77", "78", "79", _
                    "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "90", "93", "98"
                    If strValue(1).ToString.StartsWith("8") Then
                        HidColCount.Value = 25
                        Call CreatSiyou_DT(dt_position)
                    End If
                Case "96"
                    HidManifoldMode.Value = 18   '複雑モード
                    HidColCount.Value = 40
                    Call CreatSiyou_M(dt_position)
                Case "01", "02", "07", "08", "09", "10", "13", "14", "15", "16"
                    If objKtbnStrc.strcSelection.strSpecNo.ToString = "09" Then
                        If strValue(6).ToString.Length <= 0 Then Exit Sub
                    End If
                    HidManifoldMode.Value = CInt(objKtbnStrc.strcSelection.strSpecNo.ToString)   '複雑モード
                    Select Case objKtbnStrc.strcSelection.strSpecNo.ToString.Trim
                        Case "01", "02", "07", "15", "16"
                            HidColCount.Value = 40
                        Case Else
                            HidColCount.Value = 25
                    End Select
                    Call CreatSiyou_M(dt_position)
                Case "11"
                    HidManifoldMode.Value = CInt(objKtbnStrc.strcSelection.strSpecNo.ToString)   '複雑モード
                    HidColCount.Value = 20
                    Call CreatSiyou_M(dt_position)
                Case "03", "04"
                    HidManifoldMode.Value = CInt(objKtbnStrc.strcSelection.strSpecNo.ToString)   '複雑モード
                    Select Case objKtbnStrc.strcSelection.strSpecNo.ToString.Trim
                        Case "04"
                            HidColCount.Value = 40
                        Case Else
                            HidColCount.Value = 20
                    End Select
                    Call CreatSiyou_M(dt_position)
                Case "12", "18", "19", "20", "21", "22", "23", "94", "95"
                    HidColCount.Value = 12
                    HidManifoldMode.Value = 12               '複雑モード*2
                    Call CreatSiyou_M(dt_position)
                Case "17"
                    HidColCount.Value = 5
                    HidManifoldMode.Value = 17               '複雑モード*2
                    Call CreatSiyou_M(dt_position)
                Case "05", "06"
                    HidManifoldMode.Value = CInt(objKtbnStrc.strcSelection.strSpecNo.ToString)   'ISO
                    HidColCount.Value = 10
                    Call CreatSiyou_M(dt_position)
                Case "A3"
                    HidManifoldMode.Value = 1                '複雑モード MN3Q0,MT3Q0
                    HidColCount.Value = 40
                    Call CreatSiyou_M(dt_position)
            End Select
            '画面にバインドするdt_detailとdt_titleの作成    ↑↑↑↑↑↑

            '選択肢の設定
            'If Me.Session("dt_Comb") Is Nothing Then
            Dim strV As ArrayList = bllSiyou.subGetRail_Cmb(HidManifoldMode.Value, objKtbnStrc)
            If strV.Count > 0 Then
                Dim intRail As Integer = GetRailTubeIndex(0)
                If intRail > 0 Then dt_Comb(intRail + 1) = strV
            End If
            'End If

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' GridViewの設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub BindGridView()
        Try
            GridViewTitle.Columns.Clear()
            GridViewDetail.Columns.Clear()
            If DS_Title.Tables.Count = 3 Then
                '@@@@GridViewの設定

                'GridViewTitle
                CreatDataGridColumn(DS_Title.Tables("LabelName"))      'DataGridの列を作成する
                GridViewTitle.DataSource = DS_Title.Tables("title")
                GridViewTitle.DataBind()
                GridViewTitle.SelectedIndex = -1

                'GridViewDetail
                GridViewDetail.DataSource = DS_Title.Tables("data")
                GridViewDetail.DataBind()
                GridViewDetail.SelectedIndex = -1
            End If

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 単価画面から戻る時にDropDownListの初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetDropDownList()
        Try
            'タイトルの設定
            For intRow As Integer = 0 To GridViewTitle.DataSource.Rows.Count - 1
                Dim dr As DataRow = GridViewTitle.DataSource.Rows(intRow)

                If Not IsDBNull(dr("ColKata")) Then
                    If Not dr("ColKata").ToString.Equals(String.Empty) Then
                        Dim drp As New WebControls.DropDownList
                        If Me.GridViewTitle.Rows(intRow).Cells.Count > 1 Then
                            '品名が複数行の場合DropDownListは2番cellであり、単一行の場合は1番cellです
                            If TypeName(Me.GridViewTitle.Rows(intRow).Cells(1).Controls(0)).Equals("DropDownList") Then
                                drp = CType(Me.GridViewTitle.Rows(intRow).Cells(1).Controls(0), DropDownList)
                            End If
                        Else
                            If TypeName(Me.GridViewTitle.Rows(intRow).Cells(0).Controls(0)).Equals("DropDownList") Then
                                drp = CType(Me.GridViewTitle.Rows(intRow).Cells(0).Controls(0), DropDownList)
                            End If
                        End If

                        drp.SelectedValue = dr("ColKata").ToString
                    End If
                End If
            Next
            'CXA・CXBがある場合の設定
            If HidManifoldMode.Value.Equals("3") OrElse HidManifoldMode.Value.Equals("4") Then
                For intRowDetail As Integer = 0 To GridViewDetail.DataSource.Rows.Count - 1
                    Dim dr As DataRow = GridViewDetail.DataSource.Rows(intRowDetail)

                    If Not IsDBNull(dr("ColKataA")) Then
                        If Not dr("ColKataA").ToString.Trim.Equals(String.Empty) Then
                            Dim drpA As New WebControls.DropDownList
                            Dim drpB As New WebControls.DropDownList

                            If TypeName(Me.GridViewDetail.Rows(intRowDetail).Cells(0).Controls(0)).Equals("DropDownList") Then
                                drpA = CType(Me.GridViewDetail.Rows(intRowDetail).Cells(0).Controls(0), DropDownList)
                                drpB = CType(Me.GridViewDetail.Rows(intRowDetail).Cells(1).Controls(0), DropDownList)
                                drpA.SelectedValue = dr("ColKataA").ToString.Trim
                                drpB.SelectedValue = dr("ColKataB").ToString.Trim
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 単価画面から戻る時にTextBoxの初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetTextBox()
        If Not HidOther.Value.Equals(String.Empty) Then
            Dim strOthers() As String = HidOther.Value.Split(",")

            For intRow As Integer = 0 To GridViewDetail.Rows.Count - 1
                Dim dr As DataRow = GridViewDetail.DataSource.Rows(intRow)
                If Not dr("Col0").ToString.Equals(String.Empty) Then
                    '使用数がある場合
                    If Not strOthers(intRow).Trim.Equals("1") AndAlso Not strOthers(intRow).Trim.Equals("2") Then
                        '付属品の場合
                        Dim textbox As New WebControls.TextBox
                        If HidManifoldMode.Value.Equals("3") Or HidManifoldMode.Value.Equals("4") Then
                            'CXの場合は3番目は使用数
                            If Me.GridViewDetail.Rows(intRow).Cells(2).Controls.Count > 0 Then
                                textbox = CType(Me.GridViewDetail.Rows(intRow).Cells(2).Controls(0), TextBox)
                                textbox.Text = dr("Col0").ToString
                            End If
                        Else
                            If Me.GridViewDetail.Rows(intRow).Cells(0).Controls.Count > 0 Then
                                textbox = CType(Me.GridViewDetail.Rows(intRow).Cells(0).Controls(0), TextBox)
                                textbox.Text = dr("Col0").ToString
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 画面選択不可範囲の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetNoSelect()
        Try
            If HidManifoldMode.Value > 0 Then
                Dim dt_detail As DataTable = Nothing
                If Not DS_Title.Tables("data") Is Nothing Then dt_detail = DS_Title.Tables("data")
                Call subSetNoSelFlag(dt_detail)

                Dim cel As System.Web.UI.WebControls.TableCell
                Dim str_other() As String = Me.HidOther.Value.ToString.Split(",")
                For inti As Integer = 0 To GridViewDetail.Rows.Count - 1
                    If str_other.Count > inti AndAlso (str_other(inti) <> "1" And str_other(inti) <> "2") Then Continue For
                    If dt_Comb.Count > inti + 1 Then '行選択不可
                        If CType(dt_Comb(inti + 1), ArrayList).Count <= 0 Then
                            Dim intStart As Integer = 1
                            Select Case HidManifoldMode.Value
                                Case "3", "4"
                                    intStart = 3
                                Case "5", "6"
                                    Exit For
                            End Select
                            For intj As Integer = intStart To Me.GridViewDetail.Rows(inti).Cells.Count - 1
                                If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, inti, False) Then
                                    Dim GV As GridView = Me.GridViewDetail.Rows(inti).Cells(1).Controls(0)
                                    For intm As Integer = 0 To GV.Rows.Count - 1
                                        If GV.Rows(intm).Cells.Count > intj Then
                                            cel = GV.Rows(intm).Cells(intj)
                                            cel.BackColor = Drawing.Color.FromArgb(192, 192, 192)
                                            If cel.Attributes.Count > 0 Then cel.Attributes.Remove("onclick")
                                        End If
                                    Next
                                Else
                                    cel = Me.GridViewDetail.Rows(inti).Cells(intj)
                                    cel.BackColor = Drawing.Color.FromArgb(192, 192, 192)
                                    If cel.Attributes.Count > 0 Then cel.Attributes.Remove("onclick")
                                End If
                            Next
                        End If
                    End If
                Next

                'レール長さReadonly
                SetRailReadonly()

            End If
        Catch ex As Exception
            AlertMessage(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' レール長さをReadonlyに設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetRailReadonly()
        Dim celtxt As New TableCell
        Dim intRail As Integer

        intRail = GetRailTubeIndex(0) '「取付ﾚｰﾙ長さ」の行番号
        If intRail <> -1 Then
            'レール長さがある場合
            If HidManifoldMode.Value.Equals("3") OrElse HidManifoldMode.Value.Equals("4") Then
                celtxt = Me.GridViewDetail.Rows(intRail).Cells(2)
            Else
                celtxt = Me.GridViewDetail.Rows(intRail).Cells(0)
            End If
            '201502月次更新
            Dim txt As TextBox = celtxt.Controls(0)

            If objKtbnStrc.strcSelection.strSeriesKataban.Contains("MN4G") OrElse _
                objKtbnStrc.strcSelection.strSeriesKataban.Contains("MN3G") OrElse _
                objKtbnStrc.strcSelection.strSeriesKataban.Contains("M4G") OrElse _
                objKtbnStrc.strcSelection.strSeriesKataban.Contains("M3G") Then
                txt.Attributes.Add("disabled", "disabled")
            End If
        End If
    End Sub

    'RM1803032_表示サイズ調整追加
    ''' <summary>
    ''' DataGrid列の書式を設定
    ''' </summary>
    ''' <param name="subSetLbl"></param>
    ''' <remarks></remarks>
    Private Sub CreatDataGridColumn(subSetLbl As DataTable)
        Dim intColWeight As Long = CdCst.MonifoldGrid.intGridWidth
        Dim dr() As DataRow = Nothing
        Dim col_t As New Web.UI.WebControls.TemplateField
        Dim col As New Web.UI.WebControls.BoundField

        Try
            Dim intColWidth1 As Long = 0
            Dim intColWidth2 As Long = 0
            Dim intColWidthDatail As Long = 0

            'パネルサイズ変更用
            Dim intWidthPnl As Long = 980

            Select Case HidManifoldMode.Value
                Case 0
                    intColWidth1 = 130
                    intColWidth2 = 130
                Case 5, 6
                    intColWidth1 = 250
                    intColWidth2 = 250
                Case 7, 9, 13, 16, 18
                    intColWidth1 = 145
                    intColWidth2 = 185
                Case 11
                    intColWidth1 = 220
                    intColWidth2 = 220
                Case 17 'GAMD0
                    intColWidth1 = 200
                    intColWidth2 = 200
                Case 12 'VSKM
                    intColWidth1 = 190
                    intColWidth2 = 150
                Case Else
                    intColWidth1 = 150
                    intColWidth2 = 150
            End Select
            intWidthPnl -= (intColWidth1 + intColWidth2)
            PnlDetail2.Width = WebControls.Unit.Pixel(intWidthPnl)
            'スクロールバー表示切替（４０連対応）
            Select Case HidManifoldMode.Value
                Case 1, 2, 4, 7, 15, 16, 18
                    PnlDetail2.ScrollBars = ScrollBars.Horizontal
            End Select
            
            Select Case HidManifoldMode.Value
                Case 0 '簡易モード
                    dr = subSetLbl.Select("label_seq='5' AND label_div='L'")
                    col = New Web.UI.WebControls.BoundField
                    col.DataField = "ColNo"
                    col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
                    col.ItemStyle.BackColor = Color.FromArgb(192, 192, 255)
                    col.ItemStyle.Wrap = True
                    col.ItemStyle.Width = WebControls.Unit.Pixel(intColWidth1)
                    col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                    If dr.Length > 0 Then col.HeaderText = dr(0)("label_content").ToString
                    GridViewTitle.Columns.Add(col)

                    dr = subSetLbl.Select("label_seq='2' AND label_div='L'")
                    col = New Web.UI.WebControls.BoundField
                    col.DataField = "ColKata"
                    col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
                    col.ItemStyle.Wrap = True
                    col.ItemStyle.Width = WebControls.Unit.Pixel(intColWidth2)
                    col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                    If dr.Length > 0 Then col.HeaderText = dr(0)("label_content").ToString
                    GridViewTitle.Columns.Add(col)
                    GridViewTitle.Width = WebControls.Unit.Pixel(intColWidth1 + intColWidth2)
                Case Else
                    dr = subSetLbl.Select("label_seq='1' AND label_div='L'")
                    col = New Web.UI.WebControls.BoundField
                    col.DataField = "ColNo"
                    col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
                    col.HtmlEncode = False
                    col.ItemStyle.BackColor = Color.FromArgb(192, 192, 255)
                    col.ItemStyle.Wrap = True
                    col.ItemStyle.Width = WebControls.Unit.Pixel(intColWidth1)
                    col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                    If dr.Length > 0 Then col.HeaderText = dr(0)("label_content").ToString
                    GridViewTitle.Columns.Add(col)

                    dr = subSetLbl.Select("label_seq='2' AND label_div='L'")
                    col_t = New Web.UI.WebControls.TemplateField
                    col_t.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
                    col_t.ItemStyle.Wrap = True
                    col.ItemStyle.Width = WebControls.Unit.Pixel(intColWidth2)
                    col_t.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                    If dr.Length > 0 Then col_t.HeaderText = dr(0)("label_content").ToString
                    GridViewTitle.Columns.Add(col_t)
                    GridViewTitle.Width = WebControls.Unit.Pixel(intColWidth1 + intColWidth2)

                    Select Case HidManifoldMode.Value
                        Case 3, 4 'CXA,CXB
                            col = New Web.UI.WebControls.BoundField

                            col.DataField = "ColKataA"
                            col.HeaderText = "CX A"
                            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
                            col.ItemStyle.Wrap = False
                            col.ItemStyle.Width = WebControls.Unit.Pixel(70)
                            col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                            GridViewDetail.Columns.Add(col)

                            col = New Web.UI.WebControls.BoundField
                            col.DataField = "ColKataB"
                            col.HeaderText = "CX B"
                            col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
                            col.ItemStyle.Wrap = False
                            col.ItemStyle.Width = WebControls.Unit.Pixel(70)
                            col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                            GridViewDetail.Columns.Add(col)

                            intColWidthDatail += 140
                    End Select
            End Select

            '※電気接続がT0Dの場合、設置位置を逆にする
            If objKtbnStrc.strcSelection.strSeriesKataban = "LMF0" AndAlso _
                objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "T0D" Then
                col = New Web.UI.WebControls.BoundField
                col.DataField = "Col" & 0
                col.ItemStyle.Wrap = False
                col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
                dr = subSetLbl.Select("label_seq='3' AND label_div='L'")
                If dr.Length > 0 Then col.HeaderText = dr(0)("label_content").ToString
                col.ItemStyle.HorizontalAlign = HorizontalAlign.Center
                col.ItemStyle.Width = WebControls.Unit.Pixel(70)
                GridViewDetail.Columns.Add(col)

                intColWidthDatail += 70

                For inti As Integer = HidColCount.Value To 1 Step -1
                    col = New Web.UI.WebControls.BoundField
                    col.DataField = "Col" & inti
                    col.HeaderText = inti.ToString.PadLeft(2, "0")
                    col.ItemStyle.Wrap = False
                    col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                    col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
                    col.ItemStyle.BackColor = Drawing.Color.FromArgb(255, 255, 192)
                    col.ItemStyle.HorizontalAlign = HorizontalAlign.Center
                    col.ItemStyle.Width = WebControls.Unit.Pixel(intColWeight)
                    GridViewDetail.Columns.Add(col)
                    intColWidthDatail += intColWeight
                Next
            Else
                For inti As Integer = 0 To HidColCount.Value
                    col = New Web.UI.WebControls.BoundField
                    col.DataField = "Col" & inti
                    col.HeaderText = inti.ToString.PadLeft(2, "0")
                    col.ItemStyle.Wrap = False
                    col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                    col.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
                    If inti = 0 Then
                        dr = subSetLbl.Select("label_seq='3' AND label_div='L'")
                        If dr.Length > 0 Then col.HeaderText = dr(0)("label_content").ToString
                        col.ItemStyle.HorizontalAlign = HorizontalAlign.Center
                        col.ItemStyle.Width = WebControls.Unit.Pixel(70)
                        intColWidthDatail += 70
                    Else
                        col.ItemStyle.BackColor = Drawing.Color.FromArgb(255, 255, 192)
                        col.ItemStyle.HorizontalAlign = HorizontalAlign.Center
                        col.ItemStyle.Width = WebControls.Unit.Pixel(intColWeight)
                        intColWidthDatail += intColWeight
                    End If
                    GridViewDetail.Columns.Add(col)
                Next
            End If
            GridViewDetail.Width = WebControls.Unit.Pixel(intColWidthDatail)

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 簡易マニホルード仕様
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <remarks></remarks>
    Private Sub CreatSiyou_DT(ByVal dt As DataTable)
        Dim dt_title As New DataTable("title")
        Dim dt_detail As New DataTable("data")
        Try
            Dim dc_t As New DataColumn
            dc_t = New DataColumn("ColNo")            '記号
            dt_title.Columns.Add(dc_t)
            dc_t = New DataColumn("ColKata")          '形番
            dt_title.Columns.Add(dc_t)

            dc_t = New DataColumn("Col0")             '使用数
            dt_detail.Columns.Add(dc_t)
            For inti As Integer = 0 To HidColCount.Value - 1     '列生成
                dc_t = New DataColumn("Col" & (inti + 1))
                dt_detail.Columns.Add(dc_t)
            Next

            Dim dr_1() As DataRow = dt.Select("item_div='1'")
            If dr_1.Count <= 0 Then
                '対応言語がない場合はdefault言語で設定
                dr_1 = dt.Select("df_item_div='1'")
                For inti As Integer = 0 To dr_1.Length - 1
                    Dim dr As DataRow = dt_title.NewRow
                    dr("ColNo") = dr_1(inti)("df_label_content").ToString
                    dt_title.Rows.Add(dr)
                Next
            Else
                For inti As Integer = 0 To dr_1.Length - 1
                    Dim dr As DataRow = dt_title.NewRow
                    dr("ColNo") = dr_1(inti)("label_content").ToString
                    dt_title.Rows.Add(dr)
                Next
            End If

            Dim dr_2() As DataRow = dt.Select("item_div='2'")
            If dr_2.Count <= 0 Then
                '対応言語がない場合はdefault言語で設定
                dr_2 = dt.Select("df_item_div='2'")
                For inti As Integer = 0 To dr_2.Length - 1
                    If dr_1.Length > inti Then
                        dt_title.Rows(inti)("ColKata") = dr_2(inti)("df_label_content").ToString
                    Else
                        Exit For
                    End If
                Next
            Else
                For inti As Integer = 0 To dr_2.Length - 1
                    If dr_1.Length > inti Then
                        dt_title.Rows(inti)("ColKata") = dr_2(inti)("label_content").ToString
                    Else
                        Exit For
                    End If
                Next
            End If

            Dim dr_detail As DataRow = Nothing
            For inti As Integer = 0 To dt_title.Rows.Count - 1
                dr_detail = dt_detail.NewRow
                dt_detail.Rows.Add(dr_detail)
            Next

            '特殊形番の作成
            subChangeKataban(dt_title)

            DS_Title.Tables.Add(dt_title)
            DS_Title.Tables.Add(dt_detail)

            If dr_1.Length > 0 Then
                Dim strOthers As String = dr_1(0)("others").ToString.Trim
                Dim strOthersDefault As String = dr_1(0)("df_others").ToString.Trim

                If strOthers.Equals(String.Empty) Then
                    Me.HidSimpleOther.Value = strOthersDefault
                Else
                    Me.HidSimpleOther.Value = strOthers
                End If
            End If


            PnlDetail.Height = WebControls.Unit.Pixel(580)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 簡易マニホールドの特殊形番変換
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subChangeKataban(ByRef dt_title As DataTable)
        Dim strNewKataList As New ArrayList
        Dim dtSpecItem As New DataTable
        Dim dtContent As New DataTable

        With objKtbnStrc.strcSelection
            Call KHManifold.subInitTable(dtSpecItem, dtContent)
            Call SiyouDAL.subSQL_ItemMst(objCon, .strSpecNo.Trim, dtSpecItem, dtContent, Me.objLoginInfo.SelectLang)
            '特殊形番の作成
            strNewKataList = KHManifold.fncGetNewKataban(dtSpecItem, dtContent, .strSpecNo.Trim, _
                                                             .strSeriesKataban.Trim, .strOpSymbol, _
                                                             .strKeyKataban, selLang.SelectedValue.Trim)
            If strNewKataList.Count >= dt_title.Rows.Count Then
                For inti As Integer = 0 To dt_title.Rows.Count - 1
                    Dim strTmp() As String = strNewKataList(inti).ToString.Split("_")

                    If strTmp.Length >= 4 Then
                        dt_title.Rows(inti).Item("ColKata") = strTmp(1)
                    End If
                Next
            End If
        End With
    End Sub

    ''' <summary>
    ''' 複雑版仕様
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <remarks></remarks>
    Private Sub CreatSiyou_M(ByVal dt As DataTable)
        Dim dt_title As New DataTable("title")
        Dim dt_detail As New DataTable("data")

        Try
            Dim dc_t As New DataColumn
            dc_t = New DataColumn("ColNo")       '記号/品名
            dt_title.Columns.Add(dc_t)
            dc_t = New DataColumn("ColKata")     '形番
            dt_title.Columns.Add(dc_t)

            Select Case HidManifoldMode.Value
                Case 3, 4
                    dc_t = New DataColumn("ColKataA")     '継手CX、A
                    dt_detail.Columns.Add(dc_t)
                    dc_t = New DataColumn("ColKataB")     '継手CX、B
                    dt_detail.Columns.Add(dc_t)
            End Select

            dc_t = New DataColumn("Col0")        '使用数
            dt_detail.Columns.Add(dc_t)

            For inti As Integer = 0 To HidColCount.Value - 1     '列生成
                dc_t = New DataColumn("Col" & (inti + 1))
                dt_detail.Columns.Add(dc_t)
            Next

            Dim strSpecNo As String = objKtbnStrc.strcSelection.strSpecNo.ToString
            Select Case HidManifoldMode.Value
                Case 12, 17
                    Dim intKey As Integer = -1
                    Select Case strSpecNo
                        Case "12", "20", "22"
                            intKey = 9
                        Case "17"
                            intKey = 11
                        Case "18", "21"
                            intKey = 8
                        Case "19"
                            intKey = 10
                        Case "23", "95"
                            intKey = 7
                        Case "94"
                            intKey = 13
                    End Select
                    Dim dr_1() As DataRow = dt.Select("", "label_seq")
                    Dim dr As DataRow = Nothing
                    Dim dr_detail As DataRow = Nothing
                    For inti As Integer = 0 To dr_1.Length - 1
                        If dr_1(inti)("label_content").ToString.Length <= 0 Then Continue For
                        Dim lbl_seq As String = dr_1(inti)("label_seq").ToString
                        For intj As Integer = 0 To CInt(dr_1(inti)("item_num").ToString) - 1
                            dr = dt_title.NewRow
                            dr_detail = dt_detail.NewRow
                            dr("ColNo") = dr_1(inti)("label_content").ToString
                            If HidColMerge.Value.Length > 0 Then HidColMerge.Value &= ","
                            If intj = 0 Then
                                HidColMerge.Value &= CInt(dr_1(inti)("item_num").ToString)
                            Else
                                HidColMerge.Value &= 0
                            End If
                            dt_detail.Rows.Add(dr_detail)
                            dt_title.Rows.Add(dr)

                            If HidOther.Value.Length > 0 Then HidOther.Value &= ","
                            HidOther.Value &= CInt(dr_1(inti)("item_div").ToString)
                        Next
                    Next
                    dr_1 = Nothing
                Case Else
                    Dim dr_1() As DataRow = dt.Select("", "label_seq")
                    Dim dr As DataRow = Nothing
                    Dim dr_detail As DataRow = Nothing
                    For inti As Integer = 0 To dr_1.Length - 1
                        If dr_1(inti)("label_content").ToString.Length <= 0 Then Continue For
                        Dim lbl_seq As String = dr_1(inti)("label_seq").ToString
                        If HidManifoldMode.Value = 14 And lbl_seq = "2" Then
                            dr_1(inti)("label_content") = dr_1(inti)("label_content").ToString.Replace("／", "<BR>").Replace("/", "<BR>")
                        End If
                        For intj As Integer = 0 To CInt(dr_1(inti)("item_num").ToString) - 1
                            dr = dt_title.NewRow
                            dr_detail = dt_detail.NewRow
                            dr("ColNo") = dr_1(inti)("label_content").ToString

                            If intj = 0 Then
                                If objKtbnStrc.strcSelection.strSeriesKataban = "MW4GB4" And _
                                    (inti = 8 Or inti = 9 Or inti = 10) Then
                                    dr("ColNo") = String.Empty
                                Else
                                    dr("ColNo") = dr_1(inti)("label_content").ToString
                                End If
                            End If

                            If HidManifoldMode.Value = 9 AndAlso _
                                (dr_1(inti)("label_seq") = "14" Or dr_1(inti)("label_seq") = "15") Then
                                '六角穴付プラグ特殊対応(ラベル番号14、15)
                                If dr_1(inti)("label_seq") = "14" Then
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                        Case "M4TB3"
                                            dr("ColNo") &= "R1/4"
                                        Case "M4TB4"
                                            dr("ColNo") &= "R3/8"
                                    End Select
                                Else
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                        Case "M4TB3"
                                            dr("ColNo") &= "R3/8"
                                        Case "M4TB4"
                                            dr("ColNo") &= "R1/2"
                                    End Select
                                End If
                            End If

                            If HidOther.Value.Length > 0 Then HidOther.Value &= ","
                            'タグ銘板判断
                            Select Case objKtbnStrc.strcSelection.strSpecNo
                                Case "07", "96"
                                    If dr_1(inti)("label_seq") = "12" Then
                                        HidOther.Value &= "99"
                                    Else
                                        HidOther.Value &= CInt(dr_1(inti)("item_div").ToString)
                                    End If
                                Case "15"
                                    '14->タグ銘板、13->ケーブルクランプ    'ADD BY YGY 20140909
                                    If dr_1(inti)("label_seq") = "14" OrElse _
                                        dr_1(inti)("label_seq") = "13" Then
                                        HidOther.Value &= "99"
                                    Else
                                        HidOther.Value &= CInt(dr_1(inti)("item_div").ToString)
                                    End If
                                Case Else
                                    HidOther.Value &= CInt(dr_1(inti)("item_div").ToString)
                            End Select

                            If HidColMerge.Value.Length > 0 Then HidColMerge.Value &= ","
                            If intj = 0 Then
                                HidColMerge.Value &= CInt(dr_1(inti)("item_num").ToString)
                            Else
                                HidColMerge.Value &= 0
                            End If
                            dt_detail.Rows.Add(dr_detail)
                            dt_title.Rows.Add(dr)
                        Next
                    Next
                    dr_1 = Nothing
            End Select

            DS_Title.Tables.Add(dt_title)
            DS_Title.Tables.Add(dt_detail)

            If dt_title.Rows.Count * (CdCst.MonifoldGrid.intGridHeight + 2) <= 570 Then
                PnlDetail.Height = WebControls.Unit.Pixel(580)
            Else
                PnlDetail.Height = WebControls.Unit.Pixel(dt_title.Rows.Count * (CdCst.MonifoldGrid.intGridHeight + 2) + 10)
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    '''  縦列ごとの選択可否フラグ設定(例外設定のみ)
    ''' </summary>
    ''' <remarks></remarks>
    Private Function subSetNoSelFlag(dt_detail As DataTable) As Boolean
        subSetNoSelFlag = True
        Try
            Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
            Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban
            Dim strValue() As String = objKtbnStrc.strcSelection.strOpSymbol
            Select Case HidManifoldMode.Value
                Case 1
                    Dim strOptions() As String
                    Dim Tflag As Boolean = False

                    Select Case strSeriesKata
                        Case "MN3EX0", "MN4EX0"
                            strOptions = strValue(4).Split(CdCst.Sign.Delimiter.Comma)
                        Case "MN3Q0", "MT3Q0"
                            strOptions = strValue(5).Split(CdCst.Sign.Delimiter.Comma)
                        Case Else
                            strOptions = strValue(6).Split(CdCst.Sign.Delimiter.Comma)
                    End Select

                    For inti As Integer = 0 To strOptions.Length - 1
                        If strOptions(inti).ToString.Length <= 0 Then Continue For
                        If strOptions(inti).StartsWith("T") Then
                            Tflag = True
                        End If
                    Next
                    For inti As Integer = 0 To strOptions.Length - 1
                        If strOptions(inti).ToString.Length <= 0 Then Continue For
                        If Tflag Then
                            Select Case strOptions(inti)
                                Case "T30", "T30N", "T50", "T51", "T52", "T53", "T5B", "T5C", "T631", "T6A0", _
                                     "T6A1", "T6C0", "T6C1", "T6E0", "T6E1", "T6G1", "T6J0", "T6J1", "T6K1", _
                                     "T7D1", "T7D2", "T7G1", "T7G2", "T7N1", "T7N2", _
                                     "T7EC1", "T7EC2", "T7ECT1", "T7ECT2" '2016/08/23 RM1608024 T7EC Append
                                    Call SetKoteiCol(dt_detail, CdCst.Siyou_01.Elect1 - 1, 1)
                                Case "T6K1"
                                    If strSeriesKata = "MN3E0" Or strSeriesKata = "M4E0" Then
                                        Call SetKoteiCol(dt_detail, CdCst.Siyou_01.Elect1 - 1, 1)
                                    End If
                                Case "TM1A", "TM1B", "TM1C", "TM52", "T30R", "T30NR", "T50R", "T51R", "T52R", "T53R"
                                    Call SetKoteiCol(dt_detail, CdCst.Siyou_01.EndL - 1, 1)
                            End Select
                        Else
                            Call SetKoteiCol(dt_detail, CdCst.Siyou_01.EndL - 1, 1)
                        End If
                    Next
                Case 2
                    Call SetKoteiCol(dt_detail, CdCst.Siyou_02.End1 - 1, 1)
                Case 3
                    '１５行目変更不可
                    If Strings.Left(strValue(4).ToString, 1) <> "8" Then
                        SetKoteiRowOnly(CdCst.Siyou_03.Masking - 1)
                    End If
                Case 4
                    '２０行目変更不可,レール長さ
                Case 5      'CMF ISO
                    Select Case strKeyKata
                        Case "1"
                            If strValue(1).ToString <> "Z" Then
                                'ABポート口径・ABポート位置・流露遮蔽板以外の行を選択不可にする
                                For intI As Integer = CdCst.Siyou_05.ElType1 - 1 To CdCst.Siyou_05.ElType6 - 1
                                    SetKoteiRowOnly(intI)
                                Next
                                For intI As Integer = CdCst.Siyou_05.RepSpace1 - 1 To CdCst.Siyou_05.SpDecomp4 - 1
                                    SetKoteiRowOnly(intI)
                                Next
                                For inti As Integer = 1 To CInt(strValue(2))
                                    Call SetKoteiCol(dt_detail, CdCst.Siyou_05.ElType1 - 1, inti, False)
                                Next
                            Else
                                'ABポート口径・ABポート位置・流露遮蔽板以外の行を選択不可にする
                                For intI As Integer = CdCst.Siyou_05.ElType3 - 1 To CdCst.Siyou_05.ElType6 - 1
                                    SetKoteiRowOnly(intI)
                                Next
                                For intI As Integer = CdCst.Siyou_05.RepSpace1 - 1 To CdCst.Siyou_05.SpDecomp4 - 1
                                    SetKoteiRowOnly(intI)
                                Next
                            End If

                        Case "8"
                            If strValue(9).ToString.Length <= 0 Then
                                If strValue(1).ToString <> "Z" Then
                                    'ABポート口径・ABポート位置・流露遮蔽板以外の行を選択不可にする
                                    For intI As Integer = CdCst.Siyou_05.ElType1 - 1 To CdCst.Siyou_05.ElType6 - 1
                                        SetKoteiRowOnly(intI)
                                    Next
                                    For intI As Integer = CdCst.Siyou_05.RepSpace1 - 1 To CdCst.Siyou_05.SpDecomp4 - 1
                                        SetKoteiRowOnly(intI)
                                    Next

                                    For inti As Integer = 1 To CInt(strValue(2))
                                        Call SetKoteiCol(dt_detail, CdCst.Siyou_05.ElType1 - 1, inti)
                                    Next
                                Else
                                    'ABポート口径・ABポート位置・流露遮蔽板以外の行を選択不可にする
                                    For intI As Integer = CdCst.Siyou_05.ElType3 - 1 To CdCst.Siyou_05.ElType6 - 1
                                        SetKoteiRowOnly(intI)
                                    Next
                                    For intI As Integer = CdCst.Siyou_05.RepSpace1 - 1 To CdCst.Siyou_05.SpDecomp4 - 1
                                        SetKoteiRowOnly(intI)
                                    Next
                                End If
                            End If
                        Case "4", "5", "6", "7", "9"
                            Select Case strKeyKata
                                Case "4", "6"
                                    'ABポート口径・ABポート位置・流露遮蔽板以外の行を選択不可にする
                                    For intI As Integer = CdCst.Siyou_05.ElType1 - 1 To CdCst.Siyou_05.ElType6 - 1
                                        SetKoteiRowOnly(intI)
                                    Next
                                    For intI As Integer = CdCst.Siyou_05.RepSpace1 - 1 To CdCst.Siyou_05.SpDecomp4 - 1
                                        SetKoteiRowOnly(intI)
                                    Next

                                    For inti As Integer = 3 To CInt(strValue(2))
                                        Call SetKoteiCol(dt_detail, CdCst.Siyou_05.ElType1 - 1, inti)
                                    Next
                            End Select
                            Call SetKoteiColOnly(1)
                            Call SetKoteiColOnly(2)
                            Call SetKoteiMidRow(1)  'Mid固定
                            Call SetKoteiMidRow(2)  'Mid固定
                            For inti As Integer = CInt(strValue(2)) To CInt(HidColCount.Value) - 1
                                Call SetKoteiMidRow(inti)  'Mid固定
                            Next
                            For inti As Integer = CInt(strValue(2)) + 1 To CInt(HidColCount.Value)
                                Call SetKoteiColOnly(inti) '列固定
                            Next
                    End Select

                    '選択した連数より多く選択できないように ADD BY YGY 20141119
                    '正常列
                    For inti As Integer = CInt(strValue(2)) + 1 To dt_detail.Columns.Count - 1
                        Call SetKoteiColOnly(inti)
                    Next
                    'Mid列
                    For inti As Integer = CInt(strValue(2)) To dt_detail.Columns.Count - 2
                        Call SetKoteiMidRow(inti)
                    Next

                    'ABポート位置設定
                    If strValue(4).ToString.Trim <> "L" Then
                        For intI As Integer = CdCst.Siyou_05.ABPlugR - 1 To CdCst.Siyou_05.ABPlugL - 1
                            SetKoteiRowOnly(intI)
                        Next
                    End If
                    'ABポート口径設定
                    Select Case strValue(3).ToString.Trim
                        Case "02", "03", "04", "HX3", "HX4", "HX5", "HX6"
                            SetKoteiRowOnly(CdCst.Siyou_05.ABCon02 - 1)
                            SetKoteiRowOnly(CdCst.Siyou_05.ABCon03 - 1)
                            SetKoteiRowOnly(CdCst.Siyou_05.ABCon04 - 1)
                        Case "HX1"
                            SetKoteiRowOnly(CdCst.Siyou_05.ABCon04 - 1)
                        Case "HX2"
                            SetKoteiRowOnly(CdCst.Siyou_05.ABCon02 - 1)
                    End Select
                Case 6
                    '※電気接続がT0Dの場合、設置位置を逆にする
                    '****
                    '02BとC8B以外は流露遮蔽板を選択不可にする
                    Select Case strValue(3).ToString
                        Case "02B", "C8B"
                            If strValue(4).ToString = "T0D" Then
                                For inti As Integer = CInt(strValue(1)) To CInt(HidColCount.Value) - 1
                                    Call SetKoteiMidRow(10 - inti)  'Mid固定
                                Next
                            Else
                                For inti As Integer = CInt(strValue(1)) To CInt(HidColCount.Value) - 1
                                    Call SetKoteiMidRow(inti)  'Mid固定
                                Next
                            End If
                        Case Else
                            For inti As Integer = 1 To CInt(HidColCount.Value) - 1
                                Call SetKoteiMidRow(inti)  'Mid固定
                            Next
                    End Select

                    If strValue(5).ToString = String.Empty Then
                        For inti As Integer = 1 To CInt(strValue(1))
                            'T0Dの場合は画面選択位置を逆にセットする
                            If strValue(4).ToString = "T0D" Then
                                Call SetKoteiCol(dt_detail, CdCst.Siyou_06.Elect1 - 1, inti, True, True)
                            Else
                                Call SetKoteiCol(dt_detail, CdCst.Siyou_06.Elect1 - 1, inti)
                            End If

                        Next
                        For inti As Integer = CdCst.Siyou_06.Elect1 - 1 To CdCst.Siyou_06.Elect6 - 1
                            SetKoteiRowOnly(inti)
                        Next
                    End If

                    If strValue(4).ToString = "T0D" Then
                        For inti As Integer = CInt(strValue(1)) + 1 To CInt(HidColCount.Value)
                            Call SetKoteiColOnly(11 - inti) '列固定
                        Next
                    Else
                        For inti As Integer = CInt(strValue(1)) + 1 To CInt(HidColCount.Value)
                            Call SetKoteiColOnly(inti) '列固定
                        Next
                    End If

                    'ABポート位置設定
                    If strValue(2).ToString.Trim <> "XX" Then
                        For intI As Integer = CdCst.Siyou_06.ABCon01 - 1 To CdCst.Siyou_06.ABCon1Z - 1
                            SetKoteiRowOnly(intI)
                        Next
                    End If

                Case 7
                    Dim strDensen As String = ""        '電線接続
                    'If strKeyKata = "R" Or strKeyKata = "U" Then
                    If strKeyKata = "R" Or strKeyKata = "U" Or strKeyKata = "S" Or strKeyKata = "V" Then 'RM1610013
                        strDensen = strValue(5)
                    Else
                        strDensen = strValue(4)
                    End If
                    If Strings.Left(strDensen.ToString, 1) = "T" Then
                        If Strings.Left(strDensen.ToString & Space(4), 4).Substring(3, 1) = "R" Then
                            '１７行１列目固定
                            SetKoteiCol(dt_detail, CdCst.Siyou_07.EndLeft - 1, 1)
                        Else
                            '１行１列目固定
                            SetKoteiCol(dt_detail, CdCst.Siyou_07.Equip - 1, 1)
                        End If
                    Else
                        '１７行１列目固定
                        SetKoteiCol(dt_detail, CdCst.Siyou_07.EndLeft - 1, 1)
                    End If
                Case 8
                    Call SetKoteiCol(dt_detail, CdCst.Siyou_08.EndP1 - 1, 1)
                Case 9
                    Call SetKoteiCol(dt_detail, CdCst.Siyou_09.Endb1 - 1, 1)
                    If strValue(6).Trim.Equals("T10") AndAlso strValue(7).Contains("CL") Then
                    Else
                        Call SetKoteiCol(dt_detail, CdCst.Siyou_09.Wiring - 1, 2)
                    End If
                    Call SetKoteiMidRow(1)  'Mid固定 
                Case 10
                    Select Case strValue(6).ToString
                        Case "T10", "T11", "T30", "T50", "T621", "T631", "T6A0", "T6A1", "T6C0", "T6C1", "T6E0", "T6E1", "T6G1", "T6J0", "T6J1", "T6K1", "T30N"
                            SetKoteiCol(dt_detail, 0, 1)
                        Case "T10R", "T11R", "T30R", "T30NR", "T50R", "C", "C0", "C1", "C2"
                            SetKoteiCol(dt_detail, 12, 1)
                    End Select
                Case 11
                    SetKoteiCol(dt_detail, CdCst.Siyou_11.EndL - 1, 1)
                Case 12
                    Dim intMaxSeq As Integer = 0
                    Select Case strSeriesKata
                        Case "VSJM", "VSXM", "VSZM", "VSNM", "VSNM"
                            intMaxSeq = Int(strValue(8))
                        Case "VSJPM"
                            intMaxSeq = Int(strValue(7))
                        Case "VSXPM"
                            intMaxSeq = Int(strValue(6))
                        Case "VSKM"
                            intMaxSeq = Int(strValue(9))
                        Case "VSZPM", "VSNPM"
                            intMaxSeq = Int(strValue(5))
                    End Select

                    For inti As Integer = intMaxSeq + 1 To 12
                        Call SetKoteiColOnly(inti) '選択不可、色も変更
                    Next
                Case 14
                    Call SetKoteiCol(dt_detail, CdCst.Siyou_14.End1 - 1, 1)
                Case 15
                    Dim strOptionY As String = Nothing
                    Dim str() As String = strValue(6).ToString.Split(CdCst.Sign.Delimiter.Comma)
                    For intI As Integer = 0 To str.Length - 1
                        If str(intI).Contains("Y") Then
                            strOptionY = str(intI)
                        End If
                    Next

                    If strOptionY Is Nothing Then
                        If strValue(4).ToString = "R1" Then
                            SetKoteiCol(dt_detail, CdCst.Siyou_15.EndL - 1, 1)
                        Else
                            SetKoteiCol(dt_detail, CdCst.Siyou_15.Elect - 1, 1)
                        End If
                    Else
                        Select Case strOptionY
                            Case "Y10", "Y20", "Y30", "Y40", "Y01", "Y02", "Y03", "Y04", _
                                 "Y11", "Y21", "Y31", "Y41", "Y12", "Y22", "Y32", "Y42"
                                Dim intOne As Integer = Int(strOptionY.Substring(2, 1))
                                For inti As Integer = 0 To intOne - 1
                                    SetKoteiCol(dt_detail, CdCst.Siyou_15.InOut2 - 1, (inti + 1))
                                Next

                                Dim intTwo As Integer = Int(strOptionY.Substring(1, 1))
                                For inti As Integer = intOne To intOne + intTwo - 1
                                    SetKoteiCol(dt_detail, CdCst.Siyou_15.InOut1 - 1, (inti + 1))
                                Next

                                SetKoteiCol(dt_detail, CdCst.Siyou_15.Elect - 1, (intOne + intTwo + 1))
                        End Select
                    End If
                Case 16
                    Dim str() As String
                    Select Case strKeyKata
                        Case "S", "Y"
                            str = strValue(7).ToString.Split(CdCst.Sign.Delimiter.Comma)
                        Case Else
                            str = strValue(6).ToString.Split(CdCst.Sign.Delimiter.Comma)
                    End Select

                    Dim strOptionY As String = Nothing
                    'Dim str() As String = strValue(6).ToString.Split(CdCst.Sign.Delimiter.Comma)
                    For intI As Integer = 0 To str.Length - 1
                        If str(intI).Contains("Y") Then
                            strOptionY = str(intI)
                        End If
                    Next

                    If strOptionY Is Nothing Then
                        '省配線接続が「T7※1」の場合はエンドブロックLではなく配線ブロックを固定とする  RM1705010 修正 2017/05/11 
                        If strSeriesKata = "MW4GB4" Or strSeriesKata = "MW4GZ4" Then
                            'If strValue(4).ToString = "T7EC1" Or strValue(4).ToString = "T7ECP1" Or _
                            '   strValue(4).ToString = "T7EN1" Or strValue(4).ToString = "T7ENP1" Then
                            If strValue(4).ToString.StartsWith("T7") Then
                                SetKoteiCol(dt_detail, CdCst.Siyou_16.Elect - 1, 1)
                            Else
                                SetKoteiCol(dt_detail, CdCst.Siyou_16.EndL - 1, 1)
                            End If
                        Else
                            SetKoteiCol(dt_detail, CdCst.Siyou_16.EndL - 1, 1)
                        End If
                        'SetKoteiCol(dt_detail, CdCst.Siyou_16.EndL - 1, 1)
                    Else
                        Select Case strOptionY
                            Case "Y10", "Y20", "Y30", "Y40", "Y01", "Y02", "Y03", "Y04", _
                                 "Y11", "Y21", "Y31", "Y41", "Y12", "Y22", "Y32", "Y42"
                                Dim intOne As Integer = Int(strOptionY.Substring(2, 1))
                                For inti As Integer = 0 To intOne - 1
                                    SetKoteiCol(dt_detail, CdCst.Siyou_16.InOut2 - 1, (inti + 1))
                                Next

                                Dim intTwo As Integer = Int(strOptionY.Substring(1, 1))
                                For inti As Integer = intOne To intOne + intTwo - 1
                                    SetKoteiCol(dt_detail, CdCst.Siyou_16.InOut1 - 1, (inti + 1))
                                Next

                                SetKoteiCol(dt_detail, CdCst.Siyou_16.Elect - 1, (intOne + intTwo + 1))
                        End Select
                    End If
                Case 17
                    For inti As Integer = CInt(strValue(6).ToString) + 1 To 5
                        Call SetKoteiColOnly(inti) '選択不可、色も変更
                    Next
                Case 18
                    Select Case strKeyKata
                        Case "U", "R"
                            If strValue(5).ToString.StartsWith("T") Then
                                If Strings.Left(strValue(5).ToString & Space(4), 4).Substring(3, 1) = "R" Then
                                    '１７行１列目固定
                                    Call SetKoteiCol(dt_detail, CdCst.Siyou_18.EndLeft - 1, 1)
                                Else
                                    '１行１列目固定
                                    Call SetKoteiCol(dt_detail, CdCst.Siyou_18.Equip - 1, 1)
                                End If
                            Else
                                '１７行１列目固定
                                Call SetKoteiCol(dt_detail, CdCst.Siyou_18.EndLeft - 1, 1)
                            End If
                        Case Else
                            If strValue(4).ToString.StartsWith("T") Then
                                If Strings.Left(strValue(4).ToString & Space(4), 4).Substring(3, 1) = "R" Then
                                    '１７行１列目固定
                                    Call SetKoteiCol(dt_detail, CdCst.Siyou_18.EndLeft - 1, 1)
                                Else
                                    '１行１列目固定
                                    Call SetKoteiCol(dt_detail, CdCst.Siyou_18.Equip - 1, 1)
                                End If
                            Else
                                '１７行１列目固定
                                Call SetKoteiCol(dt_detail, CdCst.Siyou_18.EndLeft - 1, 1)
                            End If
                    End Select
            End Select

        Catch ex As Exception
            Call AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    '''  初期値の設定(X行X列目固定)
    ''' </summary>
    ''' <param name="dt_detail">バインドデータ</param>
    ''' <param name="intRow">行番号</param>
    ''' <param name="intCol">列番号</param>
    ''' <param name="blnRemoveClick"></param>
    ''' <param name="blnReverse">逆にセットフラグ</param>
    ''' <remarks></remarks>
    Private Sub SetKoteiCol(dt_detail As DataTable, _
                            ByVal intRow As Integer, _
                            ByVal intCol As Integer, _
                            Optional ByVal blnRemoveClick As Boolean = True, _
                            Optional ByVal blnReverse As Boolean = False)
        Try
            Dim cel As System.Web.UI.WebControls.TableCell
            For inti As Integer = 0 To Me.GridViewDetail.Rows.Count - 1
                If Me.GridViewDetail.Rows(inti).Cells.Count > intCol Then
                    Dim bolStart As Boolean = False
                    If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, inti, bolStart) AndAlso bolStart Then 'Mid選択欄
                        'DELETE BY YGY 20141126 LMF
                        'Dim GV As GridView = Me.GridViewDetail.Rows(inti).Cells(1).Controls(0)
                        'For intm As Integer = 0 To GV.Rows.Count - 1
                        '    If GV.Rows(intm).Cells.Count > intCol Then
                        '        cel = GV.Rows(intm).Cells(intCol)
                        '        'CHANGED BY YGY 20141119 CMF22-03L-04B-SB
                        '        If blnRemoveClick Then
                        '            cel.Attributes.Remove("onclick")
                        '            cel.BackColor = Drawing.Color.FromArgb(192, 192, 192)
                        '        End If
                        '    End If
                        'Next
                        Continue For
                    End If

                    If blnReverse Then
                        'LMF T0Dの場合は位置を逆に設定する
                        cel = Me.GridViewDetail.Rows(inti).Cells(11 - intCol)
                    Else
                        cel = Me.GridViewDetail.Rows(inti).Cells(intCol)
                    End If

                    'CHANGED BY YGY 20141119
                    If inti = intRow Then
                        cel.Text = strMaru
                        cel.Attributes.Remove("onclick")
                        cel.BackColor = Drawing.Color.FromArgb(192, 192, 192)
                    Else
                        If blnRemoveClick Then
                            cel.Attributes.Remove("onclick")
                            cel.BackColor = Drawing.Color.FromArgb(192, 192, 192)
                        End If
                    End If
                End If
            Next
            dt_detail.Rows(intRow)(intCol) = strMaru                          '●設置
            'HiddenFieldにも位置情報を登録
            If Not HidClick.Value.Contains(intRow + 2 & "," & intCol & ";") Then
                HidClick.Value &= intRow + 2 & "," & intCol & ";"
            End If

            '使用数を計算する
            Dim intUseCount As Long = 0
            For inti As Integer = 1 To dt_detail.Columns.Count - 1
                If dt_detail.Rows(intRow)(inti).ToString.Length > 0 Then
                    intUseCount += 1
                End If
            Next
            dt_detail.Rows(intRow)("Col0") = intUseCount
            cel = Me.GridViewDetail.Rows(intRow).Cells(0)
            cel.Text = intUseCount
            Me.Session("DS_Title") = DS_Title
        Catch ex As Exception
            Call AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 列を固定にする
    ''' </summary>
    ''' <param name="intCol"></param>
    ''' <remarks></remarks>
    Private Sub SetKoteiColOnly(ByVal intCol As Integer)
        Try
            Dim cel As System.Web.UI.WebControls.TableCell
            For inti As Integer = 0 To Me.GridViewDetail.Rows.Count - 1
                If Me.GridViewDetail.Rows(inti).Cells.Count > intCol Then
                    Dim bolStart As Boolean = False
                    If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, inti, bolStart) AndAlso bolStart Then 'Mid選択欄
                        Dim GV As GridView = Me.GridViewDetail.Rows(inti).Cells(1).Controls(0)
                        For intm As Integer = 0 To GV.Rows.Count - 1
                            If GV.Rows(intm).Cells.Count > intCol Then
                                cel = GV.Rows(intm).Cells(intCol)
                                cel.Attributes.Remove("onclick")
                                cel.BackColor = Drawing.Color.FromArgb(192, 192, 192)
                            End If
                        Next
                        Continue For
                    End If
                    cel = Me.GridViewDetail.Rows(inti).Cells(intCol)
                    cel.Attributes.Remove("onclick")
                    cel.BackColor = Drawing.Color.FromArgb(192, 192, 192)
                End If
            Next
        Catch ex As Exception
            Call AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 行を固定にする
    ''' </summary>
    ''' <param name="intRow"></param>
    ''' <remarks></remarks>
    Private Sub SetKoteiRowOnly(ByVal intRow As Integer)
        Try
            Dim cel As System.Web.UI.WebControls.TableCell
            If Me.GridViewDetail.Rows.Count > intRow Then
                For inti As Integer = 1 To Me.GridViewDetail.Rows(intRow).Cells.Count - 1
                    cel = Me.GridViewDetail.Rows(intRow).Cells(inti)
                    cel.Attributes.Remove("onclick")
                    cel.BackColor = Drawing.Color.FromArgb(192, 192, 192)
                Next
            End If
        Catch ex As Exception
            Call AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 中央行を固定にする
    ''' </summary>
    ''' <param name="intCol"></param>
    ''' <remarks></remarks>
    Private Sub SetKoteiMidRow(intCol As Integer)
        Try
            Dim cel As System.Web.UI.WebControls.TableCell
            For inti As Integer = 0 To Me.GridViewDetail.Rows.Count - 1
                Dim bolStart As Boolean = False
                If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, inti, bolStart) AndAlso bolStart Then 'Mid選択欄
                    Dim GV As GridView = Me.GridViewDetail.Rows(inti).Cells(1).Controls(0)
                    For intR As Integer = 0 To GV.Rows.Count - 1
                        cel = GV.Rows(intR).Cells(intCol)
                        cel.Attributes.Remove("onclick")
                        cel.BackColor = Drawing.Color.FromArgb(192, 192, 192)
                    Next
                    Exit For
                End If
            Next
        Catch ex As Exception
            Call AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 対象列の行intStartからintEndまで、複数を選択できず（既に選択したCellをクリアする）
    ''' </summary>
    ''' <param name="dt_detail"></param>
    ''' <param name="intRowIdx"></param>
    ''' <param name="intColIdx"></param>
    ''' <param name="intNow"></param>
    ''' <param name="intStart"></param>
    ''' <param name="intEnd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetPositionGroup(ByRef dt_detail As DataTable, ByVal intRowIdx As Integer, ByVal intColIdx As Integer, _
                                 ByVal intNow As Integer, ByVal intStart As Integer, ByVal intEnd As String) As Boolean
        '中間行の場合
        If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, intNow, False) Then Return True
        'Indexオーバーの場合
        If Me.GridViewDetail.Rows(intNow).Cells.Count <= intColIdx Then Return True
        '
        If (intRowIdx >= intStart - 1 And intRowIdx <= intEnd - 1) And _
            (intNow >= intStart - 1 And intNow <= intEnd - 1) Then
            Dim cel As New System.Web.UI.WebControls.TableCell

            If objKtbnStrc.strcSelection.strSeriesKataban = "LMF0" AndAlso _
                            objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "T0D" Then   '逆
                cel = Me.GridViewDetail.Rows(intNow).Cells(11 - intColIdx)
            Else
                cel = Me.GridViewDetail.Rows(intNow).Cells(intColIdx)
            End If

            If cel.Text.Equals(strMaru) Then
                Return False
            End If
            'cel.Text = String.Empty
            'If dt_detail.Rows(intNow)(intColIdx).ToString.Length > 0 Then
            '    dt_detail.Rows(intNow)(intColIdx) = String.Empty
            '    SetUseCount(dt_detail, intNow, False)
            'End If
        End If

        Return True
    End Function

    ''' <summary>
    ''' 対象列の行intStartからintEndまで、複数を選択できず（既に選択したCellをクリアする）
    ''' </summary>
    ''' <param name="intRowIdx">処理対象行</param>
    ''' <param name="ColumnIndex">処理対象列</param>
    ''' <param name="intNow">現在行</param>
    ''' <param name="intStart">開始行</param>
    ''' <param name="intEnd">終了行</param>
    ''' <param name="intExStart">共存開始行</param>
    ''' <param name="intExEnd">共存終了行</param>
    ''' <param name="intEx1">共存行1</param>
    ''' <param name="intEx2">共存行2</param>
    ''' <remarks></remarks>
    Private Function SetPositionGroupEx(ByRef dt_detail As DataTable, ByVal intRowIdx As Integer, ByVal ColumnIndex As Integer, _
                                 ByVal intNow As Integer, ByVal intStart As Integer, ByVal intEnd As Integer, _
                                 ByVal intExStart As Integer, ByVal intExEnd As Integer, intEx1 As Integer, intEx2 As Integer) As Boolean
        Try
            Dim cel As System.Web.UI.WebControls.TableCell

            '中間行の場合
            If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, intNow, False) Then Return True
            'Indexオーバーの場合
            If Me.GridViewDetail.Rows(intNow).Cells.Count <= ColumnIndex Then Return True
            '範囲外の場合
            If ((intRowIdx < intExStart - 1) Or (intRowIdx > intExEnd - 1)) And _
                ((intRowIdx < intEx1 - 1) Or (intRowIdx > intEx2 - 1)) Then

                cel = Me.GridViewDetail.Rows(intNow).Cells(ColumnIndex)
                If cel.Text.Equals(strMaru) Then
                    Return False
                End If
            ElseIf (intRowIdx >= intEx1 - 1) And (intRowIdx <= intEx2 - 1) Then
                If (intNow < intExStart - 1 Or intNow > intExEnd - 1) Then

                    cel = Me.GridViewDetail.Rows(intNow).Cells(ColumnIndex)
                    If cel.Text.Equals(strMaru) Then
                        Return False
                    End If
                End If
            ElseIf (intRowIdx >= intExStart - 1 And intRowIdx <= intExEnd - 1) Then
                If (intNow < intEx1 - 1) Or (intNow > intEx2 - 1) Then

                    cel = Me.GridViewDetail.Rows(intNow).Cells(ColumnIndex)
                    If cel.Text.Equals(strMaru) Then
                        Return False
                    End If
                End If
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try

        Return True
    End Function

    ''' <summary>
    ''' 使用数の設定
    ''' </summary>
    ''' <param name="dt_detail"></param>
    ''' <remarks></remarks>
    Private Sub SetUseCount(ByRef dt_detail As DataTable)
        For Each dr In dt_detail.Rows
            '使用数
            Dim intUsed As Integer = 0

            For intColumn As Integer = 1 To dt_detail.Columns.Count - 1
                If dr(intColumn).ToString.Equals(strMaru) Then
                    If objKtbnStrc.strcSelection.strSpecNo.Equals("16") Then

                    End If
                    intUsed += 1
                End If
            Next
            dr("Col0") = intUsed
        Next
    End Sub

    ''' <summary>
    ''' 取付レール長さ設定
    ''' </summary>
    ''' <param name="ds"></param>
    ''' <param name="intRailRowID">レール長さ行番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetRailLength(ds As DataSet, ByVal intRailRowID As Integer) As Decimal
        If HidManifoldMode.Value <= 0 Then
            Return 0
        Else
            Dim dblStdNum As Decimal = 0
            Dim dblRailLen As Decimal = 0
            Dim dblManiLen As Decimal = 0

            Call bllSiyou.subGetRail(ds, HidManifoldMode.Value, objKtbnStrc, intRailRowID, HidRailChangeFlg.Value, dblRailLen, dblStdNum)

            HidStdNum.Value = dblStdNum

            Return dblRailLen
        End If
    End Function

    ''' <summary>
    ''' 「取付ﾚｰﾙ長さ」,「チューブ」,「タグ銘板」の行番号
    ''' </summary>
    ''' <param name="intMode">0：Rail、1：Tube、2：Tag</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetRailTubeIndex(intMode As Integer) As Integer
        GetRailTubeIndex = -1
        Dim str() As String = Me.HidOther.Value.ToString.Split(",")

        For inti As Integer = 0 To str.Length - 1
            Select Case intMode
                Case 0 '取付ﾚｰﾙ長さ
                    If str(inti) = "5" Or str(inti) = "6" Then Return inti
                Case 1 'チューブ
                    If str(inti) = "4" Then Return inti
                Case 2 'タグ銘板
                    If str(inti) = "99" Then Return inti
            End Select
        Next
    End Function

    ''' <summary>
    ''' 特殊選択肢の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetComb_Change()
        '特殊な選択肢
        Dim list As New ArrayList
        '画面選択した形番
        Dim strSelectedKataban As String = String.Empty
        If HidSelect.Value.ToString.Equals(String.Empty) Then
            strSelectedKataban = String.Empty
        Else
            strSelectedKataban = HidSelect.Value.Split(",")(0)
        End If


        '機種によりの設定
        Select Case objKtbnStrc.strcSelection.strSeriesKataban
            Case "MN3E00", "MN4E0"
                If dt_Comb(1).Count <= 0 Then
                    list.Add("")
                    dt_Comb(2) = list
                Else
                    If strSelectedKataban.Equals("N4E0-T50") Then
                        'RM1801043_電装ブロック表示修正
                        list.Add("")
                        If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                            list.Add("N4E0-T30NR")
                        Else
                            list.Add("N4E0-T30R")
                        End If
                        list.Add("N4E0-T50R")
                        list.Add("N4E0-T51R")
                        list.Add("N4E0-T52R")
                        list.Add("N4E0-T53R")
                        list.Add("N4E0-TM1A")
                        list.Add("N4E0-TM1B")
                        list.Add("N4E0-TM1C")
                        list.Add("N4E0-TM52")
                        '201503月次更新
                        Dim str As String = String.Empty
                        If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                            str = "N4E0-Q-6N-C,N4E0-Q-6LN-C,N4E0-Q-8N-C,N4E0-Q-8LN-C," & _
                                    "N4E0-Q-6N-S-C,N4E0-Q-6LN-S-C,N4E0-Q-8N-S-C,N4E0-Q-8LN-S-C," & _
                                    "N4E0-Q-6N-SA-C,N4E0-Q-6LN-SA-C,N4E0-Q-8N-SA-C,N4E0-Q-8LN-SA-C," & _
                                    "N4E0-QK-6N-C,N4E0-QK-6LN-C,N4E0-QK-8N-C,N4E0-QK-8LN-C," & _
                                    "N4E0-QK-6N-S-C,N4E0-QK-6LN-S-C,N4E0-QK-8N-S-C,N4E0-QK-8LN-S-C," & _
                                    "N4E0-QK-6N-SA-C,N4E0-QK-6LN-SA-C,N4E0-QK-8N-SA-C,N4E0-QK-8LN-SA-C," & _
                                    "N4E0-QZ-6N-C,N4E0-QZ-6LN-C,N4E0-QZ-8N-C,N4E0-QZ-8LN-C," & _
                                    "N4E0-QZ-6N-S-C,N4E0-QZ-6LN-S-C,N4E0-QZ-8N-S-C,N4E0-QZ-8LN-S-C," & _
                                    "N4E0-QZ-6N-SA-C,N4E0-QZ-6LN-SA-C,N4E0-QZ-8N-SA-C,N4E0-QZ-8LN-SA-C," & _
                                    "N4E0-QKZ-6N-C,N4E0-QKZ-6LN-C,N4E0-QKZ-8N-C,N4E0-QKZ-8LN-C," & _
                                    "N4E0-QKZ-6N-S-C,N4E0-QKZ-6LN-S-C,N4E0-QKZ-8N-S-C,N4E0-QKZ-8LN-S-C," & _
                                    "N4E0-QKZ-6N-SA-C,N4E0-QKZ-6LN-SA-C,N4E0-QKZ-8N-SA-C,N4E0-QKZ-8LN-SA-C," & _
                                    "N4E0-QX-6N-C,N4E0-QX-6LN-C,N4E0-QX-8N-C,N4E0-QX-8LN-C," & _
                                    "N4E0-QX-6N-S-C,N4E0-QX-6LN-S-C,N4E0-QX-8N-S-C,N4E0-QX-8LN-S-C," & _
                                    "N4E0-QX-6N-SA-C,N4E0-QX-6LN-SA-C,N4E0-QX-8N-SA-C,N4E0-QX-8LN-SA-C," & _
                                    "N4E0-QKX-6N-C,N4E0-QKX-6LN-C,N4E0-QKX-8N-C,N4E0-QKX-8LN-C," & _
                                    "N4E0-QKX-6N-S-C,N4E0-QKX-6LN-S-C,N4E0-QKX-8N-S-C,N4E0-QKX-8LN-S-C," & _
                                    "N4E0-QKX-6N-SA-C,N4E0-QKX-6LN-SA-C,N4E0-QKX-8N-SA-C,N4E0-QKX-8LN-SA-C"
                        Else
                            str = "N4E0-Q-6-C,N4E0-Q-6L-C,N4E0-Q-8-C,N4E0-Q-8L-C," & _
                                    "N4E0-Q-6-S-C,N4E0-Q-6L-S-C,N4E0-Q-8-S-C,N4E0-Q-8L-S-C," & _
                                    "N4E0-Q-6-SA-C,N4E0-Q-6L-SA-C,N4E0-Q-8-SA-C,N4E0-Q-8L-SA-C," & _
                                    "N4E0-QK-6-C,N4E0-QK-6L-C,N4E0-QK-8-C,N4E0-QK-8L-C," & _
                                    "N4E0-QK-6-S-C,N4E0-QK-6L-S-C,N4E0-QK-8-S-C,N4E0-QK-8L-S-C," & _
                                    "N4E0-QK-6-SA-C,N4E0-QK-6L-SA-C,N4E0-QK-8-SA-C,N4E0-QK-8L-SA-C," & _
                                    "N4E0-QZ-6-C,N4E0-QZ-6L-C,N4E0-QZ-8-C,N4E0-QZ-8L-C," & _
                                    "N4E0-QZ-6-S-C,N4E0-QZ-6L-S-C,N4E0-QZ-8-S-C,N4E0-QZ-8L-S-C," & _
                                    "N4E0-QZ-6-SA-C,N4E0-QZ-6L-SA-C,N4E0-QZ-8-SA-C,N4E0-QZ-8L-SA-C," & _
                                    "N4E0-QKZ-6-C,N4E0-QKZ-6L-C,N4E0-QKZ-8-C,N4E0-QKZ-8L-C," & _
                                    "N4E0-QKZ-6-S-C,N4E0-QKZ-6L-S-C,N4E0-QKZ-8-S-C,N4E0-QKZ-8L-S-C," & _
                                    "N4E0-QKZ-6-SA-C,N4E0-QKZ-6L-SA-C,N4E0-QKZ-8-SA-C,N4E0-QKZ-8L-SA-C," & _
                                    "N4E0-QX-6-C,N4E0-QX-6L-C,N4E0-QX-8-C,N4E0-QX-8L-C," & _
                                    "N4E0-QX-6-S-C,N4E0-QX-6L-S-C,N4E0-QX-8-S-C,N4E0-QX-8L-S-C," & _
                                    "N4E0-QX-6-SA-C,N4E0-QX-6L-SA-C,N4E0-QX-8-SA-C,N4E0-QX-8L-SA-C," & _
                                    "N4E0-QKX-6-C,N4E0-QKX-6L-C,N4E0-QKX-8-C,N4E0-QKX-8L-C," & _
                                    "N4E0-QKX-6-S-C,N4E0-QKX-6L-S-C,N4E0-QKX-8-S-C,N4E0-QKX-8L-S-C," & _
                                    "N4E0-QKX-6-SA-C,N4E0-QKX-6L-SA-C,N4E0-QKX-8-SA-C,N4E0-QKX-8L-SA-C"
                        End If

                        Dim strL() As String = str.Split(",")
                        For inti As Integer = 0 To strL.Length - 1
                            list.Add(strL(inti))
                        Next
                        Me.dt_Comb(16) = list
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol(6) = "TX" Then
                            Select Case strSelectedKataban
                                Case "N4E0-T30", "N4E0-T30N", "N4E0-T51", "N4E0-T52", "N4E0-T53", "N4E0-T5B", "N4E0-T5C"
                                    list.Add("")
                                    If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                        list.Add("N4E0-T30NR")
                                    Else
                                        list.Add("N4E0-T30R")
                                    End If
                                    list.Add("N4E0-T51R")
                                    list.Add("N4E0-T52R")
                                    list.Add("N4E0-T53R")
                                    list.Add("N4E0-TM1A")
                                    list.Add("N4E0-TM1B")
                                    list.Add("N4E0-TM1C")
                                    list.Add("N4E0-TM52")

                                    Me.dt_Comb(2) = list
                                Case "N4E0-T50R"
                                    'RM1801043_電装ブロック表示修正
                                    list.Add("")
                                    If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                        list.Add("N4E0-T30N")
                                    Else
                                        list.Add("N4E0-T30")
                                    End If
                                    list.Add("N4E0-T50")
                                    list.Add("N4E0-T51")
                                    list.Add("N4E0-T52")
                                    list.Add("N4E0-T53")
                                    list.Add("N4E0-T5B")
                                    list.Add("N4E0-T5C")
                                    list.Add("N4E0-TM1A")
                                    list.Add("N4E0-TM1B")
                                    list.Add("N4E0-TM1C")
                                    list.Add("N4E0-TM52")

                                    Me.dt_Comb(2) = list
                                Case "N4E0-T30R", "N4E0-T30NR", "N4E0-T51R", "N4E0-T52R", "N4E0-T53R"

                                    list.Add("")
                                    If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                        list.Add("N4E0-T30N")
                                    Else
                                        list.Add("N4E0-T30")
                                    End If
                                    list.Add("N4E0-T51")
                                    list.Add("N4E0-T52")
                                    list.Add("N4E0-T53")
                                    list.Add("N4E0-T5B")
                                    list.Add("N4E0-T5C")
                                    list.Add("N4E0-TM1A")
                                    list.Add("N4E0-TM1B")
                                    list.Add("N4E0-TM1C")
                                    list.Add("N4E0-TM52")

                                    Me.dt_Comb(2) = list
                                Case "N4E0-TM1A", "N4E0-TM1B", "N4E0-TM1C", "N4E0-TM52"
                                    list.Add("")
                                    If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                        list.Add("N4E0-T30N")
                                    Else
                                        list.Add("N4E0-T30")
                                    End If
                                    list.Add("N4E0-T51")
                                    list.Add("N4E0-T52")
                                    list.Add("N4E0-T53")
                                    list.Add("N4E0-T5B")
                                    list.Add("N4E0-T5C")
                                    If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                        list.Add("N4E0-T30NR")
                                    Else
                                        list.Add("N4E0-T30R")
                                    End If
                                    list.Add("N4E0-T51R")
                                    list.Add("N4E0-T52R")
                                    list.Add("N4E0-T53R")
                                    list.Add("N4E0-TM1A")
                                    list.Add("N4E0-TM1B")
                                    list.Add("N4E0-TM1C")
                                    list.Add("N4E0-TM52")

                                    Me.dt_Comb(2) = list
                                Case "N3Q0-T30", "N3Q0-T51", "N3Q0-T53", "N3Q0-T30U", "N3Q0-T51U", "N3Q0-T53U"
                                    list.Add("")
                                    list.Add("N3Q0-T30R")
                                    list.Add("N3Q0-T51R")
                                    list.Add("N3Q0-T53R")
                                    list.Add("N3Q0-T30UR")
                                    list.Add("N3Q0-T51UR")
                                    list.Add("N3Q0-T53UR")

                                    Me.dt_Comb(2) = list
                                Case "N3Q0-T30R", "N3Q0-T51R", "N3Q0-T53R", "N3Q0-T30UR", "N3Q0-T51UR", "N3Q0-T53UR"
                                    list.Add("")
                                    list.Add("N3Q0-T30")
                                    list.Add("N3Q0-T51")
                                    list.Add("N3Q0-T53")
                                    list.Add("N3Q0-T30U")
                                    list.Add("N3Q0-T51U")
                                    list.Add("N3Q0-T53U")

                                    Me.dt_Comb(2) = list
                                Case Else
                                    Me.dt_Comb(2) = list
                            End Select
                        End If
                    End If
                End If
            Case "MN3E0", "MN3EX0", "MN3Q0", "MN4E00", "MN4EX0", "MT3Q0"
                If dt_Comb(1).Count <= 0 Then
                    list.Add("")
                    dt_Comb(2) = list
                Else
                    If objKtbnStrc.strcSelection.strOpSymbol(4) = "TX" OrElse objKtbnStrc.strcSelection.strOpSymbol(6) = "TX" Then

                        Select Case strSelectedKataban
                            Case "N4E0-T30", "N4E0-T30N", "N4E0-T51", "N4E0-T52", "N4E0-T53", "N4E0-T5B", "N4E0-T5C"
                                list.Add("")
                                If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                    list.Add("N4E0-T30NR")
                                Else
                                    list.Add("N4E0-T30R")
                                End If
                                list.Add("N4E0-T51R")
                                list.Add("N4E0-T52R")
                                list.Add("N4E0-T53R")
                                list.Add("N4E0-TM1A")
                                list.Add("N4E0-TM1B")
                                list.Add("N4E0-TM1C")
                                list.Add("N4E0-TM52")

                                Me.dt_Comb(2) = list
                            Case "N4E0-T50"
                                'RM1801043_電装ブロック表示修正
                                list.Add("")
                                If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                    list.Add("N4E0-T30NR")
                                Else
                                    list.Add("N4E0-T30R")
                                End If
                                list.Add("N4E0-T50R")
                                list.Add("N4E0-T51R")
                                list.Add("N4E0-T52R")
                                list.Add("N4E0-T53R")
                                list.Add("N4E0-TM1A")
                                list.Add("N4E0-TM1B")
                                list.Add("N4E0-TM1C")
                                list.Add("N4E0-TM52")

                                Me.dt_Comb(2) = list
                            Case "N4E0-T50R"
                                'RM1801043_電装ブロック表示修正
                                list.Add("")
                                If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                    list.Add("N4E0-T30N")
                                Else
                                    list.Add("N4E0-T30")
                                End If
                                list.Add("N4E0-T50")
                                list.Add("N4E0-T51")
                                list.Add("N4E0-T52")
                                list.Add("N4E0-T53")
                                list.Add("N4E0-T5B")
                                list.Add("N4E0-T5C")
                                list.Add("N4E0-TM1A")
                                list.Add("N4E0-TM1B")
                                list.Add("N4E0-TM1C")
                                list.Add("N4E0-TM52")

                                Me.dt_Comb(2) = list
                            Case "N4E0-T30R", "N4E0-T30NR", "N4E0-T51R", "N4E0-T52R", "N4E0-T53R"

                                list.Add("")
                                If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                    list.Add("N4E0-T30N")
                                Else
                                    list.Add("N4E0-T30")
                                End If
                                list.Add("N4E0-T51")
                                list.Add("N4E0-T52")
                                list.Add("N4E0-T53")
                                list.Add("N4E0-T5B")
                                list.Add("N4E0-T5C")
                                list.Add("N4E0-TM1A")
                                list.Add("N4E0-TM1B")
                                list.Add("N4E0-TM1C")
                                list.Add("N4E0-TM52")

                                Me.dt_Comb(2) = list
                            Case "N4E0-TM1A", "N4E0-TM1B", "N4E0-TM1C", "N4E0-TM52"
                                list.Add("")
                                If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                    list.Add("N4E0-T30N")
                                Else
                                    list.Add("N4E0-T30")
                                End If
                                list.Add("N4E0-T51")
                                list.Add("N4E0-T52")
                                list.Add("N4E0-T53")
                                list.Add("N4E0-T5B")
                                list.Add("N4E0-T5C")
                                If objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                                    list.Add("N4E0-T30NR")
                                Else
                                    list.Add("N4E0-T30R")
                                End If
                                list.Add("N4E0-T51R")
                                list.Add("N4E0-T52R")
                                list.Add("N4E0-T53R")
                                list.Add("N4E0-TM1A")
                                list.Add("N4E0-TM1B")
                                list.Add("N4E0-TM1C")
                                list.Add("N4E0-TM52")

                                Me.dt_Comb(2) = list
                            Case "N3Q0-T30", "N3Q0-T51", "N3Q0-T53", "N3Q0-T30U", "N3Q0-T51U", "N3Q0-T53U"
                                list.Add("")
                                list.Add("N3Q0-T30R")
                                list.Add("N3Q0-T51R")
                                list.Add("N3Q0-T53R")
                                list.Add("N3Q0-T30UR")
                                list.Add("N3Q0-T51UR")
                                list.Add("N3Q0-T53UR")

                                Me.dt_Comb(2) = list
                            Case "N3Q0-T30R", "N3Q0-T51R", "N3Q0-T53R", "N3Q0-T30UR", "N3Q0-T51UR", "N3Q0-T53UR"
                                list.Add("")
                                list.Add("N3Q0-T30")
                                list.Add("N3Q0-T51")
                                list.Add("N3Q0-T53")
                                list.Add("N3Q0-T30U")
                                list.Add("N3Q0-T51U")
                                list.Add("N3Q0-T53U")

                                Me.dt_Comb(2) = list
                            Case Else
                                Me.dt_Comb(2) = list
                        End Select
                    End If
                End If
            Case "LMF0"
                If objKtbnStrc.strcSelection.strOpSymbol(2) = "XX" Then
                    'Ａ・Ｂポート接続口径を選択不可にする
                    Dim arr As New ArrayList
                    arr.Add("")
                    For intI As Integer = CdCst.Siyou_06.ABCon01 To CdCst.Siyou_06.ABCon1Z
                        dt_Comb(intI) = arr
                    Next
                End If
        End Select

        'マニホールド番号によりの設定
        Call subSetSpecContaintData()
    End Sub

    ''' <summary>
    ''' Manifold6の場合はその他電圧を選択した時にオプションの変更
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetSpecContaintData()
        '
        If objKtbnStrc.strcSelection.strSpecNo.Equals("06") Then
            Dim strOptions() As String = objKtbnStrc.strcSelection.strOpSymbol

            If strOptions(5) = "9" Then
                For inti As Integer = 1 To 5
                    Dim strSelections As ArrayList = CType(dt_Comb(inti), ArrayList)

                    For intSelections As Integer = 0 To strSelections.Count - 1
                        If strSelections(intSelections).ToString.EndsWith("-" & strOptions(7)) Then
                            If strOptions.Length > 8 Then
                                dt_Comb(inti)(intSelections) = strSelections(intSelections).ToString & "-" & strOptions(8)
                            End If
                        End If
                    Next
                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' 隠しエリアからデータの設定
    ''' </summary>
    ''' <param name="dt_title"></param>
    ''' <remarks></remarks>
    Private Sub SetDataFromHid(ByRef dt_title As DataTable)
        Dim strKataSel() As String = Me.HidSelect.Value.ToString.Split(",")
        If strKataSel.Length > 0 AndAlso strKataSel.Length >= dt_title.Rows.Count Then
            For inti As Integer = 0 To dt_title.Rows.Count - 1
                dt_title.Rows(inti)("ColKata") = strKataSel(inti).ToString
            Next
        End If
    End Sub

    ''' <summary>
    ''' 選択した仕様情報をDataTableに登録
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SaveInfoToDatatable(ByRef dt_title As DataTable, ByRef dt_detail As DataTable)
        If Not HidSelect.Value.Equals(String.Empty) AndAlso Not HidClick.Value.Equals(String.Empty) Then
            '画面に保存された形番情報
            Dim strKatabans() As String = HidSelect.Value.Trim.Split(",")
            '画面に保存された使用数情報
            Dim strUse() As String = HidUse.Value.Trim.Split(",")

            '画面に保存された選択情報
            Dim strPositions() As String = (From strp In HidClick.Value.Split(";")
                                            Select strp).Distinct.ToArray
            '形番情報の設定
            For intRow As Integer = 0 To dt_title.Rows.Count - 1
                Dim dr As DataRow = dt_title.Rows(intRow)

                If intRow <= strKatabans.Count - 1 Then
                    dr("ColKata") = strKatabans(intRow)
                End If

                '使用数の設定
                dt_detail.Rows(intRow)("Col0") = strUse(intRow)
            Next

            '仕様情報の設定
            For Each strPosition In strPositions
                If Not strPosition.Equals(String.Empty) Then
                    Dim intRowIdx As Integer = 0
                    Dim intColIdx As Integer = 0
                    Dim intMidRowIdx As Integer = 0
                    Dim strX As String
                    Dim strY As String
                    Dim strMidRow As String = String.Empty

                    '選択された各座標
                    If strPosition.StartsWith("M") Then
                        'Mid行の場合
                        strMidRow = strPosition.Split(",")(0).Replace("M", String.Empty)
                        strX = strPosition.Split(",")(1)
                        strY = strPosition.Split(",")(2)
                        If Integer.TryParse(strX, intRowIdx) AndAlso _
                            Integer.TryParse(strY, intColIdx) AndAlso _
                            Integer.TryParse(strMidRow, intMidRowIdx) Then

                            Dim intUsed As Integer = 0
                            '座標の調整
                            intRowIdx = intRowIdx + intMidRowIdx - 2
                            '仕様の設定
                            dt_detail.Rows(intRowIdx)(intColIdx) = strMaru                  '●設置

                            ''使用数の取得
                            'Dim strUsed As String = IIf(IsDBNull(dt_detail.Rows(intRowIdx)("Col0")), "0", dt_detail.Rows(intRowIdx)("Col0"))
                            ''使用数の設定
                            'If Integer.TryParse(strUsed, intUsed) Then
                            '    If strPosition.StartsWith("M18") AndAlso _
                            '        objKtbnStrc.strcSelection.strSpecNo.Equals("16") AndAlso _
                            '        intRowIdx = 17 Then
                            '        'Mid第18行単位2で増加する
                            '        dt_detail.Rows(intRowIdx)("Col0") = intUsed + 2
                            '    Else
                            '        dt_detail.Rows(intRowIdx)("Col0") = intUsed + 1
                            '    End If
                            'End If
                        End If
                    Else
                        '正常の場合
                        strX = strPosition.Split(",")(0)
                        strY = strPosition.Split(",")(1)
                        If Integer.TryParse(strX, intRowIdx) AndAlso _
                            Integer.TryParse(strY, intColIdx) Then

                            Dim intUsed As Integer = 0
                            '座標の調整
                            intRowIdx -= 2
                            '仕様の設定
                            dt_detail.Rows(intRowIdx)(intColIdx) = strMaru                  '●設置

                            ''使用数の取得
                            'Dim strUsed As String = IIf(IsDBNull(dt_detail.Rows(intRowIdx)("Col0")), "0", dt_detail.Rows(intRowIdx)("Col0"))
                            ''仕様数の設定
                            'If Integer.TryParse(strUsed, intUsed) Then
                            '    dt_detail.Rows(intRowIdx)("Col0") = intUsed + 1
                            'End If
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 入力データのチェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckInput(ByVal dt_title As DataTable, ByRef dt_detail As DataTable, ByVal strPositions() As String) As Boolean

        For Each strPosition In strPositions
            If Not strPosition.Equals(String.Empty) Then
                '重複選択チェック　　↓↓↓↓↓↓
                '選択された各座標
                Dim intRowIdx As Integer = 0
                Dim intColIdx As Integer = 0
                Dim intMidRow As Integer = 0                            'Mid行
                Dim strRow As String = String.Empty
                Dim strColumn As String = String.Empty
                Dim strMidRow As String = String.Empty

                If strPosition.StartsWith("M") Then
                    'Mid行の場合
                    strRow = strPosition.Split(",")(0).Replace("M", String.Empty)
                    strMidRow = strPosition.Split(",")(1)
                    strColumn = strPosition.Split(",")(2)
                Else
                    '普通の場合
                    strRow = strPosition.Split(",")(0)
                    strColumn = strPosition.Split(",")(1)
                End If

                If Integer.TryParse(strRow, intRowIdx) AndAlso _
                    Integer.TryParse(strColumn, intColIdx) Then

                    '座標の調整
                    intRowIdx -= 2

                    '選択重複列があるかどうかを判断
                    For inti As Integer = 0 To GridViewDetail.Rows.Count - 1

                        If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, inti, False) Then Exit For

                        If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, intRowIdx, False) Then Exit For

                        '重複あるかどうかの判断
                        If Not CheckDoubleSelect(inti, intRowIdx, intColIdx, dt_detail) Then
                            Call AlertMessage("W1390") '複数選択できない
                            Return False
                        End If
                    Next
                End If

                If strRow.Equals(String.Empty) Then
                    'Mid行以外の場合
                    '形番なし選択のチェック
                    'ISOのA・Bﾎﾟｰﾄ接続口径、形番無くても位置を指定できる
                    If Not Me.GridViewDetail.Rows(intRowIdx).Cells(intColIdx).BackColor.Equals(Drawing.Color.FromArgb(192, 192, 192)) Then
                        '選択可能な場合、形番選択されないならエラー
                        If (HidManifoldMode.Value = 5 AndAlso intRowIdx >= CdCst.Siyou_05.ABCon02 - 1 AndAlso intRowIdx <= CdCst.Siyou_05.ABCon04 - 1) OrElse _
                            (HidManifoldMode.Value = 6 AndAlso intRowIdx >= CdCst.Siyou_06.ABCon01 - 1 AndAlso intRowIdx <= CdCst.Siyou_06.ABCon1Z - 1 AndAlso objKtbnStrc.strcSelection.strOpSymbol(2).ToString = "XX") Then

                        ElseIf HidManifoldMode.Value = 5 AndAlso objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "L" AndAlso (intRowIdx = 7 OrElse intRowIdx = 6) Then

                        Else
                            '選択チェック
                            If dt_title.Rows(intRowIdx)("ColKata").ToString.Trim.Length <= 0 AndAlso _
                                dt_detail.Rows(intRowIdx)(intColIdx).ToString.Length > 0 Then
                                'dt_detail.Rows(intRowIdx)(intColIdx).ToString.Length <= 0 Then      LMF010-01Z-02U-T0U-9-N-M0-AC220V
                                Call AlertMessage("W1400")
                                Return False
                            End If
                        End If
                    End If
                End If
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' メインタイトル
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GridViewTitle_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridViewTitle.RowDataBound
        Try
            If e.Row.RowIndex < 0 Then
                Exit Sub
            End If

            If objKtbnStrc.strcSelection.strSpecNo.Equals("05") Then

                Dim dtGridTitle As DataTable = CType(GridViewTitle.DataSource, DataTable)

                If e.Row.RowIndex = dtGridTitle.Rows.Count - 1 Then
                    'ISOバルブの接続ブロックの行を非表示
                    e.Row.Visible = False
                    Exit Sub
                End If

            End If

            If HidColMerge.Value.ToString.Length > 0 Then
                '複雑版マニホールド
                Dim strMerge() As String = HidColMerge.Value.ToString.Split(",")

                Select Case strMerge(e.Row.RowIndex)
                    Case 1
                        e.Row.Cells(0).Style.Add("padding-left", "3px")
                        Exit Select
                    Case 0
                        e.Row.Cells.RemoveAt(0)
                    Case Else
                        e.Row.Cells(0).RowSpan = CInt(strMerge(e.Row.RowIndex))
                        e.Row.Cells(0).Style.Add("padding-left", "3px")
                End Select

                Dim str_itemdiv() As String = Me.HidOther.Value.ToString.Split(",")
                If e.Row.RowIndex < str_itemdiv.Length Then
                    If str_itemdiv(e.Row.RowIndex) = "4" Then 'チューブ
                        Dim chk As New CheckBox
                        chk.ID = "chk1"
                        chk.Width = WebControls.Unit.Percentage(95)
                        chk.BorderStyle = BorderStyle.None
                        chk.BorderWidth = WebControls.Unit.Pixel(0)
                        chk.Font.Name = GetFontName(selLang.SelectedValue)
                        chk.Font.Bold = True
                        chk.Font.Size = WebControls.FontUnit.Point(11)
                        chk.AutoPostBack = False
                        chk.Height = WebControls.Unit.Pixel(CdCst.MonifoldGrid.intGridWidth)
                        Select Case selLang.SelectedValue
                            Case "ja"
                                chk.Text = CdCst.Manifold.UnNecessity.Japanese
                            Case Else
                                chk.Text = CdCst.Manifold.UnNecessity.English
                        End Select
                        e.Row.Cells(1).HorizontalAlign = HorizontalAlign.Center
                        If Me.HidTube.Value = "0" Then
                            chk.Checked = True
                        Else
                            chk.Checked = False
                        End If
                        e.Row.Cells(1).Controls.Add(chk)
                        chk.Attributes.Add("onclick", "TubeChecked('" & chk.ClientID & "','" & Me.ClientID & "_');")

                        'チューブ抜具使用不可
                        Select Case HidManifoldMode.Value
                            Case "4" '２４行目変更不可
                                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                    Case "M3GA1", "M3GB1", "M4GA1", "M4GB1", "M3GD1", "M3GE1", "M4GD1", "M4GE1"
                                    Case Else
                                        e.Row.Cells(1).Enabled = False
                                End Select
                            Case "18"
                                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                    Case "MN3GD1", "MN3GE1", "MN4GD1", "MN4GE1"
                                    Case Else
                                        e.Row.Cells(1).Enabled = False
                                End Select
                            Case Else
                                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                    Case "MN4GB2"
                                        e.Row.Cells(1).Enabled = False
                                End Select
                        End Select
                        Exit Sub
                    End If
                End If

                Dim drp As New DropDownList
                drp.ID = "cmbkata"
                drp.Width = WebControls.Unit.Percentage(100)
                drp.BorderStyle = BorderStyle.None
                drp.BorderWidth = WebControls.Unit.Pixel(0)
                drp.Font.Name = GetFontName(selLang.SelectedValue)
                drp.Font.Bold = True
                drp.Font.Size = WebControls.FontUnit.Point(11)
                drp.AutoPostBack = False
                drp.Height = WebControls.Unit.Pixel(CdCst.MonifoldGrid.intGridWidth)
                drp.ViewStateMode = UI.ViewStateMode.Enabled

                If dt_Comb.Count > e.Row.RowIndex + 1 Then
                    If str_itemdiv(e.Row.RowIndex) = "3" Then '検査成績書を変換する
                        Dim arr As ArrayList = CType(dt_Comb(e.Row.RowIndex + 1), ArrayList)
                        Select Case selLang.SelectedValue
                            Case "ja"
                            Case Else
                                For inti As Integer = 0 To arr.Count - 1
                                    If arr(inti).ToString.Contains(CdCst.Manifold.InspReportJp.Japanese) Then
                                        arr(inti) = CdCst.Manifold.InspReportJp.English
                                    End If
                                    If arr(inti).ToString.Contains(CdCst.Manifold.InspReportEn.Japanese) Then
                                        arr(inti) = CdCst.Manifold.InspReportEn.English
                                    End If
                                Next
                        End Select
                    End If

                    drp.DataSource = dt_Comb(e.Row.RowIndex + 1)
                    drp.DataBind()
                    If CType(dt_Comb(e.Row.RowIndex + 1), ArrayList).Count <= 0 Then
                        drp.Style.Add("background-color", "#CCFFCC")
                        drp.Enabled = False
                    Else
                        drp.Style.Add("background-color", "#FFFFCC")
                    End If
                End If

                If e.Row.Cells.Count > 1 Then
                    e.Row.Cells(1).HorizontalAlign = HorizontalAlign.Center
                    e.Row.Cells(1).Controls.Add(drp)
                Else
                    e.Row.Cells(0).HorizontalAlign = HorizontalAlign.Center
                    e.Row.Cells(0).Controls.Add(drp)
                End If

                Select Case str_itemdiv(e.Row.RowIndex)
                    Case "99", "5", "6"     'タブ銘板/取付レール長さ
                        Dim intIndex As Integer = CInt(Strings.Right(Me.GridViewTitle.Rows(0).ClientID.ToString, 2)) + e.Row.RowIndex
                        Dim strTxtID As String = Strings.Left(Me.GridViewTitle.Rows(0).ClientID, Me.GridViewTitle.Rows(0).ClientID.Length - 2)
                        strTxtID = strTxtID & intIndex.ToString.PadLeft(2, "0")
                        strTxtID = strTxtID.Replace("GridViewTitle", "GridViewDetail") & "_txtNum"
                        'イベントの設定
                        Select Case str_itemdiv(e.Row.RowIndex)
                            Case "5", "6" '取付レール長さ
                                drp.Attributes.Add("onChange", "GridViewCellSelect('" & drp.ClientID & "','" & strTxtID & "','" & str_itemdiv(e.Row.RowIndex) & "','" & HidRailChangeFlg.ClientID & "');")
                            Case "99"
                                drp.Attributes.Add("onChange", "GridViewCellSelect('" & drp.ClientID & "','" & strTxtID & "','" & str_itemdiv(e.Row.RowIndex) & "','" & String.Empty & "');")
                        End Select
                        '選択したデータの設定
                        Dim strRailSelect As String = IIf(IsDBNull(e.Row.DataItem(1)), String.Empty, e.Row.DataItem(1))
                        drp.SelectedValue = strRailSelect
                        If HidPostBack.Value.Equals("1") Then
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "SetRailData", "GridViewCellSelect('" & drp.ClientID & "','" & strTxtID & "','" & str_itemdiv(e.Row.RowIndex) & "','" & HidRailChangeFlg.ClientID & "');", True)
                        End If
                End Select

                Select Case HidManifoldMode.Value 'CXA,CXB
                    Case "3", "4"
                        Dim strKey1 As String = String.Empty
                        Dim strStartID As String = String.Empty
                        Select Case e.Row.RowIndex
                            Case 0
                                strStartID = drp.ClientID.ToString.Replace("_cmbkata", "")
                            Case Else
                                strStartID = GridViewTitle.Rows(0).ClientID.ToString
                        End Select
                        Select Case HidManifoldMode.Value 'CXA,CXB
                            Case "3"
                                strKey1 = objKtbnStrc.strcSelection.strOpSymbol(6).ToString
                                If strKey1 = "CX" Or strKey1 = "CXF" Then
                                    If e.Row.RowIndex >= CdCst.Siyou_03.Elect1 - 1 And e.Row.RowIndex <= CdCst.Siyou_03.Elect14 - 1 Then
                                        drp.Attributes.Add("onChange", "GridViewCellSelCX('" & drp.ClientID & "','" & strParent & Me.ID & "','" & e.Row.RowIndex & "','" & strStartID & "');")
                                    End If
                                End If
                            Case "4"
                                Dim arrCmb As New ArrayList
                                If e.Row.RowIndex >= CdCst.Siyou_04.Valve1 - 1 And e.Row.RowIndex <= CdCst.Siyou_04.Spacer4 - 1 Then
                                    If dt_Comb.Count > e.Row.RowIndex + 1 Then
                                        arrCmb = CType(dt_Comb(e.Row.RowIndex + 1), ArrayList)
                                        For inti As Integer = 0 To arrCmb.Count - 1
                                            If arrCmb(inti).Contains("-CX") Then
                                                strKey1 = "CX"
                                                Exit For
                                            End If
                                        Next
                                    End If
                                    If strKey1 = "CX" Then
                                        drp.Attributes.Add("onChange", "GridViewCellSelCX('" & drp.ClientID & "','" & strParent & Me.ID & "','" & e.Row.RowIndex & "','" & strStartID & "');")
                                    End If
                                End If
                        End Select
                End Select
                '電装ブロック選択変更の時選択肢を更新
                If objKtbnStrc.strcSelection.strSpecNo.Equals("01") AndAlso e.Row.RowIndex = 0 Then
                    drp.Attributes.Add("onChange", "GridViewCellSelBlock('" & strParent & Me.ID & "')")
                End If
            Else
                '簡易マニホールド
                If Not e.Row.DataItem Is Nothing Then
                    Dim lbl As New Label
                    lbl.ID = "lblkata"
                    lbl.Width = WebControls.Unit.Percentage(100)
                    lbl.BorderStyle = BorderStyle.None
                    lbl.BorderWidth = WebControls.Unit.Pixel(0)
                    lbl.Font.Name = GetFontName(selLang.SelectedValue)
                    lbl.Font.Bold = True
                    lbl.Font.Size = WebControls.FontUnit.Point(12)

                    'lbl.Height = WebControls.Unit.Pixel(CdCst.MonifoldGrid.intGridWidth)
                    lbl.ViewStateMode = UI.ViewStateMode.Enabled

                    lbl.Text = e.Row.DataItem("ColKata")
                    'lbl.Enabled = False
                    e.Row.Cells(1).HorizontalAlign = HorizontalAlign.Left
                    e.Row.Cells(1).Controls.Add(lbl)
                End If
            End If

        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' メインデータ
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GridViewDetail_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridViewDetail.RowDataBound
        If e.Row.RowIndex < 0 Then Exit Sub

        If objKtbnStrc.strcSelection.strSpecNo.Equals("05") Then
            Dim dtGridTitle As DataTable = CType(GridViewTitle.DataSource, DataTable)

            If e.Row.RowIndex = dtGridTitle.Rows.Count - 1 Then
                'ISOバルブの接続ブロックの行を非表示
                e.Row.Visible = False
                Exit Sub
            End If
        End If

        Try
            Dim str_itemdiv() As String = Me.HidOther.Value.ToString.Split(",")
            Dim str() As String = e.Row.ClientID.ToString.Split("_")
            Dim strID As String = str(str.Length - 1).Replace("ctl", "")
            Dim strKataStart As String = Me.GridViewTitle.Rows(0).ClientID
            If e.Row.RowIndex < str_itemdiv.Length Then
                Select Case str_itemdiv(e.Row.RowIndex)
                    Case "4"      'ﾁｭｰﾌﾞ抜具
                        For inti As Integer = e.Row.Cells.Count - 1 To 0 Step -1
                            e.Row.Cells.RemoveAt(inti)
                        Next
                        Exit Sub
                    Case "5", "6", "99", "3" '取付ﾚｰﾙ長さ
                        Dim dt_detail As New DataTable
                        If Not DS_Title.Tables("data") Is Nothing Then dt_detail = DS_Title.Tables("data")

                        Dim strErrorMessage As String
                        Dim txt As New TextBox
                        Dim drpCXA As New DropDownList
                        Dim drpCXB As New DropDownList

                        strErrorMessage = ClsCommon.fncGetMsg(selLang.SelectedValue, "W1002")

                        'テキストボックスの作成    ↓↓↓↓↓↓
                        txt = CreateTextBox(str_itemdiv, e.Row.RowIndex)
                        Select Case str_itemdiv(e.Row.RowIndex)
                            Case "5", "6"
                                '取付レール長さ
                                Dim strRailSelect As String = String.Empty
                                Select Case HidManifoldMode.Value
                                    Case "3", "4"
                                        strRailSelect = IIf(IsDBNull(e.Row.DataItem(2)), String.Empty, e.Row.DataItem(2))
                                    Case Else
                                        strRailSelect = IIf(IsDBNull(e.Row.DataItem(0)), String.Empty, e.Row.DataItem(0))
                                End Select

                                txt.Text = strRailSelect
                        End Select
                        'テキストボックスの作成    ↓↓↓↓↓↓
                        Select Case HidManifoldMode.Value
                            Case "3", "4"
                                drpCXA = CreateCXDropDownList("cmbCXA")
                                drpCXB = CreateCXDropDownList("cmbCXB")
                                e.Row.Cells(0).HorizontalAlign = HorizontalAlign.Center
                                e.Row.Cells(1).HorizontalAlign = HorizontalAlign.Center
                                e.Row.Cells(2).HorizontalAlign = HorizontalAlign.Center
                                e.Row.Cells(0).Controls.Add(drpCXA)
                                e.Row.Cells(1).Controls.Add(drpCXB)
                                e.Row.Cells(2).Controls.Add(txt)
                            Case Else
                                e.Row.Cells(0).HorizontalAlign = HorizontalAlign.Center
                                e.Row.Cells(0).Controls.Add(txt)
                        End Select

                        Select Case str_itemdiv(e.Row.RowIndex)
                            Case "5"
                                For inti As Integer = e.Row.Cells.Count - 1 To 2 Step -1
                                    e.Row.Cells.RemoveAt(inti)
                                Next
                                e.Row.Cells(1).Style.Add("background-color", "#C7EDCC")
                                If Me.HidManifoldMode.Value = 14 Then
                                    e.Row.Cells(1).BorderColor = Drawing.Color.Black
                                    e.Row.Cells(1).BorderStyle = BorderStyle.Solid
                                    e.Row.Cells(1).BorderWidth = WebControls.Unit.Pixel(1)
                                Else
                                    e.Row.Cells(1).BorderColor = Drawing.Color.FromArgb(202, 255, 202)
                                    e.Row.Cells(1).BorderStyle = BorderStyle.None
                                    e.Row.Cells(1).BorderWidth = WebControls.Unit.Pixel(0)
                                End If
                                e.Row.Cells(1).HorizontalAlign = HorizontalAlign.Left
                                e.Row.Cells(1).Style.Add("padding-left", "3px")
                                e.Row.Cells(1).ColumnSpan = 3
                            Case Else
                                Dim intEndIndex As Integer = 0
                                Select Case HidManifoldMode.Value
                                    Case "3", "4"
                                        intEndIndex = 3
                                    Case Else
                                        intEndIndex = 1
                                End Select
                                '不要なセルを削除
                                For inti As Integer = e.Row.Cells.Count - 1 To intEndIndex Step -1
                                    e.Row.Cells.RemoveAt(inti)
                                Next
                        End Select

                        'レール長さ更新ボタンの追加
                        Select Case str_itemdiv(e.Row.RowIndex)
                            Case "5", "6"
                                Dim btnCell As New TableCell
                                Dim btnRailUpdate As New Button

                                btnRailUpdate = CreateRailUpdateBtn()

                                btnCell.Controls.Add(btnRailUpdate)
                                btnCell.ColumnSpan = 4
                                btnCell.BorderWidth = 0
                                e.Row.Cells.Add(btnCell)
                        End Select

                        Exit Sub
                End Select
            End If

            Dim bolFirstOne As Boolean = False
            If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, e.Row.RowIndex, bolFirstOne) Then
                'Mid行の場合
                If Not bolFirstOne Then
                    For inti As Integer = e.Row.Cells.Count - 1 To 1 Step -1
                        e.Row.Cells.RemoveAt(inti)
                    Next
                Else
                    Dim intRowWidth As Integer = CdCst.MonifoldGrid.intGridWidth * e.Row.Cells.Count

                    'Mid行にあるCellsを削除
                    e.Row.Cells(1).HorizontalAlign = HorizontalAlign.Left
                    e.Row.Cells(1).ColumnSpan = e.Row.Cells.Count - 1
                    e.Row.Cells(1).RowSpan = 2
                    For inti As Integer = e.Row.Cells.Count - 1 To 2 Step -1
                        e.Row.Cells.RemoveAt(inti)
                    Next

                    'Mid行を構成する
                    Dim GridMid As New WebControls.GridView
                    GridMid.ID = "MidGridView"
                    GridMid.AutoGenerateColumns = False
                    GridMid.ShowHeader = False
                    GridMid.GridLines = GridLines.Both
                    GridMid.Font.Size = WebControls.FontUnit.Point(12)
                    GridMid.Font.Name = GetFontName(selLang.SelectedValue)
                    GridMid.CellPadding = 0
                    GridMid.CellSpacing = 0
                    GridMid.Width = WebControls.Unit.Percentage(100)

                    e.Row.Cells(1).Controls.Add(GridMid)
                    'ADD BY YGY 20141110
                    e.Row.Cells(1).Width = WebControls.Unit.Pixel(intRowWidth)

                    Dim col As New Web.UI.WebControls.BoundField
                    For inti As Integer = 0 To CInt(HidColCount.Value)
                        col = New Web.UI.WebControls.BoundField
                        col.DataField = "Col" & inti
                        col.HeaderText = inti.ToString.PadLeft(2, "0")
                        col.ItemStyle.Wrap = False
                        col.ItemStyle.Height = WebControls.Unit.Pixel(CdCst.MonifoldGrid.intGridWidth - 1)
                        col.ItemStyle.HorizontalAlign = HorizontalAlign.Center
                        If inti = 0 Then
                            'col.ItemStyle.Width = WebControls.Unit.Pixel(CdCst.MonifoldGrid.intGridWidth / 2 - 1)
                            col.ItemStyle.Width = WebControls.Unit.Percentage(100 / (HidColCount.Value * 2))
                            col.ItemStyle.BackColor = Drawing.Color.FromArgb(202, 255, 202)
                        ElseIf inti = CInt(HidColCount.Value) Then
                            col.ItemStyle.BackColor = Drawing.Color.FromArgb(202, 255, 202)
                        Else
                            col.ItemStyle.Width = WebControls.Unit.Percentage(100 / HidColCount.Value)
                            col.ItemStyle.BackColor = Drawing.Color.FromArgb(255, 255, 192)
                        End If
                        GridMid.Columns.Add(col)
                    Next

                    Dim dt As New DataTable
                    Dim dc As New DataColumn
                    For inti As Integer = 0 To CInt(HidColCount.Value)
                        dc = New DataColumn("Col" & inti)
                        dt.Columns.Add(dc)
                    Next
                    Dim dr As DataRow = Nothing
                    For inti As Integer = 0 To 1
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                    Next

                    'Mid行の位置情報を取得する
                    Dim dt_detail As New DataTable
                    If Not DS_Title.Tables("data") Is Nothing Then dt_detail = DS_Title.Tables("data")
                    For inti As Integer = e.Row.RowIndex To e.Row.RowIndex + 1
                        If objKtbnStrc.strcSelection.strSeriesKataban = "LMF0" AndAlso _
                            objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "T0D" Then   '逆
                            Dim intLoop As Integer = 1
                            For intj As Integer = CInt(HidColCount.Value) - 1 To 1 Step -1
                                dt.Rows(inti - e.Row.RowIndex)("col" & intLoop) = dt_detail(inti)("col" & intj)
                                intLoop += 1
                            Next
                        Else
                            For intj As Integer = 1 To CInt(HidColCount.Value) - 1
                                dt.Rows(inti - e.Row.RowIndex)("col" & intj) = dt_detail(inti)("col" & intj)
                            Next
                        End If
                    Next
                    GridMid.DataSource = dt
                    GridMid.DataBind()

                    '形番RowIDを取得する
                    If objKtbnStrc.strcSelection.strSeriesKataban = "LMF0" AndAlso _
                        objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "T0D" Then   '逆
                        For inti As Integer = 0 To GridMid.Rows.Count - 1
                            Dim intLoop As Integer = 1
                            For intj As Integer = GridMid.Rows(inti).Cells.Count - 2 To 1 Step -1
                                GridMid.Rows(inti).Cells(intLoop).Attributes.Add("onclick", "GridViewCellClickMid('" & strParent & Me.ID & "','" & GridMid.ClientID & "','" & CInt(strID) & "','" & _
                                                        inti & "','" & intj & "','" & strKataStart & "', '1', '0');")
                                intLoop += 1
                            Next
                        Next
                    Else
                        For inti As Integer = 0 To GridMid.Rows.Count - 1
                            For intj As Integer = 1 To GridMid.Rows(inti).Cells.Count - 2
                                'マニホールド16の18行目の使用数が2単位で増加すること
                                If objKtbnStrc.strcSelection.strSpecNo.Equals("16") AndAlso inti = 1 Then
                                    GridMid.Rows(inti).Cells(intj).Attributes.Add("onclick", "GridViewCellClickMid('" & strParent & Me.ID & "','" & GridMid.ClientID & "','" & CInt(strID) & "','" & _
                                                        inti & "','" & intj & "','" & strKataStart & "','0', '1');")
                                Else
                                    GridMid.Rows(inti).Cells(intj).Attributes.Add("onclick", "GridViewCellClickMid('" & strParent & Me.ID & "','" & GridMid.ClientID & "','" & CInt(strID) & "','" & _
                                                        inti & "','" & intj & "','" & strKataStart & "','0', '0');")
                                End If
                            Next
                        Next
                    End If
                End If
            Else
                '普通行の場合
                Dim intStart As Integer = 1
                Select Case HidManifoldMode.Value 'CXA,CXB
                    Case "17" 'GAMD0 Base
                        If e.Row.RowIndex = 5 Then
                            For inti As Integer = e.Row.Cells.Count - 1 To 0 Step -1
                                e.Row.Cells.RemoveAt(inti)
                            Next
                        End If
                    Case "3", "4"
                        intStart = 3
                        Dim drp As New DropDownList

                        drp.ID = "cmbCXA"
                        drp.Width = WebControls.Unit.Percentage(100)
                        drp.BorderStyle = BorderStyle.None
                        drp.BorderWidth = WebControls.Unit.Pixel(0)
                        drp.Font.Name = GetFontName(selLang.SelectedValue)
                        drp.Font.Bold = True
                        drp.Font.Size = WebControls.FontUnit.Point(11)
                        drp.AutoPostBack = False
                        drp.Height = WebControls.Unit.Pixel(CdCst.MonifoldGrid.intGridWidth)

                        Dim drp1 As New DropDownList
                        drp1.ID = "cmbCXB"
                        drp1.Width = WebControls.Unit.Percentage(100)
                        drp1.BorderStyle = BorderStyle.None
                        drp1.BorderWidth = WebControls.Unit.Pixel(0)
                        drp1.Font.Name = GetFontName(selLang.SelectedValue)
                        drp1.Font.Bold = True
                        drp1.Font.Size = WebControls.FontUnit.Point(11)
                        drp1.AutoPostBack = False
                        drp1.Height = WebControls.Unit.Pixel(CdCst.MonifoldGrid.intGridWidth)
                        If dt_Comb.Count > e.Row.RowIndex + 1 AndAlso CType(dt_Comb(e.Row.RowIndex + 1), ArrayList).Count <= 0 Then
                            drp.Style.Add("background-color", "#CCFFCC")
                            drp.Enabled = False
                            drp1.Style.Add("background-color", "#CCFFCC")
                            drp1.Enabled = False
                        Else
                            drp.Style.Add("background-color", "#FFFFCC")
                            drp1.Style.Add("background-color", "#FFFFCC")
                        End If

                        e.Row.Cells(0).Controls.Add(drp)
                        e.Row.Cells(1).Controls.Add(drp1)

                        'Dim strKata() As String = Me.HidSelect.Value.ToString.Split(",")  '画面で選択した形番を取得する
                        Dim strKata As ArrayList = CType(dt_Comb(e.Row.RowIndex + 1), ArrayList)

                        If drp.Enabled Then
                            'CXA,CXB設定
                            Dim intRow As Integer = e.Row.RowIndex
                            Dim CXList As ArrayList
                            '選択肢の設定
                            If HidSelect.Value.ToString.Equals(String.Empty) Then
                                CXList = bllSiyou.GetCXList(HidManifoldMode.Value, objKtbnStrc, String.Empty, intRow)
                            Else
                                Dim strKataCX As String = HidSelect.Value.ToString.Split(",")(e.Row.RowIndex)
                                CXList = bllSiyou.GetCXList(HidManifoldMode.Value, objKtbnStrc, strKataCX, intRow)
                            End If

                            drp.ViewStateMode = UI.ViewStateMode.Enabled
                            drp.DataSource = CXList
                            drp.DataBind()
                            drp1.ViewStateMode = UI.ViewStateMode.Enabled
                            drp1.DataSource = CXList
                            drp1.DataBind()

                            '選択内容の設定
                            If Not HidCXA.Value.ToString.Equals(String.Empty) AndAlso _
                                CXList.Count > 1 Then
                                Dim strCXASelected As String = HidCXA.Value.Split(",")(e.Row.RowIndex)
                                drp.SelectedValue = strCXASelected
                            End If

                            If Not HidCXB.Value.ToString.Equals(String.Empty) AndAlso _
                                CXList.Count > 1 Then
                                Dim strCXBSelected As String = HidCXB.Value.Split(",")(e.Row.RowIndex)
                                drp1.SelectedValue = strCXBSelected
                            End If
                        End If
                End Select

                '形番RowIDを取得する
                If objKtbnStrc.strcSelection.strSeriesKataban = "LMF0" AndAlso _
                    objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "T0D" Then   '逆
                    Dim intLoop As Integer = 1
                    For inti As Integer = e.Row.Cells.Count - 1 To intStart Step -1
                        e.Row.Cells(intLoop).Attributes.Add("onclick", "GridViewCellClick('" & strParent & Me.ID & "','" & _
                                                         CInt(strID) & "','" & inti & "','" & strKataStart & "', '1', '0');")
                        intLoop += 1
                    Next
                Else
                    For inti As Integer = intStart To e.Row.Cells.Count - 1
                        e.Row.Cells(inti).Attributes.Add("onclick", "GridViewCellClick('" & strParent & Me.ID & "','" & _
                                                         CInt(strID) & "','" & inti & "','" & strKataStart & "', '0', '0');")
                    Next
                End If
            End If


        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' OKボタン押す
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        btnOK.Visible = False
        Dim strUpdKataVal() As String = Nothing
        Dim strUpdKigou() As String = Nothing
        Dim intUpdUseVal() As Double
        Dim strUseVal() As String = New String() {}
        Dim strPosition() As String
        Dim clsManCommon As KHManifold
        Dim dt_detail As New DataTable
        Dim dt_title As New DataTable
        Dim dblRailLen As Double = 0D
        Dim strCXA() As String = Nothing
        Dim strCXB() As String = Nothing

        Try
            'HiddenFieldの正規化
            Call FormatAllHiddenField()
            'Me.HidPostBack.Value = "1"    'OKボタンを押すときにフラグを設定する

            Dim strPositions() As String = (From strp In HidClick.Value.Split(";")
                                            Select strp).Distinct.ToArray
            Dim str_itemdiv() As String = Me.HidOther.Value.ToString.Split(",")
            If Not DS_Title.Tables("data") Is Nothing Then dt_detail = DS_Title.Tables("data")
            If Not DS_Title.Tables("title") Is Nothing Then dt_title = DS_Title.Tables("title")

            '入力した使用数情報を保存
            Call SaveInputUsedNumber()

            'ﾚｰﾙ長さの設定
            Call SetRailFromPage(DS_Title)

            'RM1803***_手動入力時コメント
            'If Not HidRailChangeFlg.Value = "1" Then
            '手動入力していない場合は自動計算する
            Call SetRail()
            'End If


            'CX情報をDSに保存
            If (HidManifoldMode.Value = "3" OrElse HidManifoldMode.Value = "4") Then
                If Session("TestMode") Is Nothing Then
                    Call SetCXSelectedInfo(dt_detail)
                End If
            End If

            Me.Session("DS_Title") = DS_Title

            'マニホールドテスト専用
            If Not Me.Session("ManifoldItemKey") Is Nothing Then
                ManifoldTest_Siyou_OK(strCXA, strCXB, dt_title, dt_detail, strUpdKataVal, strUseVal, strUpdKigou)
            Else
                Select Case Me.HidManifoldMode.Value
                    Case 0
                        ReDim strUpdKataVal(dt_title.Rows.Count - 1)
                        ReDim strUseVal(dt_title.Rows.Count - 1)
                        ReDim strUpdKigou(dt_title.Rows.Count - 1)
                        For inti As Integer = 0 To dt_title.Rows.Count - 1
                            strUpdKataVal(inti) = dt_title.Rows(inti)("colKata").ToString
                            strUseVal(inti) = dt_detail.Rows(inti)("col0").ToString
                            strUpdKigou(inti) = dt_title.Rows(inti)("ColNo").ToString
                        Next
                    Case 3, 4
                        ReDim strUpdKataVal(dt_title.Rows.Count - 1)
                        ReDim strUseVal(dt_title.Rows.Count - 1)
                        Dim intEnd As Integer = 0
                        Select Case Me.HidManifoldMode.Value
                            Case 3
                                intEnd = CdCst.Siyou_03.Masking - 1
                            Case 4
                                intEnd = CdCst.Siyou_04.Spacer4 - 1     'RM1803032_スペーサ行追加対応
                        End Select
                        strUpdKataVal = Me.HidSelect.Value.ToString.Split(",")        '画面で選択した形番を取得する
                        'ADD BY YGY 20141125
                        '03の場合は"--"を"MP"に変換する
                        If HidManifoldMode.Value.Equals("3") Then
                            For inti As Integer = 0 To strUpdKataVal.Length - 1
                                If strUpdKataVal(inti).Trim.Equals("--") Then
                                    strUpdKataVal(inti) = "MP"
                                End If
                            Next
                        End If
                        For inti As Integer = 0 To intEnd
                            strUseVal(inti) = dt_detail.Rows(inti)("col0").ToString
                        Next
                        Dim str() As String = Me.HidUse.Value.ToString.Split(",")
                        For inti As Integer = intEnd + 1 To dt_title.Rows.Count - 1
                            strUseVal(inti) = str(inti)
                        Next
                        If Me.HidCXA.Value.Length > 0 Then
                            strCXA = Me.HidCXA.Value.ToString.Split(",")
                            strCXB = Me.HidCXB.Value.ToString.Split(",")
                            objKtbnStrc.strcSelection.strCXAKataban = strCXA
                            objKtbnStrc.strcSelection.strCXBKataban = strCXB
                        End If
                    Case Else
                        'strUpdKataVal = Me.HidSelect.Value.ToString.Split(",")        '画面で選択した形番を取得する
                        'strUseVal = Me.HidUse.Value.ToString.Split(",")               '画面で入力した使用数を取得する

                        ReDim strUpdKataVal(dt_title.Rows.Count - 1)
                        ReDim strUseVal(dt_title.Rows.Count - 1)
                        For inti As Integer = 0 To dt_title.Rows.Count - 1
                            strUpdKataVal(inti) = dt_title.Rows(inti)("colKata").ToString
                            strUseVal(inti) = dt_detail.Rows(inti)("col0").ToString
                        Next
                End Select
            End If

            'CHANGED BY YGY 20141104    M4GA210-C6-E21-16-1
            For inti As Integer = 0 To strUpdKataVal.Length - 1
                'For inti As Integer = 0 To str_itemdiv.Length - 1
                If Not str_itemdiv Is Nothing AndAlso str_itemdiv.Length > inti Then
                    Select Case str_itemdiv(inti)
                        Case "4" 'Tube
                            If Me.HidTube.Value = "0" Then '不要
                                strUpdKataVal(inti) = "0"
                                'ReDim Preserve strUpdKataVal(UBound(strUpdKataVal) + 1)
                                'strUpdKataVal(UBound(strUpdKataVal)) = "0"
                            Else
                                strUpdKataVal(inti) = "1"
                                'ReDim Preserve strUpdKataVal(UBound(strUpdKataVal) + 1)
                                'strUpdKataVal(UBound(strUpdKataVal)) = "1"
                            End If
                    End Select
                End If
            Next

            '検査成績書
            For inti As Integer = 0 To strUpdKataVal.Length - 1
                If strUpdKataVal(inti).Equals(CdCst.Manifold.InspReportJp.Japanese) OrElse _
                    strUpdKataVal(inti).Equals(CdCst.Manifold.InspReportJp.English) Then

                    strUpdKataVal(inti) = CdCst.Manifold.InspReportJp.SelectValue
                ElseIf strUpdKataVal(inti).Equals(CdCst.Manifold.InspReportEn.Japanese) OrElse _
                    strUpdKataVal(inti).Equals(CdCst.Manifold.InspReportEn.English) Then

                    strUpdKataVal(inti) = CdCst.Manifold.InspReportEn.SelectValue
                End If
            Next

            ReDim intUpdUseVal(strUseVal.Length - 1)
            For inti As Integer = 0 To strUseVal.Length - 1
                If strUseVal(inti).Length > 0 AndAlso IsNumeric(strUseVal(inti)) Then
                    If inti = GetRailTubeIndex(0) Then
                        intUpdUseVal(inti) = CDec(strUseVal(inti))
                    Else
                        intUpdUseVal(inti) = CInt(strUseVal(inti))
                    End If
                Else
                    intUpdUseVal(inti) = 0
                End If
            Next

            Dim strPos As String = String.Empty
            Dim intStart As Integer = 1
            Select Case Me.HidManifoldMode.Value
                Case 3, 4
                    intStart = 3
            End Select
            'CHANGED BY YGY 20141028
            ReDim strPosition(dt_detail.Rows.Count - 1)
            'ReDim strPosition(dt_detail.Rows.Count)
            For inti As Integer = 0 To dt_detail.Rows.Count - 1
                If Me.HidManifoldMode.Value > 0 Then
                    Select Case str_itemdiv(inti)
                        Case "5", "6" '取付レール長さ画面入力値取得
                            dblRailLen = intUpdUseVal(inti)
                    End Select
                End If
                strPos = String.Empty
                For intj As Integer = intStart To dt_detail.Columns.Count - 1
                    If dt_detail.Rows(inti)(intj).ToString.Length > 0 AndAlso _
                        (dt_detail.Rows(inti)(intj).ToString = "●" Or dt_detail.Rows(inti)(intj).ToString = "@") Then
                        strPos &= "1"
                    Else
                        strPos &= "0"
                    End If
                Next
                strPosition(inti) = strPos
            Next

            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "MN4S0", "MN3S0"
                    If intUpdUseVal(22) = 0 And strUpdKataVal(22).ToString.Length > 0 Then
                        strUpdKataVal(22) = String.Empty
                    End If
            End Select

            Select Case HidManifoldMode.Value
                Case 5
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "4", "5", "6", "7", "9"
                            For intj As Integer = CdCst.Siyou_05.ElType1 - 1 To CdCst.Siyou_05.SpDecomp4 - 1
                                strPosition(intj) = "2" & Strings.Right(strPosition(intj), strPosition(intj).Length - 1)
                                'CHANGED BY YGY 20141016
                                'strPosition(intj) = Strings.Left(strPosition(intj), 2) & "2" & Strings.Right(strPosition(intj), strPosition(intj).Length - 2)
                                strPosition(intj) = Strings.Left(strPosition(intj), 1) & "2" & Strings.Right(strPosition(intj), strPosition(intj).Length - 2)
                            Next
                            For intj As Integer = CdCst.Siyou_05.ExpCovRep - 1 To CdCst.Siyou_05.ExpCovExh - 1
                                strPosition(intj) = "2" & Strings.Right(strPosition(intj), strPosition(intj).Length - 1)
                                'CHANGED BY YGY 20141016
                                'strPosition(intj) = Strings.Left(strPosition(intj), 2) & "2" & Strings.Right(strPosition(intj), strPosition(intj).Length - 2)
                                strPosition(intj) = Strings.Left(strPosition(intj), 2) & "2" & Strings.Right(strPosition(intj), strPosition(intj).Length - 2)
                            Next
                    End Select
                Case 6
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "LMF0"
                            If objKtbnStrc.strcSelection.strOpSymbol(5).ToString = "9" Then
                                For inti As Integer = 0 To 4
                                    If strUpdKataVal(inti).ToString.Length > 0 Then
                                        'その他電圧の場合、異電圧で終了されていないなら、異電圧を追加
                                        If Not strUpdKataVal(inti).ToString.EndsWith("-" & objKtbnStrc.strcSelection.strOpSymbol(8).ToString) Then
                                            strUpdKataVal(inti) &= "-" & objKtbnStrc.strcSelection.strOpSymbol(8).ToString
                                        End If
                                    End If
                                Next
                            End If
                    End Select
                Case 12                           'ややこしいなVSとGAMD0系、価格の計算方法は違います
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "VSJM", "VSXM", "VSZM"
                            For inti As Integer = 0 To strUpdKataVal.Count - 1
                                If strUpdKataVal(inti).ToString.Length > 0 Then
                                    Dim strRes As String = String.Empty
                                    Dim str() As String = strUpdKataVal(inti).ToString.Split("-")
                                    strRes &= str(0) & ",-,"
                                    strRes &= Strings.Left(str(1), 1) & ","
                                    strRes &= Strings.Mid(str(1), 2, 2) & ","
                                    strRes &= Strings.Right(str(1), 1) & ",-,"
                                    strRes &= str(2) & ",-,"
                                    If str.Length = 4 Then
                                        strRes &= str(3)
                                    End If
                                    strUpdKataVal(inti) = strRes
                                End If
                            Next
                        Case "VSJPM", "VSXPM"
                            For inti As Integer = 0 To strUpdKataVal.Count - 1
                                If strUpdKataVal(inti).ToString.Length > 0 Then
                                    Dim strRes As String = String.Empty
                                    Dim str() As String = strUpdKataVal(inti).ToString.Split("-")
                                    strRes &= str(0) & ",-,"
                                    strRes &= Strings.Left(str(1), 1) & ","
                                    strRes &= Strings.Right(str(1), 1) & ",-,"
                                    If str.Length = 3 Then
                                        strRes &= str(2)
                                    End If
                                    strUpdKataVal(inti) = strRes
                                End If
                            Next
                        Case "VSKM"
                            For inti As Integer = 0 To strUpdKataVal.Count - 1
                                If strUpdKataVal(inti).ToString.Length > 0 Then
                                    Dim strRes As String = String.Empty
                                    Dim str() As String = strUpdKataVal(inti).ToString.Split("-")
                                    If str.Length >= 4 Then
                                        strRes &= str(0) & ",-,"
                                        strRes &= Strings.Left(str(1), 1) & ","
                                        strRes &= Strings.Mid(str(1), 2, 2) & ","
                                        strRes &= Strings.Right(str(1), 1) & ",-,"
                                        strRes &= str(2) & ",-,"
                                        strRes &= str(3) & ",-,"
                                        If str.Length = 5 Then
                                            strRes &= str(4)
                                        End If
                                        strUpdKataVal(inti) = strRes
                                    End If
                                End If
                            Next
                        Case "VSNM"
                            For inti As Integer = 0 To strUpdKataVal.Count - 1
                                If strUpdKataVal(inti).ToString.Length > 0 Then
                                    Dim strRes As String = String.Empty
                                    Dim str() As String = strUpdKataVal(inti).ToString.Split("-")
                                    strRes &= str(0) & ",-,"
                                    strRes &= Strings.Left(str(1), 1) & ","
                                    strRes &= Strings.Right(str(1), 2) & ",-,"
                                    strRes &= str(2) & ",-,"
                                    If str.Length = 4 Then
                                        strRes &= str(3)
                                    End If
                                    strUpdKataVal(inti) = strRes
                                End If
                            Next
                        Case Else    'VSZPM,VSNPM
                            For inti As Integer = 0 To strUpdKataVal.Count - 1
                                If strUpdKataVal(inti).ToString.Length > 0 Then
                                    Dim strRes As String = String.Empty
                                    Dim str() As String = strUpdKataVal(inti).ToString.Split("-")
                                    strRes &= str(0) & ",-,"
                                    strRes &= str(1) & ",-,"
                                    If str.Length = 3 Then
                                        strRes &= str(2)
                                    End If
                                    strUpdKataVal(inti) = strRes
                                End If
                            Next
                    End Select
                Case 17
                    For inti As Integer = 0 To strUpdKataVal.Count - 1
                        If inti >= 6 Then Exit For
                        If strUpdKataVal(inti).ToString.Length > 0 Then
                            Dim strRes As String = String.Empty
                            Dim str() As String = strUpdKataVal(inti).ToString.Split("-")
                            If inti <> 5 Then '単体ブロック
                                If str.Length = 4 Then
                                    strRes &= Strings.Left(str(0), 4) & ","
                                    strRes &= Strings.Mid(str(0), 5, 1) & ","
                                    strRes &= Strings.Right(str(0), 2) & ",-,"
                                    strRes &= str(1) & ",-,"
                                    If str(2).Length = 1 Then
                                        strRes &= str(2) & ","
                                    Else
                                        strRes &= Strings.Left(str(2), 1) & ","
                                        strRes &= Strings.Right(str(2), 1) & ","
                                    End If
                                    strRes &= "-," & str(3)
                                End If
                            Else  'ベース
                                If str.Length = 4 Then
                                    strRes &= str(0) & "-" & str(1) & ",-,"
                                    strRes &= str(2) & ",-,"
                                    strRes &= Strings.Left(str(3), 1) & ","
                                    strRes &= Strings.Right(str(3), 1)
                                End If
                            End If
                            strUpdKataVal(inti) = strRes
                        End If
                    Next
                    If strUpdKataVal.Length > 0 AndAlso strUpdKataVal(5).ToString.Length > 0 Then
                        intUpdUseVal(5) = 1
                    End If
            End Select

            objKtbnStrc.strcSelection.strOptionKataban = strUpdKataVal
            objKtbnStrc.strcSelection.strPositionInfo = strPosition
            objKtbnStrc.strcSelection.intQuantity = intUpdUseVal
            objKtbnStrc.strcSelection.strOpSymbol(0) = objKtbnStrc.strcSelection.strSeriesKataban

            'マニホールド情報も保存する
            Dim dt_MFHistory As New DS_History.MF_HistoryDataTable
            Dim dr_MFHistory As DataRow = dt_MFHistory.NewRow
            dr_MFHistory("UpdateDate") = Now
            Dim hostname As String = String.Empty
            Try
                'Dim strIP As String = Request.UserHostAddress
                'Dim IPhostname As Net.IPHostEntry = System.Net.Dns.GetHostEntry(strIP)
                'hostname = IPhostname.HostName
                'If hostname.Length > 0 Then hostname = Left(hostname, InStr(1, hostname, ".") - 1)
                hostname = "******"
            Catch ex As Exception
            End Try

            dr_MFHistory("UpdateComputer") = Right(hostname.PadRight(10), 10)
            dr_MFHistory("UpdateUser") = Me.objUserInfo.UserId
            dr_MFHistory("DataFlag") = "0"
            dr_MFHistory("Kataban") = objKtbnStrc.strcSelection.strFullKataban
            dr_MFHistory("GSPrice") = objKtbnStrc.strcSelection.intGsPrice
            dr_MFHistory("KataCheck") = objKtbnStrc.strcSelection.strKatabanCheckDiv
            dr_MFHistory("KataPlace") = objKtbnStrc.strcSelection.strPlaceCd
            dr_MFHistory("L1Length") = objKtbnStrc.strcSelection.decDinRailLength

            For inti As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                dr_MFHistory("Kata" & inti + 1) = objKtbnStrc.strcSelection.strOptionKataban(inti)
                Dim strPosition_ As String = String.Empty
                If objKtbnStrc.strcSelection.strPositionInfo.Length > inti AndAlso _
                    (Not objKtbnStrc.strcSelection.strPositionInfo(inti) Is Nothing) AndAlso _
                    objKtbnStrc.strcSelection.strPositionInfo(inti).ToString.Length > 0 Then
                    For intj As Integer = 0 To objKtbnStrc.strcSelection.strPositionInfo(inti).ToString.Length - 1
                        If Mid(objKtbnStrc.strcSelection.strPositionInfo(inti).ToString, intj + 1, 1) = "0" Then
                            strPosition_ &= " "
                        Else
                            strPosition_ &= "Y"
                        End If
                    Next
                End If
                If inti + 1 <= 25 Then dr_MFHistory("Position" & inti + 1) = strPosition_
                dr_MFHistory("Count" & inti + 1) = objKtbnStrc.strcSelection.intQuantity(inti)
            Next

            '入力チェック
            Dim strMsgCd As String = String.Empty
            Dim strMsg As String = String.Empty

            '入力データのチェック
            If Not CheckInput(dt_title, dt_detail, strPositions) Then
                Exit Sub
            End If

            Select Case HidManifoldMode.Value
                Case 0
                    If Not ClsInputCheck_00.fncInputChk(objKtbnStrc, Me.HidSimpleOther.Value.ToString, strMsgCd, strMsg) Then
                        dr_MFHistory("ErrorMsgCd") = strMsgCd
                        dt_MFHistory.Rows.Add(dr_MFHistory)
                        Using da As New DS_HistoryTableAdapters.MF_HistoryTableAdapter
                            If Me.Session("TestMode") Is Nothing Then da.Update(dt_MFHistory)
                        End Using
                        If strMsg.Length > 0 Then
                            AlertMessage(strMsgCd, strMsg)
                        Else
                            AlertMessage(strMsgCd)
                        End If
                        Exit Try
                    Else
                        dt_MFHistory.Rows.Add(dr_MFHistory)
                        Using da As New DS_HistoryTableAdapters.MF_HistoryTableAdapter
                            If Me.Session("TestMode") Is Nothing Then da.Update(dt_MFHistory)
                        End Using
                    End If
                Case Else
                    'Manifold画面選択した形番などのﾁｪｯｸ（WS側、Web系と共通になる）
                    Dim dblStdNum As Double = 0D
                    If Me.HidStdNum.Value.Length > 0 Then
                        Dim strStdNum() As String = Me.HidStdNum.Value.ToString.Split(",")
                        If strStdNum.Length = 2 Then
                            dblStdNum = CDbl(strStdNum(1).ToString)
                        ElseIf strStdNum.Length = 1 Then
                            dblStdNum = CDbl(strStdNum(0).ToString)
                        End If
                    End If

                    If Not SiyouBLL.InputCheck(objKtbnStrc, HidManifoldMode.Value, dblStdNum, strMsg, strMsgCd) Then

                        'テストする場合はエラーメッセージを出力
                        If Session("TestMode") IsNot Nothing Then
                            Session("EventEndFlg") = True
                        End If

                        'エラー履歴を登録
                        dr_MFHistory("ErrorMsgCd") = strMsgCd
                        dt_MFHistory.Rows.Add(dr_MFHistory)

                        Using da As New DS_HistoryTableAdapters.MF_HistoryTableAdapter
                            If Me.Session("TestMode") Is Nothing Then da.Update(dt_MFHistory)
                        End Using

                        If strMsg.StartsWith("0,") Then
                            strMsg = Strings.Right(strMsg, strMsg.Length - 2)
                        End If

                        If strMsg.Split(",").Length = 2 Then
                            AlertMessage(strMsgCd, strMsg.Split(",")(1))
                        Else
                            AlertMessage(strMsgCd, strMsg)
                        End If

                        '警告する場合は単価画面へ遷移
                        If strMsgCd <> "W1300" Then
                            Exit Try
                        End If

                        'エラー位置を赤にする
                        'If Not strMsg Is Nothing AndAlso strMsg.Length > 0 Then
                        '    Dim strErr() As String = strMsg.Split("|")    '複数あるかも
                        '    Dim strM() As String = Nothing
                        '    For inti As Integer = 0 To strErr.Length - 1
                        '        strM = strErr(inti).Split(",")
                        '        If strM.Length = 2 Then                   'XとY
                        '            If IsNumeric(strM(0)) And IsNumeric(strM(1)) Then
                        '                Dim intRow As Integer = CInt(strM(0))
                        '                Dim intColumn As Integer = CInt(strM(1))
                        '                Dim cel As System.Web.UI.WebControls.TableCell
                        '                Dim gridViewDetail As New WebControls.GridView
                        '                Dim blnIndent As Boolean

                        '                If intRow > CInt(Me.HidColCount.Value) Then Exit For
                        '                If intColumn > GridViewTitle.Rows.Count Then Exit For

                        '                If intRow > 0 And intColumn > 0 Then
                        '                    'CXA,CXBの場合
                        '                    If HidManifoldMode.Value = 3 Or HidManifoldMode.Value = 4 Then
                        '                        intColumn += 2
                        '                        'cel = Me.GridViewDetail.Rows(intRow - 1).Cells(intColumn + 2)
                        '                        'cel = Me.GridViewDetail.Rows(intRow - 1).Cells(intColumn)
                        '                    End If
                        '                    'Midの場合
                        '                    If bllSiyou.GetMidRow(HidManifoldMode.Value, intRow, blnIndent) Then
                        '                        gridViewDetail = Me.FindControl("MidGridView")
                        '                    Else
                        '                        gridViewDetail = Me.GridViewDetail
                        '                        cel = gridViewDetail.Rows(intRow - 1).Cells(intColumn)
                        '                        cel.BackColor = Color.Red
                        '                    End If


                        '                ElseIf CInt(strM(1)) = 0 Then   '形番行を選択する
                        '                    cel = Me.GridViewTitle.Rows(strM(0) - 1).Cells(0)
                        '                    cel.BackColor = Color.Red
                        '                End If
                        '            End If
                        '        End If
                        '    Next
                        'End If


                    Else
                        dt_MFHistory.Rows.Add(dr_MFHistory)
                        Using da As New DS_HistoryTableAdapters.MF_HistoryTableAdapter
                            If Me.Session("TestMode") Is Nothing Then da.Update(dt_MFHistory)
                        End Using
                    End If
            End Select

            Select Case HidManifoldMode.Value
                Case 0
                    Call SiyouBLL.subEditSpecInfo_00(objKtbnStrc, strUpdKigou)
                Case 5
                    Call SiyouBLL.subEditSpecInfoGMF(objKtbnStrc, HidManifoldMode.Value)
                Case 6
                    Call SiyouBLL.subEditSpecInfoLMF(objKtbnStrc, HidManifoldMode.Value)
            End Select

            '仕様書構成テーブル更新対象データから取付レール行を削除
            For inti As Integer = 0 To dt_detail.Rows.Count - 1
                If Me.HidManifoldMode.Value > 0 Then
                    Select Case str_itemdiv(inti)
                        Case "5", "6" '取付レール長さ画面入力値取得
                            Dim strSaveKata(strUpdKataVal.Length - 2) As String
                            Dim strSaveUseVal(intUpdUseVal.Length - 2) As Double
                            For intj As Integer = 0 To strUpdKataVal.Length - 1
                                If intj < inti Then
                                    strSaveKata(intj) = strUpdKataVal(intj)
                                ElseIf intj > inti Then
                                    strSaveKata(intj - 1) = strUpdKataVal(intj)
                                End If
                            Next
                            For intj As Integer = 0 To intUpdUseVal.Length - 1
                                If intj < inti Then
                                    strSaveUseVal(intj) = intUpdUseVal(intj)
                                ElseIf intj > inti Then
                                    strSaveUseVal(intj - 1) = intUpdUseVal(intj)
                                End If
                            Next
                            objKtbnStrc.strcSelection.strOptionKataban = strSaveKata
                            objKtbnStrc.strcSelection.intQuantity = strSaveUseVal
                            Exit For
                    End Select
                End If
            Next

            '引当情報更新
            clsManCommon = New KHManifold(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
            clsManCommon.subUpdateSelSpec(objCon, objKtbnStrc, dblRailLen, "1")

            strUpdKataVal = Nothing
            intUpdUseVal = Nothing
            strUseVal = Nothing
            strPosition = Nothing
            clsManCommon = Nothing
            dt_detail = Nothing
            dt_title = Nothing
            dblRailLen = Nothing
            strCXA = Nothing
            strCXB = Nothing

            Select Case objKtbnStrc.strcSelection.strPriceNo.Trim
                Case "89", "96", "D3"      'ISO単価画面へ
                    RaiseEvent GotoISOTanka()
                Case Else                  '普通の単価画面へ
                    RaiseEvent GotoTanka()
            End Select
        Catch ex As Exception
            AlertMessage(ex)
        Finally
            btnOK.Visible = True
        End Try
        'エラーがある場合画面を初期化すること
        'HidPostBack.Value = "1"
    End Sub

    ''' <summary>
    ''' 仕様のセット(ISOの場合)
    ''' </summary>
    ''' <param name="dt_data"></param>
    ''' <param name="dt_title"></param>
    ''' <param name="arr_zokusei"></param>
    ''' <param name="dr"></param>
    ''' <remarks></remarks>
    Private Sub SetSiyouISO_ManifoldTest(ByVal dt_data As DataTable, ByVal dt_title As DataTable, ByVal arr_zokusei As ArrayList, ByVal dr As DataRow)
        'タイトル名を取得
        Dim lstTitle As New List(Of String)
        For Each drTitle In dt_title.Rows
            If Not drTitle("colKata").ToString.Equals(String.Empty) Then
                lstTitle.Add(drTitle("colKata").ToString)
            End If
        Next

        'タイトルの設定
        For intLoop As Integer = 2 To 20
            For intj As Integer = intLoop - 2 To dt_title.Rows.Count - 1
                If intj >= 20 Then Exit For

                Dim str_z() As String = arr_zokusei(intj).ToString.Split(",")
                Dim bolZ As Boolean = False
                Dim intKPosition As Integer = -1

                '属性が一致するかどうかの判断
                For inti As Integer = 0 To str_z.Length - 1
                    If dr("属性" & intLoop).ToString = str_z(inti) Then
                        bolZ = True
                        Exit For
                    End If
                Next
                If Not bolZ Then Continue For

                'ADD BY YGY 20141010    #A320349のようなデータを対応するため
                '既にセットしたタイトルが2回セットしないように
                'タイトルの設定
                If Not lstTitle.Contains(dr("形番" & intLoop).ToString) Then
                    If dr("数量" & intLoop).ToString.Trim.Length > 0 AndAlso _
                       CLng(dr("数量" & intLoop).ToString.Trim) > 0 Then
                        If dr("形番" & intLoop).ToString.Trim.Length > 0 Then
                            Dim cel As System.Web.UI.WebControls.TableCell
                            cel = GridViewTitle.Rows(intj).Cells(GridViewTitle.Rows(intj).Cells.Count - 1)
                            If cel.Controls.Count > 0 Then
                                Dim drp As New DropDownList
                                drp = cel.Controls(0)
                                'ADD BY YGY 20141017    ↓↓↓↓↓↓
                                '#A283313    マニホールドテストの場合は正しく出力ために、形番にある電圧を削除
                                If HidManifoldMode.Value = "6" Then
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                        Case "LMF0"
                                            If objKtbnStrc.strcSelection.strOpSymbol(5).ToString = "9" Then
                                                If intj <= 4 Then
                                                    Dim strSplit() As String = dr("形番" & intLoop).ToString.Split("-")
                                                    Dim strLast As String = strSplit(UBound(strSplit))
                                                    If strLast.Equals(objKtbnStrc.strcSelection.strOpSymbol(8).ToString) Then
                                                        Dim len As Integer = dr("形番" & intLoop).ToString.Length - strLast.Length
                                                        dr("形番" & intLoop) = dr("形番" & intLoop).ToString.Substring(0, len - 1)
                                                    End If
                                                End If
                                                'For inti As Integer = 0 To 4
                                                '    If strUpdKataVal(inti).ToString.Length > 0 Then
                                                '        strUpdKataVal(inti) &= "-" & objKtbnStrc.strcSelection.strOpSymbol(8).ToString
                                                '    End If
                                                'Next
                                            End If
                                    End Select
                                End If
                                drp.Text = dr("形番" & intLoop).ToString
                                dt_title.Rows(intj)("ColKata") = dr("形番" & intLoop).ToString

                                'ADD BY YGY 20141017    ↑↑↑↑↑↑
                            End If
                            cel = Nothing
                        End If
                    End If
                End If

                'ADD BY YGY 20141015
                '対応する形番が一致する場合のみ数量と位置をセット
                If Not dr("形番" & intLoop).ToString.Equals(String.Empty) AndAlso _
                    Not dt_title.Rows(intj)("ColKata").Equals(dr("形番" & intLoop).ToString) Then
                    Continue For
                Else
                    '数量の設定
                    Dim intStart As Long = 0
                    If dr("数量" & intLoop).ToString.Trim.Length > 0 AndAlso _
                       CLng(dr("数量" & intLoop).ToString.Trim) > 0 Then
                        If Not GridViewDetail.Rows(intj) Is Nothing AndAlso _
                            GridViewDetail.Rows(intj).Controls.Count > 0 Then
                            If Not GridViewDetail.Rows(intj).Cells(intStart) Is Nothing Then
                                Dim cel As System.Web.UI.WebControls.TableCell
                                cel = GridViewDetail.Rows(intj).Cells(intStart)
                                cel.Text = dr("数量" & intLoop).ToString
                                dt_data.Rows(intj)("Col0") = dr("数量" & intLoop).ToString
                                cel = Nothing
                            End If
                        End If
                    End If
                    '位置の設定
                    If dr("位置" & intLoop).ToString.Trim.Length > 0 Then
                        Dim strPosition As String = dr("位置" & intLoop).ToString

                        'ADD BY YGY 20141016    ↓↓↓↓↓↓
                        'LMF0の場合位置を逆にする
                        If objKtbnStrc.strcSelection.strSeriesKataban.Equals("LMF0") AndAlso _
                            objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "T0D" Then
                            strPosition = StrReverse(strPosition.PadRight(dt_data.Columns.Count - 1, Space(1)))
                        End If
                        'ADD BY YGY 20141016    ↑↑↑↑↑↑

                        For intK As Integer = 1 To strPosition.Length
                            If Strings.Mid(strPosition, intK, 1).Equals("Y") Then
                                Dim intCol As Integer = -1
                                For intC As Integer = 0 To dt_data.Columns.Count - 1
                                    If dt_data.Columns(intC).ColumnName = "Col" & intK Then
                                        intCol = intC
                                        Exit For
                                    End If
                                Next
                                If intCol = -1 Then Exit For

                                '画面初期化する時に、既に選択した場合
                                If intKPosition = -1 Then intKPosition = intj
                                Dim cel As System.Web.UI.WebControls.TableCell
                                cel = GridViewDetail.Rows(intKPosition).Cells(intCol)

                                If cel.Text.ToString = "●" Then
                                    Continue For
                                End If
                                cel.Text = "●"
                                dt_data.Rows(intKPosition)("Col" & intK) = "●"
                                cel = Nothing
                            End If
                        Next
                    End If
                End If
                Exit For
            Next
        Next
    End Sub

    ''' <summary>
    ''' 仕様のセット(履歴読み出す場合)
    ''' </summary>
    ''' <param name="dt_data"></param>
    ''' <param name="dt_title"></param>
    ''' <param name="arr_zokusei"></param>
    ''' <param name="dr"></param>
    ''' <remarks></remarks>
    Private Sub SetSiyouHistory_ManifoldTest(ByVal dt_data As DataTable, ByVal dt_title As DataTable, ByVal arr_zokusei As ArrayList, ByVal dr As DataRow)

        'When do test using dirty data do not read the option
        Dim lstOther As List(Of String) = HidOther.Value.Split(",").ToList
        Dim lstOptionItemDiv As List(Of String) = New List(Of String) From {"3", "4", "5", "6", "99"}
        Dim intOptionStart As Integer = 30

        For Each strOther As String In lstOther
            If lstOptionItemDiv.Contains(strOther) Then
                intOptionStart = lstOther.IndexOf(strOther)
                Exit For
            End If
        Next

        For intLoop As Integer = 1 To intOptionStart
            For intj As Integer = intLoop - 1 To dt_title.Rows.Count - 1
                Dim intKPosition As Integer = -1

                '形番の設定
                If dr("Count" & intLoop).ToString.Trim.Length > 0 AndAlso _
                   CLng(dr("Count" & intLoop).ToString.Trim) > 0 Then
                    If dr("Kata" & intLoop).ToString.Trim.Length > 0 Then
                        Dim cel As System.Web.UI.WebControls.TableCell
                        cel = GridViewTitle.Rows(intj).Cells(GridViewTitle.Rows(intj).Cells.Count - 1)
                        If cel.Controls.Count > 0 Then
                            Dim strKataFull As String = dr("Kata" & intLoop).ToString
                            Dim strKataban As String = strKataFull


                            If strKataFull.Contains(",") Then
                                'CXが選択された場合
                                Dim strKataList As List(Of String) = strKataFull.Split(",").ToList

                                If strKataList.Count = 3 Then
                                    strKataban = strKataList.Item(0)
                                    Dim strCXA As String = strKataList.Item(1)
                                    Dim strCXB As String = strKataList.Item(2)

                                    dt_data.Rows(intj)("ColKataA") = strCXA
                                    dt_data.Rows(intj)("ColKataB") = strCXB
                                End If
                            End If

                            If TryCast(cel.Controls(0), DropDownList) Is Nothing Then
                                Dim lbl As New Label
                                lbl = cel.Controls(0)
                                lbl.Text = strKataban
                            Else

                                Dim drp As New DropDownList
                                drp = cel.Controls(0)
                                drp.Text = strKataban
                            End If

                            dt_title.Rows(intj)("ColKata") = strKataban
                        End If
                        cel = Nothing
                    End If
                End If

                Dim intStart As Long = 0

                'CXが選択された場合
                If dt_data.Columns(0).ColumnName.Equals("ColKataA") Then
                    intStart = 2
                End If

                '数量の設定
                If dr("Count" & intLoop).ToString.Trim.Length > 0 AndAlso _
                   CLng(dr("Count" & intLoop).ToString.Trim) > 0 Then
                    If Not GridViewDetail.Rows(intj) Is Nothing AndAlso _
                        GridViewDetail.Rows(intj).Controls.Count > 0 Then
                        If Not GridViewDetail.Rows(intj).Cells(intStart) Is Nothing Then
                            Dim cel As System.Web.UI.WebControls.TableCell
                            cel = GridViewDetail.Rows(intj).Cells(intStart)
                            cel.Text = dr("Count" & intLoop).ToString
                            dt_data.Rows(intj)("Col0") = dr("Count" & intLoop).ToString
                            cel = Nothing
                        End If
                    End If
                End If

                '位置の設定　※不明
                If intLoop <= 25 Then
                    If dr("Position" & intLoop).ToString.Trim.Length > 0 Then
                        Dim strPosition As String = dr("Position" & intLoop).ToString
                        For intK As Integer = 0 To strPosition.Length - 1
                            If Strings.Mid(strPosition, intK + 1, 1).Equals("Y") Then
                                Dim intCol As Integer = -1
                                For intC As Integer = 0 To dt_data.Columns.Count - 1
                                    If dt_data.Columns(intC).ColumnName = "Col" & (intK + 1) Then
                                        intCol = intC
                                        Exit For
                                    End If
                                Next
                                If intCol = -1 Then Exit For

                                '画面初期化する時に、既に選択した場合
                                If intKPosition = -1 Then intKPosition = intj
                                Dim cel As System.Web.UI.WebControls.TableCell
                                cel = GridViewDetail.Rows(intKPosition).Cells(intCol)

                                If cel.Text.ToString = "●" Then
                                    Continue For
                                End If
                                cel.Text = "●"
                                dt_data.Rows(intKPosition)("Col" & (intK + 1)) = "●"
                                cel = Nothing
                            End If
                        Next
                    End If
                End If
                Exit For
            Next
        Next
    End Sub

    ''' <summary>
    ''' 仕様のセット
    ''' </summary>
    ''' <param name="dt_data">仕様データ</param>
    ''' <param name="dt_title">仕様タイトル</param>
    ''' <param name="arr_zokusei">kh_item_mstからspec_noに対応する属性コード</param>
    ''' <param name="str_itemdiv">タグ銘板</param>
    ''' <param name="dr">DB保存データ</param>
    ''' <remarks></remarks>
    Private Sub SetSiyou_ManifoldTest(ByRef dt_data As DataTable, ByRef dt_title As DataTable, ByVal arr_zokusei As ArrayList, ByVal str_itemdiv() As String, ByVal dr As DataRow)
        Dim strWire As String = String.Empty    '個別配線形番
        Dim intWireRow As Integer = -1          '個別配線行No
        Dim strWirePosition As New List(Of String)     '個別配線の位置を記録
        'Dim intRowInti As New List(Of Integer)         '初期化された行の番号を記録
        Dim intRowSeted As New List(Of Integer)        'セットした行の番号を記録

        '1-20は部品、20以降は付属品である
        '部品のセット
        Try
            ''初期化された行を記録
            'For Each drTitleInti In dt_title.Rows
            '    If Not drTitleInti("ColKata").Equals(DBNull.Value) Then
            '        intRowInti.Add(dt_title.Rows.IndexOf(drTitleInti))
            '    End If
            'Next

            '個別配線の取得
            For Each datarow In dt_title.Rows
                If datarow("ColNo").Equals("個別配線") AndAlso _
                        datarow("ColKata").ToString.Length > 0 Then
                    Dim intIndex As Integer = dt_title.Rows.IndexOf(datarow)
                    Dim blnWireEnable As Boolean = False
                    For Each cellwire In GridViewDetail.Rows(intIndex).Cells
                        If GridViewDetail.Rows(intIndex).Cells.GetCellIndex(cellwire) = 0 Then
                        Else
                            If Not cellwire.BackColor.Equals(Drawing.Color.FromArgb(192, 192, 192)) Then
                                blnWireEnable = True
                                Exit For
                            End If
                        End If
                    Next
                    '個別配線が選択可能の場合
                    If blnWireEnable Then
                        strWire = datarow("ColKata").ToString
                        intWireRow = dt_title.Rows.IndexOf(datarow)
                        Exit For
                    End If

                End If
            Next

            'DBにより画面の設定
            For intLoop As Integer = 1 To 30
                Dim strZoukusei As String = dr("属性" & intLoop).ToString
                Dim strKataban As String = dr("形番" & intLoop).ToString
                Dim strQuantity As String = dr("数量" & intLoop).ToString
                Dim strPosition As String = dr("位置" & intLoop).ToString
                Dim intLoopEnd As Integer = IIf(dt_title.Rows.Count > 20, 20, dt_title.Rows.Count)
                Dim intPosition As New List(Of Integer)
                '選択位置
                If Not strPosition.Equals(String.Empty) Then
                    Dim chrPosition() As Char = strPosition.ToArray

                    For inttmp As Integer = 0 To chrPosition.Count - 1
                        If chrPosition(inttmp).ToString.Equals("Y") Then
                            intPosition.Add(inttmp + 1)
                        End If
                    Next
                End If

                For intj As Integer = 0 To intLoopEnd - 1
                    Dim str_z() As String = arr_zokusei(intj).ToString.Split(",")    '属性
                    Dim blnSelectEnable As Boolean = True                            '選択可能フラグ
                    Dim blnLeftSelect As Boolean = False                             'ｴﾝﾄﾞﾌﾞﾛｯｸ(左)フラグ
                    Dim strRealKataban As String = String.Empty

                    '対応する属性コードがない場合
                    If Not str_z.Contains(strZoukusei) Then
                        Continue For
                    End If

                    '既にセットされた行の場合(個別配線以外)
                    'AndAlso Not strKataban.EndsWith(strWire)
                    If intRowSeted.Contains(intj) Then
                        Continue For
                    End If

                    '選択不可の場合
                    With GridViewDetail.Rows(intj).Cells
                        For Each intp In intPosition
                            Dim intSelectedPosition As Integer
                            Dim blnIndent As Boolean = False
                            'CXがある場合は位置を調整
                            If dt_data.Columns(0).ColumnName.Equals("ColKataA") Then
                                intSelectedPosition = intp + 2
                            Else
                                intSelectedPosition = intp
                            End If
                            'セットするセルの背景が灰色の場合は選択不可と判断する
                            If bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, intj, blnIndent) Then
                                'インデントを入れた行の場合
                                Dim gridIndent As New WebControls.GridView

                                If blnIndent Then
                                    '第Ⅰ行の場合
                                    gridIndent = GridViewDetail.Rows(intj).Cells(1).Controls(0)
                                    If gridIndent.Rows(0).Cells(intSelectedPosition).BackColor.Equals(Drawing.Color.FromArgb(192, 192, 192)) Then
                                        blnSelectEnable = False
                                    End If
                                Else
                                    '第Ⅱ行の場合
                                    gridIndent = GridViewDetail.Rows(intj - 1).Cells(1).Controls(0)
                                    If gridIndent.Rows(1).Cells(intSelectedPosition).BackColor.Equals(Drawing.Color.FromArgb(192, 192, 192)) Then
                                        blnSelectEnable = False
                                    End If
                                End If
                            Else
                                'インデント以外の場合
                                '背景は灰色且つ"●"未入力の場合
                                If GridViewDetail.Rows(intj).Cells(intSelectedPosition).BackColor.Equals(Drawing.Color.FromArgb(192, 192, 192)) AndAlso _
                                    Not GridViewDetail.Rows(intj).Cells(intSelectedPosition).Text.Equals("●") Then
                                    blnSelectEnable = False
                                    Exit For
                                End If
                            End If
                        Next
                    End With
                    If blnSelectEnable.Equals(False) Then
                        Continue For
                    End If

                    'ｴﾝﾄﾞﾌﾞﾛｯｸが既に入力した場合
                    If dt_title.Rows(intj)("ColNo").ToString.StartsWith("ｴﾝﾄﾞﾌﾞﾛｯｸ") Then
                        'Colが始まる列の数
                        Dim intColumnNo As Integer = CType(dt_data.Columns(dt_data.Columns.Count - 1).ColumnName.Replace("Col", ""), Integer)

                        For intq As Integer = 1 To intColumnNo
                            If dt_data.Rows(intj)("Col" & intq).Equals("●") AndAlso _
                                Not intq.Equals(strPosition.IndexOf("Y") + 1) Then
                                blnLeftSelect = True
                                Exit For
                            End If
                        Next
                    End If
                    If blnLeftSelect Then
                        Continue For
                    End If

                    '仕切りﾌﾞﾛｯｸの場合形番で対応する項目を決める
                    If dt_title.Rows(intj)("ColNo").ToString.StartsWith("仕切りﾌﾞﾛｯｸ") OrElse _
                        dt_title.Rows(intj)("ColNo").ToString.StartsWith("仕切ﾌﾟﾗｸﾞ") Then
                        If Not dt_title.Rows(intj)("colKata").ToString.Equals(strKataban) Then
                            Continue For
                        End If
                    End If


                    Dim intKPosition As Integer = -1
                    '03と04場合、CXの可能性があります
                    If (objKtbnStrc.strcSelection.strSpecNo.ToString.Trim = "03" Or _
                                       objKtbnStrc.strcSelection.strSpecNo.ToString.Trim = "04") AndAlso _
                                   dr("形番" & (intj + 1)).ToString.Contains("-CX") AndAlso _
                               Not dr("形番" & (intj + 1)).ToString.EndsWith("-CX") Then
                        Dim cel As System.Web.UI.WebControls.TableCell
                        Dim str() As String = dr("形番" & (intj + 1)).ToString.Split("-")
                        Dim strRealKata As String = String.Empty
                        For intR As Integer = 0 To str.Length - 2
                            If strRealKata.Length > 0 Then strRealKata &= "-"
                            strRealKata &= str(intR)
                        Next
                        cel = GridViewTitle.Rows(intj).Cells(GridViewTitle.Rows(intj).Cells.Count - 1)
                        If cel.Controls.Count > 0 Then
                            Dim drp As New DropDownList
                            drp = cel.Controls(0)
                            drp.Text = strRealKata
                            dt_title.Rows(intj)("colKata") = strRealKata
                            strRealKataban = strRealKata
                            ''セットした行Noを記録する
                            'If Not intRowSeted.Contains(intj) Then
                            '    intRowSeted.Add(intj)
                            'End If
                        End If
                        cel = Nothing

                        '形番からCXAとCXBを展開
                        Dim strCX As String = str(str.Length - 1)
                        Dim CXflag As Boolean = False

                        'CHANGED BY YGY 20140826    ↓↓↓↓↓↓
                        'CX選択候補を取得
                        Dim CXList As ArrayList = bllSiyou.GetCXList(HidManifoldMode.Value, objKtbnStrc, strRealKata, intj)
                        If CXList.Count > 0 Then
                            '後から展開して、展開した結果が同時に選択候補にある場合は終了
                            For intCX As Integer = strCX.Length - 1 To 0 Step -1
                                Dim strCXB As String = Strings.Right(strCX, strCX.Length - intCX)
                                Dim strCXA As String = Strings.Left(strCX, intCX)

                                Dim celA As WebControls.TableCell = GridViewDetail.Rows(intj).Cells(0)
                                Dim celB As WebControls.TableCell = GridViewDetail.Rows(intj).Cells(1)
                                If celA.Controls.Count > 0 AndAlso celB.Controls.Count > 0 Then
                                    If CXList.Contains(strCXA) AndAlso CXList.Contains(strCXB) Then
                                        CType(celA.Controls(0), DropDownList).Text = strCXA
                                        CType(celB.Controls(0), DropDownList).Text = strCXB
                                        dt_data.Rows(intj)("ColKataA") = strCXA
                                        dt_data.Rows(intj)("ColKataB") = strCXB
                                        Exit For
                                    End If
                                End If
                            Next
                            'End If
                            'CHANGED BY YGY 20140826    ↑↑↑↑↑↑
                        End If
                    Else
                        If strKataban.Trim.Length > 0 Then
                            '個別配線の場合
                            If strWire.Length > 0 And strKataban.EndsWith(strWire) Then
                                Dim str As String = strKataban
                                If Not str.Length.Equals(strWire.Length) Then
                                    strKataban = Strings.Left(str, str.Length - strWire.Length - 1)
                                End If
                                For intPK As Integer = 0 To intj
                                    If dt_title.Rows(intPK)("ColKata").ToString = strKataban Then
                                        '個別配線を記録
                                        intKPosition = intPK
                                        strWirePosition.Add(strPosition)
                                        Exit For
                                    End If
                                Next
                            End If
                            '個別配線以外の場合
                            If intKPosition = -1 Then
                                Dim intSetRowNo As Integer = intj
                                If intRowSeted.Contains(intSetRowNo) Then
                                    intSetRowNo += 1
                                End If
                                Dim cel As System.Web.UI.WebControls.TableCell
                                cel = GridViewTitle.Rows(intSetRowNo).Cells(GridViewTitle.Rows(intSetRowNo).Cells.Count - 1)
                                If cel.Controls.Count > 0 Then
                                    If dt_title.Rows(intSetRowNo)("ColKata").Equals(DBNull.Value) OrElse _
                                        dt_title.Rows(intSetRowNo)("ColKata").Equals(String.Empty) Then
                                        Dim drp As New DropDownList
                                        drp = cel.Controls(0)
                                        drp.Text = strKataban
                                        dt_title.Rows(intSetRowNo)("ColKata") = strKataban
                                    Else

                                    End If
                                    '個別配線を記録
                                    If strWire.Length > 0 And dr("形番" & intLoop).ToString.EndsWith(strWire) Then
                                        strWirePosition.Add(strPosition)
                                    End If

                                End If
                                cel = Nothing
                            End If
                        End If
                    End If

                    Dim intStart As Long = 0
                    If objKtbnStrc.strcSelection.strSpecNo.ToString.Trim = "03" Or _
                                       objKtbnStrc.strcSelection.strSpecNo.ToString.Trim = "04" Then
                        intStart += 2
                    End If

                    If strPosition.Trim.Length > 0 Then
                        For intK As Integer = 0 To strPosition.Length - 1
                            If Strings.Mid(strPosition, intK + 1, 1).Equals("Y") Then
                                Dim intCol As Integer = -1
                                For intC As Integer = 0 To dt_data.Columns.Count - 1
                                    If dt_data.Columns(intC).ColumnName = ("Col" & (intK + 1)) Then
                                        intCol = intC
                                        Exit For
                                    End If
                                Next
                                If intCol = -1 Then Exit For

                                '画面初期化する時に、既に選択した場合
                                If intKPosition = -1 Then
                                    If intRowSeted.Contains(intj) Then
                                        intKPosition = intj + 1
                                    Else
                                        intKPosition = intj
                                    End If
                                End If
                                Dim cel As System.Web.UI.WebControls.TableCell

                                'ADD BY YGY 20140801    ↓↓↓↓↓↓
                                Dim strHide() As String = Me.HidOther.Value.ToString.Split(",")
                                Dim bolFirstLineInner As Boolean = False
                                If strHide.Count > intKPosition AndAlso strHide(intKPosition).Equals("4") Then
                                    'チューブ抜具の場合
                                    Continue For
                                ElseIf bllSiyou.GetMidRow(objKtbnStrc.strcSelection.strSpecNo, intKPosition, bolFirstLineInner) Then
                                    'インデントを入れた行の場合
                                    Dim gridVeiwDetailInner As New WebControls.GridView

                                    If Not bolFirstLineInner Then
                                        gridVeiwDetailInner = GridViewDetail.Rows(intKPosition - 1).Cells(1).Controls(0)
                                        cel = gridVeiwDetailInner.Rows(1).Cells(intCol)
                                    Else
                                        gridVeiwDetailInner = GridViewDetail.Rows(intKPosition).Cells(1).Controls(0)
                                        cel = gridVeiwDetailInner.Rows(0).Cells(intCol)
                                    End If
                                Else
                                    cel = GridViewDetail.Rows(intKPosition).Cells(intCol)
                                End If
                                'ADD BY YGY 20140801    ↑↑↑↑↑↑
                                'ラベル形番がデータ形番と不一致することを防ぐ
                                While intKPosition < dt_title.Rows.Count - 1
                                    If dt_title.Rows(intKPosition)("ColKata").Equals(strKataban) OrElse _
                                       dt_title.Rows(intKPosition)("ColKata").Equals(strRealKataban) Then
                                        Exit While
                                    Else
                                        intKPosition += 1
                                    End If
                                End While
                                If dt_title.Rows(intKPosition)("ColKata").Equals(strKataban) OrElse _
                                       dt_title.Rows(intKPosition)("ColKata").Equals(strRealKataban) Then
                                    'セットした行Noを記録する
                                    If Not intRowSeted.Contains(intKPosition) Then
                                        intRowSeted.Add(intKPosition)
                                    End If
                                    '●をセットする
                                    If cel.Text.ToString = "●" Then
                                        Continue For
                                    Else
                                        cel.Text = "●"
                                        dt_data.Rows(intKPosition)("Col" & (intK + 1)) = "●"
                                    End If
                                    cel = Nothing
                                End If
                            End If
                        Next
                    Else
                        'セットした行Noを記録する
                        If Not intRowSeted.Contains(intj) Then
                            intRowSeted.Add(intj)
                        End If
                    End If
                    Exit For
                Next
            Next
            '個別配線のデータのセット
            If (Not strWire.Equals(String.Empty)) AndAlso (Not intWireRow.Equals(-1)) AndAlso (strWirePosition.Count > 0) Then
                For Each wr In strWirePosition
                    For intp As Integer = 1 To 40
                        wr = wr.PadRight(40, Space(1))
                        If wr.Substring(intp - 1, 1).Equals("Y") Then
                            dt_data.Rows(intWireRow)("Col" & intp) = "●"
                        End If
                    Next
                Next
            End If

            '数量をセット
            For Each drow In dt_data.Rows
                Dim intCountMaru As Integer = 0

                For inti As Integer = 1 To HidColCount.Value
                    If drow("Col" & inti).Equals("●") Then
                        intCountMaru += 1
                    End If
                Next
                If intCountMaru <> 0 Then
                    drow("Col0") = intCountMaru
                Else
                    drow("Col0") = 0
                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try

        '添付品(21-30)
        Dim intCount As Integer = 31
        Dim CompleteRows As New List(Of Integer)
        Dim hasFlg As Boolean = False
        Dim IsL1Flg As Boolean = False

        Try
            While intCount <= 45
                Dim intDataRow As Integer = 0
                '属性がない場合次へ
                If dr("属性" & intCount).Equals(DBNull.Value) OrElse _
                    dr("属性" & intCount).Equals(String.Empty) Then
                    intCount += 1
                Else
                    For intj As Integer = 0 To dt_title.Rows.Count - 1
                        'DBのデータと一致する項目を探す
                        If arr_zokusei(intj).ToString.Contains(dr("属性" & intCount)) AndAlso _
                            Not CompleteRows.Contains(intj) AndAlso _
                            Not dr("形番" & intCount).Equals(DBNull.Value) Then
                            '画面上にこの属性の位置がある場合
                            hasFlg = True             '属性あるフラグ
                            intDataRow = intj         '属性記録
                            CompleteRows.Add(intDataRow) '属性位置を記録
                            Exit For
                        End If
                    Next

                    If hasFlg Then
                        hasFlg = False
                        '形番と数量をセット
                        If intDataRow < str_itemdiv.Length Then
                            Select Case str_itemdiv(intDataRow)
                                'Case "4", "5", "6", "9", "99"
                                Case "12"
                                    Dim cel As System.Web.UI.WebControls.TableCell
                                    cel = GridViewTitle.Rows(intDataRow).Cells(GridViewTitle.Rows(intDataRow).Cells.Count - 1)
                                    If cel.Controls.Count > 0 Then
                                        Dim drp As New DropDownList
                                        drp = cel.Controls(0)
                                        drp.Text = dr("形番" & intCount).ToString
                                        dt_title.Rows(intDataRow)("ColKata") = dr("形番" & intCount).ToString
                                    End If

                                Case Else
                                    Dim cel As System.Web.UI.WebControls.TableCell
                                    cel = GridViewTitle.Rows(intDataRow).Cells(GridViewTitle.Rows(intDataRow).Cells.Count - 1)
                                    If cel.Controls.Count > 0 Then
                                        'CHANGED BY YGY 20140708
                                        If cel.Controls(0).GetType.Name.Equals("DropDownList") Then
                                            Dim drp As New DropDownList

                                            If dr("属性" & intCount).ToString = "L1" Then
                                                IsL1Flg = True
                                                '形番
                                                Call SetRailData(GetRailTubeIndex(0))
                                                dt_title.Rows(intDataRow)("ColKata") = DS_Title.Tables("data").Rows(intDataRow)("Col0").ToString
                                            Else
                                                drp = cel.Controls(0)
                                                drp.Text = dr("形番" & intCount).ToString
                                                dt_title.Rows(intDataRow)("ColKata") = dr("形番" & intCount).ToString
                                                '数量
                                                If dr("数量" & intCount).ToString.Length > 0 AndAlso CLng(dr("数量" & intCount).ToString) > 0 Then
                                                    cel = New System.Web.UI.WebControls.TableCell
                                                    cel = GridViewDetail.Rows(intDataRow).Cells(0)
                                                    cel.Text = CLng(dr("数量" & intCount).ToString)
                                                    dt_data.Rows(intDataRow)("Col0") = CLng(dr("数量" & intCount).ToString)
                                                End If
                                            End If
                                        Else
                                            'チューブ抜具
                                            Dim chk As New CheckBox
                                            chk = cel.Controls(0)
                                            If dr("形番" & intCount).ToString.Equals(CdCst.strTube) Then
                                                chk.Checked = True
                                                dt_title.Rows(intDataRow)("ColKata") = dr("形番" & intCount).ToString
                                            Else
                                                chk.Checked = False
                                                dt_title.Rows(intDataRow)("ColKata") = dr("形番" & intCount).ToString
                                            End If
                                        End If

                                    End If
                            End Select
                        End If

                    End If
                    intCount += 1
                End If
            End While

            '取付レールをセット
            If IsL1Flg.Equals(False) Then
                For Each drTitle In dt_title.Rows
                    Dim intRowNo As Integer = dt_title.Rows.IndexOf(drTitle)

                    If drTitle("ColNo").Equals("取付ﾚｰﾙ長さ") Then
                        If Not GridViewDetail.Rows(intRowNo).Cells(0).BackColor.Equals(Drawing.Color.FromArgb(192, 192, 192)) Then
                            Call SetRailData(GetRailTubeIndex(0))
                            If dt_data.Rows(intRowNo)("Col0") <> 0 Then
                                drTitle("ColKata") = dt_data.Rows(intRowNo)("Col0")
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' 最後のカンマを削除する
    ''' </summary>
    ''' <param name="strArg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RemoveLastComma(ByVal strArg As String) As String
        If strArg.Equals(String.Empty) Then
            Return String.Empty
        Else
            Return strArg.Substring(0, strArg.Length - 1)
        End If
    End Function

    ''' <summary>
    ''' OKボタンを押した時、画面上のﾚｰﾙ長さを保存
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetRailFromPage(ByRef DS_Title As DataSet)
        Dim dt_detail As DataTable = DS_Title.Tables("data")
        Dim dt_title As DataTable = DS_Title.Tables("title")

        'ﾚｰﾙ長さをセット
        Dim intRail As Integer = -1
        Dim celtxt As System.Web.UI.WebControls.TableCell
        intRail = GetRailTubeIndex(0) 'すべての順番
        If intRail >= 0 Then
            Select Case Me.HidManifoldMode.Value
                Case "3", "4"
                    celtxt = Me.GridViewDetail.Rows(intRail).Cells(2)
                Case Else
                    celtxt = Me.GridViewDetail.Rows(intRail).Cells(0)
            End Select

            If celtxt.Controls.Count > 0 Then
                Dim txt As TextBox = celtxt.Controls(0)
                Dim blnHasScript As Boolean = False

                dt_title.Rows(intRail)("ColKata") = txt.Text
                dt_detail.Rows(intRail)("Col0") = txt.Text

            End If
        End If
    End Sub

    ''' <summary>
    ''' 入力が数字であるかどうかの判断Javascript
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function NumberJudgement(Optional ByVal strRail As String = "") As String
        Dim strJavascript As String = String.Empty
        Dim strErrMessage As String = ClsCommon.fncGetMsg(selLang.SelectedValue, "W1002")
        Dim strHidRailChangeFlgID As String = Me.HidRailChangeFlg.ClientID

        '数字判断ロジック
        If strRail.Equals(String.Empty) Then
            strJavascript = "if(isNaN(this.value)) {alert('" & strErrMessage & "');this.select();}"
        Else
            strJavascript = "if(isNaN(this.value)) {alert('" & strErrMessage & "');this.select();} else {fncTextRailChange('" & strHidRailChangeFlgID & "');}"
        End If

        Return strJavascript
    End Function

    ''' <summary>
    ''' データテーブルをクリアする
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <remarks></remarks>
    Private Sub ClearData(ByVal dt As DataTable)
        For inti As Integer = 0 To dt.Rows.Count - 1
            For intj As Integer = 0 To dt.Columns.Count - 1
                dt.Rows(inti)(intj) = String.Empty
            Next
        Next
    End Sub

    ''' <summary>
    ''' 使用数だけの行に追加するテキストボックスの作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreateTextBox(ByVal str_itemdiv() As String, ByVal intRowIndex As Integer) As TextBox
        '@@@@テキストボックスの作成
        Dim txt As New TextBox
        txt.ID = "txtNum"
        txt.Width = WebControls.Unit.Pixel(70 - 2)
        txt.BorderStyle = BorderStyle.None
        txt.BorderWidth = WebControls.Unit.Pixel(0)
        txt.Font.Name = GetFontName(selLang.SelectedValue)
        txt.Font.Bold = True
        txt.Font.Size = WebControls.FontUnit.Point(12)
        txt.AutoPostBack = False
        txt.ViewStateMode = UI.ViewStateMode.Enabled
        txt.Style.Add("text-align", "center")
        txt.Style.Add("margin", "0")
        txt.Style.Add("padding", "0")

        Select Case str_itemdiv(intRowIndex)
            Case "5", "6"
                'レール長さが5桁まで入力できる
                txt.MaxLength = 5
                '数字判断とﾚｰﾙ長さ変更フラグの設定
                txt.Attributes.Add("onblur", NumberJudgement("Rail"))
            Case Else
                '他のオプションが2桁まで入力できる
                txt.MaxLength = 2
                '数字判断
                txt.Attributes.Add("onblur", NumberJudgement())
        End Select

        If dt_Comb.Count > intRowIndex + 1 Then
            If CType(dt_Comb(intRowIndex + 1), ArrayList).Count <= 0 Then
                txt.Style.Add("background-color", "#CCFFCC")
                txt.ReadOnly = True
            Else
                txt.Style.Add("background-color", "#FFFFCC")
            End If
        End If

        If str_itemdiv(intRowIndex) = "99" Then 'タブ銘板入力不可、選択したら、1になる
            txt.ReadOnly = True
            txt.Style.Add("background-color", "#C7EDCC")
        End If

        Return txt
    End Function

    ''' <summary>
    ''' 使用数だけの行に追加するドロップダウンリストの作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreateCXDropDownList(ByVal strID As String) As DropDownList
        Dim drp As New DropDownList

        drp.ID = strID
        drp.Width = WebControls.Unit.Percentage(100)
        drp.BorderStyle = BorderStyle.None
        drp.BorderWidth = WebControls.Unit.Pixel(0)
        drp.Font.Name = GetFontName(selLang.SelectedValue)
        drp.Font.Bold = True
        drp.Font.Size = WebControls.FontUnit.Point(11)
        drp.AutoPostBack = False
        drp.Height = WebControls.Unit.Pixel(CdCst.MonifoldGrid.intGridWidth)
        drp.Enabled = False
        drp.Style.Add("background-color", "#CCFFCC")
        drp.Enabled = False

        Return drp
    End Function

    ''' <summary>
    ''' 重複チェック
    ''' </summary>
    ''' <param name="inti">ループ行番号</param>
    ''' <param name="intRowIdx">比較行番号</param>
    ''' <param name="intColIdx">比較列番号</param>
    ''' <param name="dt_detail">選択データ</param>
    ''' <returns>重複データあるかどうか</returns>
    ''' <remarks></remarks>
    Private Function CheckDoubleSelect(ByVal inti As Integer, ByVal intRowIdx As Integer, ByVal intColIdx As Integer, ByVal dt_detail As DataTable) As Boolean
        Dim blnDouble As Boolean = True

        'RM1803032_スペーサ行追加対応
        If inti <> intRowIdx Then '5,6はISO系、同じ列中で複数の形番を選択できる
            Select Case HidManifoldMode.Value
                Case 1
                    blnDouble = SetPositionGroupEx(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_01.Elect1, CdCst.Siyou_01.Regulat2, _
                                       CdCst.Siyou_01.Valve1, CdCst.Siyou_01.Dummy2, CdCst.Siyou_01.Wiring, CdCst.Siyou_01.Wiring)
                Case 4
                    blnDouble = SetPositionGroupEx(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_04.Valve1, CdCst.Siyou_04.Spacer4, _
                                        CdCst.Siyou_04.Valve1, CdCst.Siyou_04.MasPlate2, CdCst.Siyou_04.Spacer1, CdCst.Siyou_04.Spacer4)
                Case 5
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_05.ElType1, CdCst.Siyou_05.ElType6)
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_05.ABCon02, CdCst.Siyou_05.ABCon04)
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_05.ABPlugR, CdCst.Siyou_05.ABPlugL)
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_05.RepSpace1, CdCst.Siyou_05.RepSpace2)
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_05.ExhSpace1, CdCst.Siyou_05.ExhSpace2)
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_05.Pilot1, CdCst.Siyou_05.Pilot2)
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_05.SpDecomp1, CdCst.Siyou_05.SpDecomp4)
                Case 6
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_06.Elect1, CdCst.Siyou_06.Elect6)
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_06.ABCon01, CdCst.Siyou_06.ABCon1Z)
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_06.RepSpace1, CdCst.Siyou_06.RepSpace2)
                    blnDouble = blnDouble And SetPositionGroup(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_06.ExhSpace1, CdCst.Siyou_06.ExhSpace2)
                Case 7
                    'Elect1～Mixを選択したら、Spacer1～Spacer2を変わらずに、逆も同じ
                    blnDouble = SetPositionGroupEx(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_07.Equip, CdCst.Siyou_07.EndRight, _
                                        CdCst.Siyou_07.Elect1, CdCst.Siyou_07.Mix, CdCst.Siyou_07.Spacer1, CdCst.Siyou_07.Spacer4)
                Case 9
                    blnDouble = SetPositionGroupEx(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_09.Endb1, CdCst.Siyou_09.ExhaustSp, _
                                        CdCst.Siyou_09.ElValve1, CdCst.Siyou_09.MpValve2, CdCst.Siyou_09.SpReguP, CdCst.Siyou_09.ExhaustSp)
                Case 15
                    blnDouble = SetPositionGroupEx(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_15.InOut1, CdCst.Siyou_15.EndL, _
                                        CdCst.Siyou_15.Valve1, CdCst.Siyou_15.Valve8, CdCst.Siyou_15.Spacer1, CdCst.Siyou_15.Spacer4)
                Case 16
                    blnDouble = SetPositionGroupEx(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_16.EndL, CdCst.Siyou_16.RegulatorB, _
                                        CdCst.Siyou_16.Valve1, CdCst.Siyou_16.MPValve2, CdCst.Siyou_16.Spacer1, CdCst.Siyou_16.Spacer2)
                Case 18
                    blnDouble = SetPositionGroupEx(dt_detail, intRowIdx, intColIdx, inti, CdCst.Siyou_18.Equip, CdCst.Siyou_18.EndRight, _
                                        CdCst.Siyou_18.Elect1, CdCst.Siyou_18.Mix, CdCst.Siyou_18.Spacer1, CdCst.Siyou_18.Spacer4)
                Case Else
                    If dt_detail.Rows(inti)(intColIdx).Equals(strMaru) Then
                        blnDouble = False
                    End If
            End Select
        End If

        Return blnDouble
    End Function

    ''' <summary>
    ''' HiddenFieldの正規化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FormatAllHiddenField()
        'CHANGED BY YGY 20141029
        'Hide項目に最後のカンマを削除する
        Me.HidSelect.Value = RemoveLastComma(Me.HidSelect.Value.ToString)
        Me.HidUse.Value = RemoveLastComma(Me.HidUse.Value.ToString)
        Me.HidCXA.Value = RemoveLastComma(Me.HidCXA.Value.ToString)
        Me.HidCXB.Value = RemoveLastComma(Me.HidCXB.Value.ToString)
    End Sub

    ''' <summary>
    ''' マニホールドテスト時実行
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ManifoldTest_Siyou()
        'マニホールドテスト専用
        'ADD BY YGY 20140708
        Dim str_itemdiv() As String = Me.HidOther.Value.ToString.Split(",")

        If Not Me.Session("ManifoldKataban") Is Nothing Then
            If Not Me.Session("TestFlag") Is Nothing Then Exit Sub
            Me.Session("TestFlag") = True

            'Dim lngLoop As Long = CLng(Me.Session("ManifoldKatabanLoop"))
            Dim listKataban As New ManifoldKataban(Me.Session("ManifoldKataban"))
            Dim strKataban As String = listKataban.KATABAN.ToString
            Dim strSiyou As String = listKataban.SIYOUSYO.ToString

            '属性リストを取得する
            Dim arr_zokusei As New ArrayList
            Dim dt_zokusei As New DS_100Test.kh_item_mstDataTable
            Using da_zokusei As New DS_100TestTableAdapters.kh_item_mstTableAdapter
                da_zokusei.FillbySpec(dt_zokusei, objKtbnStrc.strcSelection.strSpecNo.ToString.Trim)
            End Using
            For inti As Integer = 0 To dt_zokusei.Rows.Count - 1
                If Not dt_zokusei.Rows(inti) Is Nothing AndAlso dt_zokusei.Rows(inti)("item_num").ToString.Length > 0 Then
                    For intj As Integer = 0 To CLng(dt_zokusei.Rows(inti)("item_num").ToString) - 1
                        arr_zokusei.Add(dt_zokusei.Rows(inti)("zokusei_cd").ToString)
                    Next
                End If
            Next

            '仕様詳細データを取得する
            Dim dt As New DataTable
            Dim strPath As String = String.Empty
            Select Case Me.Session("TestMode").ToString
                Case "1"
                    dt = New DS_100Test.MF_Siyou_ISODataTable

                    Using da As New DS_100TestTableAdapters.MF_Siyou_ISOTableAdapter
                        da.FillBy(dt, strKataban, strSiyou)
                    End Using
                Case "2"

                    Dim dr As DS_PriceTest.kh_shiyou_testRow = Me.Session("ManifoldKataban")
                    Dim drShiyouTest As DS_PriceTest.kh_shiyou_testRow

                    dt = New DS_PriceTest.kh_shiyou_testDataTable

                    drShiyouTest = dt.NewRow

                    drShiyouTest.ItemArray = dr.ItemArray.Clone

                    dt.Rows.Add(drShiyouTest)
                Case Else
                    dt = New DS_100Test.MF_SiyouDataTable

                    Using da As New DS_100TestTableAdapters.MF_SiyouTableAdapter
                        da.FillBy(dt, strKataban, strSiyou)
                    End Using
            End Select

            If dt.Rows.Count > 0 Then
                Dim dr As DataRow = dt.Rows(0)

                Dim dt_data As DataTable = New DataTable
                Dim dt_title As DataTable = New DataTable
                '
                If Not Me.Session("DS_Title") Is Nothing Then
                    If Not DS_Title.Tables("data") Is Nothing Then dt_data = DS_Title.Tables("data")
                    If Not DS_Title.Tables("title") Is Nothing Then dt_title = DS_Title.Tables("title")
                End If

                'Main(1-20)
                Select Case Me.Session("TestMode").ToString
                    Case "1"  'ISO
                        SetSiyouISO_ManifoldTest(dt_data, dt_title, arr_zokusei, dr)
                    Case "2"  '履歴
                        SetSiyouHistory_ManifoldTest(dt_data, dt_title, arr_zokusei, dr)
                    Case Else
                        SetSiyou_ManifoldTest(dt_data, dt_title, arr_zokusei, str_itemdiv, dr)
                End Select

            Else
                Exit Sub
            End If
            Me.Session("DS_Title") = DS_Title

            Call btnOK_Click(Me, Nothing)
        End If
    End Sub

    ''' <summary>
    ''' マニホールドテスト時OKボタンを押すイベント
    ''' </summary>
    ''' <param name="strCXA"></param>
    ''' <param name="strCXB"></param>
    ''' <param name="dt_title"></param>
    ''' <param name="dt_detail"></param>
    ''' <param name="strUpdKataVal"></param>
    ''' <param name="strUseVal"></param>
    ''' <param name="strUpdKigou"></param>
    ''' <remarks></remarks>
    Private Sub ManifoldTest_Siyou_OK(ByRef strCXA() As String, ByRef strCXB() As String, ByVal dt_title As DataTable, ByVal dt_detail As DataTable, ByRef strUpdKataVal() As String, ByRef strUseVal() As String, ByRef strUpdKigou() As String)
        ReDim strUpdKataVal(dt_title.Rows.Count - 1)
        ReDim strUseVal(dt_title.Rows.Count - 1)
        ReDim strUpdKigou(dt_title.Rows.Count - 1)

        Select Case Me.HidManifoldMode.Value
            Case 3, 4
                ReDim strCXA(dt_title.Rows.Count - 1)
                ReDim strCXB(dt_title.Rows.Count - 1)
                For inti As Integer = 0 To dt_title.Rows.Count - 1
                    strCXA(inti) = dt_detail.Rows(inti)("ColKataA").ToString
                    strCXB(inti) = dt_detail.Rows(inti)("ColKataB").ToString
                Next
                objKtbnStrc.strcSelection.strCXAKataban = strCXA
                objKtbnStrc.strcSelection.strCXBKataban = strCXB

                For inti As Integer = 0 To dt_title.Rows.Count - 1
                    strUpdKataVal(inti) = dt_title.Rows(inti)("colKata").ToString
                    strUseVal(inti) = dt_detail.Rows(inti)("col0").ToString
                Next
            Case 5, 6    'ISO
                'ADD BY YGY 20141010
                For Each dr In dt_title.Rows
                    Dim index As Integer = dt_title.Rows.IndexOf(dr)
                    'DBデータにA・Bポートが指定された場合、画面上に入力できない時は
                    'DBに登録されたデータを削除
                    If dr("colNo").ToString.StartsWith("A・Bﾎﾟｰﾄﾌﾟﾗｸﾞ位置") OrElse _
                        dr("colNo").ToString.StartsWith("A・Bﾎﾟｰﾄ接続口径") Then

                        Dim cellCount As Integer = Me.GridViewDetail.Rows(index).Cells.Count

                        If cellCount > 0 Then

                            Dim celDetail As System.Web.UI.WebControls.TableCell
                            Dim celTitle As System.Web.UI.WebControls.DropDownList

                            celDetail = Me.GridViewDetail.Rows(index).Cells(cellCount - 1)
                            celTitle = CType(Me.GridViewTitle.Rows(index).Cells(1).Controls(0), DropDownList)

                            '選択不可の場合
                            ' OrElse celTitle.Enabled = False
                            If (celDetail.BackColor.Equals(Drawing.Color.FromArgb(192, 192, 192))) Then
                                'DBに登録された場合
                                If Not dt_detail.Rows(index)("col0").ToString.Equals(String.Empty) Then
                                    dt_detail.Rows.RemoveAt(index)
                                    dt_detail.Rows.InsertAt(dt_detail.NewRow, index)
                                End If
                            Else
                                'ISOのA・Bﾎﾟｰﾄ接続口径、形番無くでも位置を指定できる
                                '色が正しいでも、選択不可の場合
                                If (HidManifoldMode.Value = 5 And index >= CdCst.Siyou_05.ABCon02 - 1 And _
                                    index <= CdCst.Siyou_05.ABCon04 - 1) Or _
                                    (HidManifoldMode.Value = 6 And index >= CdCst.Siyou_06.ABCon01 - 1 And _
                                     index <= CdCst.Siyou_06.ABCon1Z - 1 And _
                                     objKtbnStrc.strcSelection.strOpSymbol(2).ToString = "XX") Then
                                ElseIf HidManifoldMode.Value = 5 AndAlso _
                                    objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "L" And (index = 7 Or index = 6) Then
                                Else
                                    'DBに登録された場合
                                    If Not dt_detail.Rows(index)("col0").ToString.Equals(String.Empty) Then
                                        dt_detail.Rows.RemoveAt(index)
                                        dt_detail.Rows.InsertAt(dt_detail.NewRow, index)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next

                For inti As Integer = 0 To dt_title.Rows.Count - 1
                    strUpdKataVal(inti) = dt_title.Rows(inti)("colKata").ToString
                    strUseVal(inti) = dt_detail.Rows(inti)("col0").ToString
                Next
            Case 0
                For inti As Integer = 0 To dt_title.Rows.Count - 1
                    strUpdKataVal(inti) = dt_title.Rows(inti)("colKata").ToString
                    strUseVal(inti) = dt_detail.Rows(inti)("col0").ToString
                    strUpdKigou(inti) = dt_title.Rows(inti)("ColNo").ToString
                Next
            Case Else
                For inti As Integer = 0 To dt_title.Rows.Count - 1
                    strUpdKataVal(inti) = dt_title.Rows(inti)("colKata").ToString
                    strUseVal(inti) = dt_detail.Rows(inti)("col0").ToString
                Next
        End Select
    End Sub

    ''' <summary>
    ''' ユーザー入力した使用数を保存
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SaveInputUsedNumber()

        If Not HidUse.Value.Equals(String.Empty) Then
            Dim strHidUse() As String = HidUse.Value.Split(",")

            For intRow As Integer = 0 To strHidUse.Length - 1
                Dim strUse As String = strHidUse(intRow)
                DS_Title.Tables("data").Rows(intRow)("Col0") = strUse
            Next
        End If
    End Sub

    ''' <summary>
    ''' 選択したCX情報をDSに保存
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetCXSelectedInfo(ByVal dt_detail As DataTable)
        Dim strCXA() As String = Me.HidCXA.Value.ToString.Split(",")
        Dim strCXB() As String = Me.HidCXB.Value.ToString.Split(",")

        For inti As Integer = 0 To strCXA.Count - 1
            dt_detail.Rows(inti)("ColKataA") = strCXA(inti)
            dt_detail.Rows(inti)("ColKataB") = strCXB(inti)
        Next
    End Sub

    ''' <summary>
    ''' レール長さ更新ボタンの作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreateRailUpdateBtn() As Button
        Dim btnResult As New Button

        'オプション名称データ取得
        Dim dt_Title As DataTable = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHSiyou, selLang.SelectedValue)
        Dim dr_Title() As DataRow = dt_Title.Select("label_div = 'B' AND label_seq = 1")
        If dr_Title.Count > 0 Then
            btnResult.Text = dr_Title(0).Item("label_content")
        End If

        btnResult.UseSubmitBehavior = False
        btnResult.Width = WebControls.Unit.Percentage(100)
        btnResult.BorderWidth = 1
        btnResult.BorderStyle = BorderStyle.Outset
        '計算イベントの作成
        AddHandler btnResult.Click, AddressOf subUpdateRail

        btnResult.Attributes.Add("onclick", "UpdateRail('" & Me.ClientID & "');")
        Return btnResult
    End Function

    ''' <summary>
    ''' レール長さを更新
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subUpdateRail()
        'セッション設定
        Session.Add("RailUpdate", "TRUE")

        'HiddenFieldの正規化
        Call FormatAllHiddenField()

        Dim strPositions() As String = (From strp In HidClick.Value.Split(";")
                                        Select strp).Distinct.ToArray
        Dim str_itemdiv() As String = Me.HidOther.Value.ToString.Split(",")

        '入力した使用数情報を保存
        Call SaveInputUsedNumber()

        'ﾚｰﾙ長さの設定
        Call SetRailFromPage(DS_Title)

        '手動入力フラグをクリアする
        HidRailChangeFlg.Value = "0"

        Call SetRail()
        'End If
    End Sub

End Class
