Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports System.Drawing

Public Class WebUC_Master
    Inherits KHBase

#Region "プロパティ"
    Private strMode As String = String.Empty
    Private intColWeight As Integer = 20
    Private strKey As String = String.Empty
    Private subSetLbl As DataTable = New DataTable
    Private ListLevel As ArrayList = New ArrayList
    Private ListSeq As ArrayList = New ArrayList
#End Region

#Region "定数"
    Dim strWidth_User() As Integer = {80, 80, 80, 170, 50, 50, 50, 170, 70, 110, 70} 'タイトル幅
    Dim strWidth_CountryItem() As Integer = {150, 630, 100, 100} 'タイトル幅
    'Dim strWidth_RateLocal() As Integer = {120, 150, 100, 100, 80, 80} 'タイトル幅
    Dim strWidth_RateLocal() As Integer = {180, 180, 160, 160, 150, 150} 'タイトル幅
    'Dim strWidth_RateNet() As Integer = {120, 120, 150, 100, 80, 80} 'タイトル幅
    Dim strWidth_RateNet() As Integer = {180, 180, 160, 160, 150, 150} 'タイトル幅

    Dim intLevel() As Integer = {1, 2, 4, 8, 16, 32, 64, 128, 256, 512, 1024}        'レベル
    Dim strWidth_Check() As Integer = {90, 125, 105, 145, 105, 120, 140, 135}
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Me.OnLoad(Nothing)
        If strMode = "11" Then
            Me.RadioButton1.Visible = True
            Me.RadioButton1.Checked = True
            Me.RadioButton2.Visible = True
            Me.RadioButton2.Checked = False
            RadioButton1.Attributes.Add("onclick", "RadioClick('" & Me.ClientID & "_','1');")
            RadioButton2.Attributes.Add("onclick", "RadioClick('" & Me.ClientID & "_','2');")
        Else
            Me.RadioButton1.Checked = False
            Me.RadioButton2.Checked = False
            Me.RadioButton1.Visible = False
            Me.RadioButton2.Visible = False
        End If

        'ユーザマスタ初期化
        If strMode = "10" Then
            AspNetPager1.Visible = True
            AspNetPager1.RecordCount = 0
            AspNetPager1.CurrentPageIndex = 1
        Else
            AspNetPager1.Visible = False
        End If

        '端末認証情報の設定
        Call SetWebLoginPnl()
    End Sub

    ''' <summary>
    ''' ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub
        If Me.objUserInfo.UserId Is Nothing Then Exit Sub
        Try
            SetAttributes(Title2, 1)
            SetAttributes(lblUserID, 1)
            strMode = Me.HidMode.Value.ToString
            lblUserID.Text = Me.objUserInfo.UserId

            Call CreateSelPnl()
            Call CreatEditPnl()
            Select Case strMode
                Case "10" '10 ユーザーマスタメンテナンス
                    Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHUserMaster, selLang.SelectedValue, Me)
                    If Not Me.pnlSelect.FindControl("txtS_UserID") Is Nothing Then
                        Me.pnlSelect.FindControl("txtS_UserID").Focus()
                    End If
                    subSetLbl = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHUserMaster, selLang.SelectedValue)

                Case "11" '11 掛率マスタメンテナンス
                    Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHRateMstMnt, selLang.SelectedValue, Me)
                    If Not Me.pnlSelect.FindControl("txtS_Search") Is Nothing Then
                        Me.pnlSelect.FindControl("txtS_Search").Focus()
                    End If
                    subSetLbl = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHRateMstMnt, selLang.SelectedValue)
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "14" '14 国別生産品マスタメンテナンス
                    Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHCountryItemMstMnt, selLang.SelectedValue, Me)
                    If Not Me.pnlSelect.FindControl("txtS_CountryID") Is Nothing Then
                        Me.pnlSelect.FindControl("txtS_CountryID").Focus()
                    End If
                    subSetLbl = KHLabelCtl.fncGetPageAllLabels(objCon, CdCst.PgmId.KHCountryItemMstMnt, selLang.SelectedValue)
                Case "15" '15 マスタメンテナンス
            End Select
            'ラベルタイトル設置
            Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHMaster, selLang.SelectedValue, Me)
            Call SetAllFontName(Me)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 検索欄の生成
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub CreateSelPnl(Optional intMode As Integer = 0)
        Dim btn As Button = Nothing
        Me.pnlSelect.Controls.Clear()
        Try
            Select Case strMode
                Case "10" '10 ユーザーマスタメンテナンス
                    Call CreatTextBox(1, "txtS_UserID", 150)
                    Call CreatTextBox(2, "txtS_StdDate", 150)
                Case "11" '11 掛率マスタメンテナンス
                    If HidRateDiv.Value = "1" Then Call CreatDropDown(2, "txtS_Made", 150, 0)
                    Call CreatDropDown(3, "txtS_Sale", 150, 0)
                    Call CreatTextBox(4, "txtS_Search", 180)
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "14" '14 国別生産品マスタメンテナンス
                    Call CreatDropDown(1, "txtS_CountryID", 150, 0)
                    Call CreatTextBox(2, "txtS_Kataban", 450)
                Case "15" '15 マスタメンテナンス
            End Select
            btn = New Button
            btn.ID = "Button3"
            AddHandler btn.Click, AddressOf btnSearch
            btn.Width = WebControls.Unit.Pixel(60)
            Me.pnlSelect.Controls.Add(btn)
            btn = New Button
            btn.ID = "Button1"
            AddHandler btn.Click, AddressOf btnClear
            btn.Width = WebControls.Unit.Pixel(60)
            Me.pnlSelect.Controls.Add(btn)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 編集欄の作成
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub CreatEditPnl(Optional intMode As Integer = 0)
        Dim lbl As Label = Nothing
        Dim txt As TextBox = Nothing
        Dim btn As Button = Nothing
        Dim drp As DropDownList = Nothing
        Dim chk As CheckBox = Nothing
        Dim strWidth_Edit() As Integer = Nothing
        Me.pnlEditInput.Controls.Clear()
        Try
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    strWidth_Edit = {0, 0, 0, 130, 550, 80, 80, 0, 0, 0} 'タイトル幅
                    For inti = 3 To 6
                        lbl = New Label
                        lbl.ID = "Label" & inti
                        SetAttributes(lbl, 2)
                        lbl.Width = WebControls.Unit.Pixel(strWidth_Edit(inti))
                        lbl.Height = WebControls.Unit.Pixel(intColWeight - 2)
                        Me.pnlEditTitle.Controls.Add(lbl)

                        Select Case inti
                            Case 3
                                drp = New DropDownList
                                drp.ID = "txtEdit" & inti
                                Dim dt As DataTable = MasterBLL.fncSQL_CountryCodeList(objConBase, selLang.SelectedValue)
                                If Not dt Is Nothing Then
                                    drp.DataTextField = "country_nm"
                                    drp.DataValueField = "country_cd"
                                    drp.DataSource = dt
                                    drp.DataBind()
                                End If
                                SetAttributes(drp, 2)
                                drp.BackColor = Drawing.Color.FromArgb(255, 255, 192)
                                drp.Height = WebControls.Unit.Pixel(intColWeight + 2)
                                drp.Width = WebControls.Unit.Pixel(strWidth_Edit(inti) + 4)
                                Me.pnlEditInput.Controls.Add(drp)
                            Case Else
                                txt = New TextBox
                                txt.Text = String.Empty
                                Select Case inti
                                    Case 4
                                        txt.MaxLength = 60
                                        SetAttributes(txt, 3)
                                    Case 5  '発効日
                                        txt.Text = Now.ToString("yyyy/MM/dd")
                                        txt.MaxLength = 10
                                        SetAttributes(txt, 2)
                                    Case 6  '失効日
                                        txt.Text = "9999/12/31"
                                        txt.MaxLength = 10
                                        SetAttributes(txt, 2)
                                End Select
                                txt.ID = "txtEdit" & inti
                                txt.BackColor = Drawing.Color.FromArgb(255, 255, 192)
                                txt.Width = WebControls.Unit.Pixel(strWidth_Edit(inti))
                                txt.Height = WebControls.Unit.Pixel(intColWeight - 2)
                                Me.pnlEditInput.Controls.Add(txt)
                        End Select
                    Next
                Case "11" '11 掛率マスタメンテナンス
                    strWidth_Edit = {0, 0, 0, 0, 0, 140, 140, 200, 110, 110, 100, 80, 80} 'タイトル幅
                    For inti As Integer = 5 To 12
                        If HidRateDiv.Value = "1" Then
                            Select Case inti
                                Case 8, 9
                                    Continue For
                            End Select
                        Else
                            Select Case inti
                                Case 5, 10
                                    Continue For
                            End Select
                        End If
                        lbl = New Label
                        lbl.ID = "Label" & inti
                        SetAttributes(lbl, 2)
                        lbl.Width = WebControls.Unit.Pixel(strWidth_Edit(inti))
                        lbl.Height = WebControls.Unit.Pixel(intColWeight - 2)
                        Me.pnlEditTitle.Controls.Add(lbl)

                        Select Case inti
                            Case 5, 6
                                drp = New DropDownList
                                drp.ID = "txtEdit" & inti
                                Dim dt As DataTable = MasterBLL.fncSQL_CountryCodeList(objConBase, selLang.SelectedValue)
                                If Not dt Is Nothing Then
                                    drp.DataTextField = "country_nm"
                                    drp.DataValueField = "country_cd"
                                    drp.DataSource = dt
                                    drp.DataBind()
                                End If
                                SetAttributes(drp, 2)
                                drp.BackColor = Drawing.Color.FromArgb(255, 255, 192)
                                drp.Height = WebControls.Unit.Pixel(intColWeight + 2)
                                drp.Width = WebControls.Unit.Pixel(strWidth_Edit(inti) + 4)
                                Me.pnlEditInput.Controls.Add(drp)
                            Case Else
                                txt = New TextBox
                                txt.Text = String.Empty
                                Select Case inti
                                    Case 7
                                        txt.MaxLength = 60
                                        SetAttributes(txt, 3)
                                    Case 11  '発効日
                                        txt.Text = Now.ToString("yyyy/MM/dd")
                                        txt.MaxLength = 10
                                        SetAttributes(txt, 2)
                                    Case 12  '失効日
                                        txt.Text = "9999/12/31"
                                        txt.MaxLength = 10
                                        SetAttributes(txt, 2)
                                    Case Else
                                        txt.MaxLength = 17
                                        SetAttributes(txt, 2)
                                End Select
                                txt.ID = "txtEdit" & inti
                                txt.BackColor = Drawing.Color.FromArgb(255, 255, 192)
                                txt.Width = WebControls.Unit.Pixel(strWidth_Edit(inti))
                                txt.Height = WebControls.Unit.Pixel(intColWeight - 2)
                                Me.pnlEditInput.Controls.Add(txt)
                        End Select
                    Next
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    strWidth_Edit = {0, 0, 0, 0, 0, 80, 80, 80, 120, 50, 50, 80, 180, 70, 0, 150} 'タイトル幅
                    For inti As Integer = 5 To 15
                        If inti = 14 Then Continue For
                        lbl = New Label
                        lbl.ID = "Label" & inti
                        SetAttributes(lbl, 2)
                        lbl.Width = WebControls.Unit.Pixel(strWidth_Edit(inti))
                        lbl.Height = WebControls.Unit.Pixel(intColWeight - 2)
                        Me.pnlEditTitle.Controls.Add(lbl)
                    Next

                    For inti As Integer = 5 To 15
                        txt = New TextBox
                        txt.Text = String.Empty
                        txt.Style.Add("padding", "2px 0px 0px 2px")
                        Select Case inti
                            Case 5, 13
                                txt.MaxLength = 10
                                If inti = 5 Then txt.Style.Add("text-transform", "uppercase")
                            Case 8, 12
                                txt.MaxLength = 50
                            Case 11
                                txt.MaxLength = 2
                                txt.Style.Add("text-transform", "uppercase")
                            Case 6  '発効日
                                txt.Text = Now.ToString("yyyy/MM/dd")
                                txt.MaxLength = 10
                            Case 7  '失効日
                                txt.Text = "9999/12/31"
                                txt.MaxLength = 10
                            Case 9, 10, 15  '国コード DropDown '営業所 DropDown 'ユーザー種別 DropDown
                                drp = New DropDownList
                                drp.ID = "txtEdit" & inti
                                Dim dt_data As New DataTable
                                Select Case inti
                                    Case 9
                                        drp.DataTextField = "country_cd"
                                        dt_data = MasterBLL.fncSQL_CountryMst(objConBase)
                                    Case 10
                                        drp.DataTextField = "office_cd"
                                        dt_data = MasterBLL.fncSQL_OfficeMst(objConBase)
                                    Case 15
                                        drp.DataTextField = "user_class_nm"
                                        drp.DataValueField = "user_class"
                                        dt_data = MasterBLL.fncSQL_UserClassMst(objConBase, selLang.SelectedValue)
                                        '国内GSの場合は自動的にMACADDRESS入力
                                        drp.Attributes.Add("onChange", "SetGSClientInfo('" & Me.ClientID & "','" & drp.ClientID & "');")
                                End Select
                                If Not dt_data Is Nothing AndAlso dt_data.Rows.Count > 0 Then
                                    drp.DataSource = dt_data
                                    drp.DataBind()
                                End If
                                SetAttributes(drp, 2)
                                drp.BackColor = Drawing.Color.FromArgb(255, 255, 192)
                                drp.Height = WebControls.Unit.Pixel(intColWeight + 2)
                                drp.Width = WebControls.Unit.Pixel(strWidth_Edit(inti) + 4)
                                Me.pnlEditInput.Controls.Add(drp)
                                Continue For
                            Case 14 'パスワード有効期限（非表示）
                                Continue For
                        End Select
                        txt.ID = "txtEdit" & inti
                        SetAttributes(txt, 2)
                        txt.BackColor = Drawing.Color.FromArgb(255, 255, 192)
                        txt.Width = WebControls.Unit.Pixel(strWidth_Edit(inti))
                        txt.Height = WebControls.Unit.Pixel(intColWeight - 2)
                        Me.pnlEditInput.Controls.Add(txt)
                    Next

                    For inti As Integer = 0 To 7
                        chk = New CheckBox
                        chk.ID = "Label2" & inti
                        SetAttributes(chk)
                        chk.Width = WebControls.Unit.Pixel(strWidth_Check(inti))
                        chk.Height = WebControls.Unit.Pixel(intColWeight)
                        Me.pnlEditInput1.Controls.Add(chk)
                    Next

                    For inti As Integer = 0 To 7
                        chk = New CheckBox
                        chk.ID = "Label3" & inti
                        SetAttributes(chk)
                        chk.Width = WebControls.Unit.Pixel(strWidth_Check(inti))
                        chk.Height = WebControls.Unit.Pixel(intColWeight)
                        Me.pnlEditInput2.Controls.Add(chk)
                    Next

                    For inti As Integer = 0 To 6
                        chk = New CheckBox
                        chk.ID = "Label4" & inti
                        SetAttributes(chk)
                        chk.Width = WebControls.Unit.Pixel(strWidth_Check(inti))
                        chk.Height = WebControls.Unit.Pixel(intColWeight)
                        Me.pnlEditInput3.Controls.Add(chk)
                    Next

                    '権限選択肢を無効にする
                    pnlEditInput1.Enabled = False
                    pnlEditInput2.Enabled = False
                    pnlEditInput3.Enabled = False
                Case "15" '15 マスタメンテナンス
            End Select
            btn = New Button
            btn.ID = "Button5"
            AddHandler btn.Click, AddressOf btnAdd
            btn.Width = WebControls.Unit.Pixel(60)
            Me.pnlEditButton.Controls.Add(btn)
            Dim strMsg As String = ClsCommon.fncGetMsg(selLang.SelectedValue, "I5020")
            If Not btn Is Nothing Then btn.Attributes.Add(CdCst.JavaScript.OnClick, strConfirm(strMsg))

            btn = New Button
            btn.ID = "Button6"
            AddHandler btn.Click, AddressOf btnEdit
            btn.Width = WebControls.Unit.Pixel(60)
            btn.Enabled = False
            Me.pnlEditButton.Controls.Add(btn)
            strMsg = ClsCommon.fncGetMsg(selLang.SelectedValue, "I5030")
            If Not btn Is Nothing Then btn.Attributes.Add(CdCst.JavaScript.OnClick, strConfirm(strMsg))

            btn = New Button
            btn.ID = "Button7"
            AddHandler btn.Click, AddressOf btnDelete
            btn.Width = WebControls.Unit.Pixel(60)
            btn.Enabled = False
            Me.pnlEditButton.Controls.Add(btn)
            strMsg = ClsCommon.fncGetMsg(selLang.SelectedValue, "I5040")
            If Not btn Is Nothing Then btn.Attributes.Add(CdCst.JavaScript.OnClick, strConfirm(strMsg))
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 端末認証欄の生成
    ''' </summary>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub CreatWebLogPnl(Optional intMode As Integer = 0)

    End Sub

    ''' <summary>
    ''' 色の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetInputColor()
        Dim obj As Object = Nothing
        Try
            For inti As Integer = 3 To 15
                obj = Me.FindControl("txtEdit" & inti)
                If Not obj Is Nothing Then obj.BackColor = Drawing.Color.FromArgb(255, 255, 192)
            Next
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 登録ボタン（新規）
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnAdd(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim bolCheck As Boolean = True
        Try
            If Not UpdateCheck() Then Exit Sub
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    Dim strCountryID As String = CType(Me.FindControl("txtEdit3"), DropDownList).SelectedValue.ToString
                    Dim strKataban As String = CType(Me.FindControl("txtEdit4"), TextBox).Text.Trim.ToString
                    '新規登録(存在あれば、メッセージを出す；なければ、登録)
                    Dim dt As New DS_Master.kh_country_item_mstDataTable
                    Using da As New DS_MasterTableAdapters.kh_country_item_mstTableAdapter
                        Dim intSeq As Object = da.MaxSeqNo(strKataban, strCountryID)
                        If intSeq Is DBNull.Value OrElse intSeq Is Nothing Then
                            intSeq = 1
                        Else
                            intSeq += 1
                        End If
                        Dim dr As DataRow = dt.NewRow
                        dr(0) = strKataban.ToUpper
                        dr(1) = strCountryID
                        dr(2) = intSeq
                        dr("cost_flag") = "0"
                        dr("cost_price") = 0
                        dr("cost_rate") = 0
                        If CreatUpdateData(dr) Then
                            dt.Rows.Add(dr)
                            da.Update(dt)
                        End If
                    End Using
                    Dim txtS_CountryID As DropDownList = Me.pnlSelect.FindControl("txtS_CountryID")
                    If Not txtS_CountryID Is Nothing Then txtS_CountryID.SelectedValue = strCountryID
                    Dim txtS_Kataban As TextBox = Me.pnlSelect.FindControl("txtS_Kataban")
                    If Not txtS_Kataban Is Nothing Then txtS_Kataban.Text = strKataban
                Case "11" '11 掛率マスタメンテナンス
                    Dim strSaleCountry As String = CType(Me.FindControl("txtEdit6"), DropDownList).SelectedValue.ToString
                    Dim strKataban As String = CType(Me.FindControl("txtEdit7"), TextBox).Text.Trim.ToString

                    If HidRateDiv.Value = "1" Then  '購入価格

                        Dim strMadeCountry As String = CType(Me.FindControl("txtEdit5"), DropDownList).SelectedValue.ToString
                        Dim dt As New DS_Master.kh_country_rate_netprice_mstDataTable
                        Using da As New DS_MasterTableAdapters.kh_country_rate_netprice_mstTableAdapter
                            Dim intSeq As Object = da.MaxSeqNo(strMadeCountry, strSaleCountry, strKataban)
                            If intSeq Is DBNull.Value OrElse intSeq Is Nothing Then
                                intSeq = 1
                            Else
                                '既に存在する場合は、その有効期限をチェックすること
                                Dim dtNetPrice As New DataTable

                                dtNetPrice = da.GetDataByKata(strMadeCountry.ToUpper, strSaleCountry.ToUpper, strKataban.ToUpper)

                                For Each drNetPrice As DataRow In dtNetPrice.Rows
                                    If drNetPrice.Item("in_effective_date") <= Now AndAlso _
                                        drNetPrice.Item("out_effective_date") >= Now Then
                                        AlertMessage("W5060")
                                        Exit Sub
                                    End If
                                Next
                                intSeq += 1
                            End If
                            Dim dr As DataRow = dt.NewRow
                            dr(0) = strMadeCountry.ToUpper
                            dr(1) = strSaleCountry.ToUpper
                            dr(2) = strKataban.ToUpper
                            dr(3) = intSeq
                            If CreatUpdateData(dr) Then
                                dt.Rows.Add(dr)
                                da.Update(dt)
                            End If
                        End Using
                    Else '現地定価
                        Dim dt As New DS_Master.kh_country_rate_localprice_mstDataTable
                        Using da As New DS_MasterTableAdapters.kh_country_rate_localprice_mstTableAdapter
                            Dim intSeq As Object = da.MaxSeqNo(strSaleCountry, strKataban)
                            If intSeq Is DBNull.Value OrElse intSeq Is Nothing Then
                                intSeq = 1
                            Else
                                '既に存在する場合は、その有効期限をチェックすること
                                Dim dtNetPrice As New DataTable

                                dtNetPrice = da.GetDataByKata(strSaleCountry.ToUpper, strKataban.ToUpper)

                                For Each drNetPrice As DataRow In dtNetPrice.Rows
                                    If drNetPrice.Item("in_effective_date") <= Now AndAlso _
                                        drNetPrice.Item("out_effective_date") >= Now Then
                                        AlertMessage("W5060")
                                        Exit Sub
                                    End If
                                Next
                                intSeq += 1
                            End If
                            Dim dr As DataRow = dt.NewRow
                            dr(0) = strSaleCountry.ToUpper
                            dr(1) = strKataban.ToUpper
                            dr(2) = intSeq
                            If CreatUpdateData(dr) Then
                                dt.Rows.Add(dr)
                                da.Update(dt)
                            End If
                        End Using
                    End If
                    Dim txtS_Search As TextBox = Me.pnlSelect.FindControl("txtS_Search")
                    If Not txtS_Search Is Nothing Then txtS_Search.Text = strKataban
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    Dim strUserID As String = CType(Me.FindControl("txtEdit5"), TextBox).Text.Trim
                    '新規登録(存在あれば、SeqNo+1；なければ、直接登録)
                    Dim dt As New DS_Master.kh_user_mstDataTable
                    Using da As New DS_MasterTableAdapters.kh_user_mstTableAdapter
                        Dim intSeq As Object = da.MaxSeqNo(strUserID)
                        If intSeq Is DBNull.Value OrElse intSeq Is Nothing Then
                            intSeq = 1
                        Else
                            '既に存在する場合は、その有効期限をチェックすること
                            Dim dtUser As New DataTable
                            dtUser = da.GetDataByUserID(strUserID)

                            For Each drUser As DataRow In dtUser.Rows
                                If drUser.Item("in_effective_date") <= Now AndAlso _
                                    drUser.Item("out_effective_date") >= Now Then
                                    AlertMessage("W5060")
                                    Exit Sub
                                End If
                            Next
                            intSeq += 1
                        End If
                        Dim dr As DataRow = dt.NewRow
                        dr(0) = strUserID.ToUpper
                        dr(1) = intSeq
                        If CreatUpdateData(dr) Then
                            dt.Rows.Add(dr)
                            da.Update(dt)
                        End If

                        HidTableKey.Value = dr(0) & "," & dr(1)
                    End Using

                    '端末認証情報の保存
                    If pnlWebLog.Visible Then
                        '入力情報の取得
                        Dim strPass As String = txtPassword.Text
                        Dim strMac As String = txtMacAddress.Text
                        Dim strSerial As String = txtSerial.Text
                        'Dim dLastUsedTime As Date = CType(txtLastUsedTime.Text, Date)

                        '情報の保存
                        Dim dtUser As New DS_M_User.M_UserDataTable
                        Using daUser As New DS_M_UserTableAdapters.M_UserTableAdapter
                            Dim intNum As Integer = 0
                            intNum = daUser.GetNumByUserID(strUserID)
                            If intNum > 0 Then
                                '更新
                                daUser.UpdateByUserID(strPass, strMac, strSerial, strUserID)
                            Else
                                '追加
                                daUser.Insert(strUserID, strPass, strMac, strSerial)
                            End If
                        End Using
                    End If

                    Dim txtS_UserID As TextBox = Me.pnlSelect.FindControl("txtS_UserID")
                    If Not txtS_UserID Is Nothing Then txtS_UserID.Text = strUserID

                    'btnSearch(sender, e)

                    'Dim strScript As String = "UserMasterCellClick('ctl00_ContentDetail_WebUC_Master','ctl00_ContentDetail_WebUC_Master_GridViewMain_ctl02_ctl00_ctl02','IDH303,1');"
                    'ScriptManager.RegisterStartupScript(Page, Page.GetType, "UserMasterCellClick", strScript, False)

                    '権限選択肢を有効にする
                    pnlEditInput1.Enabled = True
                    pnlEditInput2.Enabled = True
                    pnlEditInput3.Enabled = True
                    'ボタンの設定
                    CType(Me.FindControl("button5"), Button).Enabled = False
                    CType(Me.FindControl("button6"), Button).Enabled = True
                    CType(Me.FindControl("button7"), Button).Enabled = True
                    'チェックボックスの設定
                    Dim txtS_StdDate As TextBox = Me.pnlSelect.FindControl("txtS_StdDate")
                    subSetCheckBox(txtS_UserID.Text, txtS_StdDate.Text)

                Case "15" '15 マスタメンテナンス
            End Select

            AlertMessage("I5120")
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' チェックボックスの設定
    ''' </summary>
    ''' <param name="userId"></param>
    ''' <param name="stdDate"></param>
    ''' <remarks></remarks>
    Private Sub subSetCheckBox(ByVal userId As String, stdDate As String)

        Dim dt_view As DataTable
        dt_view = MasterBLL.fncSQL_UserMstList(objConBase, userId, stdDate, Me.selLang.SelectedValue, 1, 5)

        For intType As Integer = 2 To 4
            Dim intLvl As Long = 0
            Dim ListLvl As New ArrayList

            Select Case intType
                Case 2
                    intLvl = dt_view.Rows(0)("price_disp_lvl")
                Case 3
                    intLvl = dt_view.Rows(0)("add_information_lvl")
                Case 4
                    intLvl = dt_view.Rows(0)("use_function_lvl")
            End Select

            For inti As Integer = intLevel.Length - 1 To 0 Step -1
                If intLvl >= intLevel(inti) Then
                    ListLvl.Add(intLevel(inti))
                    intLvl -= intLevel(inti)
                End If
            Next

            For intColumn As Integer = 0 To 7
                Dim chkBoxTmp As New CheckBox

                chkBoxTmp = CType(Me.FindControl("Label" & intType & intColumn), CheckBox)

                If chkBoxTmp IsNot Nothing Then
                    If ListLvl.Contains(intLevel(intColumn)) Then
                        chkBoxTmp.Checked = True
                        chkBoxTmp.BackColor = ColorTranslator.FromHtml("#CACAFF")
                        chkBoxTmp.ForeColor = Color.Red
                    Else
                        chkBoxTmp.Checked = False
                        chkBoxTmp.BackColor = ColorTranslator.FromHtml("#C7EDCC")
                        chkBoxTmp.ForeColor = Color.Black
                    End If
                End If
            Next

        Next
    End Sub

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UpdateCheck() As Boolean
        UpdateCheck = False
        Dim bolCheck As Boolean = True
        Try
            For inti As Integer = 3 To 15
                Dim txt As Object = Me.FindControl("txtEdit" & inti)
                If Not txt Is Nothing Then txt.BackColor = Drawing.Color.FromArgb(255, 255, 192)
            Next
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    '必須入力チェック
                    For inti As Integer = 3 To 6
                        Select Case inti
                            Case 3
                                Dim drp As DropDownList = Me.FindControl("txtEdit" & inti)
                                If drp.Text.Trim.Length <= 0 Then
                                    drp.BackColor = Color.Red
                                    bolCheck = False
                                End If
                            Case Else
                                Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                If Not txt Is Nothing Then
                                    If txt.Text.Trim.Length <= 0 Then
                                        txt.BackColor = Color.Red
                                        bolCheck = False
                                    End If
                                End If
                        End Select
                    Next
                    '入力内容チェック
                    For inti As Integer = 3 To 6
                        Select Case inti
                            Case 5, 6
                                Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                If Not IsDate(txt.Text) Then
                                    txt.BackColor = Color.Red
                                    bolCheck = False
                                End If
                            Case 4
                                Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                If Not IsHankaku(txt.Text) Then
                                    txt.BackColor = Color.Red
                                    bolCheck = False
                                End If
                        End Select
                    Next
                    If Not bolCheck Then Exit Function
                Case "11" '11 掛率マスタメンテナンス
                    '必須入力チェック
                    For inti As Integer = 5 To 12
                        Select Case inti
                            Case 5, 6 '国コード
                                If Not Me.FindControl("txtEdit" & inti) Is Nothing Then
                                    Dim drp As DropDownList = Me.FindControl("txtEdit" & inti)
                                    If drp Is Nothing Then Continue For
                                    If drp.Text.Trim.Length <= 0 Then
                                        drp.BackColor = Color.Red
                                        bolCheck = False
                                    End If
                                End If
                            Case Else
                                If Not Me.FindControl("txtEdit" & inti) Is Nothing Then
                                    Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                    If txt Is Nothing Then Continue For
                                    If txt.Text.Trim.Length <= 0 Then
                                        txt.BackColor = Color.Red
                                        bolCheck = False
                                    End If
                                End If
                        End Select
                    Next
                    '入力内容チェック
                    For inti As Integer = 7 To 12
                        Select Case inti
                            Case 11, 12
                                If Me.FindControl("txtEdit" & inti) Is Nothing Then Continue For
                                Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                If Not IsDate(txt.Text) Then
                                    txt.BackColor = Color.Red
                                    bolCheck = False
                                End If
                            Case 7
                                Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                If Not IsHankaku(txt.Text) Then
                                    txt.BackColor = Color.Red
                                    bolCheck = False
                                End If
                            Case 8, 9, 10
                                If Me.FindControl("txtEdit" & inti) Is Nothing Then Continue For
                                Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                If Not IsNumeric(txt.Text) Then
                                    txt.BackColor = Color.Red
                                    bolCheck = False
                                End If
                        End Select
                    Next
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    '必須入力チェック
                    For inti As Integer = 5 To 15
                        Select Case inti
                            Case 9, 10, 15 '国コードとユーザー種別
                                If inti = 10 Or inti = 14 Then Continue For
                                Dim drp As DropDownList = Me.FindControl("txtEdit" & inti)
                                If drp.Text.Trim.Length <= 0 Then
                                    drp.BackColor = Color.Red
                                    bolCheck = False
                                End If
                            Case Else
                                Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                If Not txt Is Nothing Then
                                    Select Case inti
                                        Case 5, 6, 7, 13 'ユーザーID、発効日、失効日とパスワード
                                            If txt.Text.Trim.Length <= 0 Then
                                                txt.BackColor = Color.Red
                                                bolCheck = False
                                            End If
                                    End Select
                                End If
                        End Select
                    Next
                    '入力内容チェック
                    For inti As Integer = 5 To 15
                        Select Case inti
                            Case 6, 7
                                Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                If Not IsDate(txt.Text) Then
                                    txt.BackColor = Color.Red
                                    bolCheck = False
                                End If
                            Case 5, 11, 12, 13
                                Dim txt As TextBox = Me.FindControl("txtEdit" & inti)
                                If Not IsHankaku(txt.Text) Then
                                    txt.BackColor = Color.Red
                                    bolCheck = False
                                End If
                        End Select
                    Next
                    '端末認証情報のチェック
                    If pnlWebLog.Visible Then
                        '@@@パスワードのチェック
                        Dim strPassword = txtPassword.Text.Trim

                        '長さチェック
                        If strPassword.Length = 0 OrElse strPassword.Length > 10 Then
                            txtPassword.BackColor = Color.Red
                            bolCheck = False
                        End If
                        '入力文字列のチェック
                        If Not ClsCommon.fncAlpNumChk(strPassword) Then
                            txtPassword.BackColor = Color.Red
                            bolCheck = False
                        End If

                        '@@@マックアドレスのチェック
                        Dim strMac = txtMacAddress.Text.Trim

                        '長さチェック
                        If strMac.Length > 20 Then
                            txtMacAddress.BackColor = Color.Red
                            bolCheck = False
                        End If
                        '入力文字列のチェック
                        Dim strKigo(1) As String
                        strKigo(0) = ":"
                        If Not ClsCommon.fncAlpNumChk(strMac, strKigo) Then
                            txtMacAddress.BackColor = Color.Red
                            bolCheck = False
                        End If

                        '@@@シリアルNoのチェック
                        Dim strSerial = txtSerial.Text.Trim

                        '長さチェック
                        If strSerial.Length = 0 Then
                            'ユーザー種別が国内以外(16以下)
                            Dim drpUserSerial As DropDownList = CType(FindControl("txtEdit15"), DropDownList)

                            If CType(drpUserSerial.SelectedValue, Integer) <= 17 Then
                                txtSerial.BackColor = Color.Red
                                bolCheck = False
                            End If
                        Else
                            If strSerial.Length > 7 Then
                                txtSerial.BackColor = Color.Red
                                bolCheck = False
                            End If
                        End If
                    End If
                Case "15" '15 マスタメンテナンス
            End Select
            If bolCheck Then UpdateCheck = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' 画面より登録や更新データを取得する
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <param name="intMode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreatUpdateData(ByRef dr As DataRow, Optional intMode As Integer = 0) As Boolean
        CreatUpdateData = False
        Try
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    Dim objtxt As Object = Nothing
                    For inti As Integer = 3 To 6
                        Select Case inti
                            Case 3
                                objtxt = CType(Me.FindControl("txtEdit" & inti), DropDownList)
                                If Not objtxt Is Nothing Then dr(1) = CType(objtxt, DropDownList).SelectedValue
                            Case 4
                                objtxt = CType(Me.FindControl("txtEdit" & inti), TextBox)
                                If Not objtxt Is Nothing Then dr(0) = objtxt.text.ToString.Trim.ToUpper
                            Case Else
                                objtxt = CType(Me.FindControl("txtEdit" & inti), TextBox)
                                If Not objtxt Is Nothing Then dr(inti - 2) = objtxt.text
                        End Select
                    Next
                    If intMode = 0 Then '新規
                        dr("register_person") = objUserInfo.UserId
                        dr("register_datetime") = Now
                    Else                '更新
                        dr("current_person") = objUserInfo.UserId
                        dr("current_datetime") = Now
                    End If
                Case "11" '11 掛率マスタメンテナンス
                    Dim objtxt As Object = Nothing
                    For inti As Integer = 5 To 12
                        If HidRateDiv.Value = "1" Then  '購入価格
                            Select Case inti
                                Case 10
                                    objtxt = CType(Me.FindControl("txtEdit" & inti), TextBox)
                                    If Not CheckRateFormat(objtxt.text.ToString.Trim) Then
                                        AlertMessage("W5060")
                                        Exit Function
                                    Else
                                        If Not objtxt Is Nothing Then dr("fob_rate") = objtxt.text.ToString.Trim
                                    End If

                            End Select
                        Else
                            Select Case inti
                                Case 8
                                    objtxt = CType(Me.FindControl("txtEdit" & inti), TextBox)
                                    If Not CheckRateFormat(objtxt.text.ToString.Trim) Then
                                        AlertMessage("W5060")
                                        Exit Function
                                    Else
                                        If Not objtxt Is Nothing Then dr("list_price_rate1") = objtxt.text.ToString.Trim
                                    End If

                                Case 9
                                    objtxt = CType(Me.FindControl("txtEdit" & inti), TextBox)
                                    If Not CheckRateFormat(objtxt.text.ToString.Trim) Then
                                        AlertMessage("W5060")
                                        Exit Function
                                    Else
                                        If Not objtxt Is Nothing Then dr("list_price_rate2") = objtxt.text.ToString.Trim
                                    End If

                            End Select
                        End If
                    Next
                    objtxt = CType(Me.FindControl("txtEdit11"), TextBox)
                    If Not objtxt Is Nothing Then dr("in_effective_date") = objtxt.text.ToString.Trim.ToUpper
                    objtxt = CType(Me.FindControl("txtEdit12"), TextBox)
                    If Not objtxt Is Nothing Then dr("out_effective_date") = objtxt.text.ToString.Trim.ToUpper

                    If intMode = 0 Then '新規
                        dr("register_person") = objUserInfo.UserId
                        dr("register_datetime") = Now
                    Else                '更新
                        dr("current_person") = objUserInfo.UserId
                        dr("current_datetime") = Now
                    End If
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    Dim objtxt As Object = Nothing
                    'ユーザー種別
                    Dim strUserClass As String = String.Empty
                    '国ベースコード
                    Dim strBaseCode As String = String.Empty

                    For inti As Integer = 6 To 15
                        Select Case inti
                            Case 9, 10
                                objtxt = CType(Me.FindControl("txtEdit" & inti), DropDownList)
                                If Not objtxt Is Nothing Then
                                    dr(inti - 4) = objtxt.text
                                    '国コードの場合はベースコードを取得
                                    If inti = 9 Then
                                        strBaseCode = MasterBLL.fncGetBaseCdByCountryCd(objConBase, objtxt.text)

                                    End If
                                End If
                            Case 14
                                '3か月後にパスワード期限切れ
                                dr(inti - 4) = Now.AddMonths(3).ToString("yyyy/MM/dd")
                            Case 15
                                objtxt = CType(Me.FindControl("txtEdit" & inti), DropDownList)
                                If Not objtxt Is Nothing Then
                                    dr(inti - 4) = objtxt.SelectedValue
                                    strUserClass = objtxt.SelectedValue
                                End If

                            Case Else
                                objtxt = CType(Me.FindControl("txtEdit" & inti), TextBox)
                                If Not objtxt Is Nothing Then dr(inti - 4) = objtxt.text
                        End Select
                    Next

                    Dim chk As CheckBox = Nothing
                    Dim intSaveLevel As Long = 0
                    '画面入力されていない場合はユーザ種別により権限を登録
                    Dim intUseDefaultLevel As Long = 0

                    For intj As Integer = 2 To 4
                        intSaveLevel = 0
                        For inti As Integer = 0 To 9
                            If Not Me.FindControl("Label" & intj.ToString & inti.ToString) Is Nothing Then
                                chk = CType(Me.FindControl("Label" & intj.ToString & inti.ToString), CheckBox)
                                If chk.Checked = True Then intSaveLevel += intLevel(inti)
                            End If
                        Next
                        Select Case intj
                            Case 2
                                dr("price_disp_lvl") = intSaveLevel
                            Case 3
                                dr("add_information_lvl") = intSaveLevel
                            Case 4
                                dr("use_function_lvl") = intSaveLevel
                        End Select

                        intUseDefaultLevel += intSaveLevel
                    Next
                    '画面入力されていない場合はユーザ種別により権限を登録
                    If (intUseDefaultLevel = 0) AndAlso (Not strUserClass.Equals(String.Empty)) Then
                        'ユーザー種別により権限の取得
                        Dim dtAuthority As New DS_Master.kh_authority_mstDataTable

                        Using da_authority As New DS_MasterTableAdapters.kh_authority_mstTableAdapter
                            dtAuthority = da_authority.GetDataByUserClass(strBaseCode, strUserClass)
                        End Using

                        If dtAuthority.Rows.Count > 0 Then
                            dr("price_disp_lvl") = dtAuthority.Rows(0).Item("price_disp_lvl")
                            dr("add_information_lvl") = dtAuthority.Rows(0).Item("add_information_lvl")
                            dr("use_function_lvl") = dtAuthority.Rows(0).Item("use_function_lvl")
                        End If
                    End If

                    dr("fraction_proc_div") = ""
                    If intMode = 0 Then '新規
                        dr("register_person") = objUserInfo.UserId
                        dr("register_datetime") = Now
                    Else                '更新
                        dr("current_person") = objUserInfo.UserId
                        dr("current_datetime") = Now
                    End If
                Case "15" '15 マスタメンテナンス
            End Select
            CreatUpdateData = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' 更新ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnEdit(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim strKey() As String = Me.HidTableKey.Value.Split(",")
            If Not UpdateCheck() Then Exit Sub
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    If Not strKey Is Nothing AndAlso strKey.Length = 3 Then
                        If strKey(0).Length > 0 AndAlso strKey(1).Length > 0 AndAlso strKey(2).Length > 0 Then
                            Dim dt As New DS_Master.kh_country_item_mstDataTable
                            Using da As New DS_MasterTableAdapters.kh_country_item_mstTableAdapter
                                da.Fill(dt, strKey(1), strKey(0), strKey(2))
                                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                                    Dim dr As DataRow = dt.Rows(0)
                                    If CreatUpdateData(dr, 1) Then da.Update(dt)
                                End If
                            End Using
                            Dim txtS_CountryID As DropDownList = Me.pnlSelect.FindControl("txtS_CountryID")
                            If Not txtS_CountryID Is Nothing Then txtS_CountryID.SelectedValue = strKey(0)
                            Dim txtS_Kataban As TextBox = Me.pnlSelect.FindControl("txtS_Kataban")
                            If Not txtS_Kataban Is Nothing Then txtS_Kataban.Text = strKey(1)
                        End If
                    End If
                Case "11" '11 掛率マスタメンテナンス
                    If HidRateDiv.Value = "1" Then  '購入価格
                        If Not strKey Is Nothing AndAlso strKey.Length = 4 Then
                            If strKey(0).Length > 0 AndAlso strKey(1).Length > 0 AndAlso strKey(2).Length > 0 AndAlso strKey(3).Length > 0 Then
                                Dim dt As New DS_Master.kh_country_rate_netprice_mstDataTable
                                Using da As New DS_MasterTableAdapters.kh_country_rate_netprice_mstTableAdapter
                                    da.Fill(dt, strKey(0), strKey(1), strKey(2), strKey(3))
                                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                                        Dim dr As DataRow = dt.Rows(0)
                                        If CreatUpdateData(dr, 1) Then da.Update(dt)
                                    End If
                                End Using
                            End If
                        End If
                    Else
                        If Not strKey Is Nothing AndAlso strKey.Length = 3 Then
                            If strKey(0).Length > 0 AndAlso strKey(1).Length > 0 AndAlso strKey(2).Length > 0 Then
                                Dim dt As New DS_Master.kh_country_rate_localprice_mstDataTable
                                Using da As New DS_MasterTableAdapters.kh_country_rate_localprice_mstTableAdapter
                                    da.Fill(dt, strKey(0), strKey(1), strKey(2))
                                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                                        Dim dr As DataRow = dt.Rows(0)
                                        If CreatUpdateData(dr, 1) Then da.Update(dt)
                                    End If
                                End Using
                            End If
                        End If
                    End If
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    If Not strKey Is Nothing AndAlso strKey.Length = 2 Then
                        If strKey(0).Length > 0 AndAlso strKey(1).Length > 0 Then
                            Dim dt As New DS_Master.kh_user_mstDataTable
                            'ユーザーID
                            Dim strUserID As String = strKey(0)
                            'SeqNo
                            Dim strSeqNo As String = strKey(1)

                            'マスタの更新
                            Using da As New DS_MasterTableAdapters.kh_user_mstTableAdapter
                                da.Fill(dt, strKey(0), strKey(1))
                                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                                    Dim dr As DataRow = dt.Rows(0)
                                    If CreatUpdateData(dr, 1) Then da.Update(dt)
                                End If
                            End Using

                            '端末認証情報の更新
                            If pnlWebLog.Visible Then
                                '入力情報の取得
                                Dim strPass As String = txtPassword.Text
                                Dim strMac As String = txtMacAddress.Text
                                Dim strSerial As String = txtSerial.Text
                                'Dim dLastUsedTime As Date = CType(txtLastUsedTime.Text, Date)

                                '情報の保存
                                Dim dtUser As New DS_M_User.M_UserDataTable
                                Using daUser As New DS_M_UserTableAdapters.M_UserTableAdapter
                                    Dim intNum As Integer = 0
                                    intNum = daUser.GetNumByUserID(strUserID)
                                    If intNum > 0 Then
                                        '更新
                                        daUser.UpdateByUserID(strPass, strMac, strSerial, strUserID)
                                    Else
                                        '追加
                                        daUser.Insert(strUserID, strPass, strMac, strSerial)
                                    End If
                                End Using
                            End If

                            Dim txtS_UserID As TextBox = Me.pnlSelect.FindControl("txtS_UserID")
                            If Not txtS_UserID Is Nothing Then txtS_UserID.Text = strKey(0)
                        End If
                        AspNetPager1.RecordCount = 0
                    End If

                Case "15" '15 マスタメンテナンス
            End Select
            AlertMessage("I5130")
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 削除ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnDelete(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim strKey() As String = Me.HidTableKey.Value.Split(",")
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    If Not strKey Is Nothing AndAlso strKey.Length = 3 Then
                        If strKey(0).Length > 0 AndAlso strKey(1).Length > 0 AndAlso strKey(2).Length > 0 Then
                            Dim dt As New DS_Master.kh_country_item_mstDataTable
                            Using da As New DS_MasterTableAdapters.kh_country_item_mstTableAdapter
                                da.Fill(dt, strKey(1).ToString, strKey(0).ToString, CLng(strKey(2)))
                                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                                    dt.Rows(0).Delete()
                                    da.Update(dt)
                                End If
                            End Using
                        End If
                    End If
                Case "11" '11 掛率マスタメンテナンス
                    If HidRateDiv.Value = "1" Then  '購入価格
                        If Not strKey Is Nothing AndAlso strKey.Length = 4 Then
                            If strKey(0).Length > 0 AndAlso strKey(1).Length > 0 AndAlso strKey(2).Length > 0 AndAlso strKey(3).Length > 0 Then
                                Dim dt As New DS_Master.kh_country_rate_netprice_mstDataTable
                                Using da As New DS_MasterTableAdapters.kh_country_rate_netprice_mstTableAdapter
                                    da.Fill(dt, strKey(0), strKey(1), strKey(2), strKey(3))
                                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                                        dt.Rows(0).Delete()
                                        da.Update(dt)
                                    End If
                                End Using
                            End If
                        End If
                    Else
                        If Not strKey Is Nothing AndAlso strKey.Length = 3 Then
                            If strKey(0).Length > 0 AndAlso strKey(1).Length > 0 AndAlso strKey(2).Length > 0 Then
                                Dim dt As New DS_Master.kh_country_rate_localprice_mstDataTable
                                Using da As New DS_MasterTableAdapters.kh_country_rate_localprice_mstTableAdapter
                                    da.Fill(dt, strKey(0), strKey(1), strKey(2))
                                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                                        dt.Rows(0).Delete()
                                        da.Update(dt)
                                    End If
                                End Using
                            End If
                        End If
                    End If
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    If Not strKey Is Nothing AndAlso strKey.Length = 2 Then
                        If strKey(0).Length > 0 AndAlso strKey(1).Length > 0 Then
                            'ユーザマスタの削除
                            Dim dt As New DS_Master.kh_user_mstDataTable
                            Using da As New DS_MasterTableAdapters.kh_user_mstTableAdapter
                                da.Fill(dt, strKey(0).ToString, CLng(strKey(1)))
                                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                                    dt.Rows(0).Delete()
                                    da.Update(dt)
                                End If
                            End Using

                            '端末認証情報の削除
                            Dim dtMUser As New DS_M_User.M_UserDataTable
                            Using daMUser As New DS_M_UserTableAdapters.M_UserTableAdapter
                                dtMUser = daMUser.GetDataByUserID(strKey(0))
                                If dtMUser.Rows.Count > 0 Then
                                    daMUser.Delete(strKey(0))
                                End If
                            End Using
                        End If
                    End If
                Case "15" '15 マスタメンテナンス
            End Select
            AlertMessage("I5140")
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 検索イベント
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSearch()
        Try
            Me.HidSelID.Value = String.Empty  'クリアにする
            '一覧照会処理
            Dim dt_view As New DataTable

            '列名の設定
            dt_view = SetGridViewColumn()

            '色の設定
            Call SetInputColor()

            Select Case strMode
                Case "10" '10 ユーザーマスタメンテナンス
                    Dim txtS_UserID As TextBox = Me.pnlSelect.FindControl("txtS_UserID")
                    Dim txtS_StdDate As TextBox = Me.pnlSelect.FindControl("txtS_StdDate")
                    If Not txtS_UserID Is Nothing AndAlso Not txtS_StdDate Is Nothing Then
                        If txtS_UserID.Text.Trim.Length <= 0 AndAlso txtS_StdDate.Text.Trim.Length <= 0 Then
                            AlertMessage("W0070")
                            txtS_UserID.Focus()
                            Exit Sub
                        End If
                    End If
                Case "11" '11 掛率マスタメンテナンス
                    If HidRateDiv.Value = "1" Then  '購入価格
                        Dim txtS_Made As DropDownList = Me.pnlSelect.FindControl("txtS_Made")
                        Dim txtS_Sale As DropDownList = Me.pnlSelect.FindControl("txtS_Sale")
                        Dim txtS_Search As TextBox = Me.pnlSelect.FindControl("txtS_Search")
                        If Not txtS_Made Is Nothing AndAlso Not txtS_Sale Is Nothing AndAlso Not txtS_Search Is Nothing Then
                            If txtS_Made.Text.Trim.Length <= 0 AndAlso txtS_Sale.Text.Trim.Length <= 0 AndAlso _
                                txtS_Search.Text.Trim.Length <= 0 Then
                                AlertMessage("W0070")
                                txtS_Search.Focus()
                                Exit Sub
                            End If
                        End If
                    Else '現地定価
                        Dim txtS_Sale As DropDownList = Me.pnlSelect.FindControl("txtS_Sale")
                        Dim txtS_Search As TextBox = Me.pnlSelect.FindControl("txtS_Search")
                        If Not txtS_Sale Is Nothing AndAlso Not txtS_Search Is Nothing Then
                            If txtS_Sale.Text.Trim.Length <= 0 AndAlso txtS_Search.Text.Trim.Length <= 0 Then
                                AlertMessage("W0070")
                                txtS_Search.Focus()
                                Exit Sub
                            End If
                        End If
                    End If
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "14" '14 国別生産品マスタメンテナンス
                    Dim txtS_CountryID As DropDownList = Me.pnlSelect.FindControl("txtS_CountryID")
                    Dim txtS_Kataban As TextBox = Me.pnlSelect.FindControl("txtS_Kataban")
                    If Not txtS_CountryID Is Nothing AndAlso Not txtS_CountryID Is Nothing Then
                        If txtS_CountryID.Text.Trim.Length <= 0 OrElse txtS_Kataban.Text.Trim.Length <= 0 Then
                            AlertMessage("W0070")
                            txtS_Kataban.Focus()
                            Exit Sub
                        End If
                    End If
                Case "15" '15 マスタメンテナンス
            End Select
            If Not dt_view Is Nothing Then
                'タイトルの表示
                Dim dtTitle As DataTable = dt_view.Clone
                Dim drTitle As DataRow = dtTitle.NewRow
                dtTitle.Rows.Add(drTitle)

                CreateTitle()

                '情報の表示
                Dim dt As DataTable = dt_view.Clone
                Dim dr As DataRow = dt.NewRow
                dt.Rows.Add(dr)
                Me.GridViewMain.DataSource = dt
                Me.GridViewMain.DataBind()

                Call ClearInput()

                'ボタンの設定
                CType(Me.FindControl("button5"), Button).Enabled = True
                CType(Me.FindControl("button6"), Button).Enabled = False
                CType(Me.FindControl("button7"), Button).Enabled = False
            End If
        Catch ex As Exception
            'ｴﾗｰ画面に遷移
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 検索ボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnSearch(ByVal sender As Object, ByVal e As System.EventArgs)

        Select Case strMode
            Case "10"
                'ページング設定
                AspNetPager1.CurrentPageIndex = 1
        End Select

        subSearch()
    End Sub

    ''' <summary>
    ''' 入力をクリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearInput()
        Try
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    For inti As Integer = 3 To 6
                        Dim obj As Object = Me.FindControl("txtEdit" & inti.ToString)
                        If Not obj Is Nothing Then
                            Select Case inti
                                Case 3  '国コード
                                    CType(obj, DropDownList).SelectedIndex = -1
                                Case 5  '発効日
                                    CType(obj, TextBox).Text = Now.ToString("yyyy/MM/dd")
                                Case 6  '失効日
                                    CType(obj, TextBox).Text = "9999/12/31"
                                Case Else
                                    CType(obj, TextBox).Text = String.Empty
                            End Select
                        End If
                    Next
                Case "11" '11 掛率マスタメンテナンス
                    For inti As Integer = 5 To 12
                        Dim obj As Object = Me.FindControl("txtEdit" & inti.ToString)
                        If Not obj Is Nothing Then
                            Select Case inti
                                Case 5, 6  '国コード
                                    CType(obj, DropDownList).SelectedIndex = -1
                                Case 11  '発効日
                                    CType(obj, TextBox).Text = Now.ToString("yyyy/MM/dd")
                                Case 12  '失効日
                                    CType(obj, TextBox).Text = "9999/12/31"
                                Case Else
                                    CType(obj, TextBox).Text = String.Empty
                            End Select
                        End If
                    Next
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    For inti As Integer = 5 To 15
                        Dim obj As Object = Me.FindControl("txtEdit" & inti.ToString)
                        If Not obj Is Nothing Then
                            Select Case inti
                                Case 6  '発効日
                                    CType(obj, TextBox).Text = Now.ToString("yyyy/MM/dd")
                                Case 7  '失効日
                                    CType(obj, TextBox).Text = "9999/12/31"
                                Case 9, 10, 15  '国コード DropDown '営業所 DropDown 'ユーザー種別 DropDown
                                    CType(obj, DropDownList).SelectedIndex = -1
                                Case 14 'パスワード有効期限（非表示）
                                    Continue For
                                Case Else
                                    CType(obj, TextBox).Text = String.Empty
                            End Select
                        End If
                    Next

                    For intj As Integer = 0 To 7
                        For inti As Integer = 2 To 4
                            If Not Me.FindControl("Label" & inti.ToString & intj.ToString) Is Nothing Then
                                CType(Me.FindControl("Label" & inti.ToString & intj.ToString), CheckBox).Checked = False
                            End If
                        Next
                    Next
                    '権限選択肢を無効にする
                    pnlEditInput1.Enabled = False
                    pnlEditInput2.Enabled = False
                    pnlEditInput3.Enabled = False

                    '端末認証情報のクリア
                    txtPassword.Text = String.Empty
                    txtMacAddress.Text = String.Empty
                    txtSerial.Text = String.Empty
                    txtPassword.BackColor = ColorTranslator.FromHtml("#FFFFC0")
                    txtMacAddress.BackColor = ColorTranslator.FromHtml("#FFFFC0")
                    txtSerial.BackColor = ColorTranslator.FromHtml("#FFFFC0")

                    'ボタンのリセット
                    CType(Me.FindControl("button5"), Button).Enabled = True
                    CType(Me.FindControl("button6"), Button).Enabled = False
                    CType(Me.FindControl("button7"), Button).Enabled = False
                    'txtLastUsedTime.Text = String.Empty
                Case "15" '15 マスタメンテナンス
            End Select
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' クリアボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnClear(sender As Object, e As EventArgs)
        Me.HidSelID.Value = String.Empty  'クリアにする
        Me.HidTableKey.Value = String.Empty
        Try
            Call SetInputColor()
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    If Not Me.pnlSelect.FindControl("txtS_CountryID") Is Nothing Then
                        Dim txtS_CountryID As DropDownList = Me.pnlSelect.FindControl("txtS_CountryID")
                        txtS_CountryID.SelectedIndex = -1
                    End If
                    If Not Me.pnlSelect.FindControl("txtS_Kataban") Is Nothing Then
                        Dim txtS_Kataban As TextBox = Me.pnlSelect.FindControl("txtS_Kataban")
                        txtS_Kataban.Text = String.Empty
                    End If
                Case "11" '11 掛率マスタメンテナンス
                    If Not Me.pnlSelect.FindControl("txtS_Made") Is Nothing Then
                        Dim txtS_Made As DropDownList = Me.pnlSelect.FindControl("txtS_Made")
                        txtS_Made.SelectedIndex = -1
                    End If
                    If Not Me.pnlSelect.FindControl("txtS_Sale") Is Nothing Then
                        Dim txtS_Sale As DropDownList = Me.pnlSelect.FindControl("txtS_Sale")
                        txtS_Sale.SelectedIndex = -1
                    End If
                    If Not Me.pnlSelect.FindControl("txtS_Search") Is Nothing Then
                        Dim txtS_Search As TextBox = Me.pnlSelect.FindControl("txtS_Search")
                        txtS_Search.Text = String.Empty
                    End If
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    If Not Me.pnlSelect.FindControl("txtS_UserID") Is Nothing Then
                        Dim txtS_UserID As TextBox = Me.pnlSelect.FindControl("txtS_UserID")
                        txtS_UserID.Text = String.Empty
                    End If
                    If Not Me.pnlSelect.FindControl("txtS_StdDate") Is Nothing Then
                        Dim txtS_StdDate As TextBox = Me.pnlSelect.FindControl("txtS_StdDate")
                        txtS_StdDate.Text = String.Empty
                    End If
                Case "15" '15 マスタメンテナンス
            End Select
            Call ClearInput()
            AspNetPager1.RecordCount = 0
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 表示項目の設定
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetGridViewColumn() As DataTable
        Me.GridViewMain.Columns.Clear()
        Dim strBound() As String = Nothing
        Dim strWidth() As String = Nothing
        Dim col As Web.UI.WebControls.BoundField = Nothing
        Dim dr() As DataRow = Nothing
        Dim dc As DataColumn = Nothing
        SetGridViewColumn = New DataTable
        Try
            strKey = String.Empty
            Select Case strMode
                Case "10" '10 ユーザーマスタメンテナンス
                    strKey = "user_id,in_effective_date,out_effective_date,user_nm,country_cd,"
                    strKey &= "office_cd,person_cd,mail_address,password,password_exp_date,user_class"
                Case "11" '11 掛率マスタメンテナンス
                    If HidRateDiv.Value = "1" Then  '購入価格
                        strKey = "exp_country_cd,imp_country_cd,rate_search_key,,,fob_rate,in_effective_date,out_effective_date"
                    Else '現地定価
                        strKey = ",country_cd,rate_search_key,list_price_rate1,list_price_rate2,,in_effective_date,out_effective_date"
                    End If
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "14" '14 国別生産品マスタメンテナンス
                    strKey = "country_cd,kataban,in_effective_date,out_effective_date"
                Case "15" '15 マスタメンテナンス
            End Select
            strBound = strKey.Split(",")
            For inti As Integer = 0 To strBound.Length - 1
                If strBound(inti).Length > 0 Then
                    col = New Web.UI.WebControls.BoundField
                    col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                    GridViewMain.Columns.Add(col)
                    dc = New DataColumn
                    SetGridViewColumn.Columns.Add(dc)
                End If
            Next
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' 属性の設定
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Private Sub SetAttributes(ByRef obj As Object, Optional intMode As Integer = 0)
        Select Case obj.GetType.Name.ToUpper
            Case "LABEL"
                If intMode <> 1 Then
                    obj.Style.Add("background-color", "#008000")
                    obj.Style.Add("color", "#FFFFFF")
                End If
                If intMode = 2 Then
                    'obj.Style.Add("font-size", "11pt")
                    obj.Style.Add("font-size", "10pt")
                    obj.Style.Add("border-style", "solid")
                    obj.Style.Add("border-width", "1px")
                    obj.Style.Add("border-color", "gray")
                    obj.Style.Add("text-align", "center")
                    obj.Style.Add("font-weight", "bold")
                Else
                    obj.Style.Add("font-size", "14pt")
                    obj.Style.Add("border-style", "none")
                    obj.Style.Add("border-width", "0px")
                    obj.Style.Add("text-align", "left")
                End If
                obj.Style.Add("padding-top", "2px")
                obj.Style.Add("padding-left", "2px")
                obj.Style.Add("vertical-align", "middle")
            Case "TEXTBOX"
                Select Case intMode
                    Case 2
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("border-style", "solid")
                        obj.Style.Add("border-width", "1px")
                        obj.Style.Add("border-color", "gray")
                        obj.Style.Add("text-align", "center")
                    Case 3
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("border-style", "solid")
                        obj.Style.Add("border-width", "1px")
                        obj.Style.Add("border-color", "gray")
                        obj.Style.Add("text-align", "left")
                        obj.Style.Add("text-transform", "uppercase")
                    Case Else
                        obj.Style.Add("background-color", "#FFFFCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("font-size", "14pt")
                        obj.Style.Add("border-style", "none")
                        obj.Style.Add("border-width", "0px")
                        obj.Style.Add("text-align", "left")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("text-transform", "uppercase")
                End Select
                obj.Style.Add("vertical-align", "bottom")
            Case "BUTTON"
            Case "DROPDOWNLIST"
                Select Case intMode
                    Case 2
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("border-style", "solid")
                        obj.Style.Add("border-width", "1px")
                        obj.Style.Add("border-color", "gray")
                        obj.Style.Add("text-align", "center")
                    Case 3
                        obj.Style.Add("background-color", "#FFFFCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("font-size", "14pt")
                        obj.Style.Add("border-style", "none")
                        obj.Style.Add("border-width", "0px")
                        obj.Style.Add("text-align", "left")
                    Case Else
                        obj.Style.Add("background-color", "#FFFFCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("font-size", "14pt")
                        obj.Style.Add("border-style", "none")
                        obj.Style.Add("border-width", "0px")
                        obj.Style.Add("text-align", "left")
                        obj.Style.Add("font-weight", "bold")
                End Select
                obj.Style.Add("vertical-align", "bottom")
            Case "CHECKBOX"
                obj.Style.Add("font-size", "10pt")
                obj.Style.Add("border-style", "none")
                obj.Style.Add("border-width", "0px")
                obj.Style.Add("text-align", "left")
                obj.Style.Add("padding-top", "1px")
                obj.Style.Add("padding-left", "1px")
                obj.Style.Add("vertical-align", "middle")
        End Select
    End Sub

    ''' <summary>
    ''' 表示情報の作成
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GridViewMain_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridViewMain.RowDataBound
        If e.Row.RowIndex < 0 Then Exit Sub
        Try
            Dim dt_view As New DataTable
            For inti As Integer = e.Row.Cells.Count - 1 To 1 Step -1
                e.Row.Cells.RemoveAt(inti)
            Next

            Select Case strMode
                Case "10" '10 ユーザーマスタメンテナンス
                    Dim txtS_UserID As TextBox = Me.pnlSelect.FindControl("txtS_UserID")
                    Dim txtS_StdDate As TextBox = Me.pnlSelect.FindControl("txtS_StdDate")

                    'ページング
                    Dim intStartIndex As Integer = AspNetPager1.StartRecordIndex
                    Dim intEndIndex As Integer = AspNetPager1.EndRecordIndex

                    If intStartIndex = 1 And intEndIndex = 0 Then
                        intEndIndex += AspNetPager1.PageSize
                    End If

                    '総数の取得
                    AspNetPager1.RecordCount = MasterBLL.fncSQL_UserMstCount(objConBase, txtS_UserID.Text, txtS_StdDate.Text, _
                                                           Me.selLang.SelectedValue)

                    'リスト作成
                    dt_view = MasterBLL.fncSQL_UserMstList(objConBase, txtS_UserID.Text, txtS_StdDate.Text, _
                                                           Me.selLang.SelectedValue, intStartIndex, intEndIndex)
                Case "11" '11 掛率マスタメンテナンス
                    Dim txtS_Sale As DropDownList = Nothing
                    Dim txtS_Made As DropDownList = Nothing
                    Dim txtS_Search As TextBox = Nothing
                    If Not Me.pnlSelect.FindControl("txtS_Sale") Is Nothing Then
                        txtS_Sale = Me.pnlSelect.FindControl("txtS_Sale")
                    End If
                    If Not Me.pnlSelect.FindControl("txtS_Search") Is Nothing Then
                        txtS_Search = Me.pnlSelect.FindControl("txtS_Search")
                    End If
                    If HidRateDiv.Value = "1" Then  '購入価格
                        If Not Me.pnlSelect.FindControl("txtS_Made") Is Nothing Then
                            txtS_Made = Me.pnlSelect.FindControl("txtS_Made")
                        End If
                        dt_view = MasterBLL.fncSQL_RateMstList_N(objConBase, txtS_Made.Text, txtS_Sale.Text, txtS_Search.Text)
                    Else '現地定価
                        dt_view = MasterBLL.fncSQL_RateMstList_L(objConBase, txtS_Sale.Text, txtS_Search.Text)
                    End If
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "14" '14 国別生産品マスタメンテナンス
                    Dim txtS_CountryID As DropDownList = Me.pnlSelect.FindControl("txtS_CountryID")
                    Dim txtS_Kataban As TextBox = Me.pnlSelect.FindControl("txtS_Kataban")
                    dt_view = MasterBLL.fncSQL_CountryItemMstList(objConBase, txtS_CountryID.SelectedValue, txtS_Kataban.Text)
                Case "15" '15 マスタメンテナンス
            End Select

            If dt_view.Rows.Count < 0 Then Exit Sub
            If dt_view.Rows.Count > My.Settings.MaxDispTnkCnt Then
                AlertMessage("E001", "検索対象件数：" & dt_view.Rows.Count & " 件。" & "MAX件数(" & My.Settings.MaxDispTnkCnt & "件)を超えました、検索条件を確認してください。")
                dt_view = Nothing
                Exit Sub
            End If

            e.Row.Cells(0).HorizontalAlign = HorizontalAlign.Left
            e.Row.Cells(0).ColumnSpan = dt_view.Columns.Count

            Dim GridMid As New WebControls.GridView
            AddHandler GridMid.RowDataBound, AddressOf Child_RowDataBound
            GridMid.AutoGenerateColumns = False
            GridMid.ShowHeader = False
            GridMid.GridLines = GridLines.Both
            GridMid.Font.Size = WebControls.FontUnit.Point(11)
            GridMid.Font.Name = GetFontName(selLang.SelectedValue)
            GridMid.CellPadding = 0
            GridMid.CellSpacing = 0
            e.Row.Cells(0).Controls.Add(GridMid)

            Dim col As New Web.UI.WebControls.BoundField
            Dim strBound() As String = strKey.Split(",")
            Dim intLoop As Integer = 0
            For inti As Integer = 0 To strBound.Length - 1
                If strBound(inti).Length > 0 Then
                    col = New Web.UI.WebControls.BoundField
                    col.DataField = strBound(inti)
                    col.ItemStyle.Wrap = True
                    col.ItemStyle.Height = WebControls.Unit.Pixel(intColWeight)
                    col.ItemStyle.HorizontalAlign = HorizontalAlign.Center
                    GridMid.Columns.Add(col)
                    intLoop += 1
                End If
            Next
            Me.HidTableKey.Value = String.Empty
            Me.HidSelID.Value = String.Empty
            ListLevel = New ArrayList
            ListSeq = New ArrayList
            Dim strSeqNo As String = String.Empty

            Select Case strMode
                Case "10" '10 ユーザーマスタメンテナンス
                    For inti As Integer = 0 To dt_view.Rows.Count - 1
                        ListLevel.Add(dt_view.Rows(inti)("price_disp_lvl") & "," & _
                                      dt_view.Rows(inti)("add_information_lvl") & "," & _
                                      dt_view.Rows(inti)("use_function_lvl"))
                        ListSeq.Add(dt_view.Rows(inti)("seq_no"))
                    Next
                Case "13" '13 為替率マスタメンテナンス
                Case "12" '12 情報マスタメンテナンス
                Case "11", "14" '14 国別生産品マスタメンテナンス'11 掛率マスタメンテナンス
                    For inti As Integer = 0 To dt_view.Rows.Count - 1
                        ListSeq.Add(dt_view.Rows(inti)("seq_no"))
                    Next
                Case "15" '15 マスタメンテナンス
            End Select
            GridMid.DataSource = dt_view
            GridMid.DataBind()
            'End Select
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' データバインド
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Child_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        If e.Row.RowIndex < 0 Then Exit Sub
        Dim lbl As New Label
        Dim pnl As Panel = Nothing
        Dim chk As New CheckBox
        Dim dr() As DataRow = Nothing
        Dim strLvl() As String = Nothing
        Dim strTableKey As String = String.Empty
        Try
            e.Row.Cells(0).HorizontalAlign = HorizontalAlign.Left
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    e.Row.Cells(0).ColumnSpan = strWidth_CountryItem.Length

                    strTableKey = e.Row.Cells(0).Text.Trim & "," & e.Row.Cells(1).Text.Trim & "," & ListSeq(e.Row.RowIndex)
                    pnl = New Panel
                    Dim strBound() As String = strKey.Split(",")
                    For inti As Integer = 0 To strBound.Length - 1
                        If strBound(inti).Length > 0 Then
                            lbl = New Label
                            lbl.ID = "txtEdit" & (inti + 3)
                            lbl.Text = e.Row.Cells(inti).Text
                            If inti = strBound.Length - 1 Then
                                lbl.Width = WebControls.Unit.Pixel(strWidth_CountryItem(inti) - 3)
                            Else
                                lbl.Width = WebControls.Unit.Pixel(strWidth_CountryItem(inti) - 2)
                            End If
                            lbl.Height = WebControls.Unit.Pixel(intColWeight - 4)
                            lbl.Style.Add("font-size", "10pt")
                            lbl.Style.Add("border-style", "solid")
                            lbl.Style.Add("border-width", "1px")
                            lbl.Style.Add("border-color", "gray")
                            If inti = 1 Then
                                lbl.Style.Add("text-align", "left")
                            Else
                                lbl.Style.Add("text-align", "center")
                            End If
                            'lbl.Style.Add("padding-top", "3px")
                            'lbl.Style.Add("padding-left", "1px")
                            lbl.Style.Add("padding", "3px 0px 0px 0px")
                            lbl.Style.Add("vertical-align", "middle")
                            pnl.Controls.Add(lbl)
                        End If
                    Next
                    pnl.ID = "pnlData"
                    If (e.Row.RowIndex + 1) Mod 2 = 0 Then
                        pnl.BackColor = Color.FromArgb(192, 192, 255)
                    Else
                        pnl.BackColor = Color.White
                    End If
                    e.Row.Cells(0).Controls.Add(pnl)
                    e.Row.Cells(0).BackColor = Drawing.Color.FromArgb(255, 255, 192)
                    e.Row.Cells(0).ID = "Cell0"

                    e.Row.Style.Add("background-color", "#C7EDCC")
                    Dim strRowID As String = e.Row.ClientID
                    e.Row.Attributes.Add("onclick", "CountryMasterCellClick('" & Me.ClientID & "','" & strRowID & "','" & strTableKey & "');")
                Case "11" '11 掛率マスタメンテナンス
                    Dim intWidth() As Integer = Nothing
                    If HidRateDiv.Value = "1" Then  '購入価格
                        intWidth = strWidth_RateNet
                        strTableKey = e.Row.Cells(0).Text.Trim & "," & e.Row.Cells(1).Text.Trim & "," & e.Row.Cells(2).Text.Trim & "," & ListSeq(e.Row.RowIndex)
                    Else '現地定価
                        intWidth = strWidth_RateLocal
                        strTableKey = e.Row.Cells(0).Text.Trim & "," & e.Row.Cells(1).Text.Trim & "," & ListSeq(e.Row.RowIndex)
                    End If

                    e.Row.Cells(0).ColumnSpan = intWidth.Length
                    pnl = New Panel
                    Dim strBound() As String = strKey.Split(",")
                    Dim intLoop As Integer = 0
                    For inti As Integer = 0 To strBound.Length - 1
                        If strBound(inti).Length > 0 Then
                            lbl = New Label
                            lbl.ID = "txtEdit" & (inti + 5)
                            lbl.Text = e.Row.Cells(intLoop).Text
                            If inti = strBound.Length - 1 Or intLoop = 0 Then
                                lbl.Width = WebControls.Unit.Pixel(intWidth(intLoop) - 3)
                            Else
                                lbl.Width = WebControls.Unit.Pixel(intWidth(intLoop) - 2)
                            End If
                            lbl.Height = WebControls.Unit.Pixel(intColWeight - 4)
                            lbl.Style.Add("font-size", "10pt")
                            lbl.Style.Add("border-style", "solid")
                            lbl.Style.Add("border-width", "1px")
                            lbl.Style.Add("border-color", "gray")
                            If HidRateDiv.Value = "1" Then  '購入価格
                                If intLoop <= 2 Then
                                    lbl.Style.Add("text-align", "left")
                                Else
                                    lbl.Style.Add("text-align", "center")
                                End If
                            Else '現地定価
                                If intLoop <= 1 Then
                                    lbl.Style.Add("text-align", "left")
                                Else
                                    lbl.Style.Add("text-align", "center")
                                End If
                            End If
                            'lbl.Style.Add("padding-top", "3px")
                            'lbl.Style.Add("padding-left", "1px")
                            lbl.Style.Add("padding", "3px 0px 0px 0px")
                            lbl.Style.Add("vertical-align", "middle")
                            pnl.Controls.Add(lbl)
                            intLoop += 1
                        End If
                    Next
                    pnl.ID = "pnlData"
                    If (e.Row.RowIndex + 1) Mod 2 = 0 Then
                        pnl.BackColor = Color.FromArgb(192, 192, 255)
                    Else
                        pnl.BackColor = Color.White
                    End If
                    e.Row.Cells(0).Controls.Add(pnl)
                    e.Row.Cells(0).BackColor = Drawing.Color.FromArgb(255, 255, 192)
                    e.Row.Cells(0).ID = "Cell0"

                    e.Row.Style.Add("background-color", "#C7EDCC")
                    Dim strRowID As String = e.Row.ClientID
                    e.Row.Attributes.Add("onclick", "CountryMasterCellClick('" & Me.ClientID & "','" & strRowID & "','" & strTableKey & "');")
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    If Not ListLevel(e.Row.RowIndex) Is Nothing AndAlso ListLevel(e.Row.RowIndex).Length > 0 Then
                        strLvl = ListLevel(e.Row.RowIndex).Split(",")
                    End If
                    e.Row.Cells(0).ColumnSpan = strWidth_User.Length

                    If Not strLvl Is Nothing AndAlso strLvl.Length = 3 Then
                        strTableKey = e.Row.Cells(0).Text.Trim & "," & ListSeq(e.Row.RowIndex)
                        pnl = New Panel
                        Dim strBound() As String = strKey.Split(",")
                        For inti As Integer = 0 To strBound.Length - 1
                            If strBound(inti).Length > 0 Then
                                lbl = New Label
                                If inti + 5 >= 14 Then
                                    lbl.ID = "txtEdit" & (inti + 6)
                                Else
                                    lbl.ID = "txtEdit" & (inti + 5)
                                End If

                                lbl.Text = e.Row.Cells(inti).Text
                                If inti = strBound.Length - 1 Then
                                    lbl.Width = WebControls.Unit.Pixel(strWidth_User(inti) - 3)
                                Else
                                    lbl.Width = WebControls.Unit.Pixel(strWidth_User(inti) - 2)
                                End If
                                lbl.Height = WebControls.Unit.Pixel(intColWeight - 4)
                                lbl.Style.Add("font-size", "10pt")
                                lbl.Style.Add("border-style", "solid")
                                lbl.Style.Add("border-width", "1px")
                                lbl.Style.Add("border-color", "gray")
                                lbl.Style.Add("text-align", "center")
                                'lbl.Style.Add("padding-top", "3px")
                                lbl.Style.Add("padding", "3px 0px 0px 0px")
                                lbl.Style.Add("vertical-align", "middle")
                                pnl.Controls.Add(lbl)
                            End If
                        Next
                        pnl.ID = "pnlData"
                        pnl.BackColor = Color.FromArgb(192, 192, 255)
                        e.Row.Cells(0).Controls.Add(pnl)

                        For intj As Integer = 0 To strLvl.Length - 1
                            pnl = New Panel
                            Dim intLvl As Long = strLvl(intj)
                            Select Case intj
                                Case 0 '価格
                                    dr = subSetLbl.Select("len(label_seq) = 2 AND label_seq >='20' AND label_seq <'30' AND label_div='L'")
                                Case 1 '単価画面の項目制御
                                    dr = subSetLbl.Select("len(label_seq) = 2 AND label_seq >='30' AND label_seq <'40' AND label_div='L'")
                                Case 2 'マスタリンクボタンの表示
                                    dr = subSetLbl.Select("len(label_seq) = 2 AND label_seq >='40' AND label_seq <'50' AND label_div='L'")
                            End Select

                            If Not dr Is Nothing Then
                                'lbl = New Label
                                'lbl.ID = "Level" & intj
                                'lbl.Text = "Level:" & intLvl.ToString.PadLeft(3, " ")
                                'lbl.Style.Add("font-size", "11pt")
                                'lbl.Style.Add("border-style", "none")
                                'lbl.Style.Add("text-align", "left")
                                'lbl.Style.Add("font-weight", "bold")
                                'lbl.Style.Add("padding-top", "1px")
                                'lbl.Style.Add("padding-left", "1px")
                                'lbl.Style.Add("vertical-align", "middle")
                                'lbl.Width = WebControls.Unit.Pixel(70)
                                'pnl.Controls.Add(lbl)

                                Dim ListLvl As New ArrayList
                                For inti As Integer = intLevel.Length - 1 To 0 Step -1
                                    If intLvl >= intLevel(inti) Then
                                        ListLvl.Add(intLevel(inti))
                                        intLvl -= intLevel(inti)
                                    End If
                                Next

                                '権限オプションの作成
                                For inti As Integer = 0 To 7
                                    If intj = 2 And inti = 7 Then Exit For
                                    chk = New CheckBox
                                    chk.ID = "Label" & (intj + 2).ToString & inti.ToString
                                    chk.Text = dr(inti)("label_content").ToString
                                    SetAttributes(chk)
                                    chk.Width = WebControls.Unit.Pixel(strWidth_Check(inti))
                                    chk.Height = WebControls.Unit.Pixel(intColWeight)
                                    If ListLvl.Contains(intLevel(inti)) Then
                                        chk.Checked = True
                                    Else
                                        chk.Checked = False
                                    End If
                                    chk.Enabled = False
                                    pnl.Controls.Add(chk)
                                Next

                                '端末認証情報の作成
                                If pnlWebLog.Visible Then
                                    Dim strUserID As String = e.Row.DataItem("user_id")
                                    Call CreateHidWebLogInfo(pnl, e.Row.RowIndex, strUserID)
                                End If
                                e.Row.Cells(0).Controls.Add(pnl)
                            End If
                        Next
                        e.Row.Cells(0).BackColor = Drawing.Color.FromArgb(255, 255, 192)
                        e.Row.Cells(0).ID = "Cell0"
                    End If

                    e.Row.Style.Add("background-color", "#C7EDCC")
                    Dim strRowID As String = e.Row.ClientID
                    e.Row.Attributes.Add("onclick", "UserMasterCellClick('" & Me.ClientID & "','" & strRowID & "','" & strTableKey & "');")
                Case "15" '15 マスタメンテナンス
            End Select
            For inti As Integer = e.Row.Cells.Count - 1 To 1 Step -1
                e.Row.Cells.RemoveAt(inti)
            Next
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' テキストボックスの作成
    ''' </summary>
    ''' <param name="intlblID"></param>
    ''' <param name="steTxtID"></param>
    ''' <param name="intTxtWidth"></param>
    ''' <remarks></remarks>
    Private Sub CreatTextBox(intlblID As Integer, steTxtID As String, intTxtWidth As Integer)
        Dim lbl As Label = Nothing
        Dim txt As TextBox = Nothing
        lbl = New Label
        lbl.ID = "Label" & intlblID
        SetAttributes(lbl)
        Me.pnlSelect.Controls.Add(lbl)
        txt = New TextBox
        txt.ID = steTxtID
        SetAttributes(txt)
        txt.Width = WebControls.Unit.Pixel(intTxtWidth)
        Me.pnlSelect.Controls.Add(txt)
    End Sub

    ''' <summary>
    ''' ドロップダウンの作成
    ''' </summary>
    ''' <param name="intlblID"></param>
    ''' <param name="steDrpID"></param>
    ''' <param name="intDrpWidth"></param>
    ''' <param name="intAtt"></param>
    ''' <remarks></remarks>
    Private Sub CreatDropDown(intlblID As Integer, steDrpID As String, intDrpWidth As Integer, intAtt As Integer)
        Dim lbl As Label = Nothing
        Dim drp As DropDownList = Nothing
        Dim dt As DataTable = Nothing
        lbl = New Label
        lbl.ID = "Label" & intlblID
        SetAttributes(lbl)
        Me.pnlSelect.Controls.Add(lbl)
        drp = New DropDownList
        drp.ID = steDrpID
        dt = New DataTable
        dt = MasterBLL.fncSQL_CountryCodeList(objConBase, selLang.SelectedValue)
        If Not dt Is Nothing Then
            drp.DataTextField = "country_nm"
            drp.DataValueField = "country_cd"
            drp.DataSource = dt
            drp.DataBind()
        End If
        SetAttributes(drp, intAtt)
        drp.Width = WebControls.Unit.Pixel(intDrpWidth)
        Me.pnlSelect.Controls.Add(drp)
    End Sub

    ''' <summary>
    ''' ラジオボタンの作成
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As System.EventArgs) Handles _
        RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        Call btnClear(Nothing, Nothing)
    End Sub

    ''' <summary>
    ''' ユーザーマスタ画面の権限選択可否の設定
    ''' </summary>
    ''' <param name="blnEnabled"></param>
    ''' <remarks></remarks>
    Private Sub subSetCheckBoxEnable(ByVal blnEnabled As Boolean)
        Dim chk As New CheckBox

        For intj As Integer = 2 To 4
            For inti As Integer = 0 To 9
                If Not Me.FindControl("Label" & intj.ToString & inti.ToString) Is Nothing Then
                    chk = CType(Me.FindControl("Label" & intj.ToString & inti.ToString), CheckBox)
                    chk.Enabled = blnEnabled
                End If
            Next
        Next

    End Sub

    ''' <summary>
    ''' タイトルの作成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateTitle()
        Dim strBound() As String = strKey.Split(",")
        Dim dr() As DataRow = Nothing
        Dim intAddCount As Integer = 0
        Dim intWidth() As Integer = Nothing
        Select Case strMode
            Case "14" '14 国別生産品マスタメンテナンス
                intAddCount = 3
                intWidth = strWidth_CountryItem
            Case "11" '11 掛率マスタメンテナンス
                intAddCount = 5
                If HidRateDiv.Value = "1" Then  '購入価格
                    intWidth = strWidth_RateNet
                Else '現地定価
                    intWidth = strWidth_RateLocal
                End If
            Case "12" '12 情報マスタメンテナンス
            Case "13" '13 為替率マスタメンテナンス
            Case "10" '10 ユーザーマスタメンテナンス
                intAddCount = 5
                intWidth = strWidth_User
            Case "15" '15 マスタメンテナンス
        End Select
        Dim intLoop As Integer = 0

        Dim headerRow As TableRow = New TableRow()
        For inti As Integer = 0 To strBound.Length - 1
            If strBound(inti).Length > 0 Then
                dr = subSetLbl.Select("label_seq='" & inti + intAddCount & "' AND label_div='L'")
                If dr.Length > 0 Then
                    Dim headerCell As TableCell = New TableCell()
                    headerCell.Text = dr(0)("label_content").ToString
                    headerCell.Width = WebControls.Unit.Pixel(intWidth(intLoop))
                    If inti = strBound.Length - 1 Then
                        headerCell.Width = WebControls.Unit.Pixel(intWidth(intLoop) - 3)
                    Else
                        headerCell.Width = WebControls.Unit.Pixel(intWidth(intLoop) - 2)
                    End If
                    headerCell.BackColor = Color.Green
                    headerCell.ForeColor = Color.White
                    headerCell.BorderStyle = BorderStyle.Solid
                    headerCell.BorderWidth = WebControls.Unit.Pixel(1)
                    headerCell.Font.Name = GetFontName(selLang.SelectedValue)
                    headerCell.Font.Bold = False
                    headerCell.Font.Size = WebControls.FontUnit.Point(10)
                    headerCell.Height = WebControls.Unit.Pixel(intColWeight - 2)
                    headerCell.Style.Add("text-align", "center")
                    headerRow.Cells.Add(headerCell)
                    intLoop += 1
                End If
            End If
        Next

        tblTitle.Controls.Add(headerRow)

    End Sub

    ''' <summary>
    ''' タイトルの作成
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub GridViewTitle_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        If e.Row.RowType.Equals(DataControlRowType.Header) Then
            Dim strBound() As String = strKey.Split(",")
            Dim dr() As DataRow = Nothing
            Dim intAddCount As Integer = 0
            Dim intWidth() As Integer = Nothing
            Select Case strMode
                Case "14" '14 国別生産品マスタメンテナンス
                    intAddCount = 3
                    intWidth = strWidth_CountryItem
                Case "11" '11 掛率マスタメンテナンス
                    intAddCount = 5
                    If HidRateDiv.Value = "1" Then  '購入価格
                        intWidth = strWidth_RateNet
                    Else '現地定価
                        intWidth = strWidth_RateLocal
                    End If
                Case "12" '12 情報マスタメンテナンス
                Case "13" '13 為替率マスタメンテナンス
                Case "10" '10 ユーザーマスタメンテナンス
                    intAddCount = 5
                    intWidth = strWidth_User
                Case "15" '15 マスタメンテナンス
            End Select
            Dim intLoop As Integer = 0

            Dim headerRow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            For inti As Integer = 0 To strBound.Length - 1
                If strBound(inti).Length > 0 Then
                    dr = subSetLbl.Select("label_seq='" & inti + intAddCount & "' AND label_div='L'")
                    If dr.Length > 0 Then
                        Dim headerGrid As GridView = CType(sender, GridView)

                        Dim headerCell As TableCell = New TableCell()
                        headerCell.Text = dr(0)("label_content").ToString
                        'headerCell.ColumnSpan = 1
                        If inti = 0 Then
                            headerCell.Width = WebControls.Unit.Pixel(intWidth(intLoop) - 2)
                        Else
                            headerCell.Width = WebControls.Unit.Pixel(intWidth(intLoop) - 3)
                        End If
                        headerCell.BackColor = Color.Green
                        headerCell.ForeColor = Color.White
                        headerCell.BorderStyle = BorderStyle.Solid
                        headerCell.BorderWidth = WebControls.Unit.Pixel(1)
                        headerCell.Font.Name = GetFontName(selLang.SelectedValue)
                        headerCell.Font.Bold = False
                        headerCell.Font.Size = WebControls.FontUnit.Point(10)
                        headerCell.Height = WebControls.Unit.Pixel(intColWeight - 2)
                        headerCell.Style.Add("text-align", "center")
                        headerRow.Cells.Add(headerCell)
                        intLoop += 1
                    End If
                End If
            Next
            'GridViewTitle.Controls(0).Controls.AddAt(0, headerRow)
        End If
    End Sub

    ''' <summary>
    ''' 端末認証情報の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetWebLoginPnl()
        Dim blnWebLogin As Boolean = False

        blnWebLogin = My.Settings.LoginCheck
        'blnWebLogin = True
        If blnWebLogin AndAlso strMode = "10" Then
            pnlWebLog.Visible = True
            'txtLastUsedTime.Text = Now.ToString
        Else
            pnlWebLog.Visible = False
        End If
    End Sub

    ''' <summary>
    ''' 端末認証情報の隠しエリアを作成
    ''' </summary>
    ''' <param name="strRowID"></param>
    ''' <param name="strUserID"></param>
    ''' <remarks></remarks>
    Private Sub CreateHidWebLogInfo(ByRef pnl As Panel, ByVal strRowID As String, ByVal strUserID As String)

        Dim objHidden As HiddenField

        Dim dtUser As New DS_M_User.M_UserDataTable
        Using daUser As New DS_M_UserTableAdapters.M_UserTableAdapter
            dtUser = daUser.GetDataByUserID(strUserID)
        End Using

        If dtUser.Rows.Count > 0 Then
            '端末認証パスワード
            objHidden = New HiddenField
            objHidden.ID = "HdnWebPassNo"
            objHidden.Value = Trim(dtUser.Rows(0).Item("Password").ToString)
            pnl.Controls.Add(objHidden)

            'マックアドレス
            objHidden = New HiddenField
            objHidden.ID = "HdnMacNo"
            objHidden.Value = Trim(dtUser.Rows(0).Item("MacAddress").ToString)
            pnl.Controls.Add(objHidden)

            'シリアルNo
            objHidden = New HiddenField
            objHidden.ID = "HdnSerialNo"
            objHidden.Value = Trim(dtUser.Rows(0).Item("SerialNo").ToString)
            pnl.Controls.Add(objHidden)

            '最終利用実績
            objHidden = New HiddenField
            objHidden.ID = "HdnLastDateNo"
            objHidden.Value = Trim(dtUser.Rows(0).Item("LoginTime").ToString)
            pnl.Controls.Add(objHidden)
        End If

    End Sub

    ''' <summary>
    ''' ページング
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub AspNetPager1_PageChanged(sender As Object, e As EventArgs) Handles AspNetPager1.PageChanged
        subSearch()
    End Sub

    ''' <summary>
    ''' 税率の格式を確認
    ''' </summary>
    ''' <param name="strRate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckRateFormat(ByVal strRate As String) As Boolean
        Dim regex As Regex = New Regex("^[0-9]{0,3}(.[0-9]{0,8})?$")

        Return regex.IsMatch(strRate)
    End Function
End Class