Imports WebKataban.ClsCommon
Imports WebKataban.KHCodeConstants

Public Class UC_ISOTanka
    Inherits KHBase

    Private CST_BLANK As String = ""
    Private CST_SPACE As String = " "
    Private CST_COMMA As String = ","
    Private CST_PIRIOD As String = "."
    Private CST_SLASH As String = "/"
    Private CST_PIPE As String = "|"

    'ﾌﾟﾛﾊﾟﾃｨ設定値
    Private strLangCd As String
    Private strTtlCnt As String
    Private strDataNo As String
    Private strPriceLst() As String
    Private strShipPlc As String
    Private strKtbnChk As String
    Private strCurr As String
    Private strAmnt As String
    Private strPrcFncNm As String
    Private strEditDiv As String
    Private strDispDiv() As String
    '掛単価
    Public Property RatePrice As String
    '単価
    Public Property UnitPrice As String
    '掛率
    Public Property Rate As String
    '数量
    Public Property Quantity As String
    '金額
    Public Property Price As String
    '消費税
    Public Property Tax As String
    '合計
    Public Property Total As String


#Region "プロパティ"
    Public Property IsFirst As Boolean = False
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
        txt_Rate.Attributes.Add(CdCst.JavaScript.OnKeyUp, "ISOTanka_OnKeyup(event,'1','" & Me.ClientID & "_');")
        txt_UnitPrc.Attributes.Add(CdCst.JavaScript.OnKeyUp, "ISOTanka_OnKeyup(event,'2','" & Me.ClientID & "_');")
        txt_Amount.Attributes.Add(CdCst.JavaScript.OnKeyUp, "ISOTanka_OnKeyup(event,'3','" & Me.ClientID & "_');")
    End Sub

    ''' <summary>
    ''' ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.lblName.Text = OptionNm
            Me.lblKataName.Text = OptionKtbn

            Call subSetInitScript()
            Call SetAttributes(txt_Amount, 1)
            Call SetAttributes(txt_DtlPrc, 1)
            Call SetAttributes(txt_Price, 1)
            Call SetAttributes(txt_Rate, 0)
            Call SetAttributes(txt_Tax, 1)
            Call SetAttributes(txt_Total, 1)
            Call SetAttributes(txt_UnitPrc, 0)
            'Call SetAttributes(txt_KtbnChk, 9)
            Call SetAttributes(txt_KtbnChk, 8)
            Call SetAttributes(txt_ChkZ, 1)

            Call SetAttributes(txt_Place, 9)
            Call SetAttributes(Label21, 3)
            Call SetAttributes(Label22, 3)
            Call SetAttributes(Label23, 3)
            Call SetAttributes(Label24, 3)
            Call SetAttributes(Label25, 3)
            Call SetAttributes(Label26, 3)
            Call SetAttributes(Label27, 3)
            Call SetAttributes(Label28, 3)
            Call SetAttributes(lblNo, 4)
            Call SetAttributes(lblKataName, 4)
            Call SetAttributes(lblName, 4)

            Call Me.subSetInitScreen() '初期画面設定
            Call Me.subSetData() 'ﾃﾞｰﾀｾｯﾄ

            'ラベルタイトル設置
            KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHISOTanka, strLangCd, Me)
            HidSelRowID.Value = String.Empty
            Call SetAllFontName(Me)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 初期画面設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInitScreen()
        Try
            If DispDiv Is Nothing OrElse DispDiv.Length <= 0 Then Exit Sub
            '--形番ﾁｪｯｸ説明
            'Me.Label15.Visible = DispDiv(2)
            'Me.Label15.Visible = False

            If DispDiv(3) Then
                '--消費税
                Me.txt_Tax.Visible = False
                '--合計
                Me.Label28.Visible = False
                Me.txt_Total.Visible = False
            End If

            If DispDiv(4) Then
                '--掛率
                Me.Label24.Visible = False
                Me.txt_Rate.Visible = False
                '--単価
                Me.Label25.Visible = False
                Me.txt_UnitPrc.Visible = False
                '--単価（少数1桁）
                Me.txt_DtlPrc.Visible = False
                '--数量
                Me.Label26.Visible = False
                Me.txt_Amount.Visible = False
                '--合計
                Me.Label27.Visible = False
                Me.txt_Price.Visible = False
                '--消費税
                Me.txt_Tax.Visible = False
                '--合計
                Me.Label28.Visible = False
                Me.txt_Total.Visible = False
                '--反映チェックボックス
                '別途記述
            End If

            '--形番ﾁｪｯｸ
            If DispDiv(0) Then
                With Me.txt_KtbnChk
                    .Text = CST_BLANK
                    .ReadOnly = True
                    .Width = 65
                    .Visible = True
                End With
                Me.Label21.Visible = True

                'チェック区分の最初にZを表示する
                With Me.txt_ChkZ
                    .Text = "Z"
                    .ReadOnly = True
                    .Width = 65
                    .Visible = True
                End With

            Else
                Me.txt_KtbnChk.Visible = False
                Me.Label21.Visible = False
                'Me.Label15.Visible = False
                Me.txt_ChkZ.Visible = False

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), txt_ChkZ.Text, _
                         "fnclblCheck('" & "ClsChk" & "');", True)

            End If
            '--出荷場所
            If DispDiv(1) Then
                With Me.txt_Place
                    .Text = CST_BLANK
                    .ReadOnly = True                                    '出荷場所をReadOnlyにする
                    .Width = 130
                    .Visible = True
                End With
                Me.Label22.Visible = True
            Else
                Me.txt_Place.Visible = False
                Me.Label22.Visible = False
                '保管場所＆評価タイプ
                Me.lblStrageLocation.Visible = False
            End If
            '--掛率
            With Me.txt_Rate
                .Text = CST_BLANK
                .ReadOnly = False
                .Width = 105
                .DecLen = 3
                .DispComma = True
                .EditDiv = strEditDiv
                .AllowZero = True
                .AllowMinus = False
            End With
            '--単価
            With Me.txt_UnitPrc
                .Text = CST_BLANK
                .ReadOnly = False
                .Width = 125
                .DecLen = 2
                .DispComma = True
                .EditDiv = strEditDiv
                .AllowZero = True
                .AllowMinus = False
            End With
            '--数量
            With Me.txt_Amount
                .Text = CST_BLANK
                If DataNo = 1 Then
                    .ReadOnly = False
                End If
                .Width = 125
                .DecLen = 0
                .DispComma = True
                .EditDiv = strEditDiv
                .AllowZero = True
                .AllowMinus = False
            End With

            '一番目のユーザーコントロールだけが編集できる
            Me.txt_Amount.ReadOnly = Not Me.IsFirst

            '--金額
            With Me.txt_Price
                .Text = CST_BLANK
                .Width = 125
                .DecLen = 2
                .DispComma = True
                .EditDiv = strEditDiv
                .AllowZero = True
                .AllowMinus = False
            End With
            '--消費税
            With Me.txt_Tax
                .Text = CST_BLANK
                .Width = 125
                .DecLen = 2
                .DispComma = True
                .EditDiv = strEditDiv
                .AllowZero = True
                .AllowMinus = False
            End With
            '--合計
            With Me.txt_Total
                .Text = CST_BLANK
                .Width = 125
                .DispComma = True
                .EditDiv = strEditDiv
                .AllowZero = True
                .AllowMinus = False
            End With
            '--単価(詳細)
            With Me.txt_DtlPrc
                .Text = CST_BLANK
                '.ReadOnly = True
                .Width = 125
                .DecLen = 2
                .DispComma = True
                .EditDiv = strEditDiv
                .AllowZero = True
                .AllowMinus = False
            End With
            '--反映チェックボックス
            If DispDiv(4) Then
                Me.Label29.Visible = False
                Me.Label30.Visible = False
                Me.Label31.Visible = False
                Me.ChkUnitList.Visible = False
            Else
                If DataNo = 1 Then
                    Me.Label29.Visible = True
                    Me.Label30.Visible = False
                Else
                    Me.Label29.Visible = False
                    Me.Label30.Visible = True
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 画面に値を設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetData()
        Try

            '単価ﾘｽﾄ設定
            subSetPriceList()

            'その他の項目設定
            subSetISOInfo()

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 単価ﾘｽﾄ設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetPriceList()
        Dim strValue As String = String.Empty
        'ｶﾝﾏ区切りのﾘｽﾄ内容を配列にｾｯﾄ
        Dim maxLst As Integer = 0
        If strPriceLst IsNot Nothing Then maxLst = strPriceLst.Length

        '価格テーブル
        Dim dt_price As New DataTable
        Dim strColumnNames As List(Of String) = New List(Of String) From {"strText", "strPrice", "ColumnKBN"}
        dt_price = fncCreateTableByColumnNames(strColumnNames)

        Dim dr As DataRow = Nothing
        For rowIdx As Integer = 0 To maxLst - 1
            '価格区分
            Dim strColumnKbn As String = String.Empty

            dr = dt_price.NewRow()

            '価格タイトル
            dr("strText") = strPriceLst(rowIdx).Split(CST_PIPE)(0).Trim
            '価格
            strValue = strPriceLst(rowIdx).Split(CST_PIPE)(1).Trim
            '通貨
            strCurr = strPriceLst(rowIdx).Split(CST_PIPE)(2).Trim
            '価格区分
            strColumnKbn = strPriceLst(rowIdx).Split(CST_PIPE)(3).Trim
            '金額
            If IsNumeric(strValue) Then
                '金額をｶﾝﾏ編集する
                dr("strPrice") = fncSetComma(strValue, strCurr) & "(" & strCurr & ")"
            Else
                dr("strPrice") = strValue
            End If

            '価格区分
            dr("ColumnKBN") = fncConvertColumnKBN(strColumnKbn)

            dt_price.Rows.Add(dr)
        Next
        GVPrice.DataSource = dt_price
        GVPrice.DataBind()
    End Sub

    ''' <summary>
    ''' ISO情報の設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetISOInfo()
        '形番
        Me.lblNo.Text = DataNo & CST_SLASH & TtlCnt
        'チェック区分
        Me.txt_KtbnChk.Text = strKtbnChk
        '数量
        Me.txt_Amount.Text = Me.Quantity
        '金額
        Me.txt_Price.Text = Me.Price
        '消費税
        Me.txt_Tax.Text = Me.Tax
        '合計
        Me.txt_Total.Text = Me.Total
        '出荷場所
        Me.txt_Place.Text = strShipPlc
        '掛単価
        Me.txt_DtlPrc.Text = Me.RatePrice
        '単価
        Me.txt_UnitPrc.Text = Me.UnitPrice
        '掛率
        Me.txt_Rate.Text = Me.Rate
        '編集可否
        Me.txt_Rate.EditDiv = strEditDiv
        Me.txt_UnitPrc.EditDiv = strEditDiv
        Me.txt_DtlPrc.EditDiv = strEditDiv
        '保管場所
        Select Case strShipPlc
            Case "1002", "1003", "1004", "1005"
                Me.lblStrageLocation.Text = "A***"
            Case Else
                Me.lblStrageLocation.Text = String.Empty
        End Select



        ''掛率/掛単価/単価にデフォルト値をセットする
        'If EditDiv = "0" Then
        '    Me.txt_Rate.Text = CdCst.UnitPrice.DefaultNmlRate
        '    Me.txt_DtlPrc.Text = CdCst.UnitPrice.DefaultNmlRateUnitPrice
        'Else
        '    Me.txt_Rate.Text = CdCst.UnitPrice.DefaultOtrRate
        '    Me.txt_DtlPrc.Text = CdCst.UnitPrice.DefaultOtrRateUnitPrice
        'End If
        'Me.txt_UnitPrc.Text = CdCst.UnitPrice.DefaultUnitPrice
    End Sub

    ''' <summary>
    ''' データバインド
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub GVPrice_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVPrice.RowDataBound
        If e.Row.RowIndex < 0 Then Exit Sub
        '価格があるものを選択できる
        Dim str() As String = e.Row.Cells(1).Text.ToString.Split("(")
        If str.Length = 2 Then
            Dim strPrice As String = str(0).ToString.Replace(",", "").Replace(".", "")
            Dim strCurr As String = str(1).ToString.Replace(")", "").Replace("(", "")
            Dim strName As String = Me.ClientID & "_"
            Dim intStartID As Integer = 0
            If e.Row.RowIndex = 0 Then
                intStartID = CInt(Strings.Right(e.Row.ClientID, 2))
            Else
                intStartID = CInt(Strings.Right(GVPrice.Rows(0).ClientID, 2))
            End If
            If (e.Row.RowIndex + 1) Mod 2 = 0 Then
                'e.Row.BackColor = Drawing.Color.FromArgb(173, 205, 207)
                e.Row.BackColor = Drawing.Color.FromArgb(204, 204, 255)
            Else
                e.Row.BackColor = Drawing.Color.White
            End If
            e.Row.Attributes.Add(CdCst.JavaScript.OnClick, "fncGridClick('" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',2);")
            e.Row.Attributes.Add(CdCst.JavaScript.OnKeyUp, "fncGrid_OnKeyup(event, '" & strName & "','" & e.Row.ClientID & "','" & intStartID & "',2);")
        End If
    End Sub

    ''' <summary>
    ''' カンマ編集
    ''' </summary>
    ''' <param name="strValue"></param>
    ''' <param name="strCurr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSetComma(ByVal strValue As String, ByVal strCurr As String) As String
        Dim strRtnVal As String = String.Empty
        Try
            '小数/整数の区切り文字と桁区切り文字を設定
            If strEditDiv = CdCst.EditDiv.Normal Then
                Dim str() As String = strValue.Split(".")
                If str.Length >= 2 Then
                    If String.IsNullOrEmpty(str(1).Trim("0").Trim) Then
                        strRtnVal = FormatNumber(strValue, 0)
                    Else
                        strRtnVal = FormatNumber(strValue, str(1).Length)
                    End If
                Else
                    strRtnVal = FormatNumber(strValue, 0)
                End If
            Else
                Dim intPlus As Integer '余り
                Dim decQuot As Decimal '商
                Dim intQuot As Integer
                Dim strDecDel As String
                Dim strDigitDel As String

                '小数/整数の区切り文字と桁区切り文字を設定
                If strEditDiv = CdCst.EditDiv.Normal Then
                    strDecDel = CdCst.Sign.Dot
                    strDigitDel = CdCst.Sign.Comma
                Else
                    strDecDel = CdCst.Sign.Comma
                    strDigitDel = CdCst.Sign.Dot
                End If

                '編集の必要がなかった場合
                If strValue.Length < 4 Then
                    Return strValue
                End If

                intPlus = strValue.Length Mod 3
                decQuot = strValue.Length / 3
                If CStr(decQuot).IndexOf(strDecDel) > -1 Then
                    intQuot = CStr(decQuot).Split(strDecDel)(0)
                    strRtnVal = Mid(strValue, 1, intPlus)
                Else
                    intQuot = decQuot - 1
                    strRtnVal = Mid(strValue, 1, 3)
                    intPlus = 3
                End If
                For idx As Integer = 1 To intQuot
                    strRtnVal = strRtnVal & strDigitDel & Mid(strValue, (intPlus + 1) + (idx - 1) * 3, 3)
                Next
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return strRtnVal
    End Function

    ''' <summary>
    ''' JavaScript生成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInitScript()


        Try
            'Shift+Dイベント設定
            Dim strShiftD As String = "if(event.keyCode == 68 && event.shiftKey==true){frmShiftD('ctl00_ContentDetail_WebUC_ISOTanka_');}"
            'Enterキーイベント設定
            Dim strEnterJS = "if (event.keyCode == 13){return false;}else{return true;}"

            Me.txt_DtlPrc.Attributes.Add(CdCst.JavaScript.OnKeyDown, strShiftD)
            Me.txt_Rate.Attributes.Add(CdCst.JavaScript.OnKeyDown, strEnterJS)
            Me.txt_UnitPrc.Attributes.Add(CdCst.JavaScript.OnKeyDown, strEnterJS)
            Me.txt_Amount.Attributes.Add(CdCst.JavaScript.OnKeyDown, strEnterJS)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

#Region "属性"
    ''' <summary>
    ''' Attributeの設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DispDiv() As String()
        Get
            Return strDispDiv
        End Get
        Set(ByVal Value As String())
            strDispDiv = Value
        End Set
    End Property

    ''' <summary>
    ''' 編集区分の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EditDiv() As String
        Get
            Return strEditDiv
        End Get
        Set(ByVal value As String)
            strEditDiv = value
        End Set
    End Property

    ''' <summary>
    ''' データNoの設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DataNo() As String
        Get
            Return Me.strDataNo
        End Get
        Set(ByVal value As String)
            Me.strDataNo = value
        End Set
    End Property

    ''' <summary>
    ''' データ数の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TtlCnt() As String
        Get
            Return Me.strTtlCnt
        End Get
        Set(ByVal value As String)
            Me.strTtlCnt = value
        End Set
    End Property

    ''' <summary>
    ''' オプション名称の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OptionNm() As String

    ''' <summary>
    ''' オプション形番の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OptionKtbn() As String

    ''' <summary>
    ''' 形番チェックの設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property KtbnChk() As String
        Get
            Return strKtbnChk
        End Get
        Set(ByVal value As String)
            strKtbnChk = value
        End Set
    End Property

    ''' <summary>
    ''' 出荷場所の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ShipPlace() As String
        Get
            Return strShipPlc
        End Get
        Set(ByVal value As String)
            strShipPlc = value
        End Set
    End Property

    ''' <summary>
    ''' 表示言語の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LangCd() As String
        Get
            Return Me.strLangCd
        End Get
        Set(ByVal value As String)
            Me.strLangCd = value
        End Set
    End Property

    ''' <summary>
    ''' Attributeの設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AttPrcLst() As String
        Get
            Return strPrcFncNm
        End Get
        Set(ByVal Value As String)
            strPrcFncNm = Value
        End Set
    End Property

    ''' <summary>
    ''' Attributeの設定・取得
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AttTxtPrc(ByVal key As String) As String
        Get
            Return Me.txt_UnitPrc.Attributes(key)
        End Get
        Set(ByVal Value As String)
            Me.txt_UnitPrc.Attributes.Add(key, Value)
        End Set
    End Property

    ''' <summary>
    ''' Attributeの設定・取得
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AttTxtRate(ByVal key As String) As String
        Get
            Return Me.txt_Rate.Attributes(key)
        End Get
        Set(ByVal Value As String)
            Me.txt_Rate.Attributes.Add(key, Value)
        End Set
    End Property

    ''' <summary>
    ''' Attributeの設定・取得
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AttTxtAmnt(ByVal key As String) As String
        Get
            Return Me.txt_Amount.Attributes(key)
        End Get
        Set(ByVal Value As String)
            Me.txt_Amount.Attributes.Add(key, Value)
        End Set
    End Property

    ''' <summary>
    ''' Attributeの設定・取得
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AttChkUnitList(ByVal key As String) As String
        Get
            Return Me.ChkUnitList.Attributes(key)
        End Get
        Set(ByVal Value As String)
            Me.ChkUnitList.Attributes.Add(key, Value)
        End Set
    End Property

    ''' <summary>
    ''' ドロップダウンリストの設定値の設定・取得(ﾘｽﾄ内容を配列ﾊﾟｲﾌﾟ区切りで渡す)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PriceLst() As String()
        Get
            Return strPriceLst
        End Get
        Set(ByVal value As String())
            strPriceLst = value
        End Set
    End Property
#End Region

End Class