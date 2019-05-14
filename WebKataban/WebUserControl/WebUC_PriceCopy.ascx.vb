Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class WebUC_PriceCopy
    Inherits KHBase

#Region "プロパティ"
    Public Event BackToTanka(intMode As Integer)
    'クリップボード用テキスト
    Private strCopyPrice As String = String.Empty
    Private strPriceHead As String = String.Empty  'テキストヘッダ
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="_intMode"></param>
    ''' <remarks></remarks>
    Public Sub frmInit(_intMode As Integer)
        HidMode.Value = _intMode
        'Me.OnLoad(Nothing)
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
        If HidMode.Value.Length <= 0 Then Exit Sub

        Try
            Call Me.subSetInit()
            lblSeriesKat.Text = objKtbnStrc.strcSelection.strFullKataban
            lblSeriesNm.Text = objKtbnStrc.strcSelection.strGoodsNm

            strPriceHead = String.Empty
            Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHPriceCopy, selLang.SelectedValue, Me) 'Label取得
            For inti As Integer = 1 To 8
                'クリップボード用テキストヘッダ
                Dim obj As Label = Me.tblPriceList.FindControl("Label" & inti)
                If obj Is Nothing Then
                    Continue For
                Else
                    If inti < 8 Then
                        strPriceHead &= obj.Text + "\t"
                    Else
                        strPriceHead &= obj.Text + "\r\n"
                    End If
                End If
            Next
            strCopyPrice = strPriceHead & strCopyPrice
            Me.btnOK.Attributes.Add(CdCst.JavaScript.OnClick, "Clip_Copy('" & strCopyPrice & "');")
            Call SetAllFontName(Me)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInit()
        Dim objRow As TableRow
        Dim objCell As TableCell
        Dim objLabel As Label

        Try
            '表項目を作成
            objRow = New TableRow
            strPriceHead = ""
            For i As Integer = 1 To 8
                'セルの作成
                objCell = New TableCell
                If i = 1 Then
                    objCell.Width = Unit.Percentage(30)
                Else
                    objCell.Width = Unit.Percentage(10)
                End If

                'ラベルの作成
                objLabel = New Label
                With objLabel
                    .ID = CdCst.Lbl.Name.Label & i
                    Call SetAttributes(objLabel, 1)
                    .Width = Unit.Percentage(100)
                End With

                objCell.Controls.Add(objLabel)
                objRow.Cells.Add(objCell)
            Next
            Me.tblPriceList.Rows.Add(objRow)

            '価格一覧を作成
            Select Case HidMode.Value
                Case 1
                    Call fncGetItemizedPrice()
                Case 2
                    Call fncGetItemizedPriceISO()
            End Select
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' OKボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOK_Click(sender As Object, e As System.EventArgs) Handles btnOK.Click
        RaiseEvent BackToTanka(HidMode.Value)
    End Sub

    ''' <summary>
    ''' Cancelボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        RaiseEvent BackToTanka(HidMode.Value)
    End Sub

    ''' <summary>
    ''' 価格の取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub fncGetItemizedPrice()
        Dim objRow As TableRow
        Dim objCell As TableCell
        Dim objLabel As Label

        Try
            '初期値
            Dim strKey As String = ""
            strCopyPrice = strPriceHead

            Dim strCurrency As String = String.Empty
            Dim objPrice As New KHUnitPrice

            Dim dt_math As DataTable = objPrice.fncGetCurrMathAll(objConBase)
            Dim dr_math() As DataRow = Nothing
            Dim intType As Integer = 0
            Dim dblPos As Double = 0D

            Dim dt_price As DataTable = subGetAllPrice()   '価格データを取得する
            Dim strPriceKey() As String = {"ls_price", "rg_price", "ss_price", "bs_price", "gs_price", "ps_price"}
            For inti As Integer = 0 To dt_price.Rows.Count - 1
                '行の追加
                objRow = New TableRow
                objCell = New TableCell

                '価格キー
                objLabel = New Label
                If Not IsDBNull(dt_price.Rows(inti)("kataban")) Then
                    strKey = dt_price.Rows(inti)("kataban").ToString
                    If Len(strKey) > 30 Then
                        objLabel.Text = Mid(strKey, 1, 30) & vbCrLf & Mid(strKey, 31, 30)
                    Else
                        objLabel.Text = strKey
                    End If
                End If
                SetAttributes(objLabel, 6)
                objLabel.Style.Add("text-align", "left")
                objLabel.Font.Name = GetFontName(selLang.SelectedValue)
                objLabel.Width = Unit.Percentage(100)
                objCell.Controls.Add(objLabel)
                objRow.Cells.Add(objCell)

                'クリップボード用テキストへ価格キーを保存
                strCopyPrice = strCopyPrice + strKey + "\t"

                If Not IsDBNull(dt_price.Rows(inti)("currency_cd")) Then
                    strCurrency = dt_price.Rows(inti).ToString
                End If
                dr_math = dt_math.Select("currency_cd='" & strCurrency & "'")

                For intj As Integer = 0 To strPriceKey.Length - 1
                    objCell = New TableCell
                    objLabel = New Label
                    If Not IsDBNull(dt_price.Rows(inti)(strPriceKey(intj))) Then
                        If dr_math.Length > 0 Then
                            intType = dr_math(0)("math_Type")
                            dblPos = dr_math(0)("math_Pos")
                            objLabel.Text = CDec(objPrice.subFractionProc(dt_price.Rows(inti)(strPriceKey(intj)), _
                                                                 intType, dblPos))
                        Else
                            objLabel.Text = Math.Round(dt_price.Rows(inti)(strPriceKey(intj)), 2)
                        End If
                    End If
                    SetAttributes(objLabel, 6)
                    objLabel.Width = Unit.Percentage(100)
                    objLabel.Font.Name = GetFontName(selLang.SelectedValue)
                    objCell.Controls.Add(objLabel)
                    objRow.Cells.Add(objCell)

                    'クリップボード用テキストへ定価を保存
                    strCopyPrice = strCopyPrice + objLabel.Text + "\t"
                Next

                '数量
                objCell = New TableCell
                objLabel = New Label
                If Not IsDBNull(dt_price.Rows(inti)("amount")) Then
                    objLabel.Text = dt_price.Rows(inti)("amount")
                End If
                SetAttributes(objLabel, 6)
                objLabel.Width = Unit.Percentage(100)
                objLabel.Font.Name = GetFontName(selLang.SelectedValue)
                objCell.Controls.Add(objLabel)
                objRow.Cells.Add(objCell)
                'クリップボード用テキストへ数量を保存
                strCopyPrice = strCopyPrice + objLabel.Text + "\r\n"
                Me.tblPriceList.Rows.Add(objRow)
            Next

            If dt_price.Rows.Count <= 10 Then
                Me.pnlMain.Height = WebControls.Unit.Pixel(650)
            Else
                Me.pnlMain.Height = WebControls.Unit.Pixel(650 + (dt_price.Rows.Count - 10) * 25)
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 価格情報の取得ISO
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub fncGetItemizedPriceISO()
        Dim objRow As TableRow
        Dim objCell As TableCell
        Dim objLabel As Label
        Dim objPrice As New KHUnitPrice
        Dim bolReturn1 As Boolean
        Dim bolReturn2 As Boolean
        Dim bolSpecItemJudge As Boolean
        Dim strRefKataban() As String = Nothing
        Dim strRefSeqNo() As String = Nothing
        Dim decRefAmount() As Decimal = Nothing
        Dim strRefPriceKey() As String = Nothing
        Dim strRetKatabanCheckDiv As String = String.Empty
        Dim strRetPlaceCd As String = String.Empty
        Dim htRetPriceInfo As New Hashtable

        Try
            '形番と仕様書構成順序を取得
            subGetKataban(strRefKataban, strRefSeqNo)
            '初期値
            strCopyPrice = ""

            Dim dt_math As DataTable = objPrice.fncGetCurrMathAll(objConBase)
            Dim dr_math() As DataRow = Nothing
            Dim intType As Integer = 0
            Dim dblPos As Double = 0D

            For i As Integer = 1 To UBound(strRefKataban)
                '仕様書項目かどうかを判定
                If objKtbnStrc.strcSelection.strSeriesKataban = "CMF" Or _
                   objKtbnStrc.strcSelection.strSeriesKataban = "GMF" Then
                    Select Case strRefSeqNo(i)
                        Case 1
                            bolSpecItemJudge = True  'ベース
                        Case 2, 3, 4, 5, 6, 7
                            bolSpecItemJudge = True  '電磁弁形式
                        Case 13, 14
                            bolSpecItemJudge = True  '給気スペーサ
                        Case 15, 16
                            bolSpecItemJudge = True  '排気スペーサ
                        Case 17, 18
                            bolSpecItemJudge = True  'パイロットチェック弁
                        Case 19, 20, 21, 22
                            bolSpecItemJudge = True  'スペーサ形減圧弁
                        Case 23, 24
                            bolSpecItemJudge = True  '流露遮蔽板
                        Case Else
                            bolSpecItemJudge = False
                    End Select
                End If
                If objKtbnStrc.strcSelection.strSeriesKataban = "LMF0" Then
                    Select Case strRefSeqNo(i)
                        Case 1
                            bolSpecItemJudge = True
                        Case 2, 3, 4, 5, 6, 7
                            bolSpecItemJudge = True
                        Case 13, 14
                            bolSpecItemJudge = True
                        Case 15, 16
                            bolSpecItemJudge = True
                        Case 17
                            bolSpecItemJudge = True
                        Case 18, 19
                            bolSpecItemJudge = True
                        Case Else
                            bolSpecItemJudge = False
                    End Select
                End If

                Dim strCurrency As String = String.Empty
                Dim strMadeCountry As String = String.Empty

                If bolSpecItemJudge = True Then

                    '価格キー取得
                    bolReturn1 = objKtbnStrc.fncISOGetPriceKey(i, strRefKataban(i), strRefPriceKey, decRefAmount)
                    If bolReturn1 = True Then
                        'ヘッダ情報を追加
                        objRow = New TableRow
                        objCell = New TableCell

                        '価格キーヘッダ
                        objLabel = New Label
                        objLabel.Text = "＜＜" + strRefKataban(i) + "＞＞"
                        SetAttributes(objLabel, 6)
                        objLabel.Style.Add("text-align", "left")
                        objLabel.Style.Add("font-weight", "bold")
                        objLabel.Width = Unit.Percentage(100)
                        objLabel.Font.Name = GetFontName(selLang.SelectedValue)

                        objCell.Controls.Add(objLabel)
                        objCell.ColumnSpan = 8
                        objRow.Cells.Add(objCell)

                        Me.tblPriceList.Rows.Add(objRow)

                        'ヘッダ指定
                        strPriceHead = "<<" + strRefKataban(i) + ">>\r\n"

                        For n As Integer = 1 To strRefPriceKey.Length - 1
                            '単価情報読み込み
                            bolReturn2 = objPrice.fncSelectPrice(objCon, strRefPriceKey(n), _
                                                                strRetKatabanCheckDiv, strRetPlaceCd, _
                                                                htRetPriceInfo, objKtbnStrc.strcSelection.strCurrency, strMadeCountry)
                            '積上単価情報読み込み
                            If Not bolReturn2 Then
                                bolReturn2 = objPrice.fncSelectAccumulatePrice(objCon, strRefPriceKey(n), _
                                                                              strRetKatabanCheckDiv, strRetPlaceCd, _
                                                                              htRetPriceInfo, objKtbnStrc.strcSelection.strCurrency)
                            End If

                            If strCurrency.Length <= 0 Then strCurrency = objKtbnStrc.strcSelection.strCurrency
                            dr_math = dt_math.Select("currency_cd='" & strCurrency & "'")

                            '行の追加
                            objRow = New TableRow
                            objCell = New TableCell

                            '価格キー
                            objLabel = New Label
                            objLabel.Text = strRefPriceKey(n)
                            SetAttributes(objLabel, 6)
                            objLabel.Style.Add("text-align", "left")
                            objLabel.Width = Unit.Percentage(100)
                            objLabel.Font.Name = GetFontName(selLang.SelectedValue)
                            objCell.Controls.Add(objLabel)
                            objRow.Cells.Add(objCell)

                            'クリップボード用テキストへ価格キーを保存
                            strCopyPrice = strCopyPrice + objLabel.Text + "\t"

                            Dim strPriceKey() As String = {"ListPrice", "RegPrice", "SsPrice", "BsPrice", "GsPrice", "PsPrice"}
                            For inti As Integer = 0 To strPriceKey.Length - 1
                                objCell = New TableCell
                                objLabel = New Label
                                If dr_math.Length > 0 Then
                                    intType = dr_math(0)("math_Type")
                                    dblPos = dr_math(0)("math_Pos")
                                    objLabel.Text = CDec(objPrice.subFractionProc(htRetPriceInfo(strPriceKey(inti)), _
                                                                         intType, dblPos))
                                Else
                                    objLabel.Text = htRetPriceInfo(strPriceKey(inti))
                                End If
                                SetAttributes(objLabel, 6)
                                objLabel.Width = Unit.Percentage(100)
                                objLabel.Font.Name = GetFontName(selLang.SelectedValue)
                                objCell.Controls.Add(objLabel)
                                objRow.Cells.Add(objCell)
                                'クリップボード用テキストへ定価を保存
                                strCopyPrice = strCopyPrice + objLabel.Text + "\t"
                            Next

                            '数量
                            objCell = New TableCell
                            objLabel = New Label
                            objLabel.Text = decRefAmount(n).ToString
                            SetAttributes(objLabel, 6)
                            objLabel.Width = Unit.Percentage(100)
                            objLabel.Font.Name = GetFontName(selLang.SelectedValue)
                            objCell.Controls.Add(objLabel)
                            objRow.Cells.Add(objCell)

                            'クリップボード用テキストへ数量を保存
                            strCopyPrice = strCopyPrice + objLabel.Text + "\r\n"

                            Me.tblPriceList.Rows.Add(objRow)
                        Next
                    End If
                End If
            Next

            If strRefKataban.Length <= 10 Then
                Me.pnlMain.Height = WebControls.Unit.Pixel(650)
            Else
                Me.pnlMain.Height = WebControls.Unit.Pixel(650 + (strRefKataban.Length - 10) * 25)
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 価格情報の検索
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function subGetAllPrice() As DataTable
        Dim sbSql As New StringBuilder
        subGetAllPrice = New DataTable
        Dim objCmd As SqlCommand
        Dim objAdp As SqlDataAdapter
        Try
            'SQL Query生成
            sbSql.Append(" SELECT  kataban, ")
            sbSql.Append("         ls_price, ")
            sbSql.Append("         rg_price, ")
            sbSql.Append("         ss_price, ")
            sbSql.Append("         bs_price, ")
            sbSql.Append("         gs_price, ")
            sbSql.Append("         ps_price, ")
            sbSql.Append("         amount, currency_cd ")
            sbSql.Append(" FROM    kh_sel_acc_prc_strc ")
            sbSql.Append(" WHERE   user_id             = @UserId ")
            sbSql.Append(" AND     session_id          = @SessionId ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = Me.objUserInfo.UserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = Me.objLoginInfo.SessionId
            End With

            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(subGetAllPrice)
        Catch ex As Exception
            AlertMessage(ex)
        Finally
            sbSql = Nothing
            objAdp = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 価格明細取得
    ''' </summary>
    ''' <param name="strRefKataban"></param>
    ''' <param name="strRefSeqNo"></param>
    ''' <remarks>積上単価テーブルより価格明細を取得する</remarks>
    Private Sub subGetKataban(ByRef strRefKataban As String(), ByRef strRefSeqNo As String())
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        Try

            sbSql.Append(" SELECT ")
            sbSql.Append("     ISNULL(b.kataban , '') AS kataban, ")
            sbSql.Append("     ISNULL(a.spec_strc_seq_no , '') AS spec_strc_seq_no ")
            sbSql.Append(" FROM ")
            sbSql.Append("           sales.kh_sel_spec_strc a ")
            sbSql.Append(" LEFT JOIN sales.kh_sel_acc_prc_strc b ")
            sbSql.Append(" ON    a.user_id          = b.user_id ")
            sbSql.Append(" AND   a.session_id       = b.session_id ")
            sbSql.Append(" AND   a.spec_strc_seq_no = b.disp_seq_no ")
            sbSql.Append(" WHERE a.user_id          = @UserId ")
            sbSql.Append(" AND   a.session_id       = @SessionId ")
            sbSql.Append(" AND   a.option_kataban  <> '' ")
            sbSql.Append(" AND   a.quantity         > 0 ")
            sbSql.Append(" ORDER BY a.spec_strc_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = Me.objUserInfo.UserId
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = Me.objLoginInfo.SessionId
            End With

            objRdr = objCmd.ExecuteReader
            '配列定義
            ReDim strRefKataban(0)
            ReDim strRefSeqNo(0)
            While objRdr.Read()
                '配列再定義
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve strRefSeqNo(UBound(strRefSeqNo) + 1)
                '形番
                strRefKataban(UBound(strRefKataban)) = objRdr.GetValue(objRdr.GetOrdinal("kataban"))
                '仕様書構成順序
                strRefSeqNo(UBound(strRefSeqNo)) = objRdr.GetValue(objRdr.GetOrdinal("spec_strc_seq_no"))
            End While
        Catch ex As Exception
            AlertMessage(ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try

    End Sub
End Class