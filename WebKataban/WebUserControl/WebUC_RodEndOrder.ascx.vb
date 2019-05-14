Imports WebKataban.ClsCommon

Public Class WebUC_RodEndOrder
    Inherits KHBase

#Region "プロパティ"
    Public Event GotoTanka()

    Private Structure CompData
        Public strSrsKataban As String     'シリーズ形番
        Public strKeyKataban As String     'キー形番
        Public strFullKataban As String    'フル形番
        Public strGoodsNm As String        '商品名
        Public strRodEndOption As String   'ロッド先端オプション
        Public strConCaliber As String     '接続口径
        Public strRodEndPett As String     'ロッド先端パターン
    End Structure

    Private strcCompData As CompData
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Me.OnLoad(Nothing)
        txtRodEndSize.Text = String.Empty
    End Sub

    ''' <summary>
    ''' ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub
        Try
            txtRodEndSize.Style.Add("text-transform", "uppercase")
            txtRodEndSize.Focus()
            '形番情報取得
            Call Me.subGetCompData()
            Me.lblSeriesNm.Text = strcCompData.strGoodsNm
            'ラベルタイトル設置
            Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHRodEndOrder, selLang.SelectedValue, Me)
            Call SetAllFontName(Me)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 構成情報取得
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subGetCompData()
        Try
            objKtbnStrc = New KHKtbnStrc
            '引当データ取得
            objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

            With strcCompData
                .strSrsKataban = objKtbnStrc.strcSelection.strSeriesKataban
                .strKeyKataban = objKtbnStrc.strcSelection.strKeyKataban
                .strGoodsNm = objKtbnStrc.strcSelection.strGoodsNm
                .strRodEndOption = objKtbnStrc.strcSelection.strRodEndOption

                'ロッド先端パターン保持
                Select Case strcCompData.strSrsKataban
                    Case "SSD"
                        If strcCompData.strKeyKataban = "" Then
                            strcCompData.strRodEndPett = objKtbnStrc.strcSelection.strOpSymbol(22).Trim
                        ElseIf strcCompData.strKeyKataban = "K" Then
                            strcCompData.strRodEndPett = objKtbnStrc.strcSelection.strOpSymbol(20).Trim
                        End If
                    Case "CMK2"
                        If strcCompData.strKeyKataban = "" Then
                            strcCompData.strRodEndPett = objKtbnStrc.strcSelection.strOpSymbol(17).Trim
                        End If
                    Case "SCM"
                        If strcCompData.strKeyKataban = "" Then
                            strcCompData.strRodEndPett = objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                        ElseIf strcCompData.strKeyKataban = "B" Then
                            strcCompData.strRodEndPett = objKtbnStrc.strcSelection.strOpSymbol(19).Trim
                        End If
                    Case "SCA2"
                        If strcCompData.strKeyKataban = "" Or strcCompData.strKeyKataban = "V" Then
                            strcCompData.strRodEndPett = objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                        ElseIf strcCompData.strKeyKataban = "B" Then
                            strcCompData.strRodEndPett = objKtbnStrc.strcSelection.strOpSymbol(19).Trim
                        End If
                        '接続口径保持
                        strcCompData.strConCaliber = objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                End Select
            End With
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
        Try
            Dim dalKtbnStrc As New KtbnStrcDAL

            '必須入力チェック
            If Not Me.fncIndispCheck() Then Exit Sub

            'SCA2の場合
            If strcCompData.strSrsKataban = "SCA2" Then If Not fncInpSizeCheck() Then Exit Sub

            '引当シリーズ形番更新(オプション情報)
            Call dalKtbnStrc.subSelSrsKtbnOptionUpd(objCon, Me.objUserInfo.UserId, _
                                                    Me.objLoginInfo.SessionId, _
                                                    Me.txtRodEndSize.Text.ToUpper.Trim)

            '引当シリーズ形番情報更新
            Call objKtbnStrc.subFullKatabanCreate(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)

            'ページ遷移(単価見積画面)
            RaiseEvent GotoTanka()
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 表示チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncIndispCheck() As Boolean
        Dim bolFlag As Boolean = False
        fncIndispCheck = False
        Try
            'パターンチェック
            Select Case strcCompData.strSrsKataban
                Case "SSD"
                    If strcCompData.strKeyKataban = "" Or strcCompData.strKeyKataban = "K" Then
                        If strcCompData.strRodEndPett = "N11" Or strcCompData.strRodEndPett = "N13" Then
                            bolFlag = True
                        End If
                    End If
                Case "CMK2"
                    If strcCompData.strKeyKataban = "" Then
                        If strcCompData.strRodEndPett = "N13" Then bolFlag = True
                    End If
                Case "SCM"
                    If strcCompData.strKeyKataban = "" Or strcCompData.strKeyKataban = "B" Then
                        If strcCompData.strRodEndPett = "N13" Then bolFlag = True
                    End If
                Case "SCA2"
                    If strcCompData.strKeyKataban = "" Or strcCompData.strKeyKataban = "V" Or _
                       strcCompData.strKeyKataban = "B" Then
                        If strcCompData.strRodEndPett = "N13" Then bolFlag = True
                    End If
            End Select

            '上のパターンに当てはまる場合、ロッド先端寸法を必須入力とする
            If Len(Trim(Me.txtRodEndSize.Text)) = 0 And bolFlag Then
                AlertMessage("W0110")
                fncIndispCheck = False
            Else
                fncIndispCheck = True
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' サイズチェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncInpSizeCheck() As Boolean
        Dim strInput As String
        Dim strTarget As String
        Dim strWFSize As String = ""
        Dim strASize As String = ""
        Dim strComma As String = CdCst.Sign.DecPoint.Comma
        Dim decWFSize As Decimal
        Dim decASize As Decimal
        Dim decStdSize() As Decimal

        fncInpSizeCheck = False
        Try
            'ロッド先端寸法取得
            strInput = Me.txtRodEndSize.Text.ToUpper

            '漢字が入力されていたらエラー
            If ClsCommon.fncCnvNarrow(strInput) = False Then
                AlertMessage("W0060")
                Exit Function
            End If

            If strInput.Contains("WF") Then
                '入力文字列から"WF"以降の文字を取得する
                strTarget = strInput.Substring(strInput.IndexOf("WF") + 2)

                '入力文字列中に2つ以上"WF"が存在したらエラー
                If strTarget.Contains("WF") Then
                    AlertMessage("W0150")
                    fncInpSizeCheck = False
                    Exit Function
                End If

                '数値部分のチェック
                strWFSize = Me.fncGetNumPart(strTarget)
                If Not IsNumeric(strWFSize) Then
                    AlertMessage("W0130")
                    fncInpSizeCheck = False
                    Exit Function
                End If

                If Left(strWFSize, 1) = "0" Or strWFSize.Contains(strComma) Then
                    AlertMessage("W0150")
                    fncInpSizeCheck = False
                    Exit Function
                End If
            End If

            If strInput.Contains("A") Then
                '入力文字列から"A"以降の文字を取得する
                strTarget = strInput.Substring(strInput.IndexOf("A") + 1)

                '入力文字列中に2つ以上"A"が存在したらエラー
                If strTarget.Contains("A") Then
                    AlertMessage("W0160")
                    fncInpSizeCheck = False
                    Exit Function
                End If

                '数値部分のチェック
                strASize = Me.fncGetNumPart(strTarget)
                If Not IsNumeric(strASize) Then
                    AlertMessage("W0140")
                    fncInpSizeCheck = False
                    Exit Function
                End If
                If Left(strASize, 1) = "0" Or strASize.Contains(strComma) Then
                    AlertMessage("W0160")
                    fncInpSizeCheck = False
                    Exit Function
                End If
            End If

            decStdSize = Me.fncGetStdValue()

            If Len(Trim(strWFSize)) = 0 Then
                decWFSize = decStdSize(0)
            Else
                decWFSize = CDec(strWFSize)
            End If

            If Len(Trim(strASize)) = 0 Then
                decASize = decStdSize(1)
            Else
                decASize = CDec(strASize)
            End If

            'WF寸法とA寸法の合計が最大寸法を超える場合、エラー
            If decWFSize + decASize > decStdSize(2) Then
                AlertMessage("W0120")
                fncInpSizeCheck = False
                Exit Function
            End If
            fncInpSizeCheck = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' 数字の取得
    ''' </summary>
    ''' <param name="strTarget"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetNumPart(ByVal strTarget As String) As String
        Dim sbReturn As New StringBuilder
        Dim strComma As String = CdCst.Sign.DecPoint.Comma
        Dim strDot As String = CdCst.Sign.DecPoint.Dot
        fncGetNumPart = String.Empty
        Try
            For intI As Integer = 0 To strTarget.Length - 1
                Select Case strTarget.Substring(intI, 1)
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", strComma, strDot
                        sbReturn.Append(strTarget.Substring(intI, 1))
                    Case Else
                        Exit For
                End Select
            Next
            fncGetNumPart = sbReturn.ToString
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' 標準の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetStdValue() As Decimal()
        Dim decStdSize(2) As Decimal
        fncGetStdValue = Nothing
        Try

            Select Case strcCompData.strConCaliber
                Case 40
                    decStdSize(0) = 33.5
                    decStdSize(1) = 22
                    If strcCompData.strRodEndPett = "N13" Or strcCompData.strRodEndPett = "N15" Then
                        decStdSize(2) = 133.5
                    Else
                        decStdSize(2) = 155.5
                    End If
                Case 50
                    decStdSize(0) = 37
                    decStdSize(1) = 28
                    If strcCompData.strRodEndPett = "N13" Or strcCompData.strRodEndPett = "N15" Then
                        decStdSize(2) = 137
                    Else
                        decStdSize(2) = 165
                    End If
                Case 63
                    decStdSize(0) = 35
                    decStdSize(1) = 28
                    If strcCompData.strRodEndPett = "N13" Or strcCompData.strRodEndPett = "N15" Then
                        decStdSize(2) = 135
                    Else
                        decStdSize(2) = 163
                    End If
                Case 80
                    decStdSize(0) = 48
                    decStdSize(1) = 36
                    If strcCompData.strRodEndPett = "N13" Or strcCompData.strRodEndPett = "N15" Then
                        decStdSize(2) = 148
                    Else
                        decStdSize(2) = 184
                    End If
                Case 100
                    decStdSize(0) = 53
                    decStdSize(1) = 45
                    If strcCompData.strRodEndPett = "N13" Or strcCompData.strRodEndPett = "N15" Then
                        decStdSize(2) = 153
                    Else
                        decStdSize(2) = 198
                    End If
            End Select
            fncGetStdValue = decStdSize
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function
End Class