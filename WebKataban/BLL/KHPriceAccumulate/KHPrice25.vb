'************************************************************************************
'*  ProgramID  ：KHPrice25
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*
'*  概要       ：スーパーロッドレスシリンダ　ＳＲＬ２
'*             ：ブレーキ付ロッドレスシリンダ　ＳＲＢ２
'*             ：ガイド付ロッドレスシリンダ　ＳＲＧ
'*             ：ガイド付ロッドレスシリンダ　ＳＲＧ３
'*
'*  更新履歴   ：                       更新日：2009/02/05   更新者：T.Yagyu
'*               ・RM0811134:SRG3機種追加
'************************************************************************************
Module KHPrice25

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)



        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False

        'RM0811134:SRG3 T.Y
        Dim strMountingStyle As String '支持形式
        Dim strBoreSize As String 'チューブ内径
        Dim strPipeThreadType As String '配管ねじ種類 ソース内では未使用
        Dim strCushion As String 'クッション
        Dim strStrokeLen As String 'ストローク
        Dim strSwModelNo As String 'スイッチ形番
        Dim strLeadWireLen As String 'リード線長さ
        Dim strSwQuantity As String 'スイッチ数
        Dim strOption As String 'オプション
        Dim bolC5Flag As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'RM0811134:SRG3 T.Y
            'SRG3のときだけstrOpSbl3（配管ねじ種類）が発生する
            'プログラムを共通化するためにobjKtbnStrc.strcSelection.strOpSymbolの値を変数にセットし利用する
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "SRG3"
                    strMountingStyle = objKtbnStrc.strcSelection.strOpSymbol(1)
                    strBoreSize = objKtbnStrc.strcSelection.strOpSymbol(2)
                    strPipeThreadType = objKtbnStrc.strcSelection.strOpSymbol(3)
                    strCushion = objKtbnStrc.strcSelection.strOpSymbol(4)
                    strStrokeLen = objKtbnStrc.strcSelection.strOpSymbol(5)
                    strSwModelNo = objKtbnStrc.strcSelection.strOpSymbol(6)
                    strLeadWireLen = objKtbnStrc.strcSelection.strOpSymbol(7)
                    strSwQuantity = objKtbnStrc.strcSelection.strOpSymbol(8)
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(9)
                Case Else
                    strMountingStyle = objKtbnStrc.strcSelection.strOpSymbol(1)
                    strBoreSize = objKtbnStrc.strcSelection.strOpSymbol(2)
                    strCushion = objKtbnStrc.strcSelection.strOpSymbol(3)
                    strStrokeLen = objKtbnStrc.strcSelection.strOpSymbol(4)
                    strSwModelNo = objKtbnStrc.strcSelection.strOpSymbol(5)
                    strLeadWireLen = objKtbnStrc.strcSelection.strOpSymbol(6)
                    strSwQuantity = objKtbnStrc.strcSelection.strOpSymbol(7)
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(8)
            End Select

            'RM1306001 2013/06/06
            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            'ストローク取得
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(strBoreSize.Trim), _
                                                  CInt(strStrokeLen.Trim))

            '基本価格キー
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban, 6, 1) = "Q" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                           strBoreSize.Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                'RM1306001 2013/06/05 追加
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & _
                                                           strBoreSize.Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                'RM1306001 2013/06/05 追加
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーションQ加算価格キー
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban, 6, 1) = "Q" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & _
                                                           strBoreSize.Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '支持形式加算価格キー
            If strMountingStyle.Trim <> "00" Then
                Select Case True
                    Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 4, 1) = CdCst.Sign.Hypen
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                                   strMountingStyle.Trim & CdCst.Sign.Hypen & _
                                                                   strBoreSize.Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 5, 1) = CdCst.Sign.Hypen
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                   strMountingStyle.Trim & CdCst.Sign.Hypen & _
                                                                   strBoreSize.Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & _
                                                                   strMountingStyle.Trim & CdCst.Sign.Hypen & _
                                                                   strBoreSize.Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            'スイッチ加算価格キー
            If strSwModelNo.Trim <> "" Then
                Select Case True
                    Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 4, 1) = CdCst.Sign.Hypen
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                                   strSwModelNo.Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                    Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 5, 1) = CdCst.Sign.Hypen
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                   strSwModelNo.Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & _
                                                                   strSwModelNo.Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                End Select

                'リード線長さ加算価格キー
                If strLeadWireLen.Trim <> "" Then
                    Select Case True
                        Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 4, 1) = CdCst.Sign.Hypen
                            Select Case Mid(strSwModelNo.Trim, 4, 1)
                                Case "F", "M"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                                               strLeadWireLen.Trim & CdCst.Sign.Hypen & "FM"
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                                Case "D"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                                               strLeadWireLen.Trim & CdCst.Sign.Hypen & _
                                                                               strSwModelNo.Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                                               strLeadWireLen.Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                            End Select
                        Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 5, 1) = CdCst.Sign.Hypen
                            Select Case Mid(strSwModelNo.Trim, 4, 1)
                                Case "F", "M"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               strLeadWireLen.Trim & CdCst.Sign.Hypen & "FM"
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                                Case "D"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               strLeadWireLen.Trim & CdCst.Sign.Hypen & _
                                                                               strSwModelNo.Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               strLeadWireLen.Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                            End Select
                        Case Else
                            Select Case Mid(strSwModelNo.Trim, 4, 1)
                                Case "F", "M"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & _
                                                                               strLeadWireLen.Trim & CdCst.Sign.Hypen & "FM"
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                                Case "D"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & _
                                                                               strLeadWireLen.Trim & CdCst.Sign.Hypen & _
                                                                               strSwModelNo.Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & _
                                                                               strLeadWireLen.Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwQuantity.Trim)
                            End Select
                    End Select
                End If
            End If

            'オプション・付属品価格キー
            strOpArray = Split(strOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        If Left(strOpArray(intLoopCnt).Trim, 1) = "L" Or Left(strOpArray(intLoopCnt).Trim, 1) = "N" Then
                            Select Case True
                                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 6, 1) = "Q" And Left(strOpArray(intLoopCnt).Trim, 1) = "A"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & "1" & CdCst.Sign.Hypen & _
                                                                               strBoreSize.Trim
                                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 4, 1) = CdCst.Sign.Hypen
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & "1" & CdCst.Sign.Hypen & _
                                                                               strBoreSize.Trim
                                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 5, 1) = CdCst.Sign.Hypen
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & "1" & CdCst.Sign.Hypen & _
                                                                               strBoreSize.Trim
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & "1" & CdCst.Sign.Hypen & _
                                                                               strBoreSize.Trim
                            End Select

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            decOpAmount(UBound(decOpAmount)) = Val(Mid(strOpArray(intLoopCnt).Trim, 2, 1))

                        Else
                            Select Case True
                                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 6, 1) = "Q" And Left(strOpArray(intLoopCnt).Trim, 1) = "A"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               strBoreSize.Trim
                                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 4, 1) = CdCst.Sign.Hypen
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               strBoreSize.Trim
                                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 5, 1) = CdCst.Sign.Hypen
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               strBoreSize.Trim
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               strBoreSize.Trim
                            End Select

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1

                        End If
                End Select
                'RM1306001 2013/06/05 追加
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
