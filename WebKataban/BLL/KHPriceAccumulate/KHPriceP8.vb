'************************************************************************************
'*  ProgramID  ：KHPriceP8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2008/06/10   作成者：M.Kojima
'*
'*  概要       ：軽量クランプシリンダ　CACシリーズ
'*
'************************************************************************************
Module KHPriceP8
    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '基本価格キーの設定
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = _
                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '付属品加算
            If (objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "") Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Length <> 0 Then

                'RM1801025_オプション追加対応
                'タイロッド方式
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        "TIEROD" & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '取付部品加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                        "TIEROD"
                    decOpAmount(UBound(decOpAmount)) = 1

                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                End If

                'リード線長さ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T2YD" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & "T2YD" & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim

                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T2YDT" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                         objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                         CdCst.Sign.Hypen & "T2YDT" & CdCst.Sign.Hypen & _
                         objKtbnStrc.strcSelection.strOpSymbol(5).Trim

                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T2JH" Or _
                    objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T2JV" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                         objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                         CdCst.Sign.Hypen & "T2J" & CdCst.Sign.Hypen & _
                         objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                         objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                         CdCst.Sign.Hypen & "T" & CdCst.Sign.Hypen & _
                         objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub
End Module
