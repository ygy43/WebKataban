'************************************************************************************
'*  ProgramID  ：KHPriceP7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2008/06/10   作成者：M.Kojima
'*
'*  概要       ：メカニカルパワーシリンダ　MCPシリーズ
'*
'************************************************************************************
Module KHPriceP7
    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim strSuiryoku As String '推力
        Dim strStroke As String 'ストローク
        Dim strLead As String 'リード線長さ

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            strSuiryoku = objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            strStroke = objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            strLead = objKtbnStrc.strcSelection.strOpSymbol(5).Trim

            '基本価格キーの設定
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            If (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "MCP-W") Then
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & "00" & CdCst.Sign.Hypen & _
                    strSuiryoku & CdCst.Sign.Hypen & _
                    strStroke
            ElseIf (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "MCP-S") Then
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & "00" & CdCst.Sign.Hypen & _
                    strSuiryoku
            End If
            decOpAmount(UBound(decOpAmount)) = 1

            'FA加算
            If (objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "FA") Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & "FA" & CdCst.Sign.Hypen & _
                    strSuiryoku
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'ロッド先端おねじ(N)加算
            If (objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Length <> 0) Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & _
                    strSuiryoku & _
                    CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Length <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                    strSuiryoku & CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'リード線長さ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                        CdCst.Sign.Hypen & _
                        strLead
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub
End Module
