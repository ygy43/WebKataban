'************************************************************************************
'*  ProgramID  ：KHPriceM5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ショックキラー　ＮＣＫ
'*
'************************************************************************************
Module KHPriceM5

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim bolOptionN1 As Boolean = False
        Dim bolOptionC As Boolean = False
        Dim bolC5Flag As Boolean

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'C5チェック                      
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)

            'オプション選択判定
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "N1"
                        bolOptionN1 = True
                    Case "C"
                        bolOptionC = True
                End Select
            Next

            '基本価格キー
            If bolOptionC = True Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-00-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-C"
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-00-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If

            '支持形式FA(フランジ金具)加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "FA" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'ストップナット＋六角ナット加算価格キー
            If bolOptionN1 = True Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-N1"
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'RM1001045 2010/02/24 Y.Miura 二次電池機器追加
            '二次電池加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-00-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-OP-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If
        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
