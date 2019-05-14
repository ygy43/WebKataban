'************************************************************************************
'*  ProgramID  ：KHPriceQ5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2009/02/02   作成者：T.Yagyu
'*                                      更新日：             更新者：
'*
'*  概要       ：LADシリーズ
'*
'* 変更
'*              二次電池対応             RM1004012 2010/04/23 Y.Miura 
'************************************************************************************
Module KHPriceQ5

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOptionKataban As String = ""
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If Len(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '2010/08/23 MOD RM1008009(9月VerUP) START--->
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 9) & CdCst.Sign.Hypen & _
                'objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 9) & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                '2010/08/23 MOD RM1008009(9月VerUP) <--- END
                decOpAmount(UBound(decOpAmount)) = 1
            End If
            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "B"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            Case "10A", "15A"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 9) & CdCst.Sign.Hypen & _
                                "S-B"
                            Case "20A", "25A"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 9) & CdCst.Sign.Hypen & _
                                "L-B"
                        End Select
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "1"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 9) & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-1"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '二次電池加算
            'RM1004012 2010/04/23 Y.Miura
            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 5 Then
                If objKtbnStrc.strcSelection.strOpSymbol(5) <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If


        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

