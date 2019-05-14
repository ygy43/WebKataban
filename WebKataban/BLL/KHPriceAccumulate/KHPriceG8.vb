'************************************************************************************
'*  ProgramID  ：KHPriceG8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/06   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ファインレベルスイッチ　ＭＸＫＭＬ
'*
'************************************************************************************
Module KHPriceG8

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intLoopCnt As Integer

        Dim intKML50Qty As Integer = 0
        Dim intKML604Qty As Integer = 0
        Dim intKML606Qty As Integer = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'マニホールドスイッチ配列(1)～(5)個数集計
            For intLoopCnt = 3 To 7
                Select Case objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                    Case "0"
                    Case "1"
                        intKML50Qty = intKML50Qty + 1
                    Case "4"
                        intKML604Qty = intKML604Qty + 1
                    Case "6"
                        intKML606Qty = intKML606Qty + 1
                End Select
            Next

            '混合搭載するKML50形番の判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "00-0"
                    'KML60-4-0単品価格キー
                    If intKML604Qty <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "KML60-4-0"
                        decOpAmount(UBound(decOpAmount)) = intKML604Qty
                    End If

                    'KML60-6-0単品価格キー
                    If intKML606Qty <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "KML60-6-0"
                        decOpAmount(UBound(decOpAmount)) = intKML606Qty
                    End If

                    'マニホールド用取付板価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "KML60-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "REN"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    '1点式KML50混合搭載基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "REN-" & _
                                                               intKML50Qty.ToString & "MIX-KML50-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "-0"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'KML60-4-0単品価格キー
                    If intKML604Qty <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "KML60-4-0"
                        decOpAmount(UBound(decOpAmount)) = intKML604Qty
                    End If

                    'KML60-6-0単品価格キー
                    If intKML606Qty <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "KML60-6-0"
                        decOpAmount(UBound(decOpAmount)) = intKML606Qty
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
