'************************************************************************************
'*  ProgramID  ：KHPriceD9
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/07   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スロースタートバルブ
'*             ：Ｖ３３０１／３３２１
'*             ：Ｖ３３０１－Ｗ／ Ｖ３３２１－Ｗ 
'*
'*  更新履歴   ：                       更新日：2008/01/22   更新者：NII A.Takahashi
'*               ・V3301-W/V3321-Wを追加したため、単価見積りロジック変更
'************************************************************************************
Module KHPriceE2

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intOperatePos As Integer
        Dim intElePos As Integer
        Dim intVoltagePos As Integer
        Dim intOptionPos As Integer
        Dim bolWFlg As Boolean
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "W" Then
                bolWFlg = True
                intOperatePos = 3
                intElePos = 4
                intVoltagePos = 5
                intOptionPos = 6
            Else
                bolWFlg = False
                intOperatePos = 2
                intElePos = 3
                intVoltagePos = 4
                intOptionPos = 5
            End If

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            If bolWFlg = True Then
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "W"
            Else
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
            End If
            decOpAmount(UBound(decOpAmount)) = 1

            '手動操作加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(intOperatePos).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(intOperatePos).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '電線接続加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(intElePos).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(intElePos).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '電圧オプション加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(intVoltagePos).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(intVoltagePos).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
