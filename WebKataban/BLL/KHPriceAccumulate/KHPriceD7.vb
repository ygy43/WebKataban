'************************************************************************************
'*  ProgramID  ：KHPriceD7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ピンクランプシリンダ　ＰＣＣ
'*
'************************************************************************************
Module KHPriceD7

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim))

            '基本価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "PCC"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'ロケートピン形状加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)

                'リード線長さ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                End If
            End If

            'シム加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "2" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'スパッタ対策加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
