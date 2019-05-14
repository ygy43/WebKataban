'************************************************************************************
'*  ProgramID  ：KHPriceG1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ロボットシリンダ　ＭＦＣ
'*
'*  更新履歴   ：原価積算C5追加         更新日：2016/10/14   更新者：K.Ohwaki　RM1610011
'************************************************************************************
Module KHPriceG1

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0) 'RM1610011 追加　C5加算
            Dim bolC5Flag As Boolean 'RM1610011 追加　C5加算

            'C5加算　RM1610011 追加　
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            'ストローク取得
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim))

            '基本価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("K") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1) 'RM1610011 追加　C5加算
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-K-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1) 'RM1610011 追加　C5加算
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'C5加算　RM1610011 追加　
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If

            'バリエーション「B」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("B") >= 0 And _
               objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("BS") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1) 'RM1610011 追加　C5加算
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-B-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'バリエーション「BS」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("BS") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1) 'RM1610011 追加　C5加算
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-BS-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    Case "A", "B"
                        '端子箱
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1) 'RM1610011 追加　C5加算
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1) 'RM1610011 追加　C5加算
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                End Select

                '取り付けレール長さ
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1) 'RM1610011 追加　C5加算
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-RAIL-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'リード線長さ加算価格キー
                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    Case "3", "5"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1) 'RM1610011 追加　C5加算
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(1)-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                End Select
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
