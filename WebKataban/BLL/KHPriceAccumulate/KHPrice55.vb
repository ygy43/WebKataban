'************************************************************************************
'*  ProgramID  ：KHPrice55
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ブロックマニホールド用電磁弁単体　Ｎ４ＫＢ１／Ｎ４ＫＢ２
'*
'************************************************************************************
Module KHPrice55

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)



        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intQuantity As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '数量設定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "1"
                    intQuantity = 1
                Case "2"
                    intQuantity = 2
                Case "3"
                    intQuantity = 2
                Case "4"
                    intQuantity = 2
                Case "5"
                    intQuantity = 2
            End Select

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '手動装置加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = intQuantity
            End If

            '電線接続加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = intQuantity
            End If

            'その他オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                End Select
            Next

            '切削油対応仕様加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '電圧加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                '2010/08/31 ADD RM0808112(異電圧対応) START--->
                strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                                               strCountryCd, strOfficeCd)
                'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                '2010/08/31 ADD RM0808112(異電圧対応) <---END

                Select Case strStdVoltageFlag
                    Case CdCst.VoltageDiv.Standard
                    Case CdCst.VoltageDiv.Options
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "OPT"
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                    Case CdCst.VoltageDiv.Other
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "OTH"
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                End Select
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
