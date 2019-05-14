'************************************************************************************
'*  ProgramID  ：KHPrice04
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*
'*  概要       ：低圧損形エアオペレイト式２ポート弁
'*             ：ＣＶＥ２／ＣＶＳＥ２シリーズ
'*
'*  更新履歴   ：                       更新日：2008/11/06   更新者：T.Sato
'*  ・受付No.RM0810052 電磁弁上面搭載オプション（Ｔ）追加
'* 
'************************************************************************************
Module KHPrice09

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '価格キー設定
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'コイルオプション加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "CVE2", "CVSE2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "", "2C"
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Else
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "", "2G"
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case Else
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "", "2G"
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

            'その他オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "B", "B2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "S", "ST"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-S"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '電圧加算価格キー
            Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4)
                Case "CVSE"
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                        '電圧区分取得
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                        '標準電圧以外の場合は電圧加算
                        If strStdVoltageFlag <> CdCst.VoltageDiv.Standard Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OTHER-VOL"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
