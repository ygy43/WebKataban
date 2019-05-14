'************************************************************************************
'*  ProgramID  ：KHPrice02
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：直動式２ポート弁　ＧＡＢ／ＧＡＧ
'*
'*  修正履歴   ：
'*                                      更新日：2008/03/27   更新者：NII A.Takahashi
'*  　・G/NPTねじ追加により、ロジック変更(ねじ加算対応)
'*    ・二次電池対応                     RM1004012 2010/04/22 Y.Miura
'************************************************************************************
Module KHPrice02

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)



        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim intStationQty As Integer = 0
        Dim bolOptionZ As Boolean = False
        Dim bolScrew As Boolean
        Dim intOptionPos As Integer = 0     'RM1004012

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '機種判定
            '二次電池対応機種は電圧オプション位置以降を+1する
            'RM1004012 2010/04/22 Y.Miura
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "GAB422", "GAB462"
                    intOptionPos = 0
                Case Else
                    intOptionPos = 1
            End Select

            'オプション選択判定
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "Z"
                        bolOptionZ = True
                End Select
            Next

            'ねじ判定
            If InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "G") <> 0 Or _
               InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "N") <> 0 Then
                bolScrew = True
            Else
                bolScrew = False
            End If

            '数量セット
            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Or _
                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "0" Then
                intStationQty = 1
            Else
                intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
            End If

            '基本価格キー
            If bolOptionZ = True Then
                'ドライエア用基本価格
                If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_07Z-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = intStationQty
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_08Z-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = intStationQty
                End If
            Else
                If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "0" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) & "-0"
                        decOpAmount(UBound(decOpAmount)) = intStationQty
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) & "-0-" & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1)
                        decOpAmount(UBound(decOpAmount)) = intStationQty
                    End If
                Else
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "0" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) & CdCst.Sign.Hypen & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) & "-0"
                        decOpAmount(UBound(decOpAmount)) = intStationQty
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) & CdCst.Sign.Hypen & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) & "-0-" & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1)
                        decOpAmount(UBound(decOpAmount)) = intStationQty
                    End If
                End If
            End If

            'コイルハウジング加算価格キー
            If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1) <> "" Then
                If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1) = "2" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    'RM1004012 2010/04/22 Y.Miura
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_04" & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(9 + intOptionPos).Trim, 2)
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_05" & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(9 + intOptionPos).Trim, 2)
                    End If
                    decOpAmount(UBound(decOpAmount)) = intStationQty
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_04" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_05" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = intStationQty
                End If
            End If

            '手動操作加算価格キー
            If Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 1) <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_05" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Else
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_06" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                End If
                decOpAmount(UBound(decOpAmount)) = intStationQty
            End If

            'オプション加算(1)価格キー
            If Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_06" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                Else
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_07" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                End If
                decOpAmount(UBound(decOpAmount)) = intStationQty
            End If

            'オプション加算(2)価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "S"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "" Or _
                           Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "00" Or _
                           Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "3A" Or _
                           Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "4A" Or _
                           Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "6C" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_07" & _
                                                                           strOpArray(intLoopCnt).Trim & "0"
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_08" & _
                                                                           strOpArray(intLoopCnt).Trim & "0"
                            End If
                            decOpAmount(UBound(decOpAmount)) = intStationQty
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_07" & _
                                                                           strOpArray(intLoopCnt).Trim
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_08" & _
                                                                           strOpArray(intLoopCnt).Trim
                            End If
                            decOpAmount(UBound(decOpAmount)) = intStationQty
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_07" & _
                                                                       strOpArray(intLoopCnt).Trim
                        Else
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_08" & _
                                                                       strOpArray(intLoopCnt).Trim
                        End If
                        decOpAmount(UBound(decOpAmount)) = intStationQty
                End Select
            Next

            'オプション２　Ｐ４加算　二次電池対応機器
            'RM1004012 2010/04/22 Y.Miura
            If intOptionPos > 0 Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case "P4", "P40"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next
            End If

            '電圧加算価格キー
            '2010/08/25 MOD RM0808112(海外異電圧削除対応) START --->
            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                           objKtbnStrc.strcSelection.strOpSymbol(9 + intOptionPos).Trim, _
                                                           strCountryCd, strOfficeCd)
            ''RM1004012 2010/04/22 Y.Miura　+1
            'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
            '                                               objKtbnStrc.strcSelection.strOpSymbol(9 + intOptionPos).Trim)
            '2010/08/25 MOD RM0808112(海外異電圧削除対応) <--- END

            Select Case strStdVoltageFlag
                Case CdCst.VoltageDiv.Standard
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_08" & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(9 + intOptionPos).Trim, 2)
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_09" & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(9 + intOptionPos).Trim, 2)
                    End If
                    decOpAmount(UBound(decOpAmount)) = intStationQty
            End Select

            '電圧オプション加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                Case "D", "E", "F", "R", "L", "M", "N"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "00" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "2E" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "2G" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "2H" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "STG" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "STO" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case Else
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "00" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "2E" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "2G" Or _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "2H" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "ODG" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "ODO" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

            'ねじ加算価格キー
            If bolScrew Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "MULTI-SCREW-" & Right(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1)
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
                If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "B" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "0" Then
                        decOpAmount(UBound(decOpAmount)) = 0
                    Else
                        decOpAmount(UBound(decOpAmount)) = intStationQty + 2
                    End If
                Else
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "0" Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = intStationQty * 2 + 2
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
