'************************************************************************************
'*  ProgramID  ：KHPrice18
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/24   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＡＰＫ／ＡＤＫ
'*
'*  修正履歴   ：
'*                                      更新日：2008/03/27   更新者：NII A.Takahashi
'*  　・G/NPTねじ追加により、ロジック変更(ねじ加算対応)
'*    ・二次電池対応                     RM1004012 2010/04/23 Y.Miura
'************************************************************************************
Module KHPrice18

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)


        Dim strStdVoltageFlag As String
        Dim strOpArray() As String
        Dim strOption As String
        Dim strScrewType As String
        Dim bolScrew As Boolean
        Dim intLoopCnt As Integer
        Dim intOptionPos As Integer = 0     'RM1004012

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '機種判定
            '二次電池対応機種は電圧オプション位置以降を+1する
            'RM1004012 2010/05/01 Y.Miura
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "ADK11"
                    intOptionPos = 1
                Case "APK21"
                    If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                        intOptionPos = 1
                    Else
                        intOptionPos = 0
                    End If
                Case Else
                    intOptionPos = 0
            End Select

            'ねじ判定
            If InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "G") <> 0 Or _
               InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "N") <> 0 Then
                strScrewType = Right(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1)
                bolScrew = True
            Else
                strScrewType = ""
                bolScrew = False
            End If

            '基本価格キー
            Select Case True
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "H"
                    strOption = "0"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "J"
                    strOption = "B"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "K"
                    strOption = "C"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "L"
                    strOption = "D"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "M"
                    strOption = "E"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "N"
                    strOption = "F"
                Case Else
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            End Select

            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "W"    '屋外シリーズ
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & strOption & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション価格
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           strOpArray(intLoopCnt).Trim & "-W"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    If bolScrew Then
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, (InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1).Trim, strScrewType)) - 1) & _
                                                                   CdCst.Sign.Hypen & strOption
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & strOption
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                    'ボディシール加算
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "0", "B", "C", "D", "E", "F"
                        Case "H"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "HK"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'コイルハウジング加算
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "2" Then
                        'RM1004012 2010/04/23 Y.Miura
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 2) & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(5 + intOptionPos).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
           
                    'オプション価格
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "S"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Case "2C", "3A", "4A"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "S0"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                    '二次電池P4加算
                    'RM1004012 2010/04/23 Y.Miura
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "ADK11"
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "APK21"
                            If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select

                    '電圧加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5 + intOptionPos).Trim <> "" Then
                        '電圧取得
                        '2010/08/27 ADD RM0808112(異電圧対応) START--->
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5 + intOptionPos).Trim, _
                                                                       strCountryCd, strOfficeCd)
                        'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(5 + intOptionPos).Trim)
                        '2010/08/27 ADD RM0808112(異電圧対応) <--- END
                        If strStdVoltageFlag <> CdCst.VoltageDiv.Standard Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & Left(objKtbnStrc.strcSelection.strOpSymbol(5 + intOptionPos).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If

            End Select

            'ねじ加算価格キー
            'If bolScrew Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = "MULTI-SCREW-" & Right(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1)
            '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
            '    decOpAmount(UBound(decOpAmount)) = 2
            'End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
