'************************************************************************************
'*  ProgramID  ：KHPriceB6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/09   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：真空エジェクタユニットマニホールド
'*             ：真空切替ユニットマニホールド
'*
'************************************************************************************
Module KHPriceB6

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim objOption As New KHOptionCtl
        Dim intLoopCnt As Integer
        Dim strOption() As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ミックス構成が選択されているかチェックする
            If objOption.fncVaccumMixCheck(objKtbnStrc) = True Then
                '真空エジェクタ・真空切換ユニット価格
                For intLoopCnt = 1 To 8
                    If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                        strOption = Split(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), CdCst.Sign.Delimiter.Comma)

                        '機種毎に価格キーを設定する
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "VSKM"
                                'VSKM-**A
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**" & _
                                                                           strOption(4).Trim

                                '真空ポート
                                If Left(strOption(6).Trim, 1) = "T" Then
                                    'VSKM-**A-T*
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-T*"
                                Else
                                    'VSKM-**A-*"
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                                End If

                                'バルブタイプ
                                If strOption(8).Trim <> "" Then
                                    'VSKM-**A-*"
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-**"
                                End If

                                '真空センサ仕様
                                Select Case strOption(4).Trim
                                    Case "E", "F", "L", "M", "R", "W"
                                        'VSKM-**A-*"
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                                End Select

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            Case "VSJM"
                                'VSXM-***-**S-*
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-***-**" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "-*"

                                '真空センサ仕様
                                If strOption(8).Trim <> "" Then
                                    ' VSXM-***-**S-*-*
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                                End If

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            Case "VSNM"
                                '真空センサ仕様
                                If strOption(7).Trim <> "" Then
                                    'VSNM-**-****-3-V1
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               "-**-****-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                               CdCst.Sign.Hypen & _
                                                                               strOption(7).Trim
                                Else
                                    'VSNM-**-***S-3
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-***" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                End If

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            Case "VSXM"
                                'VSJM-***-*-*
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-***-*-*"

                                '真空センサ仕様
                                If strOption(8).Trim <> "" Then
                                    'VSJM-***-*-*-DW
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                               strOption(8).Trim
                                End If

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            Case "VSZM"
                                'VSZM-**-*
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-*"

                                '真空センサ仕様
                                If strOption(8).Trim <> "" Then
                                    'VSZM-**-*-DW
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                               strOption(8).Trim
                                End If

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            Case "VSJPM"
                                'VSJP-****-*
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-****-*"

                                '真空センサ仕様
                                If strOption(5).Trim <> "" Then
                                    'VSJP-****-*
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                                End If

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            Case "VSNPM"
                                '真空センサ仕様
                                If strOption(4).Trim <> "" Then
                                    'VSNPM-***-3-V1
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               "-***-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                               CdCst.Sign.Hypen & _
                                                                               strOption(4).Trim
                                Else
                                    'VSNPM-***-3
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               "-***-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                End If

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            Case "VSXPM"
                                'VSXPM-D*-*
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOption(2).Trim & "*-*"

                                '真空センサ仕様
                                If strOption(5).Trim <> "" Then
                                    'VSXPM-D*-*-DW
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                               strOption(5).Trim
                                End If

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            Case "VSZPM"
                                'VSZPM-*
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-*"

                                '真空センサ仕様
                                If strOption(4).Trim <> "" Then
                                    'VSZPM-*-DW
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                               strOption(4).Trim
                                End If

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                        End Select
                    End If
                Next
                'マスキングブロック価格
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "VSKM"
                        For intLoopCnt = 9 To 10
                            If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "VSKM-MB-*"
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If
                        Next
                End Select
            Else
                ' 機種毎に価格キーを設定
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "VSKM"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim

                        '真空ポート
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "T" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-T*"
                        Else
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                        End If

                        'バルブタイプ
                        If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-**"
                        End If

                        '真空センサ仕様
                        If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                    Case "VSJM"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-***-**" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "-*"

                        '真空センサ仕様
                        If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                    Case "VSNM"
                        '真空センサ仕様
                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       "-**-****-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                       CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       CdCst.Sign.Hypen & _
                                                                       "**" & CdCst.Sign.Hypen & "***" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim & _
                                                                       CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                    Case "VSXM"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-***-*-*"

                        '真空センサ仕様
                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                    Case "VSJPM"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-****-*"

                        '真空センサ仕様
                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)

                    Case "VSNPM"
                        '真空センサ仕様
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       "-***-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                       CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       CdCst.Sign.Hypen & _
                                                                       "***" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)

                    Case "VSXPM"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "*-*"

                        '真空センサ仕様
                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    Case "VSZM"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-*"

                        '真空センサ仕様
                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                    Case "VSZPM"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-*"

                        '真空センサ仕様
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                End Select
            End If

            ' バルブユニット価格加算キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "VSZM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-V*-3"
                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                Case "VSZPM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-V-3"
                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
            End Select

            ' マニホールド単体加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "VSKM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-***-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSJM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-MANIHOLD-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSNM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-MANIHOLD"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSXM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSJPM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-MANIHOLD-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSNPM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-MANIHOLD"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSXPM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSZM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim & "-**"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSZPM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim & "-**"
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
