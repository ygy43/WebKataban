'************************************************************************************
'*  ProgramID  ：KHPriceB1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/23   作成者：NII K.Sudoh
'*
'*  概要       ：タイトシリンダ　ＣＭＫ２　両ロッドベース
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'*                                      更新日：2008/03/24   更新者：T.Sato
'*               ・受付No.RM0707057 CMK2ロッド先端特注対応
'*  　             WFサイズから単価キーを作成するロジックを追加
'************************************************************************************
Module KHPriceB1

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer
        Dim bolC5Flag As Boolean

        Dim decLength As Decimal
        Dim decWFLength As Decimal
        Dim strStdWFLength As String = String.Empty

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            'ストローク取得
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim))

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-BASE-D-" & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If

            'バリエーション加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    Case "DQ", "DC", "DZ", "DH", "DT", _
                         "DT2", "DO", "DG", "DG1", "DG4"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-" & _
                                                                   Mid(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2, 2) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    Case "DG2", "DG3"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-" & _
                                                                   Mid(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2, 2) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   intStroke.ToString
                    Case "DM"
                        Select Case True
                            Case CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 300
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-M-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-25-300"
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-M-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-301-500"
                        End Select
                End Select

                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    Case "DG", "DG1", "DG2", "DG3", "DG4"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '支持形式加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "00", "LB", "LS", "FA", "FB"
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SUPPORT-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'クッション・ねじ加算価格キー
            'ゴムエアクッション加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                Case "C", "GC", "NC"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-*C-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)

                'リード線長さ加算価格キー
                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    Case "3", "5"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "T2H", "T2V", "T2YH", "T2YV", "T3H", _
                                 "T3V", "T3YH", "T3YV", "T0H", "T0V", _
                                 "T5H", "T5V", "T1H", "T1V", "T8H", "T8V", _
                                 "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(1)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                                 "T2YMV", "T3YMH", "T3YMV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(2)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            Case "T2JH", "T2JV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(3)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(4)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        End Select
                End Select
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "J", "L", "M"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   intStroke.ToString
                        decOpAmount(UBound(decOpAmount)) = 2
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "P6"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 2
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "A2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If

                        If objKtbnStrc.strcSelection.strRodEndOption.Trim = "" Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "P7", "P71"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-P7*-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 2
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "M0", "M1"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "F", "FE"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("S") >= 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "S"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If
                End Select
            Next

            '付属品加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "IY"
                        'I加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-I"
                        decOpAmount(UBound(decOpAmount)) = 1
                        'Y加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-Y"
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case "FP1"
                        '食品製造工程向け商品
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                  strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                   strOpArray(intLoopCnt).Trim

                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I", "Y"
                                'I/Yの場合はロッド先端特注選択の有無により加算変更
                                If objKtbnStrc.strcSelection.strRodEndOption.Trim = "" Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next


            '付属品加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "IY"
                        'I加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-I"
                        decOpAmount(UBound(decOpAmount)) = 1
                        'Y加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-Y"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                   strOpArray(intLoopCnt).Trim

                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I", "Y"
                                'I/Yの場合はロッド先端特注選択の有無により加算変更
                                If objKtbnStrc.strcSelection.strRodEndOption.Trim = "" Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next


            'ロッド先端オーダーメイド加算価格キー
            If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                If InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") = 0 Then
                    decWFLength = 1
                Else
                    For intLoopCnt = InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") + 2 To Len(objKtbnStrc.strcSelection.strRodEndOption.Trim)
                        If Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "0" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "1" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "2" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "3" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "4" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "5" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "6" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "7" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "8" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "9" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "." Then
                            If intLoopCnt = Len(objKtbnStrc.strcSelection.strRodEndOption.Trim) Then
                                decLength = intLoopCnt - (InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") + 2) + 1
                            End If
                        Else
                            decLength = intLoopCnt - (InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") + 2) + 1
                            Exit For
                        End If
                    Next

                    decWFLength = CDec(Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") + 2, decLength)) - objKtbnStrc.strcSelection.strRodEndWFStdVal
                End If

                Select Case True
                    Case 1 <= decWFLength And decWFLength <= 100
                        strStdWFLength = "100"
                    Case 101 <= decWFLength And decWFLength <= 200
                        strStdWFLength = "200"
                    Case 201 <= decWFLength And decWFLength <= 300
                        strStdWFLength = "300"
                    Case 301 <= decWFLength And decWFLength <= 400
                        strStdWFLength = "400"
                    Case 401 <= decWFLength
                        strStdWFLength = "500"
                End Select

                '↓RM1111020  2011/12/01 Y.Tachi  修正
                If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "N13-N11" Then
                    If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "N11-N13" Then

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-TIP-OF-ROD-" & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & strStdWFLength
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If

                    End If
                End If
                '↑RM1111020  2011/12/01 Y.Tachi  修正
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
