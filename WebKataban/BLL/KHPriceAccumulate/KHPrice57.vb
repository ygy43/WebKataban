'************************************************************************************
'*  ProgramID  ：KHPrice57
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線マニホールド　Ｍ４ＴＢ３／Ｍ４ＴＢ４
'*
'************************************************************************************
Module KHPrice57

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStationQty As Integer = 0
        Dim intQuantity As Integer = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '配線方式が選択されている時は、ブロックマニホールドの引当(要素選択画面→仕様書入力画面→単価積上画面)
            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                'バルブブロック連数
                intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)

                For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                       objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                            Case CdCst.Manifold.InspReportJp.Japanese, CdCst.Manifold.InspReportJp.English, _
                                 CdCst.Manifold.InspReportEn.Japanese, CdCst.Manifold.InspReportEn.English
                                '加算なし
                            Case Else
                                Select Case intLoopCnt
                                    Case 1 To 2
                                        'エンドブロック
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 4 To 8
                                        '電磁弁付バルブブロック
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, 7) & "-L"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 9 To 10
                                        'MPV付バルブブロック
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, 9)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 3, 11 To 15
                                        '配線ブロック,Sレギュレータ(P,A,B),単独給・排気スペーサ
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 16
                                        '仕切プラグ(P)
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 17
                                        '仕切プラグ(R)
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-R"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    Case 18 To 19, 22
                                        'サイレンサ(樹脂,メタル),ケーブルクランプ
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case 20
                                        '六角穴付プラグ(上)
                                        If objKtbnStrc.strcSelection.strSeriesKataban = "M4TB3" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & "-PLUG-R1/4"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & "-PLUG-R3/8"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    Case 21
                                        '六角穴付プラグ(下)
                                        If objKtbnStrc.strcSelection.strSeriesKataban = "M4TB3" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & "-PLUG-R3/8"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & "-PLUG-R1/2"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                End Select

                                '電磁弁ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ部の時,
                                If (Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) = "M4TB3" Or _
                                    Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) = "M4TB4") And _
                                    Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) >= "0" And _
                                    Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) <= "9" Then
                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                        intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    End If
                                End If
                        End Select
                    End If
                Next
            Else
                '単体基本価格キーのみ作成
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-L"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '手動装置加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                    decOpAmount(UBound(decOpAmount)) = intQuantity
                Else
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = 1 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                End If
            End If

            '表示・保護回路減算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "L" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           "-MINUS-L"

                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                    decOpAmount(UBound(decOpAmount)) = intQuantity
                Else
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = 1 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                End If
            End If

            'その他オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "K", "A"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                            decOpAmount(UBound(decOpAmount)) = intStationQty
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "P"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim & "O"
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                            decOpAmount(UBound(decOpAmount)) = intStationQty
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "CL", "CR"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '切削油対応加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                    decOpAmount(UBound(decOpAmount)) = intStationQty
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '電圧加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "" Then
                strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                Select Case strStdVoltageFlag
                    Case CdCst.VoltageDiv.Standard
                    Case CdCst.VoltageDiv.Options
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OPT"
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                decOpAmount(UBound(decOpAmount)) = 2
                            End If
                        End If
                    Case CdCst.VoltageDiv.Other
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OTH"
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                decOpAmount(UBound(decOpAmount)) = 2
                            End If
                        End If
                End Select
            End If

            'ケーブル長さ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(11).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
