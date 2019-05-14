Module KHNewHandleCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  ニューハンドリングシステム＆ハイブリロボチェック
    '*【概要】
    '*  ニューハンドリングシステム＆ハイブリロボをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Try

            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "NSR"
                    Dim intNSRStroke As Integer

                    'ヘッド数が２ケのとき
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2" Then
                        '可搬質量判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            Case "10", "15"
                                'ヘッド間ピッチの範囲判定
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 200 And _
                                   CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 999 Then
                                Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8260"
                                    fncCheckSelectOption = False
                                End If
                            Case "30"
                                'ヘッド間ピッチの範囲判定
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 260 And _
                                   CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 999 Then
                                Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8310"
                                    fncCheckSelectOption = False
                                End If
                            Case "50"
                                'ヘッド間ピッチの範囲判定
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 310 And _
                                   CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 999 Then
                                Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8300"
                                    fncCheckSelectOption = False
                                End If
                        End Select

                        'Ｘ軸ストローク＋ヘッド間ピッチ（１ｍｍピッチ）の値を求める
                        intNSRStroke = CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                        'X軸＋ヘッド間ピッチの合計値判定
                        If intNSRStroke <= 2000 Then
                        Else
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W8270"
                            fncCheckSelectOption = False
                        End If
                    End If
                Case "NHS"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "H"
                            Dim intNHSStroke As Integer

                            'Ｘ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 50 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 2000 Then
                            Else
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8280"
                                fncCheckSelectOption = False
                            End If

                            'Ｚ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) >= 10 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 600 Then
                            Else
                                intKtbnStrcSeqNo = 5
                                'RM1210067 2013/02/01 Y.Tachi ローカル版との差異修正(W8250→W8350)
                                strMessageCd = "W8350"
                                fncCheckSelectOption = False
                            End If

                            'ヘッド数が２ケのとき
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2" Then
                                '可搬質量判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "1005", "1007", "1505", "1507", "1510", "1512"
                                        'ヘッド間ピッチの範囲判定
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 200 And _
                                           CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                        Else
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8260"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "3010", "3012", "3020"
                                        'ヘッド間ピッチの範囲判定
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 269 And _
                                           CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                        Else
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8340"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "5010", "5012", "5020", "5033"
                                        'ヘッド間ピッチの範囲判定
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 310 And _
                                           CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                        Else
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8300"
                                            fncCheckSelectOption = False
                                        End If
                                End Select

                                'Ｘ軸ストローク＋ヘッド間ピッチ（１ｍｍピッチ）の値を求める
                                intNHSStroke = CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                                'X軸＋ヘッド間ピッチの合計値判定
                                If intNHSStroke <= 2000 Then
                                Else
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W8270"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        Case "C"
                            Dim intNHSStroke As Integer

                            'ヘッド数が２ケのとき
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2" Then
                                '可搬質量判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "1004", "1006", "1504", "1506", "1510", "1512"
                                        'ヘッド間ピッチの範囲判定
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 200 And _
                                           CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                        Else
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8260"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "3010", "3012"
                                        'ヘッド間ピッチの範囲判定
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 260 And _
                                           CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                        Else
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8310"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "5010", "5012"
                                        'ヘッド間ピッチの範囲判定
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 310 And _
                                           CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                        Else
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8300"
                                            fncCheckSelectOption = False
                                        End If
                                End Select

                                'Ｘ軸ストローク＋ヘッド間ピッチ（１ｍｍピッチ）の値を求める
                                intNHSStroke = CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                                'X軸＋ヘッド間ピッチの合計値判定
                                If intNHSStroke <= 2000 Then
                                Else
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W8270"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        Case "S"
                            Dim intNHSStroke As Integer

                            'Ｘ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 50 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 2000 Then
                            Else
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8280"
                                fncCheckSelectOption = False
                            End If

                            'Ｚ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) >= 30 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 200 Then
                            Else
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8340"
                                fncCheckSelectOption = False
                            End If

                            'Ｚ軸ストロークのストローク１ｍｍ毎製作不可の判定
                            Select Case Right(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1)
                                Case "0", "5"
                                Case "1", "2", "3", "4", "6", "7", "8", "9"
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8320"
                                    fncCheckSelectOption = False
                            End Select

                            'ヘッド数が２ケのとき
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2" Then
                                '可搬質量判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "1003", "1503", "1507", "1512"
                                        'ヘッド間ピッチの範囲判定
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 200 And _
                                           CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                        Else
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8260"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "3007", "3012"
                                        'ヘッド間ピッチの範囲判定
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 260 And _
                                           CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                        Else
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8310"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "5007", "5012", "5033"
                                        'ヘッド間ピッチの範囲判定
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 310 And _
                                           CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                        Else
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8270"
                                            fncCheckSelectOption = False
                                        End If
                                End Select

                                'Ｘ軸ストローク＋ヘッド間ピッチ（１ｍｍピッチ）の値を求める
                                intNHSStroke = CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                                'X軸＋ヘッド間ピッチの合計値判定
                                If intNHSStroke <= 2000 Then
                                Else
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W8270"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        Case "L"
                            Dim intNHSStroke As Integer

                            'Ｘ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 50 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 2000 Then
                            Else
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8280"
                                fncCheckSelectOption = False
                            End If

                            'Ｚ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) >= 15 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 100 Then
                            Else
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8290"
                                fncCheckSelectOption = False
                            End If

                            'ヘッド数が２ケのとき
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2" Then
                                'ヘッド間ピッチの範囲判定
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 200 And _
                                   CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <= 999 Then
                                Else
                                    intKtbnStrcSeqNo = 6
                                    strMessageCd = "W8260"
                                    fncCheckSelectOption = False
                                End If

                                'Ｘ軸ストローク＋ヘッド間ピッチ（１ｍｍピッチ）の値を求める
                                intNHSStroke = CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                                'X軸＋ヘッド間ピッチの合計値判定
                                If intNHSStroke <= 2000 Then
                                Else
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W8270"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select
                Case "HR"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "B"
                            'Ｒ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 30 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 400 Then
                            Else
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8230"
                                fncCheckSelectOption = False
                            End If

                            'Ｚ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 30 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 300 Then
                            Else
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8240"
                                fncCheckSelectOption = False
                            End If
                        Case "G"
                            'Ｘ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 30 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 400 Then
                            Else
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8180"
                                fncCheckSelectOption = False
                            End If

                            'Ｚ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 30 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 300 Then
                            Else
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8190"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "HRL"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "A"
                            'Ｒ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 25 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 300 Then
                            Else
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8250"
                                fncCheckSelectOption = False
                            End If

                            'Ｚ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) >= 25 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 150 Then
                            Else
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8210"
                                fncCheckSelectOption = False
                            End If
                        Case "S"
                            'Ｘ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 25 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 300 Then
                            Else
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8220"
                                fncCheckSelectOption = False
                            End If

                            'Ｚ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) >= 25 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 150 Then
                            Else
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8210"
                                fncCheckSelectOption = False
                            End If
                        Case "G"
                            'Ｘ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 25 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 200 Then
                            Else
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8200"
                                fncCheckSelectOption = False
                            End If

                            'Ｚ軸ストロークの範囲判定
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) >= 25 And _
                               CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 150 Then
                            Else
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8210"
                                fncCheckSelectOption = False
                            End If
                        Case "1"
                            '基本形状判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "", "F"
                                    'ストロークの範囲判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 25 And _
                                       CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 300 Then
                                    Else
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W8160"
                                        fncCheckSelectOption = False
                                    End If
                                Case "L", "LF"
                                    'ストロークの範囲判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 301 And _
                                       CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 600 Then
                                    Else
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W8170"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                    End Select
                Case "AMD3" '2008/08/18 追加
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "1", "2"
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) = 0 And _
                            objKtbnStrc.strcSelection.strOpSymbol(8) = "R" Then
                                intKtbnStrcSeqNo = 8
                                strMessageCd = "W8670" '"オリフィス指示なしは補強リング付きを単独で選択できません。"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "AMD4", "AMD5" '2008/08/18 追加
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) = 0 And _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(8) = "R" Then
                                intKtbnStrcSeqNo = 8
                                strMessageCd = "W8670" '"オリフィス指示なしは補強リング付きを単独で選択できません。"
                                fncCheckSelectOption = False
                            End If
                    End Select
            End Select
        Catch ex As Exception

            fncCheckSelectOption = False

            Throw ex

        End Try

    End Function

End Module
