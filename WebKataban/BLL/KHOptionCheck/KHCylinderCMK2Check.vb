Module KHCylinderCMK2Check

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダチェック
    '*【概要】
    '*  シリンダＣＭＫ２シリーズをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*  更新履歴   ：                       
    '*  ・受付No：RM0908030  二次電池対応機器　
    '*                                      更新日：2009/09/04   更新者：Y.Miura
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Try

            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "CMK2"
                    '基本ベース毎にチェック
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 --->
                        'Case ""
                        Case "", "4"
                            '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 <---
                            '基本ベースチェック
                            If fncStandardBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "5"
                            '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 <---
                            '基本ベースチェック(食品製造工程向け)
                            If fncStandardBaseFP1Check(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "D"
                            '両ロッドベースチェック
                            If fncDoubleRodBaseCheck(objKtbnStrc, _
                                                     intKtbnStrcSeqNo, _
                                                     strOptionSymbol, _
                                                     strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "E"
                            '両ロッドベースチェック(食品製造工程向け)
                            If fncDoubleRodBaseFP1Check(objKtbnStrc, _
                                                     intKtbnStrcSeqNo, _
                                                     strOptionSymbol, _
                                                     strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                    '共通チェック
                    'RM0908030 2009/09/04 Y.Miura 二次電池対応
                    If fncCommonCheck(objKtbnStrc, _
                                      intKtbnStrcSeqNo, _
                                      strOptionSymbol, _
                                      strMessageCd) = False Then
                        fncCheckSelectOption = False
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncStandardBaseCheck
    '*【処理】
    '*  基本ベースチェック
    '*【概要】
    '*  基本ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '*                                          更新日：2008/04/07      更新者：T.Sato
    '*  ・受付No：RM0802088対応　ジャバラの有無により最大ストロークが変わるように修正
    '********************************************************************************************
    Private Function fncStandardBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean



        Try

            fncStandardBaseCheck = True

            'バリエーション「Q」＋ジャバラ「J」「L」は原価積算対応
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("L") >= 0 Then
                    intKtbnStrcSeqNo = 15
                    strMessageCd = "W0580"
                    fncStandardBaseCheck = False
                    Exit Try
                End If
            End If


            '↓RM1311065 2013/11/22 修正
            'オプション判定:Jを選択した場合は25以上であること
            'S1
            If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                        intKtbnStrcSeqNo = 5
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                End If
            End If
            'オプション判定:J・Lを選択した場合は25以上であること
            'S2
            If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                End If
            End If
            '↑RM1311065 2013/11/22 修正

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'バリエーション判定
            'RM1403023 E.MURATA
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Or _
                objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("R") >= 0 Then

                '2017/1/26 斉藤追加　バリエーションSRの場合、最少ストロークは5mm
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "SR" Then
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 5 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Else
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                End If
            End If
            'S1:スイッチ形番判定
            'バリエーションにBを含む時のみチェックする
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                'S1:スイッチ形番判定
                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    Case ""
                        'オプション判定:Jを選択した場合は25以上であること
                        If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Then
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        End If
                    Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                         "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                         "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                         "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", _
                         "T2JH", "T2JV", "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV"
                        'S1:スイッチ個数で判定
                        Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Case "1"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                            Case "2"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "T2WH", "T2WV", "T3WH", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 30 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV"
                                        'RM1403023 E.MURATA
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 35 Then ' スイッチ２個（"D"）
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            Case Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    'RM1403023 E.MURATA
                                    Case "T2WH", "T2WV", "T3WH", "T3WV", "T1H", "T1V", "T8H", _
                                        "T8V", "T2YH", "T2YV", "T3YH", "T3YV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 55 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                        End Select
                    Case "R1", "R2", "R2Y", "R3", "R3Y", _
                         "R0", "R4", "R5", "R6", "R1B", _
                         "R2B", "R2YB", "R3B", "R3YB", "R0B", _
                         "R4B", "R5B", "R6B", "R1A", "R2A", _
                         "R3A", "R0A", "R4A", "R5A", "R6A"
                        'S1:スイッチ個数で判定
                        Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Case "1"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                            Case Else
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 15 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                        End Select
                    Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                        'S1:スイッチ個数で判定
                        Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Case "1"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                            Case "2"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                            Case Else
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                        End Select
                End Select
            End If

            'S2:スイッチ形番判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                Case ""
                    'オプション判定:J・Lを選択した場合は25以上であること
                    If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                            intKtbnStrcSeqNo = 9
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
                    '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 --->
                    'Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                    '     "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                    '     "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                    '     "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", _
                    '     "T2JH", "T2JV", "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV"
                Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                     "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                     "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                     "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", _
                     "T2JH", "T2JV", "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV", _
                     "SW11", "SW12", "SW13", "SW14", "SW15", "SW16", "SW17", "SW21", "SW22", "SW23", _
                     "SW24", "SW25", "SW26", "SW27", "SW29", "SW30", "SW31", "SW32", "SW33", "SW34", _
                     "SW35", "SW36", "SW37", "SW40", "SW41", "SW42", "SW43", "SW44", "SW45", "SW46", "SW47", "SW48"
                    '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 <---

                    'S1:スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 10 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case "2"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 --->
                                'Case "T2WH", "T2WV", "T3WH", "T3WV"
                                Case "T2WH", "T2WV", "T3WH", "T3WV", "SW37", "SW40", "SW47", "SW48"
                                    '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 <---
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 30 Then ' スイッチ２個（"D"）
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                    End If
                                Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV"
                                    'RM1403023 E.MURATA
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 35 Then ' スイッチ２個（"D"）
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                        Case Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                'RM1403023 E.MURATA
                                Case "T2WH", "T2WV", "T3WH", "T3WV", "T1H", "T1V", "T8H", _
                                    "T8V", "T2YH", "T2YV", "T3YH", "T3YV"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 55 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 50 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                    End Select
                Case "R1", "R2", "R2Y", "R3", "R3Y", _
                     "R0", "R4", "R5", "R6", "R1B", _
                     "R2B", "R2YB", "R3B", "R3YB", "R0B", _
                     "R4B", "R5B", "R6B", "R1A", "R2A", _
                     "R3A", "R0A", "R4A", "R5A", "R6A"
                    'S1:スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 10 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 15 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                    End Select
                Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                    'S1:スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 10 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case "2"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 50 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                    End Select
            End Select

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            'S1:ストローク
            'バリエーションにBを含む時のみチェックする
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                ' バリエーション判定
                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    Case "S", "SR", "R", "SB", "SRB", _
                         "RM", "RT", "RO", "RF", "RG", "RG1", _
                         "RG2", "RG3", "RG4"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 300 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    Case Else
                End Select
            End If

            '支持形式判定:LSを選択した場合は50まで
            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "LS" Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 50 Then
                    intKtbnStrcSeqNo = 9
                    strMessageCd = "W0200"
                    fncStandardBaseCheck = False
                    Exit Try
                End If
            End If

            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "S", "SR", "SB", "SRB"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 300 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case "R", "RM", "RT", "RO", "RF", "RG", "RG1", "RG2", "RG3", "RG4"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 750 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case "P", "PQ", "PM", "PZ", _
                     "PH", "PT", "PO"
                    ' 口径判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "20", "25", "32"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 430 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case "40"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 400 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                    End Select
                Case Else
            End Select

            'ジャバラあり時の最大ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case ""
                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(4), "C") <> 0 Then
                        ' Ｓ２：ストローク
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(9)) > 650 Then
                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(15), "J") <> 0 Or _
                               InStr(1, objKtbnStrc.strcSelection.strOpSymbol(15), "L") <> 0 Then

                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        End If
                    End If
                Case "B", "R", "M", "C", "Z", "H", "T", "T2", "O", "G", "G1", "G4", _
                     "BZ", "BH", "BT", "BT2", "BO", "BG", "BG1", "BG4", _
                     "RM", "RT", "RO", "RG", "RG1", "RG4"
                    '背合わせ型(バリエーションにBを含む時)はＳ１も最大ストロークを判定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 650 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Or _
                               objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("L") >= 0 Then

                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        End If
                    End If
                    'Ｓ２：ストローク
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 650 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("L") >= 0 Then

                            intKtbnStrcSeqNo = 9
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
                Case Else
            End Select

            'ジャバラあり時の最大ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "P", "PM", "PZ", "PH", "PT", "PO"
                    'Ｓ２：ストローク
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 350 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("L") >= 0 Then

                            intKtbnStrcSeqNo = 9
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncStandardBaseFP1Check
    '*【処理】
    '*  基本ベースチェック
    '*【概要】
    '*  基本ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '*                                          更新日：2008/04/07      更新者：T.Sato
    '*  ・受付No：RM0802088対応　ジャバラの有無により最大ストロークが変わるように修正
    '********************************************************************************************
    Private Function fncStandardBaseFP1Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Dim objPrice As New KHUnitPrice

        Try

            fncStandardBaseFP1Check = True

            'バリエーション「Q」＋ジャバラ「J」「L」は原価積算対応
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("L") >= 0 Then
                    intKtbnStrcSeqNo = 15
                    strMessageCd = "W0580"
                    fncStandardBaseFP1Check = False
                    Exit Try
                End If
            End If


            '↓RM1311065 2013/11/22 修正
            'オプション判定:Jを選択した場合は25以上であること
            'S1
            If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                        intKtbnStrcSeqNo = 5
                        strMessageCd = "W0200"
                        fncStandardBaseFP1Check = False
                        Exit Try
                    End If
                End If
            End If
            'オプション判定:J・Lを選択した場合は25以上であること
            'S2
            If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseFP1Check = False
                        Exit Try
                    End If
                End If
            End If
            '↑RM1311065 2013/11/22 修正

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'バリエーション判定
            'RM1403023 E.MURATA
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Or _
                objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("R") >= 0 Then
                '2017/1/26 斉藤修正　バリエーションSRの場合最少ストローク5mm
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "SR" Then
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 5 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseFP1Check = False
                        Exit Try
                    End If
                Else
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseFP1Check = False
                        Exit Try
                    End If
                End If
            End If
            'S1:スイッチ形番判定
            'バリエーションにBを含む時のみチェックする
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                'S1:スイッチ形番判定
                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    Case ""
                        'オプション判定:Jを選択した場合は25以上であること
                        If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Then
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                        End If
                    Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                         "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                         "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                         "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", _
                         "T2JH", "T2JV", "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV"
                        'S1:スイッチ個数で判定
                        Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Case "1"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseFP1Check = False
                                    Exit Try
                                End If
                            Case "2"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "T2WH", "T2WV", "T3WH", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 30 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseFP1Check = False
                                            Exit Try
                                        End If
                                    Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV"
                                        'RM1403023 E.MURATA
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 35 Then ' スイッチ２個（"D"）
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseFP1Check = False
                                            Exit Try
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseFP1Check = False
                                            Exit Try
                                        End If
                                End Select
                            Case Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    'RM1403023 E.MURATA
                                    Case "T2WH", "T2WV", "T3WH", "T3WV", "T1H", "T1V", "T8H", _
                                        "T8V", "T2YH", "T2YV", "T3YH", "T3YV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 55 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseFP1Check = False
                                            Exit Try
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0200"
                                            fncStandardBaseFP1Check = False
                                            Exit Try
                                        End If
                                End Select
                        End Select
                    Case "R1", "R2", "R2Y", "R3", "R3Y", _
                         "R0", "R4", "R5", "R6", "R1B", _
                         "R2B", "R2YB", "R3B", "R3YB", "R0B", _
                         "R4B", "R5B", "R6B", "R1A", "R2A", _
                         "R3A", "R0A", "R4A", "R5A", "R6A"
                        'S1:スイッチ個数で判定
                        Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Case "1"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseFP1Check = False
                                    Exit Try
                                End If
                            Case Else
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 15 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseFP1Check = False
                                    Exit Try
                                End If
                        End Select
                    Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                        'S1:スイッチ個数で判定
                        Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Case "1"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseFP1Check = False
                                    Exit Try
                                End If
                            Case "2"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseFP1Check = False
                                    Exit Try
                                End If
                            Case Else
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W0200"
                                    fncStandardBaseFP1Check = False
                                    Exit Try
                                End If
                        End Select
                End Select
            End If

            'S2:スイッチ形番判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                Case ""
                    'オプション判定:J・Lを選択した場合は25以上であること
                    If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                            intKtbnStrcSeqNo = 9
                            strMessageCd = "W0200"
                            fncStandardBaseFP1Check = False
                            Exit Try
                        End If
                    End If
                    '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 --->
                    'Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                    '     "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                    '     "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                    '     "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", _
                    '     "T2JH", "T2JV", "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV"
                Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                     "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                     "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                     "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", _
                     "T2JH", "T2JV", "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV", _
                     "SW11", "SW12", "SW13", "SW14", "SW15", "SW16", "SW17", "SW21", "SW22", "SW23", _
                     "SW24", "SW25", "SW26", "SW27", "SW29", "SW30", "SW31", "SW32", "SW33", "SW34", _
                     "SW35", "SW36", "SW37", "SW40", "SW41", "SW42", "SW43", "SW44", "SW45", "SW46", "SW47", "SW48"
                    '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 <---

                    'S1:スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 10 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                        Case "2"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 --->
                                'Case "T2WH", "T2WV", "T3WH", "T3WV"
                                Case "T2WH", "T2WV", "T3WH", "T3WV", "SW37", "SW40", "SW47", "SW48"
                                    '2010/06/24 T.Fuji RM1005004(CMK2,STR2-Qシリーズ)対応 <---
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 30 Then ' スイッチ２個（"D"）
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseFP1Check = False
                                        Exit Try
                                    End If
                                Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV"
                                    'RM1403023 E.MURATA
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 35 Then ' スイッチ２個（"D"）
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseFP1Check = False
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseFP1Check = False
                                        Exit Try
                                    End If
                            End Select
                        Case Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                'RM1403023 E.MURATA
                                Case "T2WH", "T2WV", "T3WH", "T3WV", "T1H", "T1V", "T8H", _
                                    "T8V", "T2YH", "T2YV", "T3YH", "T3YV"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 55 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseFP1Check = False
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 50 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0200"
                                        fncStandardBaseFP1Check = False
                                        Exit Try
                                    End If
                            End Select
                    End Select
                Case "R1", "R2", "R2Y", "R3", "R3Y", _
                     "R0", "R4", "R5", "R6", "R1B", _
                     "R2B", "R2YB", "R3B", "R3YB", "R0B", _
                     "R4B", "R5B", "R6B", "R1A", "R2A", _
                     "R3A", "R0A", "R4A", "R5A", "R6A"
                    'S1:スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 10 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 15 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                    End Select
                Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                    'S1:スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 10 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                        Case "2"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 25 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) < 50 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                    End Select
            End Select

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            'S1:ストローク
            'バリエーションにBを含む時のみチェックする
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                ' バリエーション判定
                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    Case "S", "SR", "R", "SB", "SRB", _
                         "RM", "RT", "RO", "RF", "RG", "RG1", _
                         "RG2", "RG3", "RG4"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 300 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncStandardBaseFP1Check = False
                            Exit Try
                        End If
                    Case Else
                End Select
            End If

            '支持形式判定:LSを選択した場合は50まで
            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "LS" Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 50 Then
                    intKtbnStrcSeqNo = 9
                    strMessageCd = "W0200"
                    fncStandardBaseFP1Check = False
                    Exit Try
                End If
            End If

            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "S", "SR", "SB", "SRB"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 300 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseFP1Check = False
                        Exit Try
                    End If
                Case "R", "RM", "RT", "RO", "RF", "RG", "RG1", "RG2", "RG3", "RG4"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 750 Then
                        intKtbnStrcSeqNo = 9
                        strMessageCd = "W0200"
                        fncStandardBaseFP1Check = False
                        Exit Try
                    End If
                Case "P", "PQ", "PM", "PZ", _
                     "PH", "PT", "PO"
                    ' 口径判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "20", "25", "32"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 430 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                        Case "40"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 400 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                    End Select
                Case Else
            End Select

            'ジャバラあり時の最大ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case ""
                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(4), "C") <> 0 Then
                        ' Ｓ２：ストローク
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(9)) > 650 Then
                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(15), "J") <> 0 Or _
                               InStr(1, objKtbnStrc.strcSelection.strOpSymbol(15), "L") <> 0 Then

                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                        End If
                    End If
                Case "B", "R", "M", "C", "Z", "H", "T", "T2", "O", "G", "G1", "G4", _
                     "BZ", "BH", "BT", "BT2", "BO", "BG", "BG1", "BG4", _
                     "RM", "RT", "RO", "RG", "RG1", "RG4"
                    '背合わせ型(バリエーションにBを含む時)はＳ１も最大ストロークを判定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 650 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Or _
                               objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("L") >= 0 Then

                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncStandardBaseFP1Check = False
                                Exit Try
                            End If
                        End If
                    End If
                    'Ｓ２：ストローク
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 650 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("L") >= 0 Then

                            intKtbnStrcSeqNo = 9
                            strMessageCd = "W0200"
                            fncStandardBaseFP1Check = False
                            Exit Try
                        End If
                    End If
                Case Else
            End Select

            'ジャバラあり時の最大ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "P", "PM", "PZ", "PH", "PT", "PO"
                    'Ｓ２：ストローク
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 350 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("J") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("L") >= 0 Then

                            intKtbnStrcSeqNo = 9
                            strMessageCd = "W0200"
                            fncStandardBaseFP1Check = False
                            Exit Try
                        End If
                    End If
                Case Else
            End Select

        Catch ex As Exception

            Throw ex

        Finally

            objPrice = Nothing

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncDoubleRodBaseCheck
    '*【処理】
    '*  両ロッドベースチェック
    '*【概要】
    '*  両ロッドベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '********************************************************************************************
    Private Function fncDoubleRodBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                           ByRef intKtbnStrcSeqNo As Integer, _
                                           ByRef strOptionSymbol As String, _
                                           ByRef strMessageCd As String) As Boolean

        Dim objPrice As New KHUnitPrice

        Try

            fncDoubleRodBaseCheck = True

            'バリエーション「Q」＋ジャバラ「J」「L」は原価積算対応
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("J") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("L") >= 0 Then
                    intKtbnStrcSeqNo = 10
                    strMessageCd = "W0580"
                    fncDoubleRodBaseCheck = False
                    Exit Try
                End If
            End If

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'SW形番判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                Case ""
                    'オプション判定:J・Lを選択した場合は25以上であること
                    If objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("J") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("L") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                        End If
                    End If
                Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                     "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                     "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                     "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", _
                     "T2JH", "T2JV", "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV"
                    'S1:スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                        Case "2"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "T2WH", "T2WV", "T3WH", "T3WV"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 30 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseCheck = False
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                    End Select

                Case "R1", "R2", "R2Y", "R3", "R3Y", _
                     "R0", "R4", "R5", "R6", "R1B", _
                     "R2B", "R2YB", "R3B", "R3YB", "R0B", _
                     "R4B", "R5B", "R6B", "R1A", "R2A", _
                     "R3A", "R0A", "R4A", "R5A", "R6A"
                    'スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 15 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                    End Select
                Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                    'スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                        Case "2"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                    End Select
            End Select

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            'バリエーション判定:LSを選択した場合は50Stまで
            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "LS" Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 50 Then
                    intKtbnStrcSeqNo = 5
                    strMessageCd = "W0200"
                    fncDoubleRodBaseCheck = False
                    Exit Try
                End If
            End If

            'オプション判定:J・Lを選択した場合は300Stまで
            If objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("J") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("L") >= 0 Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 300 Then
                    intKtbnStrcSeqNo = 5
                    strMessageCd = "W0200"
                    fncDoubleRodBaseCheck = False
                    Exit Try
                End If
            End If

        Catch ex As Exception

            Throw ex

        Finally

            objPrice = Nothing
        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncDoubleRodBaseFP1Check
    '*【処理】
    '*  両ロッドベースチェック
    '*【概要】
    '*  両ロッドベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '********************************************************************************************
    Private Function fncDoubleRodBaseFP1Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                           ByRef intKtbnStrcSeqNo As Integer, _
                                           ByRef strOptionSymbol As String, _
                                           ByRef strMessageCd As String) As Boolean



        Try

            fncDoubleRodBaseFP1Check = True

            'バリエーション「Q」＋ジャバラ「J」「L」は原価積算対応
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("J") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("L") >= 0 Then
                    intKtbnStrcSeqNo = 10
                    strMessageCd = "W0580"
                    fncDoubleRodBaseFP1Check = False
                    Exit Try
                End If
            End If

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'SW形番判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                Case ""
                    'オプション判定:J・Lを選択した場合は25以上であること
                    If objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("J") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("L") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncDoubleRodBaseFP1Check = False
                            Exit Try
                        End If
                    End If
                Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                     "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                     "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                     "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", _
                     "T2JH", "T2JV", "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV"
                    'S1:スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseFP1Check = False
                                Exit Try
                            End If
                        Case "2"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "T2WH", "T2WV", "T3WH", "T3WV"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 30 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseFP1Check = False
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseFP1Check = False
                                        Exit Try
                                    End If
                            End Select
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseFP1Check = False
                                Exit Try
                            End If
                    End Select

                Case "R1", "R2", "R2Y", "R3", "R3Y", _
                     "R0", "R4", "R5", "R6", "R1B", _
                     "R2B", "R2YB", "R3B", "R3YB", "R0B", _
                     "R4B", "R5B", "R6B", "R1A", "R2A", _
                     "R3A", "R0A", "R4A", "R5A", "R6A"
                    'スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseFP1Check = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 15 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseFP1Check = False
                                Exit Try
                            End If
                    End Select
                Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                    'スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        Case "1"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseFP1Check = False
                                Exit Try
                            End If
                        Case "2"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseFP1Check = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncDoubleRodBaseFP1Check = False
                                Exit Try
                            End If
                    End Select
            End Select

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            'バリエーション判定:LSを選択した場合は50Stまで
            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "LS" Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 50 Then
                    intKtbnStrcSeqNo = 5
                    strMessageCd = "W0200"
                    fncDoubleRodBaseFP1Check = False
                    Exit Try
                End If
            End If

            'オプション判定:J・Lを選択した場合は300Stまで
            If objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("J") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(10).IndexOf("L") >= 0 Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 300 Then
                    intKtbnStrcSeqNo = 5
                    strMessageCd = "W0200"
                    fncDoubleRodBaseFP1Check = False
                    Exit Try
                End If
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncCommonCheck
    '*【処理】
    '*  共通チェック
    '*【概要】
    '*  共通のチェックを行う
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2009/09/04      更新者：Y.Miura
    '********************************************************************************************
    Private Function fncCommonCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncCommonCheck = True

            '二次電池対応機器　Ｐ４の必須チェック
            If objKtbnStrc.strcSelection.strKeyKataban = "4" Then
                Dim bolOptionP4 As Boolean = False

                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(15), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case "P4", "P40"
                            bolOptionP4 = True
                    End Select
                Next

                If Not bolOptionP4 Then
                    intKtbnStrcSeqNo = 15
                    strMessageCd = "W8770"
                    fncCommonCheck = False
                    Exit Try
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

End Module
