Module KHCylinderStrokeCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダストロークチェック
    '*【概要】
    '*  シリンダのストロークをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                      更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '*                                      更新日：2007/09/07      更新者：NII A.Takahashi
    '*  ・F2/F3スイッチ追加(機種SSG)に伴い、ストロークチェックロジックを修正
    '*                                          更新日：2008/04/07      更新者：T.Sato
    '*  ・受付No：RM0802088対応　ジャバラの有無により最大ストロークが変わるように修正
    '*  ・受付No：RM0811044対応 2008/12/15 T.Y　STGシリーズ　T1H,T1V,T8H,T8V追加
    '*  ・受付No：RM0907002  ペンシルシリンダSCP*2,ULKP 最小ストロークの規制を追加
    '*                                      更新日：2009/08/10   更新者：Y.Miura
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String = Nothing
        Dim intLoopCnt As Integer

        Try

            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "CKV2", "CKV2-M"
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            Case "R", "H"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "T2WH", "T2WV", "T3WH", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 30 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 50 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            'ジャバラを選択している場合
                            Case "J", "L"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                Case "JSM2"
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "T0H", "T0V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 27 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 51 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T1H", "T1V"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 49 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T8H", "T8V"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 23 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 47 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T2WH", "T2WV", "T3WH", "T3WV"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 31 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 55 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T2YH", "T2YV", "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", "T2JH", "T2JV"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 49 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 45 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If
                Case "JSM2-V"
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "T0H", "T0V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 27 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 51 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T1H", "T1V"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 49 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T8H", "T8V"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 23 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 47 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T2WH", "T2WV", "T3WH", "T3WV"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 31 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 55 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T2YH", "T2YV", "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", "T2JH", "T2JV"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 49 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "R", "H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T", "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 45 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If
                Case "JSK2"
                    ' スイッチ判定
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) = 0 Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "R", "H"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV", "T2JH", "T2JV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 35 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2WH", "T2WV", "T3WH", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 30 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV", "T2JH", "T2JV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 55 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If

                    ' オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            'ジャバラを選択している場合
                            Case "J"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 600 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "L"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 600 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                Case "JSK2-V"
                    ' スイッチ判定
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) = 0 Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "R", "H"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV", "T2JH", "T2JV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 35 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2WH", "T2WV", "T3WH", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 30 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV", "T2JH", "T2JV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 55 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            'ジャバラを選択している場合
                            Case "J"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 600 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "L"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 600 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                Case "MSDG-L", "MSD-KL", "MSD-L", "MSD-XL", "MSD-YL"
                    ' スイッチ判定
                    Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2)
                        Case "F0"
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "D" Then
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select
                    '2014/04/29
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "F3PH", "F3PV"
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "D" Or _
                                objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "H" Then
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select
                Case "MVC"
                    ' スイッチ判定
                    Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 2)
                        Case "F0"
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "D" Then
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 2
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select
                    '2014/04/29
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "F3PH", "F3PV"
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "D" Or _
                                objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "H" Then
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 2
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select
                Case "ULK", "ULK-V"
                    ' スイッチ判定
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) = 0 Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "R", "H"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV", "T2JH", "T2JV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 35 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2WH", "T2WV", "T3WH", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 30 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "T1H", "T1V", "T8H", "T8V", "T2YH", "T2YV", "Y3YH", "T3YV", "T2JH", "T2JV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 55 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If

                    ' オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            'ジャバラを選択している場合
                            Case "J", "L"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                Case "SRT"
                    ' 口径
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "12", "16", "20"
                            'スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    'スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    'スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    'スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    'スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                        Case "25", "32", "40"
                            'スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    'スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                        Case "50", "63"
                            ' スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                    End Select
                Case "SRM", "SRM-Q"
                    ' スイッチ個数
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        Case "R", "L"
                            ' スイッチ形番
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "T0V", "T5V", "T2YV", "T3YV", "T2YFV", _
                                     "T3YFV", "T2YMV", "T3YMV", "T8V", "T2WV", "T3WV"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0190"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T0H", "T5H", "T2YH", "T3YH", "T2YFH", _
                                     "T3YFH", "T2YMH", "T3YMH", "T2YD", "T8H", "T2WH", "T3WH"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0190"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        Case "D"
                            ' スイッチ形番
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "T0V", "T5V", "T2YV", "T3YV", "T2YFV", _
                                     "T3YFV", "T2YMV", "T3YMV", "T8V", "T2WV", "T3WV"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 45 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0190"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T0H", "T5H", "T2YH", "T3YH", "T2YFH", _
                                     "T3YFH", "T2YMH", "T3YMH", "T2YD", "T8H", "T2WH", "T3WH"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 50 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0190"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        Case "T"
                            ' スイッチ形番
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "T0V", "T5V", "T2YV", "T3YV", "T2YFV", _
                                     "T3YFV", "T2YMV", "T3YMV", "T8V", "T2WV", "T3WV"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 90 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0190"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T0H", "T5H", "T2YH", "T3YH", "T2YFH", _
                                     "T3YFH", "T2YMH", "T3YMH", "T2YD", "T8H", "T2WH", "T3WH"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 100 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0190"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        Case "4"
                            ' スイッチ形番
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "T0V", "T5V", "T2YV", "T3YV", "T2YFV", _
                                     "T3YFV", "T2YMV", "T3YMV", "T8V", "T2WV", "T3WV"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 135 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0190"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T0H", "T5H", "T2YH", "T3YH", "T2YFH", _
                                     "T3YFH", "T2YMH", "T3YMH", "T2YD", "T8H", "T2WH", "T3WH"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 150 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0190"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                    End Select
                Case "SRG"
                    ' 口径
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "12", "16", "20"
                            ' スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                        Case "25"
                            ' スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                    End Select
                Case "SRB2"
                    ' 口径
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "25", "40"
                            ' スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                        Case "63"
                            ' スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", _
                                             "T2YD", "T2YDT"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                    End Select
                Case "SRL2", "SRL2-G", "SRL2-GQ", "SRL2-J", "SRL2-Q"
                    ' 口径
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "12", "16", "20"
                            ' スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                        Case "25", "32", "40"
                            ' スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                        Case "50", "63"
                            ' スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                        Case "80", "100"
                            ' スイッチ個数
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "R", "L"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "D"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "T"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "4"
                                    ' スイッチ形番
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "M0V", "M2V", "M2WV", "M3V", "M3WV", _
                                             "M5V"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "M0H", "M2H", "M3H", "M5H"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YLV", _
                                             "T3YLV", "T2WV", "T3WV", "T2YV", "T3YV"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YLH", _
                                             "T3YLH", "T2YD", "T2YDT", "T2WH", "T3WH", "T2YH", "T3YH"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 150 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0190"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                    End Select
                Case "MDC2-L", "MDC2-XL", "MDC2-YL"
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "R", "H"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 4 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "6" Or objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "8" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "F3PH" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "F3PV" Then
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 8 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    End If
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "10" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "F3PH" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "F3PV" Then
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    End If
                                End If
                            Case Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "F0H", "F0V"
                                        ' 口径
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "4"
                                            Case "6"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 6 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "8"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 8 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "10"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 6 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "F3PH", "F3PV"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "4"
                                            Case "6"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 8 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "8"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 8 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "10"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 4 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If
                Case "UCA2-L", "UCA2-BL"
                    'RM1305005 2013/05/30 ローカル版との差異修正
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                Case "T"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 75 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0190"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        End If
                    End If
                Case "CAV2"
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                        If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 1 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1)
                            Case "R"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "LB", "FA", "CA"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "R", "H"
                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"
                                                        If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"
                                                        If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 70 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select

                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "R", "H"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                            Case "100"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 110 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                        End Select
                                                    Case Else
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 140 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                            Case "100"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 150 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                        End Select
                                                End Select
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                            Case "100"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 110 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                        End Select
                                                    Case Else
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 140 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                            Case "100"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 150 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                        End Select
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 120 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                            Case "100"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 130 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                        End Select
                                                    Case Else
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 140 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                            Case "100"
                                                                If Val(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 150 Then
                                                                    intKtbnStrcSeqNo = 5
                                                                    strMessageCd = "W0190"
                                                                    fncCheckSelectOption = False
                                                                End If
                                                        End Select
                                                End Select
                                        End Select
                                End Select
                        End Select
                    End If

                    'Case "CMA2-E"
                    '    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "LS" Then
                    '        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 700 Then
                    '            intKtbnStrcSeqNo = 3
                    '            strMessageCd = "W0190"
                    '            fncCheckSelectOption = False
                    '        End If
                    '    End If

                Case "CMA2", "CMA2-D", "CMA2-H", "CMA2-T", "CMA2-E"
                    If objKtbnStrc.strcSelection.strSeriesKataban = "CMA2-E" AndAlso _
                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "LS" Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 700 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    End If

                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "CMA2", "CMA2-D", "CMA2-E", "CMA2-H"
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "T0H", "T0V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                            Case "R", "H"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 27 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 51 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "T1H", "T1V"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                            Case "R", "H"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 49 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "T8H", "T8V"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                            Case "R", "H"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 23 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 47 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "T2WH", "T2WV", "T3WH", "T3WV"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                            Case "R", "H"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 31 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 55 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "T2YH", "T2YV", "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", "T2JH", "T2JV"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                            Case "R", "H"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 49 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case Else
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                            Case "R", "H"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            End If
                    End Select

                    'オプションで、ジャバラ(J)を選択した時は、最小値25mm
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "CMA2", "CMA2-D", "CMA2-H", "CMA2-T"
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                Case "CMA2", "CMA2-D", "CMA2-H"
                                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                                Case "CMA2-T"
                                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                            End Select

                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "J"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Next

                            ' 支持形式がLSの場合MAX=50
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "LS" Then
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 50 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select
                Case "COVN2", "COVP2"
                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 1 Then
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                        End If
                    Else
                        ' スイッチ有り
                        Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1)
                            Case "R"
                                ' 支持形式により判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "LB", "FA", "CA"

                                        ' スイッチ・個数により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "R", "H"
                                                ' スイッチが1個の場合
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 20 Then
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                                End If
                                            Case "D"
                                                ' スイッチが2個の場合

                                                ' リード線により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"
                                                        ' グロメットの場合
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 20 Then
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 5
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        End If
                                                    Case Else
                                                        ' 端子箱A,Bの場合
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 5
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        End If
                                                End Select
                                            Case "T"
                                                ' スイッチが3個の場合

                                                ' リード線により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"
                                                        ' グロメットの場合
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 40 Then
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 5
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        End If
                                                    Case Else
                                                        ' 端子箱A,Bの場合
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 70 Then
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 5
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        End If
                                                End Select
                                        End Select

                                    Case "TC", "TF"
                                        ' スイッチ・個数により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "R", "H"
                                                ' スイッチが1個の場合

                                                ' リード線により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"

                                                        ' 口径により判定
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                            Case "100"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 110 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                        End Select
                                                    Case Else

                                                        ' 口径により判定
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 140 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                            Case "100"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 150 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                        End Select
                                                End Select
                                            Case "D"
                                                ' スイッチが2個の場合

                                                ' リード線により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"

                                                        ' 口径により判定
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                            Case "100"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 110 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                        End Select
                                                    Case Else

                                                        ' 口径により判定
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 140 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                            Case "100"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 150 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                        End Select
                                                End Select
                                            Case "T"
                                                ' スイッチが3個の場合

                                                ' リード線により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                    Case "", "3", "5"

                                                        ' 口径により判定
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 120 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                            Case "100"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 130 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                        End Select
                                                    Case Else

                                                        ' 口径により判定
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                            Case "50", "75"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 140 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                            Case "100"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 150 Then
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 100 Then
                                                                        intKtbnStrcSeqNo = 5
                                                                        strMessageCd = "W0190"
                                                                        fncCheckSelectOption = False
                                                                    End If
                                                                End If
                                                        End Select
                                                End Select
                                        End Select
                                End Select
                        End Select
                    End If
                Case "FCS-L"
                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                ' スイッチが2個の場合

                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "25", "32"
                                        '口径が25,32の時はDを選択できない
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W0570"
                                        fncCheckSelectOption = False
                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If
                Case "FCH-L"
                    ' FCHシリーズ
                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り

                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                ' スイッチが2個の場合
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "25", "32"
                                        '口径が25,32の時はDを選択できない

                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W0570"
                                        fncCheckSelectOption = False

                                    Case Else
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If
                Case "FCD-DL", "FCD-KL", "FCD-L"
                    ' FCDシリーズ
                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り
                        ' 口径により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "25", "32", "40", "50"

                                ' スイッチ・個数により判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R", "H"
                                        ' スイッチが1個の場合
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        ' スイッチが2個の場合
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T"
                                        ' スイッチが3個の場合
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case Else
                                ' スイッチ・個数により判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R", "H"
                                        ' スイッチが1個の場合
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        ' スイッチが2個の場合
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T"
                                        ' スイッチが3個の場合
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 45 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If
                Case "GLC", "GLC-L2"
                    ' GLCシリーズ
                    ' 中間ストロークチェック(5mm毎)
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) Mod 5 <> 0 Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0510"
                        fncCheckSelectOption = False
                    End If

                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "R0", "R4", "R5", "R6", "R1", "R2", "R2Y", "R3", "R3Y"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "H0"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50", "63"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case Else
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            Case "D"
                                ' スイッチが2個の場合

                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "R0", "R4", "R5", "R6"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "50", "63", "80"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "R1", "R2", "R2Y", "R3", "R3Y"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50", "63"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "H0"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T"
                                ' スイッチが3個の場合

                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "R0", "R4", "R5", "R6"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "50", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "63", "80"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 50 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "R1", "R2", "R2Y", "R3", "R3Y"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50", "63"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "H0"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 30 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            Case "4"
                                ' スイッチが4個の場合

                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "R0", "R4", "R5", "R6"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 60 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "50"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 65 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 70 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "R1", "R2", "R2Y", "R3", "R3Y"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50", "63"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 60 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "80"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 65 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 70 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "H0"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50", "63"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            Case "5"
                                ' スイッチが5個の場合

                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "R0", "R4", "R5", "R6"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 80 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "50"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 85 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "63", "80"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 95 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 90 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "R1", "R2", "R2Y", "R3", "R3Y"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 75 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "63"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 80 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "80"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 85 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 90 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "H0"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 50 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "63"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 55 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 60 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                        End Select
                    End If
                Case "HCA"
                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 1 Then
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                ' スイッチが2個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "T"
                                ' スイッチが3個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If
                Case "MFC", "MFC-B", "MFC-BK", "MFC-BKL", "MFC-BL", _
                     "MFC-BS", "MFC-BSK", "MFC-K", "MFC-KL", "MFC-L"
                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D", "T"
                                ' スイッチが2個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If
                Case "MRL2", "MRL2-G", "MRL2-GL", "MRL2-L", "MRL2-W", "MRL2-WL"
                    'オプション分解
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            'オプションに"C"を選択している時
                            Case "C"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next

                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 1 Then
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "R", "H", "L"
                                ' スイッチが1個の場合

                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                ' スイッチが2個の場合
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "T2V", "T3V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2H", "T3H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YV", "T3YV", "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T1V", "T2WV", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YH", "T3YH", "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T1H", "T2WH", "T3WH"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 70 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T"
                                ' スイッチが3個の場合
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "T2V", "T3V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2H", "T3H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 85 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YV", "T3YV", "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T1V", "T2WV", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 71 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YH", "T3YH", "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T1H", "T2WH", "T3WH"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "4"
                                ' スイッチが4個の場合
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "T2V", "T3V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2H", "T3H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 120 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YV", "T3YV", "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T1V", "T2WV", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 101 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YH", "T3YH", "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T1H", "T2WH", "T3WH"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 160 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If

                    '最大ストロークチェック
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "F" Then
                        'バリエーションに"F"が選択されている場合
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban
                            Case "MRL2", "MRL2-G", "MRL2-W"
                                '機種形番の末尾が"L"でない場合
                                ' 口径により判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "6"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "10"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) > 500 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "16", "20", "25"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) > 800 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "32"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) > 700 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case Else
                                ' 口径により判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "6"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "10"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "16"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) > 500 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "20", "25", "32"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) > 700 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If
                Case "MRG2"
                    ' スイッチ判定
                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 50 Then
                            intKtbnStrcSeqNo = 2
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合

                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 50 Then
                                    intKtbnStrcSeqNo = 2
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                ' スイッチが2個の場合
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Case "T2V", "T3V", "T0V", "T5V", "T2H", "T3H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T0H", "T5H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 100 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YV", "T3YV", "T1V", "T2WV", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YH", "T3YH", "T1H", "T2WH", "T3WH"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 100 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "T"
                                ' スイッチが3個の場合
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Case "T2V", "T3V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T0V", "T5V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 50 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2H", "T3H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 100 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T0H", "T5H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 100 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YV", "T3YV", "T1V", "T2WV", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 100 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YH", "T3YH", "T1H", "T2WH", "T3WH"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 150 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "4"
                                ' スイッチが4個の場合
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Case "T2V", "T3V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 100 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T0V", "T5V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 100 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2H", "T3H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 150 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T0H", "T5H"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 150 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2YV", "T3YV", "T1V", "T2WV", "T3WV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 150 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2YH", "T3YH", "T1H", "T2WH", "T3WH"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) < 200 Then
                                            intKtbnStrcSeqNo = 2
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If
                Case "SHC", "SHC-K", "SHC-K-L2", "SHC-L2"

                    ' 中間ストロークチェック(5mm毎)
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) Mod 5 <> 0 Then
                        intKtbnStrcSeqNo = 4
                        strMessageCd = "W0510"
                        fncCheckSelectOption = False
                    End If

                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り
                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            Case "R", "H", "D", "T"
                                ' スイッチが1個、2個、3個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "4"
                                ' スイッチが4個の場合
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R0", "R4", "R5", "R6"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "50", "63"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 55 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "R1", "R2", "R2Y", "R3", "R3Y"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "50", "63", "80"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "H0"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "5"
                                ' スイッチが5個の場合
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "R0", "R4", "R5", "R6"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 65 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "50", "63"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 70 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "80"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 75 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 80 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "R1", "R2", "R2Y", "R3", "R3Y"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 65 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "63", "80"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 65 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "H0"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "40", "50"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 65 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                        End Select
                    End If
                Case "SMD2", "SMD2-L", "SMD2-M", "SMD2-ML", "SMD2-X", "SMD2-XL", "SMD2-Y", "SMD2-YL"
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "SMD2-L"
                            'スイッチ判定
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                                ' スイッチ無し
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 1 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Else
                                ' スイッチ有り
                                ' 支持形式により判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "DA"

                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "DB", "DC"

                                        ' スイッチにより判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            'RM0908030 2009/09/08 Y.Miura　二次電池対応
                                            'Case "K0H", "K5H", "K2H", "K3H"
                                            Case "K0H", "K5H", "K2H", "K3H", _
                                                 "SW58", "SW51", "SW52", "SW53", "SW61", "SW62", "SW63"
                                                ' 口径により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                    Case "6"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "10", "16", "20"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "25", "32"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                                'RM0908030 2009/09/08 Y.Miura　二次電池対応
                                                'Case "K0V", "K5V", "K2V", "K3V"
                                            Case "K0V", "K5V", "K2V", "K3V", _
                                                 "SW54", "SW55", "SW56", "SW64", "SW65", "SW66"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "K2YH", "K3YH"
                                                ' 口径により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                    Case "6", "10"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "16", "20"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "25", "32"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "K2YV", "K3YV", "K2YFH", "K3YFH", "K2YMH", "K3YMH", "K2YFV", "K3YFV", "K2YMV", "K3YMV"
                                                ' 口径により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                    Case "6", "10", "16"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "20", "25", "32"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            End If
                        Case "SMD2-XL"
                            ' スイッチ判定
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                                ' スイッチ無し
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Else
                                ' スイッチ有り

                                ' 支持形式により判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "DA"

                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "DB", "DC"

                                        ' スイッチにより判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "K0H", "K5H", "K2H", "K3H"
                                                ' 口径により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "6"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "10", "16", "20"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "25", "32"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "K0V", "K5V", "K2V", "K3V"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "K2YH", "K3YH"
                                                ' 口径により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "6", "10"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "16", "20"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "25", "32"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "K2YV", "K3YV", "K2YFH", "K3YFH", "K2YMH", "K3YMH", "K2YFV", "K3YFV", "K2YMV", "K3YMV"
                                                ' 口径により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "6", "10", "16"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "20", "25", "32"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            End If
                        Case "SMD2-YL"
                            ' スイッチ判定
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                                ' スイッチ無し
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Else
                                ' スイッチ有り

                                ' 支持形式により判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "DA"

                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "DB", "DC"

                                        ' スイッチにより判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "K0H", "K5H", "K2H", "K3H", "K2YH", "K3YH", "K2YV", "K3YV"
                                                ' 口径により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "6", "10"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "16", "20", "25", "32"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "K0V", "K5V", "K2V", "K3V"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "K2YFH", "K3YFH", "K2YMH", "K3YMH", "K2YFV", "K3YFV", "K2YMV", "K3YMV"
                                                ' 口径により判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "6"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "10", "16", "20", "25", "32"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 3
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            End If
                        Case "SMD2-ML"
                            ' スイッチ判定
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                                ' スイッチ無し
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 1 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Else
                                ' スイッチ有り
                                ' スイッチにより判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "K0H", "K5H", "K2H", "K3H"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "6"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "10", "16", "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "25", "32"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "K0V", "K5V", "K2V", "K3V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "K2YH", "K3YH"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "6", "10"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 15 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "16", "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "25", "32"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "K2YV", "K3YV", "K2YFH", "K3YFH", "K2YMH", "K3YMH", "K2YFV", "K3YFV", "K2YMV", "K3YMV"
                                        ' 口径により判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "6", "10", "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "20", "25", "32"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 3
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            End If
                    End Select
                Case "STR2-B", "STR2-M"
                    ' 中間ストロークチェック
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", "O", "F"
                            ' 5mm毎
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) Mod 5 <> 0 Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W0510"
                                fncCheckSelectOption = False
                            End If
                        Case "Q", "D"
                            ' 標準ストローク
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) Mod 10 <> 0 Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W0170"
                                fncCheckSelectOption = False
                            End If
                    End Select

                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                        ' スイッチ無し
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "", "O", "F"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "Q", "D"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Else
                        ' スイッチ有り

                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    End If

                    'オプションによる最大Stチェック
                    'オプション分解
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            'オプションに"R"を選択している時
                            Case "R"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "6", "10", "32"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 50 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "16"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 70 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "20", "25"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 60 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    Next
                Case "ULKP", "ULKP-L"

                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り

                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 5 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                ' スイッチが2個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "T"
                                ' スイッチが3個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 38 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If

                    'RM0907002 2009/08/10 Y.Miura
                Case "SCPD2-L", "SCPD2-DL", "SCPD2-ZL", "SCPD2-KL", "SCPD2-ML", _
                     "SCPD2-OL", "SCPD2-VL", "SCPH2-L", "SCPS2-L", "SCPS2-ML", "SCPS2-VL"
                    Dim intSt As Integer
                    Dim intSw As Integer
                    Dim intSwQty As Integer
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.TrimEnd
                        Case "SCPD2-L"
                            intSt = 4
                            intSw = 6
                            intSwQty = 8
                        Case "SCPD2-DL", "SCPD2-ZL"
                            intSt = 3
                            intSw = 4
                            intSwQty = 6
                        Case "SCPD2-KL", "SCPD2-ML", "SCPD2-OL", "SCPD2-VL", _
                             "SCPH2-L", "SCPS2-L", "SCPS2-ML", "SCPS2-VL"
                            intSt = 3
                            intSw = 5
                            intSwQty = 7
                    End Select

                    'スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(intSw).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 5 Then
                            intKtbnStrcSeqNo = intSt
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        'スイッチ有り

                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(intSwQty).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 5 Then
                                    intKtbnStrcSeqNo = intSt
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                ' スイッチが2個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 10 Then
                                    intKtbnStrcSeqNo = intSt
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "T"
                                ' スイッチが3個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 38 Then
                                    intKtbnStrcSeqNo = intSt
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If

                Case "SCPG2-L", "SCPG2-DL", "SCPG2-XL", "SCPG2-YL", "SCPG2-ML", _
                 "SCPG2-XML"
                    Dim intSt As Integer
                    Dim intSw As Integer
                    Dim intSwQty As Integer
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.TrimEnd
                        Case "SCPG2-L", "SCPG2-XL", "SCPG2-YL"
                            intSt = 4
                            intSw = 6
                            intSwQty = 8
                        Case "SCPG2-DL"
                            intSt = 3
                            intSw = 4
                            intSwQty = 6
                        Case "SCPG2-ML", "SCPG2-XML"
                            intSt = 3
                            intSw = 5
                            intSwQty = 7
                    End Select

                    'スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(intSw).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 5 Then
                            intKtbnStrcSeqNo = intSt
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        'スイッチ有り

                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(intSwQty).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 5 Then
                                    intKtbnStrcSeqNo = intSt
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                ' スイッチが2個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 10 Then
                                    intKtbnStrcSeqNo = intSt
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "T"
                                ' スイッチが3個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 38 Then
                                    intKtbnStrcSeqNo = intSt
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If

                Case "SCPD3-L", "SCPD3-DL", "SCPD3-ML", "SCPD3-LF", _
                     "SCPD3-OL", "SCPH3-L", "SCPS3-L", "SCPS3-ML"
                    Dim intSt As Integer
                    Dim intSw As Integer
                    Dim intSwQty As Integer
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.TrimEnd
                        Case "SCPD3-L"
                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "C" Then
                                intSt = 4
                                intSw = 6
                                intSwQty = 8
                            Else
                                intSt = 3
                                intSw = 5
                                intSwQty = 7
                            End If
                        Case "SCPD3-LF"
                            intSt = 3
                            intSw = 4
                            intSwQty = 6
                        Case "SCPD3-DL"
                            intSt = 3
                            intSw = 4
                            intSwQty = 6
                        Case "SCPD3-ML", "SCPD3-OL", _
                             "SCPH3-L", "SCPS3-L", "SCPS3-ML"
                            intSt = 3
                            intSw = 5
                            intSwQty = 7
                    End Select

                    'スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(intSw).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 5 Then
                            intKtbnStrcSeqNo = intSt
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        'スイッチ有り

                        ' スイッチ・個数により判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(intSwQty).Trim
                            Case "R", "H"
                                ' スイッチが1個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 5 Then
                                    intKtbnStrcSeqNo = intSt
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "D"
                                ' スイッチが2個の場合
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 10 Then
                                    intKtbnStrcSeqNo = intSt
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                            Case "T"
                                ' スイッチが3個の場合
                                'If CInt(objKtbnStrc.strcSelection.strOpSymbol(intSt).Trim) < 38 Then
                                '    intKtbnStrcSeqNo = intSt
                                '    strMessageCd = "W0190"
                                '    fncCheckSelectOption = False
                                'End If
                        End Select
                    End If

                Case "USC", "USC-G1", "USC-G1L2", "USC-L2"
                    ' スイッチ判定
                    ' スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" Then
                        ' スイッチ無し
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 1 Then
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    Else
                        ' スイッチ有り
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "R1", "R2", "R2Y", "R3", "R3Y", _
                                 "R0", "R4", "R5", "R6", "H0", "H0Y"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "00", "LB", "FA", "FB", "FC", _
                                         "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 35 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63", "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63", "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 55 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TC", "TF"
                                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "B" Then
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                        Case "40", "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 66 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 71 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 76 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 86 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                        Case "40", "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 92 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 97 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 102 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 112 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                    End Select
                                            End Select
                                        Else
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                        Case "40", "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 86 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 91 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 96 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 106 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                        Case "40", "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 92 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 97 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 102 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 112 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                    End Select
                                            End Select
                                        End If
                                    Case "TA", "TD", "TB", "TE"
                                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "B" Then
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 28 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 26 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 31 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 34 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                    End Select
                                            End Select
                                        Else
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 38 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 36 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 41 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 44 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 4
                                                                strMessageCd = "W0190"
                                                                fncCheckSelectOption = False
                                                            End If
                                                    End Select
                                            End Select
                                        End If
                                End Select
                            Case "T0H", "T5H"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "00", "LB", "FA", "FB", "FC", _
                                         "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 65 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 70 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R", "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 125 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T", "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 175 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 125 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 55 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            Case "T0V", "T5V"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "00", "LB", "FA", "FB", "FC", _
                                         "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 65 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 70 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R", "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 95 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 85 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 95 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T", "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 145 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 105 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            Case "T2H", "T3H"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "00", "LB", "FA", "FB", "FC", _
                                         "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R", "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 105 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 105 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 125 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T", "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 165 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 105 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 125 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 55 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            Case "T2V", "T3V"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "00", "LB", "FA", "FB", "FC", _
                                         "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63", "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63", "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R", "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 75 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 75 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 80 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 85 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 95 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T", "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 75 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 85 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 35 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 35 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                                'RM1305005 2013/05/31 ローカル版との差異修正
                            Case "T2YH", "T3YH", "T2JH", "T2YD", "T2YDT", "T2YDU", _
                                 "T2YLH", "T3YLH", "T1H", "T2WH", "T3WH"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "00", "LB", "FA", "FB", "FC", _
                                         "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63", "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R", "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 105 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 105 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 120 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T", "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 165 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 105 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 120 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 55 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            Case "T2YV", "T3YV", "T2YFV", "T3YFV", "T2YMV", _
                                 "T3YMV", "T2JV", "T2YLV", "T3YLV", "T1V", "T8V", "T2WV", "T3WV"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "00", "LB", "FA", "FB", "FC", _
                                         "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63", "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R", "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 75 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 70 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 75 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 80 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T", "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 75 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 85 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 90 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 35 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 30 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 35 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            Case "T8H"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "00", "LB", "FA", "FB", "FC", _
                                         "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 65 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R", "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 95 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 95 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 100 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T", "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 155 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 125 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 55 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            Case "T8V"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "00", "LB", "FA", "FB", "FC", _
                                         "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50", "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 15 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 20 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 45 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40", "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80", "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 65 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R", "D"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 85 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 75 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 70 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 80 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "T", "4"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 125 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 135 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 110 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 115 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 125 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "H", "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                                    Case "40"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "50"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 50 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "63"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 35 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "80"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 35 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "100"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 40 Then
                                                            intKtbnStrcSeqNo = 4
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                        End Select
                                End Select
                        End Select
                    End If
                Case "SSG"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "GN", "NN", "GD", "ND"
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(1).Trim) = 0 Then
                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "32" And CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 6 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W8420"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select

                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "H", "R", "D"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "T0H", "T0V"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "12", "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 6 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If

                                            Case "20", "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "T1H", "T1V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T2H", "T2V", "T3H", "T3V"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T5H", "T5V"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "12", "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 6 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If

                                            Case "20", "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 4
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV", "T2YH", "T2YV", _
                                         "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                                         "T2YMV", "T3YMH", "T3YMV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "F2H", "F2V", "F3H", "F3V", "F3PH", "F3PV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 5 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "F2YH", "F2YV", "F3YH", "F3YV"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                End Select

                            Case "T"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "12", "16", "20"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "25", "32", "40", "50", "63", "80", "100"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 35 Then
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    End If
                Case "STG-B", "STG-M", "STG-K"
                    'RM1305005 2013/05/30 ローカル版との差異修正
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            Case "H", "R"
                                If InStr(objKtbnStrc.strcSelection.strOpSymbol(1), "Q") = 0 Then
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        'RM0811044 2008/12/15 T.Y "T1H", "T1V", "T8H", "T8V"追加
                                        Case "T0H", "T0V", "T1H", "T1V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V", "T8H", "T8V", _
                                             "SW11", "SW12", "SW13", "SW14", "SW15", "SW16", _
                                             "SW21", "SW22", "SW23", "SW24", "SW25", "SW26", "SW27"
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                            'Case "T2WH", "T2WV", "T3WH", "T3WV", "T2YH", "T2YV", "T3YH", "T3YV", _
                                            '     "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU", _
                                            '     "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                                        Case Else
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Else
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        Case "T0H", "T2H", "T3H", "T5H", "T2WH", "T3WH", _
                                             "SW11", "SW12", "SW13", "SW21", "SW22", "SW23", "SW27", _
                                             "SW37", "SW40", "SW47", "SW48"
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 20 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T1H", "T8H", "T2YH", "T3YH", "T2JH", "T2YD", "T2YDT", "T2YDU", _
                                             "SW31", "SW32", "SW33", "SW41", "SW42", "SW43"
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T1V", "T8V", "T2YV", "T3YV", "T2JV", _
                                             "SW34", "SW35", "SW36", "SW44", "SW45", "SW46"
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 15 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case Else
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                End If
                            Case "D"
                                If InStr(objKtbnStrc.strcSelection.strOpSymbol(1), "Q") = 0 Then
                                    '↓RM1310004 2013/10/01 ローカル版と差異修正
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        'RM0811044 2008/12/15 T.Y "T1H", "T1V", "T8H", "T8V"追加
                                        Case "T0H", "T0V", "T1H", "T1V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V", "T8H", "T8V", _
                                             "SW11", "SW12", "SW13", "SW14", "SW15", "SW16", _
                                             "SW21", "SW22", "SW23", "SW24", "SW25", "SW26", "SW27"
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T2WH", "T2WV", "T3WH", "T3WV", "T2YH", "T2YV", "T3YH", "T3YV", _
                                             "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU", _
                                             "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Else
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        Case "T0H", "T2H", "T3H", "T5H", "T2WH", "T3WH", _
                                             "SW11", "SW12", "SW13", "SW21", "SW22", "SW23", "SW40", "SW47"
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 20 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T1H", "T8H", "T2YH", "T3YH", "T2YFH", "T3YFH", "T2YMH", "T3YMH", "T2JH", _
                                             "T2YD", "T2YDT", "T2YDU", _
                                             "SW31", "SW32", "SW33", "SW41", "SW42", "SW43"
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 30 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T0V", "T5V", "T8V", "T2YV", "T3YV", "T2YFV", "T3YFV", "T2YMV", "T3YMV", "T2JV", _
                                             "SW34", "SW35", "SW36", "SW44", "SW45", "SW46"
                                            'RM0811044 2008/12/15 T.Y "T1V" ⇒ "T0V", "T5V"に変更
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 15 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                            '(↓2008/8/11 追加)
                                        Case "T1V", "T2V", "T3V", "T2WV", "T3WV", _
                                             "SW14", "SW15", "SW16", "SW24", "SW25", "SW26", "SW37", "SW38"
                                            'RM0811044 2008/12/15 T.Y "T0V", "T5V" ⇒ "T1V"に変更
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 5 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                            '(↑2008/8/11 追加)
                                        Case Else
                                            If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 10 Then
                                                intKtbnStrcSeqNo = 4
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                End If
                            Case "T"
                                If CDbl(objKtbnStrc.strcSelection.strOpSymbol(4)) < 25 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0200"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If
                Case "LFC-KL"
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "T"
                                If CDbl(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 35 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0200"
                                    fncCheckSelectOption = False
                                End If
                            Case "4"
                                If CDbl(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 50 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0200"
                                    fncCheckSelectOption = False
                                End If
                            Case "5"
                                If CDbl(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 65 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0200"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If
            End Select



            'If InStr(6, SrsPartsNo, "L") > 0 Then   'スイッチ付(L)の場合

            '    ' スイッチ判定
            '    If Len(Trim(ItemCode(intSw))) = 0 Then
            '        ' スイッチ無し
            '        If Val(Trim(ItemCode(intSt))) < 5 Then
            '            If SwitchFlag = True Then
            '                Call CtErrMsgDsp("DEER211") ' ストロークが製作可能な範囲にありません。
            '                f1.txtpartsno(intSt).SetFocus()
            '            End If

            '            Youso_Pattern_Chk2_CYL_STRCHK2 = False
            '            Exit Function
            '        End If
            '    Else
            '        ' スイッチ有り

            '        ' スイッチ・個数により判定
            '        Select Case Trim(ItemCode(intSwQty))
            '            Case "R", "H"
            '                ' スイッチが1個の場合
            '                If Val(Trim(ItemCode(intSt))) < 5 Then
            '                    If SwitchFlag = True Then
            '                        Call CtErrMsgDsp("DEER211") ' ストロークが製作可能な範囲にありません。
            '                        f1.txtpartsno(intSt).SetFocus()
            '                    End If

            '                    Youso_Pattern_Chk2_CYL_STRCHK2 = False
            '                    Exit Function
            '                End If
            '            Case "D"
            '                ' スイッチが2個の場合
            '                If Val(Trim(ItemCode(intSt))) < 10 Then
            '                    If SwitchFlag = True Then
            '                        Call CtErrMsgDsp("DEER211") ' ストロークが製作可能な範囲にありません。
            '                        f1.txtpartsno(intSt).SetFocus()
            '                    End If

            '                    Youso_Pattern_Chk2_CYL_STRCHK2 = False
            '                    Exit Function
            '                End If
            '            Case "T"
            '                ' スイッチが3個の場合
            '                If Val(Trim(ItemCode(intSt))) < 38 Then
            '                    If SwitchFlag = True Then
            '                        Call CtErrMsgDsp("DEER211") ' ストロークが製作可能な範囲にありません。
            '                        f1.txtpartsno(intSt).SetFocus()
            '                    End If

            '                    Youso_Pattern_Chk2_CYL_STRCHK2 = False
            '                    Exit Function
            '                End If
            '        End Select
            '    End If
            'End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

End Module
