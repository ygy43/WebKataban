Module KHCylinderSSDCheck

#Region " Definition "

    Private bolS1StrChkSkipFlg As Boolean = False
    Private bolS2StrChkSkipFlg As Boolean = False
    Private bolOptionA2 As Boolean = False
    Private bolOptionN As Boolean = False
    Private bolOptionP4 As Boolean = False      'RM0906034 2009/08/05 Y.Miura　二次電池対応
    Private bolOptionP5 As Boolean = False
    Private bolOptionP7 As Boolean = False
    Private bolOptionS As Boolean = False

#End Region

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダチェック
    '*【概要】
    '*  シリンダＳＳＤシリーズをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*  ・受付No：RM0906034  二次電池対応機器対応
    '*                                      更新日：2009/08/05   更新者：Y.Miura
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Try

            fncCheckSelectOption = True
            bolOptionA2 = False
            bolOptionN = False
            bolOptionP4 = False     'RM0906034 2009/08/05 Y.Miura　二次電池対応機種
            bolOptionP5 = False
            bolOptionP7 = False
            bolOptionS = False

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "SSD"
                    '基本ベース毎にチェック
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        'RM0906034 2009/09/08 Y.Miura　二次電池対応機種
                        'Case ""
                        Case "", "4"
                            '基本ベースチェック
                            If fncStandardBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "D", "E"
                            '両ロッドベースチェック
                            If fncDoubleRodBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "K"
                            '高荷重ベースチェック
                            If fncHighLoadBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                            'RM0907070 2009/08/20 Y.Miura　二次電池対応
                        Case "P"
                            '高荷重ベースチェック
                            If fncHighLoadBaseP4Check(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If

                    End Select
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
    '*                                          更新日：2008/05/02      更新者：T.Sato
    '*  ・受付No.RM0804074 スイッチによる最小ストローク変更
    '*  ・受付№ RM0906034 2009/09/08 Y.Miura　二次電池対応
    '********************************************************************************************
    Private Function fncStandardBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncStandardBaseCheck = True

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "", " ", "L", "L1"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", " ", "B", "BQ", "BM", "BMO", "BT", "BTG1", "BT1", "BT1G1", "BT2", "BT2G1", _
                             "BO", "BG", "BG1", "BG2", "BG3", "BG4", "W", "WM", "WMO", "WT", "WT1", "WT2", _
                             "WO", "Q", "M", "MO", "T", "TG1", "T1", "T1G1", "T2", "T2G1", "O", "G", "G1", _
                             "G2", "G3", "G4", "G5"
                            'S1判定:最小パターン①
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                'スイッチ有無判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim.Length
                                    Case 0
                                        '内径判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "12", "16", "20", "25", "32", _
                                                 "40", "50", "63", "80", "100"
                                                '1mmから製作可能

                                                '↓2012/10/30 追加
                                                'ただしQ、BQは5mmから製作可能
                                                If Right(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) = "Q" And _
                                                   CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                                '↑2012/10/30 追加
                                            Case "125", "140", "160"
                                                '5mmから製作可能
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case Else
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            'RM0906034 2009/09/08 Y.Miura　二次電池対応
                                            'Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                            '     "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU"
                                            Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                                 "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU", _
                                                 "SW31", "SW32", "SW33", "SW34", "SW35", "SW36", "SW41", "SW42", "SW43", "SW44", "SW45", "SW46", _
                                                 "SW37", "SW40", "SW47", "SW48"
                                                '2色表示／予防保全出力SWの場合
                                                '内径判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                    Case "12", "16", "20", "25", "32", _
                                                         "40", "50", "63", "80", "100", _
                                                         "125", "140", "160"
                                                        '10mmから製作可能
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 7
                                                            strMessageCd = "W0200"
                                                            fncStandardBaseCheck = False
                                                            Exit Try
                                                        End If
                                                End Select
                                            Case Else
                                                'その他のスイッチの場合
                                                '内径判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                    Case "12", "16", "20", "25", "32", _
                                                         "40", "50", "63", "80", "100", _
                                                         "125", "140", "160"
                                                        '5mmから製作可能
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 7
                                                            strMessageCd = "W0200"
                                                            fncStandardBaseCheck = False
                                                            Exit Try
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            End If

                            'S2判定:最小パターン①
                            'スイッチ有無判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim.Length
                                Case 0
                                    'スイッチなし
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20", "25", "32", _
                                             "40", "50", "63", "80", "100"
                                            '1mmから製作可能

                                            '↓2012/10/30 追加
                                            'ただしQ、BQは5mmから製作可能
                                            If Right(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) = "Q" And _
                                               CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 14
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                            '↑2012/10/30 追加
                                        Case "125", "140", "160"
                                            '5mmから製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 14
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                Case Else
                                    'スイッチ有り
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                                        'RM0906034 2009/09/08 Y.Miura　二次電池対応
                                        'Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                        '     "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU"
                                        Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                            "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU", _
                                            "SW31", "SW32", "SW33", "SW34", "SW35", "SW36", "SW41", "SW42", "SW43", "SW44", "SW45", "SW46", _
                                            "SW37", "SW40", "SW47", "SW48"
                                            '2色表示／予防保全出力SWの場合
                                            '内径判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                Case "12", "16", "20", "25", "32", _
                                                     "40", "50", "63", "80", "100", _
                                                     "125", "140", "160"
                                                    '10mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 14
                                                        strMessageCd = "W0200"
                                                        fncStandardBaseCheck = False
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case Else
                                            'その他のスイッチの場合
                                            '内径判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                Case "12", "16", "20", "25", "32", _
                                                     "40", "50", "63", "80", "100", _
                                                     "125", "140", "160"
                                                    '5mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 5 Then
                                                        intKtbnStrcSeqNo = 14
                                                        strMessageCd = "W0200"
                                                        fncStandardBaseCheck = False
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                            End Select

                        Case "X", "XB", "XBT", "XBT2", "XM", "XT", "XT2", _
                             "Y", "YB", "YBT", "YBT2", "YM", "YT", "YT2"
                            'S1判定:最小パターン③
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                'スイッチ有無判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim.Length
                                    Case 0
                                        'スイッチなし
                                        '内径判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "12", "16", "20", "25", "32"
                                                '5mmから製作可能
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "40", "50"
                                                '10mmから製作可能
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case Else
                                        'スイッチ有り
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                                 "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU"
                                                '２色表示／予防保全出力ｓｗの場合
                                                '内径判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                    Case "12", "16", "20", "25", "32", _
                                                         "40", "50"
                                                        '10mmから製作可能
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 7
                                                            strMessageCd = "W0200"
                                                            fncStandardBaseCheck = False
                                                            Exit Try
                                                        End If
                                                End Select
                                            Case Else
                                                'その他のスイッチの場合
                                                '内径判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                    Case "12", "16", "20", "25", "32"
                                                        '5mmから製作可能
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 7
                                                            strMessageCd = "W0200"
                                                            fncStandardBaseCheck = False
                                                            Exit Try
                                                        End If
                                                    Case "40", "50"
                                                        '10mmから製作可能
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 7
                                                            strMessageCd = "W0200"
                                                            fncStandardBaseCheck = False
                                                            Exit Try
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            End If

                            'S2判定:最小パターン③
                            'スイッチ有無判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim.Length
                                Case 0
                                    'スイッチなし
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20", "25", "32"
                                            '5mmから製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 14
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "40", "50"
                                            '10mmから製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 14
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                Case Else
                                    'スイッチ有り
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                                        'RM0906034 2009/09/08 Y.Miura　二次電池対応
                                        'Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                        '     "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU"
                                        Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                             "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU", _
                                             "SW31", "SW32", "SW33", "SW34", "SW35", "SW36", "SW41", "SW42", "SW43", "SW44", "SW45", "SW46", _
                                             "SW37", "SW40", "SW47", "SW48"
                                            '2色表示／予防保全出力SWの場合
                                            '内径判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                Case "12", "16", "20", "25", "32", "40", "50"
                                                    '10mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 14
                                                        strMessageCd = "W0200"
                                                        fncStandardBaseCheck = False
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case Else
                                            'その他のSWの場合
                                            '内径判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                Case "12", "16", "20", "25", "32"
                                                    '5mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 5 Then
                                                        intKtbnStrcSeqNo = 14
                                                        strMessageCd = "W0200"
                                                        fncStandardBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "40", "50"
                                                    '10mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 14
                                                        strMessageCd = "W0200"
                                                        fncStandardBaseCheck = False
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                            End Select
                        Case "BT1L", "BG1T1L"
                            'S1
                            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim.Length = 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "20", "25"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "32", "40", "50", "63", "80", "100"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                                    Case "R", "H"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20", "25"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 15 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 35 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            End If

                            'S2
                            If objKtbnStrc.strcSelection.strOpSymbol(16).Trim.Length = 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "20", "25"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "32", "40", "50", "63", "80", "100"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(18).Trim
                                    Case "R", "H"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20", "25"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 15 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 35 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            End If
                        Case "T1L", "G1T1L"
                            If objKtbnStrc.strcSelection.strOpSymbol(16).Trim.Length = 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "20", "25"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "32", "40", "50", "63", "80", "100"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(18).Trim
                                    Case "R", "H"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20", "25"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 15 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 35 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 14
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            End If
                    End Select

                    'スイッチでL1を選択した場合は、最小ストロークは10mmになる
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "L1"
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                            End If
                            If objKtbnStrc.strcSelection.strOpSymbol(14).Trim <> "" Then
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 10 Then
                                    intKtbnStrcSeqNo = 14
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                            End If
                    End Select
                Case "L4"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", "B", "BQ", "BM", "BMO", "BT2", "BT2G1", "BO", _
                             "BG", "BG1", "BG4", "W", "WM", "WMO", "WT2", "WO", "Q", _
                             "M", "MO", "T2", "T2G1", "O", "G", "G1", "G2", "G3", "G4"
                            'S1判定:最小パターン④
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "40", "50", "63", "80", "100"
                                        '20mmから製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 20 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If

                            'S2判定
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "40", "50", "63", "80", "100"
                                    '20mmから製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) < 20 Then
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                    End Select
            End Select

            'ADD BY YGY 20140919    ↓↓↓↓↓↓
            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            Dim listOfSeries() As String = {"G1", "G2", "G3", "G4"}
            Dim blnContainGFlg As Boolean = False
            'バリエーションに「G1,G2,G3,G4」の有無判定
            For Each strSeries As String In listOfSeries
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim.Contains(strSeries) Then
                    blnContainGFlg = True
                    Exit For
                End If
            Next

            If blnContainGFlg Then
                'S1判定:最小パターン
                'ストローク有無判定
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                    'スイッチ有無判定
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim.Length > 0 Then
                        If Not fncGMinStrokeCheck(objKtbnStrc, "fncStandardBaseCheck", "S1") Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
                End If
                'S2判定:最小パターン
                'ストローク有無判定
                If objKtbnStrc.strcSelection.strOpSymbol(14).Trim <> "" Then
                    'スイッチ有無判定
                    If objKtbnStrc.strcSelection.strOpSymbol(16).Trim.Length > 0 Then
                        If Not fncGMinStrokeCheck(objKtbnStrc, "fncStandardBaseCheck", "S2") Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
                End If
            End If
            'ADD BY YGY 20140919    ↑↑↑↑↑↑

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            'S1,S2バリエーションで高荷重形(K,KM)が選択された場合は、スイッチバリエーションに関係なく高荷重形の最大ストロークを適用
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Case "K", "KM"
                    'S1判定:最大パターン⑨
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                        '内径判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "12", "16", "20"
                                '100mmまで製作可能
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 100 Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                            Case "25", "32", "40", "50"
                                '150mmまで製作可能
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 150 Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                            Case "63", "80", "100"
                                '200mmまで製作可能
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W0200"
                                    fncStandardBaseCheck = False
                                    Exit Try
                                End If
                        End Select

                        bolS1StrChkSkipFlg = True '以降のＳ１最大ストロークチェックはパスさせる
                    End If
            End Select

            Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                Case "K", "KM"
                    'S2判定:最大パターン⑨
                    '内径判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "12", "16", "20"
                            '100mmまで製作可能
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 100 Then
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case "25", "32", "40", "50"
                            '150mmまで製作可能
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 150 Then
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case "63", "80", "100"
                            '200mmまで製作可能
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 200 Then
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                    End Select

                    bolS2StrChkSkipFlg = True '以降のS2最大ストロークチェックはパスさせる
            End Select

            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "", "L"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", "G5"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン①
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16"
                                            '30mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 30 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "20", "25", "32", "40", "50", "63", _
                                             "80", "100"
                                            '50mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "125", "140", "160"
                                            '300mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン①
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16"
                                        '30mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 30 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "20", "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '50mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "125", "140", "160"
                                        '300mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "M", "MO", "T", "TG1", "T1", "T1G1", "T2", "T2G1", "O", _
                             "G", "G1", "G2", "G3", "G4"

                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン①
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20"
                                            '30mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 30 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50", "63", _
                                             "80", "100"
                                            '50mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "125", "140", "160"
                                            '300mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン①
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16", "20"
                                        '30mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 30 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '50mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "125", "140", "160"
                                        '300mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "X", "XB", "XBT", "XBT2", "XM", "XT", "XT2", _
                             "Y", "YB", "YBT", "YBT2", "YM", "YT", "YT2"

                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン③
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20", "25", "32"
                                            '標準ストロークの5mmと10mmのみ製作可能
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                Case "5", "10"
                                                Case Else
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                            End Select
                                        Case "40", "50"
                                            '標準ストロークの10mmと20mmのみ製作可能
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                Case "10", "20"
                                                Case Else
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                            End Select
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン③
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16", "20", "25", "32"
                                        '標準ストロークの5mmと10mmのみ製作可能
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                                            Case "5", "10"
                                            Case Else
                                                intKtbnStrcSeqNo = 14
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                        End Select
                                    Case "40", "50"
                                        '標準ストロークの10mmと20mmのみ製作可能
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                                            Case "10", "20"
                                            Case Else
                                                intKtbnStrcSeqNo = 14
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                        End Select
                                End Select
                            End If
                        Case "B", "BM", "BMO", "BT", "BTG1", "BT1", "BT1G1", "BT2", "BT2G1", "BO", "BG", _
                             "BG1", "BG2", "BG3", "BG4", "W", "WM", "WMO", "WT", "WT1", "WT2", "WO"

                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑤
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20"
                                            '30mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 30 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50", "63", _
                                             "80", "100"
                                            '50mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "125", "140", "160"
                                            '300mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑤
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16", "20"
                                        '30mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 30 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '50mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "125"
                                        '120mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 120 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "140"
                                        '110mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 110 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "160"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "Q"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑥
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "16"
                                            '100mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 100 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "20"
                                            '200mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50", "63", _
                                             "80", "100"
                                            '300mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑥
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16"
                                        '100mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 100 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "20"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '300mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "BQ"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑨
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20"
                                            '100mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 100 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50"
                                            '150mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 150 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "63", "80", "100"
                                            '200mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑨
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16", "20"
                                        '100mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 100 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50"
                                        '150mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 150 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "63", "80", "100"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "BT1L", "BG1T1L", "T1L", "G1T1L"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑩
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "16", "20"
                                            '30mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 30 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50", "63", _
                                             "80", "100"
                                            '50mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑩
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16", "20"
                                        '30mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 30 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '50mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                    End Select
                Case "L1"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", " ", "M", "MO", "T", "TG1", "T1", "T1G1", "T2", "T2G1", "O", _
                             "G", "G1", "G2", "G3", "G4"

                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン①
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20"
                                            '30mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 30 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50", "63", _
                                             "80", "100"
                                            '50mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "125", "140", "160"
                                            '300mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン①
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16", "20"
                                        '30mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 30 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '50mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "125", "140", "160"
                                        '300mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "X", "XB", "XBT", "XBT2", "XM", "XT", "XT2", _
                             "Y", "YB", "YBT", "YBT2", "YM", "YT", "YT2"

                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン③
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20", "25", "32"
                                            '標準ストロークの5mmと10mmのみ製作可能
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                Case "5", "10"
                                                Case Else
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                            End Select
                                        Case "40", "50"
                                            '標準ストロークの10mmと20mmのみ製作可能
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                Case "10", "20"
                                                Case Else
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncStandardBaseCheck = False
                                                    Exit Try
                                            End Select
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン③
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16", "20", "25", "32"
                                        '標準ストロークの5mmと10mmのみ製作可能
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                                            Case "5", "10"
                                            Case Else
                                                intKtbnStrcSeqNo = 14
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                        End Select
                                    Case "40", "50"
                                        '標準ストロークの10mmと20mmのみ製作可能
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                                            Case "10", "20"
                                            Case Else
                                                intKtbnStrcSeqNo = 14
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                        End Select
                                End Select
                            End If
                        Case "B", "BM", "BMO", "BT", "BTG1", "BT1", "BT1G1", "BT2", "BT2G1", "BO", "BG", _
                             "BG1", "BG2", "BG3", "BG4", "W", "WM", "WMO", "WT", "WT1", "WT2", "WO"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑤
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20"
                                            '30mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 30 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50", "63", _
                                             "80", "100"
                                            '50mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "125", "140", "160"
                                            '300mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑤
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16", "20"
                                        '30mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 30 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '50mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "125"
                                        '120mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 120 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "140"
                                        '110mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 110 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "160"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "Q"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑥
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "16"
                                            '100mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 100 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "20"
                                            '200mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50", "63", _
                                             "80", "100"
                                            '300mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑥
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16"
                                        '100mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 100 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "20"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '300mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "BQ"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑨
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20"
                                            '100mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 100 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50"
                                            '150mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 150 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "63", "80", "100"
                                            '200mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑨
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16", "20"
                                        '100mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 100 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50"
                                        '150mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 150 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "63", "80", "100"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "BT1L", "BG1T1L", "T1L", "G1T1L"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑩
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "16", "20"
                                            '30mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 30 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "25", "32", "40", "50", "63", _
                                             "80", "100"
                                            '50mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑩
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16", "20"
                                        '30mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 30 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '50mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                    End Select
                Case "L4"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", " ", "B", "BM", "BMO", "BT2", "BT2G1", "BO", "BG", "BG1", "BG4", "W", "WM", _
                             "WMO", "WT2", "WO", "M", "MO", "T2", "T2G1", "O", "G", "G1", "G2", "G3", "G4"

                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑦
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "40", "50", "63", "80", "100"
                                            '50mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑦
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "40", "50", "63", "80", "100"
                                        '50mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "Q"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑧
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "40", "50", "63", "80", "100"
                                            '300mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑧
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "40", "50", "63", "80", "100"
                                        '300mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                        Case "BQ"
                            If bolS1StrChkSkipFlg = False Then
                                'S1判定:最大パターン⑬
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "40", "50"
                                            '150mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 150 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "63", "80", "100"
                                            '200mmまで製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                                intKtbnStrcSeqNo = 7
                                                strMessageCd = "W0200"
                                                fncStandardBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                End If
                            End If

                            If bolS2StrChkSkipFlg = False Then
                                'S2判定:最大パターン⑬
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "40", "50"
                                        '150mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 150 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "63", "80", "100"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 14
                                            strMessageCd = "W0200"
                                            fncStandardBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If
                    End Select
            End Select

            '二段形の時、S1とS2の大小関係をチェックする
            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "W") <> 0 Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) Then
                Else
                    intKtbnStrcSeqNo = 14
                    strMessageCd = "W0200"
                    fncStandardBaseCheck = False
                    Exit Try
                End If
            End If

            '*-----<< Ⅲ．オプションチェック >>-----*
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "A2"
                        bolOptionA2 = True
                    Case "N"
                        bolOptionN = True
                    Case "P5", "P51"
                        bolOptionP5 = True
                    Case "P7", "P71"
                        bolOptionP7 = True
                    Case "S"
                        bolOptionS = True
                End Select
            Next

            If bolOptionA2 = True Then
                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "N" And bolOptionN = False Then
                    If InStr(strOptionSymbol, "N13") <> 0 Or _
                       InStr(strOptionSymbol, "N15") <> 0 Then
                    Else
                        intKtbnStrcSeqNo = 19
                        strMessageCd = "W0790"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                End If
            End If

            If bolOptionP5 = True Then
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "12", "16", "20", "25"
                        Case Else
                            intKtbnStrcSeqNo = 19
                            strMessageCd = "W0800"
                            fncStandardBaseCheck = False
                            Exit Try
                    End Select
                End If
            End If

            If bolOptionS = True Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "12", "16"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                            Case "5", "10", "15", "20", "25", "30"
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0830"
                                fncStandardBaseCheck = False
                                Exit Try
                        End Select
                    Case "20", "25", "32", "40", "50"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                            Case "5", "10", "15", "20", "25", "30", _
                                 "40", "50"
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0830"
                                fncStandardBaseCheck = False
                                Exit Try
                        End Select
                    Case "63", "80", "100"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                            Case "5", "10", "20", "30", "40", "50"
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0830"
                                fncStandardBaseCheck = False
                                Exit Try
                        End Select
                    Case "125", "140", "160"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                            Case "10", "20", "30", "40", "50", "60", _
                                 "70", "80", "90", "100", "110", "120", _
                                 "130", "140", "150", "160", "170", "180", _
                                 "190", "200", "210", "220", "230", "240", _
                                 "250", "260", "270", "280", "290", "300"
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0830"
                                fncStandardBaseCheck = False
                                Exit Try
                        End Select
                End Select
            End If

            Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                Case "T2YH", "T2YV", "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                     "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", "T2JH", "T2JV", "T1H", "T1V"
                    'スイッチでL1を選択し、かつオプションでP5,P51,P7,P71を選択した場合エラーメッセージを表示
                    If bolOptionP5 = True Or bolOptionP7 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L1" Then
                            intKtbnStrcSeqNo = 2
                            strMessageCd = "W0810"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If

                    'スイッチでLを選択し、かつ口径で12,16を選択し、かつオプションでP5,P51,P7,P71を選択せず、かつバリエーションでQを含まない場合はエラーメッセージを表示
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L" And _
                       (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "12" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "16") And _
                       (bolOptionP5 = False And bolOptionP7 = False) And _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") < 0 Then
                        intKtbnStrcSeqNo = 2
                        strMessageCd = "W0820"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
            End Select

            Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                Case "T2YH", "T2YV", "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                     "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", "T2JH", "T2JV", "T1H", "T1V"
                    'スイッチでL1を選択し、かつオプションでP5,P51,P7,P71を選択した場合エラーメッセージを表示
                    If bolOptionP5 = True Or bolOptionP7 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L1" Then
                            intKtbnStrcSeqNo = 2
                            strMessageCd = "W0810"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If

                    'スイッチでLを選択し、かつ口径で12,16を選択し、かつオプションでP5,P51,P7,P71を選択せず、かつバリエーションでQを含まない場合はエラーメッセージを表示
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L" And _
                       (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "12" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "16") And _
                       (bolOptionP5 = True And bolOptionP7 = True) And _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") < 0 Then
                        intKtbnStrcSeqNo = 2
                        strMessageCd = "W0820"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
            End Select

            '付属品「I」「Y」選択時はオプション「N」、もしくはロッド先端"N13","N15"を選択しなければいけない
            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(21).Trim, "I") <> 0 Or _
               InStr(1, objKtbnStrc.strcSelection.strOpSymbol(21).Trim, "Y") <> 0 Then
                If bolOptionN = True Or _
                (Len(Trim(objKtbnStrc.strcSelection.strOpSymbol(12).Trim)) <> 0 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "N") Or _
                   (InStr(1, strOptionSymbol, "N13") <> 0 Or _
                    InStr(1, strOptionSymbol, "N15") <> 0) Then
                Else
                    intKtbnStrcSeqNo = 21
                    strMessageCd = "W8590"
                    fncStandardBaseCheck = False
                    Exit Try
                End If
            End If

            'RM0906034 2009/09/08 Y.Miura　二次電池対応機種追加
            If objKtbnStrc.strcSelection.strKeyKataban.Equals("4") Then
                '二次電池対応
                If fncP4Check(objKtbnStrc, _
                                        intKtbnStrcSeqNo, _
                                        strOptionSymbol, _
                                        strMessageCd, _
                                        19) = False Then
                    fncStandardBaseCheck = False
                    Exit Try
                End If
            End If

        Catch ex As Exception

            Throw ex

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
    '*                                          更新日：2008/05/02      更新者：T.Sato
    '*  ・受付No.RM0804074 スイッチによる最小ストローク変更
    '********************************************************************************************
    Private Function fncDoubleRodBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncDoubleRodBaseCheck = True

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "", " ", "L"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "D", "DQ", "DM", "DMO", "DT", "DTG1", "DT1", "DT1G1", _
                            "DT2", "DT2G1", "DO", "DG", "DG1", "DG2", "DG3", "DG4", "KD", "KDM", _
                            "KDMO", "KDT", "KDT1", "KDTG1", "KDT1G1", "KDT2", "KDT2G1", "KDO", "KDG", _
                            "KDG1", "KDG2", "KDG3", "KDG4"
                            '判定:最小パターン②
                            'スイッチ有無判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length
                                Case 0
                                    'スイッチなし
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20", "25", "32", _
                                             "40", "50", "63", "80", "100"
                                            '1mmから製作可能

                                            '↓2012/10/30 追加
                                            'ただしDQは5mmから製作可能
                                            If Right(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) = "Q" And _
                                               CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncDoubleRodBaseCheck = False
                                                Exit Try
                                            End If
                                            '↑2012/10/30 追加
                                        Case "125", "140", "160"
                                            '10mmから製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 10 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncDoubleRodBaseCheck = False
                                            End If
                                    End Select
                                Case Else
                                    'スイッチ有り
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                             "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU"
                                            '2色表示／予防保全出力SWの場合
                                            '内径判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                Case "12", "16", "20", "25", "32", _
                                                     "40", "50", "63", "80", "100", _
                                                     "125", "140", "160"
                                                    '10mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncDoubleRodBaseCheck = False
                                                    End If
                                            End Select
                                        Case Else
                                            'その他のスイッチの場合
                                            '内径判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                Case "12", "16", "20", "25", "32", _
                                                     "40", "50", "63", "80", "100"
                                                    '5mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 5 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncDoubleRodBaseCheck = False
                                                    End If
                                                Case "125", "140", "160"
                                                    '10mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncDoubleRodBaseCheck = False
                                                    End If
                                            End Select
                                    End Select
                            End Select
                        Case "DT1L", "DG1T1L"
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length = 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncDoubleRodBaseCheck = False
                                        End If
                                    Case "20", "25"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncDoubleRodBaseCheck = False
                                        End If
                                    Case "32", "40", "50", "63", "80", "100"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncDoubleRodBaseCheck = False
                                        End If
                                End Select
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "R", "H"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncDoubleRodBaseCheck = False
                                                End If
                                            Case "20", "25"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 15 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncDoubleRodBaseCheck = False
                                                End If
                                            Case "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncDoubleRodBaseCheck = False
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncDoubleRodBaseCheck = False
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncDoubleRodBaseCheck = False
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncDoubleRodBaseCheck = False
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 35 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncDoubleRodBaseCheck = False
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncDoubleRodBaseCheck = False
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncDoubleRodBaseCheck = False
                                                End If
                                        End Select
                                End Select
                            End If
                    End Select
                Case "L4"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "D", "DQ", "DM", "DMO", "DT2", "DT2G1", "DO", "DG", "DG1", "DG2", "DG3", _
                            "DG4", "KD", "KDM", "KDMO", "KDT2", "KDT2G1", "KDO", "KDG", "KDG1", "KDG4"
                            '判定:最小パターン④
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "40", "50", "63", "80", "100"
                                    '20mmから製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 20 Then
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseCheck = False
                                    End If
                            End Select
                    End Select
            End Select

            'ADD BY YGY 20140919    ↓↓↓↓↓↓
            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            Dim listOfSeries() As String = {"G1", "G2", "G3", "G4"}
            'バリエーションに「G1,G2,G3,G4」の有無判定
            Dim blnContainGFlg As Boolean = False
            For Each strSeries As String In listOfSeries
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim.Contains(strSeries) Then
                    blnContainGFlg = True
                    Exit For
                End If
            Next

            If blnContainGFlg Then
                'S1判定:最小パターン
                'ストローク有無判定
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                    'スイッチ有無判定
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length > 0 Then
                        If Not fncGMinStrokeCheck(objKtbnStrc, "fncDoubleRodBaseCheck", "S1") Then
                            intKtbnStrcSeqNo = 6
                            strMessageCd = "W0200"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                        End If
                    End If
                End If
            End If
            'ADD BY YGY 20140919    ↑↑↑↑↑↑

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "", " ", "L"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "D", "DQ", "DM", "DMO", "DT", "DTG1", "DT1", "DT1G1", "DT1L", "DG1T1L", _
                            "DT2", "DT2G1", "DO", "DG", "DG1", "DG2", "DG3", "DG4"
                            '判定:最大パターン④
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16"
                                    '100mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 100 Then
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseCheck = False
                                    End If
                                Case "20"
                                    '200mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 200 Then
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseCheck = False
                                    End If
                                Case "25", "32", "40", "50", "63", _
                                     "80", "100", "125", "140", "160"
                                    '300mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 300 Then
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseCheck = False
                                    End If
                            End Select
                        Case "KD", "KDM", "KDMO", "KDT", "KDT1", "KDTG1", "KDT1G1", "KDT2", "KDT2G1", _
                            "KDO", "KDG", "KDG1", "KDG2", "KDG3", "KDG4"
                            '判定:最大パターン⑫
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "40", "50", "63", "80", "100"
                                    '300mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 300 Then
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseCheck = False
                                    End If
                            End Select
                    End Select
                Case "L4"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "D", "DQ", "DM", "DMO", "DT2", "DT2G1", "DO", "DG", "DG1", "DG2", "DG3", "DG4"
                            '判定:最大パターン⑫
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "40", "50", "63", "80", "100"
                                    '300mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 300 Then
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseCheck = False
                                    End If
                            End Select
                        Case "KD", "KDM", "KDMO", "KDT2", "KDT2G1", "KDO", "KDG", "KDG1", "KDG4"
                            '判定:最大パターン⑧
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "40", "50", "63", "80", "100"
                                    '300mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 300 Then
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W0200"
                                        fncDoubleRodBaseCheck = False
                                    End If
                            End Select
                    End Select
            End Select

            '*-----<< Ⅲ．オプションチェック >>-----*
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "N"
                        bolOptionN = True
                    Case "A2"
                        bolOptionA2 = True
                    Case "P5", "P51"
                        bolOptionP5 = True
                End Select
            Next

            If bolOptionA2 = True Then
                If bolOptionN = False And strOptionSymbol.Trim.Length = 0 Then
                    intKtbnStrcSeqNo = 11
                    strMessageCd = "W0790"
                    fncDoubleRodBaseCheck = False
                    Exit Try
                End If
            End If

            If bolOptionP5 = True Then
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "12", "16", "20", "25"
                        Case Else
                            intKtbnStrcSeqNo = 11
                            strMessageCd = "W0800"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                    End Select
                End If
            End If

            '付属品「I」「Y」選択時はオプション「N」、もしくはロッド先端"N13","N15"を選択しなければいけない
            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(13).Trim, "I") <> 0 Or _
               InStr(1, objKtbnStrc.strcSelection.strOpSymbol(13).Trim, "Y") <> 0 Then
                If bolOptionN = True Or _
                   (InStr(1, strOptionSymbol, "N13-N11") <> 0 Or _
                    InStr(1, strOptionSymbol, "N11-N13") <> 0) Then
                Else
                    intKtbnStrcSeqNo = 13
                    strMessageCd = "W8590"
                    fncDoubleRodBaseCheck = False
                    Exit Try
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncHighLoadBaseCheck
    '*【処理】
    '*  高荷重ベースチェック
    '*【概要】
    '*  高荷重ベースをチェックする
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
    '*                                          更新日：2008/05/02      更新者：T.Sato
    '*  ・受付No.RM0804074 スイッチによる最小ストローク変更
    '********************************************************************************************
    Private Function fncHighLoadBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncHighLoadBaseCheck = True

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "", " ", "L"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K", "KB", "KBM", "KBMO", "KBT", "KBTG1", "KBT1", "KBT1G1", "KBT2", _
                            "KBT2G1", "KBO", "KBU", "KBG", "KBG1", "KBG2", "KBG3", "KBG4", "KW", "KWM", _
                            "KWMO", "KWT", "KWT1", "KWT2", "KWO", "KM", "KMO", "KT", "KTG1", "KT1", _
                            "KT1G1", "KT2", "KT2G1", "KO", "KU", "KG", "KG1", "KG2", "KG3", "KG4", "KG5"
                            'S1判定:最小パターン①
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                'スイッチ有無判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length
                                    Case 0
                                        'スイッチなし
                                        '内径判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "12", "16", "20", "25", "32", _
                                                 "40", "50"
                                                '1mmから製作可能
                                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "KU" Then
                                                    'KUのみ5mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                        intKtbnStrcSeqNo = 7
                                                        strMessageCd = "W0200"
                                                        fncHighLoadBaseCheck = False
                                                        Exit Try
                                                    End If
                                                End If
                                                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "C" Then
                                                    'C(ゴムエアクッション付)は5mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                        intKtbnStrcSeqNo = 7
                                                        strMessageCd = "W0200"
                                                        fncHighLoadBaseCheck = False
                                                        Exit Try
                                                    End If
                                                End If
                                            Case "63", "80", "100"
                                                '1mmから製作可能
                                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "KU" Then
                                                    'KUのみ5mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                        intKtbnStrcSeqNo = 7
                                                        strMessageCd = "W0200"
                                                        fncHighLoadBaseCheck = False
                                                        Exit Try
                                                    End If
                                                End If
                                                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "C" Then
                                                    'C(ゴムエアクッション付)は10mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 7
                                                        strMessageCd = "W0200"
                                                        fncHighLoadBaseCheck = False
                                                        Exit Try
                                                    End If
                                                End If
                                            Case "125", "140", "160"
                                                '5mmから製作可能
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case Else
                                        'スイッチ有り
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                                 "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU"
                                                '2色表示／予防保全出力SWの場合
                                                '内径判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                    Case "12", "16", "20", "25", "32", _
                                                         "40", "50", "63", "80", "100"
                                                        '10mmから製作可能
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 10 Then
                                                            intKtbnStrcSeqNo = 7
                                                            strMessageCd = "W0200"
                                                            fncHighLoadBaseCheck = False
                                                            Exit Try
                                                        End If
                                                End Select
                                            Case Else
                                                'その他のスイッチの場合
                                                '内径判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                    Case "12", "16", "20", "25", "32", _
                                                         "40", "50", "63", "80", "100"
                                                        '5mmから製作可能
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 5 Then
                                                            intKtbnStrcSeqNo = 7
                                                            strMessageCd = "W0200"
                                                            fncHighLoadBaseCheck = False
                                                            Exit Try
                                                        End If
                                                End Select
                                        End Select
                                End Select
                            End If

                            'S2判定:最小パターン①
                            'スイッチ有無判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim.Length
                                Case 0
                                    'スイッチなし
                                    '内径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "12", "16", "20", "25", "32", _
                                              "40", "50"
                                            '1mmから製作可能
                                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "KU" Then
                                                'KUのみ5mmから製作可能
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            End If
                                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "C" Then
                                                'C(ゴムエアクッション付)は5mmから製作可能
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            End If
                                        Case "63", "80", "100"
                                            '1mmから製作可能
                                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "KU" Then
                                                'KUのみ5mmから製作可能
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 5 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            End If
                                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "C" Then
                                                'C(ゴムエアクッション付)は10mmから製作可能
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 7
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            End If
                                        Case "125", "140", "160"
                                            '5mmから製作可能
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 5 Then
                                                intKtbnStrcSeqNo = 13
                                                strMessageCd = "W0200"
                                                fncHighLoadBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                Case Else
                                    'スイッチ有り
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                                        Case "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                                             "T2WH", "T2WV", "T3WH", "T3WV", "T2JH", "T2JV", "T2YD", "T2YDT", "T2YDU"
                                            '2色表示／予防保全出力SWの場合
                                            '内径判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                Case "12", "16", "20", "25", "32", _
                                                     "40", "50", "63", "80", "100"
                                                    '10mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 13
                                                        strMessageCd = "W0200"
                                                        fncHighLoadBaseCheck = False
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case Else
                                            'その他のスイッチの場合
                                            '内径判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                                Case "12", "16", "20", "25", "32", _
                                                     "40", "50", "63", "80", "100"
                                                    '5mmから製作可能
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 5 Then
                                                        intKtbnStrcSeqNo = 13
                                                        strMessageCd = "W0200"
                                                        fncHighLoadBaseCheck = False
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                            End Select
                        Case "KT1L"
                            If objKtbnStrc.strcSelection.strOpSymbol(14).Trim.Length = 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 13
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "20", "25"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 15 Then
                                            intKtbnStrcSeqNo = 13
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "32", "40", "50", "63", "80", "100"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 13
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                                    Case "R", "H"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 13
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20", "25"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 15 Then
                                                    intKtbnStrcSeqNo = 13
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 10 Then
                                                    intKtbnStrcSeqNo = 13
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 13
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 13
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 13
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "16"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 35 Then
                                                    intKtbnStrcSeqNo = 13
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "20"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 45 Then
                                                    intKtbnStrcSeqNo = 13
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "25", "32", "40", "50", "63", "80", "100"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 13
                                                    strMessageCd = "W0200"
                                                    fncHighLoadBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            End If
                    End Select
                Case "L4"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K", "KB", "KBM", "KBMO", "KBT2", "KBT2G1", "KBO", "KBU", _
                            "KBG", "KBG1", "KBG4", "KW", "KWM", "KWMO", "KWT2", "KWO", _
                            "KM", "KMO", "KT2", "KT2G1", "KO", "KU", "KG", "KG1", "KG4"
                            'S1判定:最小パターン④
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "40", "50", "63", "80", "100"
                                        '20mmから製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 20 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If

                            'S2判定:最小パターン④
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "40", "50", "63", "80", "100"
                                    '20mmから製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) < 20 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                    End Select
            End Select

            'ADD BY YGY 20140919    ↓↓↓↓↓↓
            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            Dim listOfSeries() As String = {"G1", "G2", "G3", "G4"}
            'バリエーションに「G1,G2,G3,G4」の有無判定
            Dim blnContainGFlg As Boolean = False
            For Each strSeries As String In listOfSeries
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim.Contains(strSeries) Then
                    blnContainGFlg = True
                    Exit For
                End If
            Next

            If blnContainGFlg Then
                'S1判定:最小パターン
                'ストローク有無判定
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                    'スイッチ有無判定
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length > 0 Then
                        If Not fncGMinStrokeCheck(objKtbnStrc, "fncHighLoadBaseCheck", "S1") Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncHighLoadBaseCheck = False
                            Exit Try
                        End If
                    End If
                End If
                'S2判定:最小パターン
                'ストローク有無判定
                If objKtbnStrc.strcSelection.strOpSymbol(13).Trim <> "" Then
                    'スイッチ有無判定
                    If objKtbnStrc.strcSelection.strOpSymbol(14).Trim.Length > 0 Then
                        If Not fncGMinStrokeCheck(objKtbnStrc, "fncHighLoadBaseCheck", "S2") Then
                            intKtbnStrcSeqNo = 13
                            strMessageCd = "W0200"
                            fncHighLoadBaseCheck = False
                            Exit Try
                        End If
                    End If
                End If

            End If
            'ADD BY YGY 20140919    ↑↑↑↑↑↑

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "", " ", "L"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K", "KM", "KMO", "KT", "KTG1", "KT1", "KT1G1", "KT2", "KT2G1", "KO", "KU", _
                            "KG", "KG1", "KG2", "KG3", "KG4", "KG5"
                            'S1判定:最大パターン②
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16"
                                        '100mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 100 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "20"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '300mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If

                            'S2判定:最大パターン②
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16"
                                    '100mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 100 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                                Case "20"
                                    '200mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 200 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                                Case "25", "32", "40", "50", "63", _
                                     "80", "100"
                                    '300mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 300 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                        Case "KB", "KBM", "KBMO", "KBT", "KBTG1", "KBT1", "KBT1G1", "KBT2", "KBT2G1", _
                            "KBO", "KBU", "KBG", "KBG1", "KBG2", "KBG3", "KBG4", "KW", "KWM", "KWMO", _
                            "KWT", "KWT1", "KWT2", "KWO"
                            'S1判定:最大パターン⑨
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "12", "16", "20"
                                        '100mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 100 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50"
                                        '150mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 150 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "63", "80", "100"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If

                            'S2判定:最大パターン⑨
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16", "20"
                                    '100mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 100 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                                Case "25", "32", "40", "50"
                                    '150mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 150 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                                Case "63", "80", "100"
                                    '200mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 200 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                        Case "KT1L"
                            'S1判定:最大パターン⑪
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16"
                                        '100mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 100 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "20"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "25", "32", "40", "50", "63", _
                                         "80", "100"
                                        '300mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If

                            'S2判定:最大パターン⑪
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "16"
                                    '100mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 100 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                                Case "20"
                                    '200mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 200 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                                Case "25", "32", "40", "50", "63", _
                                     "80", "100"
                                    '300mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 300 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                    End Select
                Case "L4"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K", "KM", "KMO", "KT2", "KT2G1", "KO", "KU", "KG", "KG1", "KG4"
                            'S1判定:最大パターン⑧
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "40", "50", "63", "80", "100"
                                        '300mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 300 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If

                            'S2判定:最大パターン⑧
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "40", "50", "63", "80", "100"
                                    '300mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 300 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                        Case "KB", "KBM", "KBMO", "KBT2", "KBT2G1", "KBO", "KBU", "KBG", "KBG1", "KBG4", _
                            "KW", "KWM", "KWMO", "KWT2", "KWO"
                            'S1判定:最大パターン⑬
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "40", "50"
                                        '150mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 150 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "63", "80", "100"
                                        '200mmまで製作可能
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 200 Then
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0200"
                                            fncHighLoadBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            End If

                            'S2判定:最大パターン⑬
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "40", "50"
                                    '150mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 150 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                                Case "63", "80", "100"
                                    '200mmまで製作可能
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 200 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0200"
                                        fncHighLoadBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                    End Select
            End Select

            '二段形の時、S1とS2の大小関係をチェックする
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) Then
                Else
                    intKtbnStrcSeqNo = 13
                    strMessageCd = "W0610"
                    fncHighLoadBaseCheck = False
                    Exit Try
                End If
            End If

            '*-----<< Ⅲ．オプションチェック >>-----*
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "N"
                        bolOptionN = True
                    Case "A2"
                        bolOptionA2 = True
                    Case "P5", "P51"
                        bolOptionP5 = True
                    Case "S"
                        bolOptionS = True
                End Select
            Next

            If bolOptionA2 = True Then
                If objKtbnStrc.strcSelection.strOpSymbol(11).Trim <> "N" And bolOptionN = False Then
                    If InStr(strOptionSymbol, "N13") <> 0 Or _
                       InStr(strOptionSymbol, "N15") Then
                    Else
                        intKtbnStrcSeqNo = 17
                        strMessageCd = "W0790"
                        fncHighLoadBaseCheck = False
                        Exit Try
                    End If
                End If
            End If

            If bolOptionP5 = True Then
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "12", "16", "20", "25"
                        Case Else
                            intKtbnStrcSeqNo = 17
                            strMessageCd = "W0800"
                            fncHighLoadBaseCheck = False
                            Exit Try
                    End Select
                End If
            End If

            If bolOptionS = True Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "12", "16"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                            Case "5", "10", "15", "20", "25", "30", "40", "50", "60", "70", "80", "90", "100"
                                intKtbnStrcSeqNo = 13
                                strMessageCd = "W0830"
                                fncHighLoadBaseCheck = False
                                Exit Try
                        End Select
                    Case "20"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                            Case "5", "10", "15", "20", "25", "30", "40", "50", "60", "70", "80", "90", "100", _
                                 "110", "120", "130", "140", "150", "160", "170", "180", "190", "200"
                                intKtbnStrcSeqNo = 13
                                strMessageCd = "W0830"
                                fncHighLoadBaseCheck = False
                                Exit Try
                        End Select
                    Case "25", "32", "40", "50"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                            Case "10", "15", "20", "25", "30", "40", "50", "60", "70", "80", "90", "100", _
                                 "110", "120", "130", "140", "150", "160", "170", "180", "190", "200", _
                                 "210", "220", "230", "240", "250", "260", "270", "280", "290", "300"
                                intKtbnStrcSeqNo = 13
                                strMessageCd = "W0830"
                                fncHighLoadBaseCheck = False
                                Exit Try
                        End Select
                    Case "63", "80", "100"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                            Case "10", "20", "30", "40", "50", "60", "70", "80", "90", "100", _
                                 "110", "120", "130", "140", "150", "160", "170", "180", "190", "200", _
                                 "210", "220", "230", "240", "250", "260", "270", "280", "290", "300"
                                intKtbnStrcSeqNo = 13
                                strMessageCd = "W0830"
                                fncHighLoadBaseCheck = False
                                Exit Try
                        End Select
                End Select
            End If

            '付属品「I」「Y」選択時はオプション「N」、もしくはロッド先端"N13","N15"を選択しなければいけない
            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(19).Trim, "I") <> 0 Or _
               InStr(1, objKtbnStrc.strcSelection.strOpSymbol(19).Trim, "Y") <> 0 Then
                If bolOptionN = True Or _
                (Len(Trim(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)) <> 0 And objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "N") Or _
                   (InStr(1, strOptionSymbol, "N13") <> 0 Or _
                    InStr(1, strOptionSymbol, "N15") <> 0) Then
                Else
                    intKtbnStrcSeqNo = 19
                    strMessageCd = "W8590"
                    fncHighLoadBaseCheck = False
                    Exit Try
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncHighLoadBaseP4Check
    '*【処理】
    '*  高荷重ベースチェック 二次電池対応
    '*【概要】
    '*  高荷重ベース 二次電池対応をチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*  ・受付No：RM0907070  二次電池対応機器対応　新規追加
    '*                                      更新日：2009/08/20   更新者：Y.Miura
    '********************************************************************************************
    Private Function fncHighLoadBaseP4Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncHighLoadBaseP4Check = True
            bolOptionP4 = False
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOptionP4 = True
                End Select
            Next
            'P4の必須チェック
            If Not bolOptionP4 Then
                intKtbnStrcSeqNo = 17
                strMessageCd = "W8770"
                fncHighLoadBaseP4Check = False
                Exit Try
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncP4Check
    '*【処理】
    '*  二次電池対応機器チェック
    '*【概要】
    '*  二次電池が含まれるかをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*  <Integer>      intOptionPos         要素位置　　　　　   
    '*【戻り値】
    '*  <Boolean>
    '*【更新】
    '*  ・受付No：RM0906034  二次電池対応機器対応　新規追加
    '*                                      更新日：2009/09/08   更新者：Y.Miura
    '********************************************************************************************
    Private Function fncP4Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String, _
                                          ByVal intOptionPos As Integer) As Boolean

        Try

            fncP4Check = True

            '二次電池対応
            Dim bolOpP4 As Boolean = False
            Dim strOpArray() As String
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt As Integer = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOpP4 = True
                End Select
            Next
            'P4の必須チェック
            If Not bolOpP4 Then
                intKtbnStrcSeqNo = intOptionPos
                strMessageCd = "W8770"
                fncP4Check = False
                Exit Try
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncGMinStrokeCheck
    '*【処理】
    '*  にG1,G2,G3,G4を含む場合最少ストロークのチェック
    '*【概要】
    '*  にG1,G2,G3,G4を含む場合最少ストロークをチェックする
    '*【戻り値】
    '*  <Boolean>
    '*【更新】
    '*  ・受付No：RM0906034  二次電池対応機器対応　新規追加
    '*                                      更新日：2014/09/19   更新者：YGY
    '********************************************************************************************
    Private Function fncGMinStrokeCheck(ByVal objKtbnStrc As KHKtbnStrc, ByVal strFunctionName As String, ByVal strType As String) As Boolean
        fncGMinStrokeCheck = True
        Dim intBore As Integer = 4
        Dim intStroke As Integer
        Dim intNumber As Integer

        '要素位置の設定
        Select Case strFunctionName
            Case "fncStandardBaseCheck"
                Select Case strType
                    Case "S1"
                        intStroke = 7
                        intNumber = 11
                    Case "S2"
                        intStroke = 14
                        intNumber = 18
                End Select
            Case "fncDoubleRodBaseCheck"
                intStroke = 6
                intNumber = 10
            Case "fncHighLoadBaseCheck"
                Select Case strType
                    Case "S1"
                        intStroke = 7
                        intNumber = 10
                    Case "S2"
                        intStroke = 13
                        intNumber = 16
                End Select
        End Select
        '最小ストロークの判断
        Select Case objKtbnStrc.strcSelection.strOpSymbol(intNumber).Trim
            Case "R", "H"
                Select Case objKtbnStrc.strcSelection.strOpSymbol(intBore).Trim
                    Case "16", "20", "25", "32", "40", "50", "63", "80", "100"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(intStroke).Trim) < 10 Then
                            fncGMinStrokeCheck = False
                        End If
                End Select
            Case "D"
                Select Case objKtbnStrc.strcSelection.strOpSymbol(intBore).Trim
                    Case "16", "20", "25", "32", "40", "50", "63", "80", "100"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(intStroke).Trim) < 10 Then
                            fncGMinStrokeCheck = False
                        End If
                End Select
            Case "T"
                Select Case objKtbnStrc.strcSelection.strOpSymbol(intBore).Trim
                    Case "16"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(intStroke).Trim) < 25 Then
                            fncGMinStrokeCheck = False
                        End If
                    Case "25", "32", "40", "50", "63", "80", "100"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(intStroke).Trim) < 35 Then
                            fncGMinStrokeCheck = False
                        End If
                End Select
        End Select
    End Function
End Module
