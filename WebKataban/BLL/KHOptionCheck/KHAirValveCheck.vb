Module KHAirValveCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  空圧バルブチェック
    '*【概要】
    '*  空圧バルブをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/31      更新者：NII A.Takahashi
    '*  ・PDV3において、コイルオプション2CS/2HS/2ES/3RSを選定した場合、電圧AC100V/AC200V/DC12V/DC24V以外は
    '*  　選択できないよう修正
    '*                                          更新日：2007/07/06      更新者：NII A.Takahashi
    '*  ・MW4GB2,MW4GZ2(省配線のみ)において、AC電圧を選択し、かつ「W」配線を選択していない場合はエラーにする
    '*  　よう修正
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
                Case "MN4GB1", "MN4GB2", "MN4GBX12"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then

                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "1" And _
                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "8" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Contains("CL") Then
                                If Not objKtbnStrc.strcSelection.strOpSymbol(8).Contains("L") Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W9160"
                                    fncCheckSelectOption = False
                                End If
                            End If

                        End If
                    End If
                Case "N4GB1", "N4GB2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then

                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "1" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Contains("CL") Then
                                If Not objKtbnStrc.strcSelection.strOpSymbol(9).Contains("L") Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W9160"
                                    fncCheckSelectOption = False
                                End If
                            End If

                        End If
                    End If

                Case "MW4GB4", "MW4GZ4"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then
                        Dim intStationNo As Integer = Nothing
                        Dim strOptionY As String = ""

                        '連数設定
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                'オプションにＹを選択している時
                                Case "Y10", "Y20", "Y30", "Y40", "Y01", _
                                     "Y02", "Y03", "Y04", "Y11", "Y21", _
                                     "Y31", "Y41", "Y12", "Y22", "Y32", "Y42"
                                    strOptionY = strOpArray(intLoopCnt).Trim
                            End Select
                        Next

                        '**************************************************************
                        '* 省配線接続T7G7と入出ブロックオプション"Y**"関連チェック
                        '* （T8G7,T8D7選択時は、入出ブロックオプションの"Y**"の選択が必須）
                        '**************************************************************
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "T7ECPB7", "T7ECB7"
                                If strOptionY.Trim.Length = 0 Then
                                    intKtbnStrcSeqNo = 4
                                    'strMessageCd = "W9120"
                                    strMessageCd = "W9230"
                                    fncCheckSelectOption = False
                                End If
                            Case "T8G7"
                                If strOptionY.Trim.Length = 0 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W9210"
                                    fncCheckSelectOption = False
                                End If
                            Case "T8D7"
                                If strOptionY.Trim.Length = 0 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W9220"
                                    fncCheckSelectOption = False
                                End If
                            Case "T7ENB7", "T7ENPB7"
                                If strOptionY.Trim.Length = 0 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W9240"
                                    fncCheckSelectOption = False
                                End If

                            Case "T7EBB7", "T7EBPB7"
                                If strOptionY.Trim.Length = 0 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W9270"

                                    fncCheckSelectOption = False
                                End If

                            Case "T7EPB7", "T7EPPB7"
                                If strOptionY.Trim.Length = 0 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W9280"
                                    fncCheckSelectOption = False
                                End If

                        End Select
                    End If
                Case "MW3GA2", "MW4GA2", "MW4GB2", "MW4GZ2", "MW3GB2", "MW3GZ2"
                    '食品製造対応品追加のため  RM1702019  追加
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "T" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "X" Then
                        Dim intStationNo As Integer = Nothing
                        Dim strOptionY As String = ""

                        '連数設定
                        intStationNo = CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                'オプションにＹを選択している時
                                Case "Y10", "Y20", "Y30", "Y40", "Y01", _
                                     "Y02", "Y03", "Y04", "Y11", "Y21", _
                                     "Y31", "Y41", "Y12", "Y22", "Y32", "Y42"
                                    strOptionY = strOpArray(intLoopCnt).Trim
                            End Select
                        Next

                        '**************************************************************
                        '* 省配線接続T7G7と入出ブロックオプション"Y**"関連チェック
                        '* （T8G7,T8D7選択時は、入出ブロックオプションの"Y**"の選択が必須）
                        '**************************************************************
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "MW3GA2" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "T8G7", "T8D7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W8040"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T7ECBP7", "T7ECB7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W9120"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T7EBB7", "T7EBPB7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W9250"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T7EPB7", "T7EPPB7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W9260"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "T8G7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W9210"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T8D7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W9220"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T7ECB7", "T7ECPB7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W9230"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T7ENB7", "T7ENPB7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W9240"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T7EBB7", "T7EBPB7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W9250"
                                        fncCheckSelectOption = False
                                    End If
                                Case "T7EPB7", "T7EPPB7"
                                    If strOptionY.Trim.Length = 0 Then
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W9260"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        End If

                        '*********************
                        '* 連数製作不可チェック
                        '*********************
                        'オプションに入出力ブロック（Y**）を選択している時
                        If strOptionY.Trim.Length <> 0 Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "T8G1", "T8D1"
                                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "W" Then  '標準配線
                                        Select Case strOptionY
                                            Case "Y01"
                                                If intStationNo > 12 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "Y02"
                                                If intStationNo > 8 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Else                        'ダブル配線
                                        Select Case strOptionY
                                            Case "Y01"
                                                If intStationNo > 6 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "Y02"
                                                If intStationNo > 4 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    End If
                                Case "T8G2", "T8D2"
                                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "W" Then  '標準配線
                                        Select Case strOptionY
                                            Case "Y01", "Y02", "Y03", "Y04"
                                                'チェックなし
                                        End Select
                                    Else                        'ダブル配線
                                        Select Case strOptionY
                                            Case "Y01"
                                                If intStationNo > 14 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "Y02"
                                                If intStationNo > 12 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "Y03"
                                                If intStationNo > 10 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "Y04"
                                                If intStationNo > 8 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    End If
                                Case "T8G7", "T8D7"
                                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "W" Then  '標準配線
                                        Select Case strOptionY
                                            Case "Y10", "Y20", "Y30", "Y40"
                                                'チェックなし
                                            Case "Y11", "Y21", "Y31", "Y41"
                                                If intStationNo > 12 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "Y12", "Y22", "Y32", "Y42"
                                                If intStationNo > 8 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Else                        'ダブル配線
                                        Select Case strOptionY
                                            Case "Y10", "Y20", "Y30", "Y40"
                                                If intStationNo > 8 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "Y11", "Y21", "Y31", "Y41"
                                                If intStationNo > 6 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "Y12", "Y22", "Y32", "Y42"
                                                If intStationNo > 4 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    End If
                                Case "T8M6"
                                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "W" Then  '標準配線
                                        Select Case strOptionY
                                            Case "Y10", "Y20"
                                                'チェックなし
                                            Case "Y01", "Y11", "Y21"
                                                If intStationNo > 4 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Else                        'ダブル配線
                                        Select Case strOptionY
                                            Case "Y10", "Y20"
                                                If intStationNo > 4 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "Y01", "Y11", "Y21"
                                                If intStationNo > 2 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    End If
                                Case "T8MA"
                                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "W" Then  '標準配線
                                        Select Case strOptionY
                                            Case "Y10"
                                                'チェックなし
                                        End Select
                                    Else                        'ダブル配線
                                        Select Case strOptionY
                                            Case "Y10"
                                                If intStationNo > 2 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8030"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    End If
                            End Select
                        End If

                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "MW4GB2", "MW4GZ2", "MW4GA2", "MW3GA2", "MW3GB2", "MW3GZ2"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                    Case "1"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                            Case "2", "3", "4", "5"
                                            Case Else
                                                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "W" Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8430"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                        End Select
                        'RM1805036_二次電池シリーズ追加
                    ElseIf objKtbnStrc.strcSelection.strKeyKataban.Trim = "P" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "Y" Then
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "MW4GB2", "MW4GZ2", "MW4GA2", "MW3GA2"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                    Case "1"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                            Case "2", "3", "4", "5"
                                            Case Else
                                                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "W" Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W8430"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                        End Select
                    End If
                Case "AB41"
                    If (objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "03" Or objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "04" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "3N" Or objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "4N" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "3G" Or objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "4G") And _
                       (objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "8") And _
                       (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "V" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "W") Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 2) = "DC" Then
                            intKtbnStrcSeqNo = 9
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
                    '↓RM1110032 2011/11/05 Y.Tachi 
                    If (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3M" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3I" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3N" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3J") And _
                        (objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z") Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "AC" Then
                            intKtbnStrcSeqNo = 10
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
                    If (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5A" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5M" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5N" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5I" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5J") Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "DC" Then
                            intKtbnStrcSeqNo = 10
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
                Case "AB31"
                    'If (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3M" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3I" Or _
                    '    objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3N" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3J") And _
                    '    (objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z") Then
                    '    If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "AC" Then
                    '        intKtbnStrcSeqNo = 10
                    '        strMessageCd = "W8020"
                    '        fncCheckSelectOption = False
                    '    End If
                    'End If
                    If (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5A" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5M" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5N" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5I" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5J") Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "DC" Then
                            intKtbnStrcSeqNo = 10
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    'RM1402099 2014/02/25
                    If (objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z") Then

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            Case "DC5V", "DC6V", "DC12V", "DC14V", "DC21V", "DC24V", "DC25V", "DC26V", "DC48V", _
                                 "DC85V", "DC88V", "DC90V", "DC100V", "DC110V", "DC124V", "DC125V", "DC176V", _
                                 "DC230V", "DC240V"
                            Case "AC100V", "AC110V", "AC115V", "AC200V", "AC220V"
                            Case Else
                                intKtbnStrcSeqNo = 10
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                        End Select

                    End If


                    '↑RM1110032 2011/11/05 Y.Tachi 
                Case "PDV2"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case ""
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", _
                                     "AC120V", "AC200V", "AC220V", "AC230V", "AC240V", _
                                     "AC380V", "AC400V", "AC415V", "AC440V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2E", "2G"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", _
                                     "AC120V", "AC200V", "AC220V", "DC12V", "DC24V", _
                                     "DC48V", "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2H"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC200V", "DC12V", "DC24V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3A"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", _
                                     "AC120V", "AC200V", "AC220V", "AC230V", "AC240V", _
                                     "AC380V", "AC400V", "AC415V", "AC440V", "DC12V", _
                                     "DC24V", "DC48V", "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3K"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", _
                                     "AC120V", "AC200V", "AC220V", "DC12V", "DC24V", _
                                     "DC48V", "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3H"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC200V", _
                                     "AC220V", "DC24V", "DC100V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4A"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", _
                                     "AC120V", "AC200V", "AC220V", "AC230V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                    End Select
                Case "PDV3"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "2C", "2CG", "2CH"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC200V", "AC220V", "AC230V", "AC240V", "AC380V", _
                                     "AC400V", "AC415V", "AC440V", "DC12V", "DC24V", _
                                     "DC48V", "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2E", "2G"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC200V", "AC220V", "DC12V", "DC24V", "DC48V", _
                                     "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2H"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "AC100V", "AC110V", "AC200V", "AC220V", "DC12V", _
                                     "DC24V"
                                Case Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3T"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC200V", "AC220V", "DC12V", "DC24V", "DC48V", _
                                     "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3R"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC200V", _
                                     "DC12V", "DC24V"
                                Case Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4A"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "AC24V", "AC100V", "AC110V", "AC200V", "AC220V", _
                                     "AC230V"
                                Case Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2CS", "3RS"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "AC100V", "AC200V", "DC12V", "DC24V"
                                Case Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2HS", "2ES"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "AC100V", "AC200V", "AC220V", "DC12V", "DC24V"
                                Case Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                    End Select
                Case "PVS", "PKA", "PKW"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "3N"
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) = "DC" Then
                                If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 3)) >= 100 And _
                               Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 3)) <= 220 Then
                                Else
                                    If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 3)) = 24 Then
                                    Else
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                            Else
                                If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 3)) <= 220 Then
                                Else
                                    intKtbnStrcSeqNo = 6
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        Case "4N"
                            If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 3)) <= 220 Then
                            Else
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                            End If
                        Case "4M", "3M"
                            If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 3)) <= 400 Then
                            Else
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "PKS"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "3N"
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = "DC" Then
                                If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 3)) >= 100 And _
                               Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 3)) <= 220 Then
                                Else
                                    If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 3)) = 24 Then
                                    Else
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                            Else
                                If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 3)) <= 220 Then
                                Else
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        Case "4N"
                            If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 3)) <= 220 Then
                            Else
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                            End If
                        Case "4M", "3M"
                            If Val(Mid(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 3)) <= 400 Then
                            Else
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                            End If
                    End Select
                    'RM0912039 2009/12/39 Y.Miura 配線接続オプションの必須チェック漏れ対応
                    'Case "MN3E0", "MN4E0"
                Case "MN3E0", "MN4E0", "MN3E00", "MN4E00"
                    Dim bolOptionT As Boolean = False
                    Dim bolOptionD As Boolean = False

                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "TM1A", "TM1B", "TM1C", "TM52", "T30", _
                                 "T30R", "T50", "T50R", "T51", "T51R", _
                                 "T52", "T52R", "T53", "T53R", "T5B", _
                                 "T5C", "TX", "T631", "T6A0", "T6A1", _
                                 "T6C0", "T6C1", "T6E0", "T6E1", "T6G1", _
                                 "T6J0", "T6J1", "T6K1", "T7D1", "T7D2", _
                                 "T7G1", "T7G2", "T7N1", "T7N2", "T30N", "T30NR", _
                                 "T7EC1", "T7EC2", "T7ECT1", "T7ECT2" '2016/08/23 RM1608024 T7EC Append
                                bolOptionT = True
                            Case "D2", "D20", "D21", "D22", "D23", _
                                 "D2N", "D3"
                                bolOptionD = True
                        End Select

                        If bolOptionT = False Then
                            If bolOptionD = False Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W8010"
                                fncCheckSelectOption = False
                            End If
                        End If
                    Next
                    'RM0912039 2009/12/39 Y.Miura 配線接続オプションの必須チェック漏れ対応
                Case "MN3EX0", "MN4EX0"
                    Dim bolOptionT As Boolean = False
                    Dim bolOptionD As Boolean = False

                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "TM1A", "TM1B", "TM1C", "TM52", "T30", _
                                 "T30R", "T50", "T50R", "T51", "T51R", _
                                 "T52", "T52R", "T53", "T53R", "T5B", _
                                 "T5C", "TX", "T631", "T6A0", "T6A1", _
                                 "T6C0", "T6C1", "T6E0", "T6E1", "T6G1", _
                                 "T6J0", "T6J1", "T6K1", "T7D1", "T7D2", _
                                 "T7G1", "T7G2", "T7N1", "T7N2", "T30N", "T30NR", _
                                 "T7EC1", "T7EC2", "T7ECT1", "T7ECT2" '2016/08/23 RM1608024 T7EC Append
                                bolOptionT = True
                            Case "D2", "D20", "D21", "D22", "D23", _
                                 "D2N", "D3"
                                bolOptionD = True
                        End Select

                        If bolOptionT = False Then
                            If bolOptionD = False Then
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8010"
                                fncCheckSelectOption = False
                            End If
                        End If
                    Next
                Case "ADK21"
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                        intKtbnStrcSeqNo = 5
                        strMessageCd = "W8020"
                        fncCheckSelectOption = False
                    End If
                Case "APK21"
                    If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" Then
                            intKtbnStrcSeqNo = 6
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
            End Select
            '↓RM1110032 2011/11/05 Y.Tachi 
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "AG31", "AG33", "AG34"
                    'RM1402099 2014/02/05
                    If (objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z") Then

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            Case "DC5V", "DC6V", "DC12V", "DC14V", "DC21V", "DC24V", "DC25V", "DC26V", "DC48V", _
                                 "DC85V", "DC88V", "DC90V", "DC100V", "DC110V", "DC124V", "DC125V", "DC176V", _
                                 "DC230V", "DC240V"
                            Case "AC100V", "AC110V", "AC115V", "AC200V", "AC220V"
                            Case Else
                                intKtbnStrcSeqNo = 10
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                        End Select

                    End If
                Case "AG41", "AG43", "AG44"
                    If (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3M" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3I" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3N" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "3J") And _
                        (objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z") Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "AC" Then
                            intKtbnStrcSeqNo = 10
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
                    If (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5A" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5M" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5N" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5I" Or objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "5J") And _
                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z" Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "DC" Then
                            intKtbnStrcSeqNo = 10
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
                Case "AG41E4", "AG43E4", "AG44E4"
            End Select

            Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4)
                Case "GAB4", "GAB3"
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "GAB4" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "3M" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "3I" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "3N" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "3J") And _
                            (objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z") Then
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "GAB312", "GAB352", "GAB412", "GAB452"
                                    If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "AC" Then
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                    End If
                                Case "GAB422", "GAB462"
                                    If Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 2) = "AC" Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        End If
                    End If
                    If (objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5A" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5M" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5N" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5I" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5J") Then
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "GAB312", "GAB352", "GAB412", "GAB452"
                                If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "DC" Then
                                    intKtbnStrcSeqNo = 10
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                                End If
                            Case "GAB422", "GAB462"
                                If Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 2) = "DC" Then
                                    intKtbnStrcSeqNo = 9
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If

                    'RM1402099 2014/02/05
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "GAB3" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z") Then

                            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                Case "DC5V", "DC6V", "DC12V", "DC14V", "DC21V", "DC24V", "DC25V", "DC26V", "DC48V", _
                                     "DC85V", "DC88V", "DC90V", "DC100V", "DC110V", "DC124V", "DC125V", "DC176V", _
                                     "DC230V", "DC240V"
                                Case "AC100V", "AC110V", "AC115V", "AC200V", "AC220V"
                                Case Else
                                    intKtbnStrcSeqNo = 10
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select

                        End If
                    End If

                Case "GAG4", "GAG3"
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "GAG4" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "3M" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "3I" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "3N" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "3J") And _
                            (objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z") Then
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "AC" Then
                                intKtbnStrcSeqNo = 10
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                    If (objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5A" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5M" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5N" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5I" Or objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "5J") And _
                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z" Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) = "DC" Then
                            intKtbnStrcSeqNo = 10
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    'RM1402099 2014/02/05
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "GAG3" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z") Then

                            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                Case "DC5V", "DC6V", "DC12V", "DC14V", "DC21V", "DC24V", "DC25V", "DC26V", "DC48V", _
                                     "DC85V", "DC88V", "DC90V", "DC100V", "DC110V", "DC124V", "DC125V", "DC176V", _
                                     "DC230V", "DC240V"
                                Case "AC100V", "AC110V", "AC115V", "AC200V", "AC220V"
                                Case Else
                                    intKtbnStrcSeqNo = 10
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select

                        End If
                    End If
            End Select

            Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5)
                Case "ADK11"
                    If (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3M" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3I" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3N" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3J") And _
                       (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "Z") Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) = "AC" Then
                            intKtbnStrcSeqNo = 6
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
                    If (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "5A" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "5M" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "5N" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "5I" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "5J") And _
                       (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "Z") Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) = "DC" Then
                            intKtbnStrcSeqNo = 6
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    'DC24Vは標準でサージキラー内蔵のため選定不可
                    If objKtbnStrc.strcSelection.strOpSymbol(4).IndexOf("S") >= 0 Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "2E", "2G", "3A", "3M", "3N", "3I", "3J"
                            Case Else
                                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "DC24V" Then
                                    intKtbnStrcSeqNo = 6
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If

            End Select
            '↑RM1110032 2011/11/05 Y.Tachi 
        Catch ex As Exception

            fncCheckSelectOption = False

            Throw ex

        End Try

    End Function

End Module
