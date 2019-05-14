Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports WebKataban.CdCst

Public Class KHCombinationOut

    ''' <summary>
    ''' 全ての情報をメモリに読み込む
    ''' </summary>
    ''' <param name="strSeriesKata"></param>
    ''' <param name="strKeyKata"></param>
    ''' <param name="strPriceNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCacheTable(strSeriesKata As String, strKeyKata As String, _
                                         Optional strPriceNo As String = "") As DataSet
        GetCacheTable = New DataSet
        Dim dt_Option As New DS_KatOut.DT_OptionDataTable
        Dim dt_ElePattern As New DS_KatOut.kh_ele_patternDataTable
        Dim dt_VolStd As DS_KatOut.kh_std_voltage_mstDataTable

        Dim dt_fullPrice As New DS_KatOut.kh_priceDataTable
        Dim dt_accPrice As New DS_KatOut.kh_accumulate_priceDataTable
        Dim dt_screPrice As New DS_KatOut.kh_screw_kataban_mstDataTable

        Using da As New DS_KatOutTableAdapters.DT_OptionTableAdapter
            dt_Option = da.GetOptnameData(strSeriesKata, strKeyKata, Now, "en", "ja")
        End Using
        dt_Option.TableName = "Option"
        Using da As New DS_KatOutTableAdapters.kh_ele_patternTableAdapter
            dt_ElePattern = da.GetElePatternData(strSeriesKata, strKeyKata, Now)
        End Using
        dt_ElePattern.TableName = "ElePattern"
        Using da As New DS_KatOutTableAdapters.kh_std_voltage_mstTableAdapter
            dt_VolStd = da.GetDataBy(strSeriesKata, strKeyKata, Now)
        End Using
        dt_VolStd.TableName = "dt_VolStd"

        Using da As New DS_KatOutTableAdapters.kh_priceTableAdapter
            dt_fullPrice = da.GetData(strSeriesKata & "%", Now)
        End Using
        dt_fullPrice.TableName = "dt_fullPrice"
        Using da As New DS_KatOutTableAdapters.kh_accumulate_priceTableAdapter
            dt_accPrice = da.GetData(strSeriesKata & "%", Now)
        End Using
        dt_accPrice.TableName = "dt_accPrice"
        Using da As New DS_KatOutTableAdapters.kh_screw_kataban_mstTableAdapter
            dt_screPrice = da.GetData(strSeriesKata & "%")
        End Using
        dt_screPrice.TableName = "dt_screPrice"

        Dim dt_Acc As New DS_KatOut.kh_accumulate_priceDataTable
        Using da As New DS_KatOutTableAdapters.kh_accumulate_priceTableAdapter
            Select Case strPriceNo
                Case "O0"
                    dt_Acc = da.GetData("LCX" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("LCG" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("LCR" & "%", Now)
                Case "02", "03", "17", "18"
                    dt_Acc = da.GetData("MULTI-SCREW" & "%", Now)
                Case "06"
                    dt_Acc = da.GetData("M" & "%", Now)
                Case "E5"
                    dt_Acc = da.GetData("UCAC2" & "%", Now)
                Case "32"
                    dt_Acc = da.GetData("CKV2" & "%", Now)
                Case "33"
                    If strSeriesKata = "COVN2" Then
                        dt_Acc = da.GetData("COVP2" & "%", Now)
                        If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    End If
                    dt_Acc = da.GetData("C*V2" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("CAV2" & "%", Now)
                Case "35", "36"
                    dt_Acc = da.GetData("A*4*" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("C*V2" & "%", Now)
                Case "60"
                    dt_Acc = da.GetData("4G4" & "%", Now)
                Case "65"
                    dt_Acc = da.GetData("SCG" & "%", Now)
                Case "23", "85", "R9", "S1"
                    dt_Acc = da.GetData("SCP" & "%", Now)
                Case "24"
                    dt_Acc = da.GetData("MDC2" & "%", Now)
                Case "26"
                    dt_Acc = da.GetData("UCA2" & "%", Now)
                Case "29"
                    dt_Acc = da.GetData("FC" & "%", Now)
                Case "31"
                    dt_Acc = da.GetData("GRC" & "%", Now)
                Case "47"
                    dt_Acc = da.GetData("LCY" & "%", Now)
                Case "48"
                    dt_Acc = da.GetData("GLC" & "%", Now)
                Case "49"
                    dt_Acc = da.GetData("F" & "%", Now)
                Case "12"
                    dt_Acc = da.GetData("CMA2" & "%", Now)
                Case "15"
                    dt_Acc = da.GetData("LCS" & "%", Now)
                Case "52"
                    dt_Acc = da.GetData("P51" & "%", Now)
                Case "66", "67"
                    dt_Acc = da.GetData("RV" & "%", Now)
                Case "K4"
                    dt_Acc = da.GetData("SCP" & "%", Now)
                Case "K5"
                    dt_Acc = da.GetData("USC" & "%", Now)
                Case "D2"
                    dt_Acc = da.GetData("SRL" & "%", Now)
                Case "72"
                    dt_Acc = da.GetData("SRM" & "%", Now)
                Case "62"
                    dt_Acc = da.GetData("STG" & "%", Now)
                Case "79"
                    dt_Acc = da.GetData("STR2" & "%", Now)
                Case "82"
                    dt_Acc = da.GetData("ULKP" & "%", Now)
                Case "Q0"
                    dt_Acc = da.GetData("BBS-" & "%", Now)
                Case "P3"
                    dt_Acc = da.GetData("CXU-" & "%", Now)
                Case "P7"
                    dt_Acc = da.GetData("MCP" & "%", Now)
                Case "H0"
                    dt_Acc = da.GetData("MSD" & "%", Now)
                Case "H6"
                    dt_Acc = da.GetData("DT" & "%", Now)
                Case "D5"
                    dt_Acc = da.GetData("EV" & "%", Now)
                Case "N3"
                    dt_Acc = da.GetData("LCM" & "%", Now)
                Case "K1"
                    dt_Acc = da.GetData("F6" & "%", Now)
                Case "F9"
                    dt_Acc = da.GetData("FCK" & "%", Now)
                Case "G1"
                    dt_Acc = da.GetData("MFC" & "%", Now)
                Case "G8"
                    dt_Acc = da.GetData("KML" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("2R" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("3R" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("4R" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("5R" & "%", Now)
                Case "I8"
                    dt_Acc = da.GetData("FD" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("FS" & "%", Now)
                Case "L0", "L1"
                    dt_Acc = da.GetData("NR" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("RB" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("RJ" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                Case "C7"
                    dt_Acc = da.GetData("MRL" & "%", Now)
                Case "D7"
                    dt_Acc = da.GetData("PCC" & "%", Now)
                Case "J7"
                    dt_Acc = da.GetData("PF" & "%", Now)
                Case "A7", "N7", "N4", "B4"
                    dt_Acc = da.GetData("NW" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("W3G" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("W4G" & "%", Now)
                Case "F1", "F4", "F3"
                    dt_Acc = da.GetData("NSR" & "%", Now)
                Case "H4"
                    dt_Acc = da.GetData("RCC2" & "%", Now)
                Case "Q4"
                    dt_Acc = da.GetData("RG" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("PC" & "%", Now)
                Case "C8"
                    dt_Acc = da.GetData("P1100" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("P4100" & "%", Now)
                Case "A8"
                    dt_Acc = da.GetData("RG" & "%", Now)
                Case "Q6"
                    dt_Acc = da.GetData("SFR" & "%", Now)
                Case "87", "46"
                    dt_Acc = da.GetData("SMD2" & "%", Now)
                Case "83"
                    dt_Acc = da.GetData("USSD" & "%", Now)
                Case "B7"
                    dt_Acc = da.GetData("VSXM" & "%", Now)
                Case "M1"
                    dt_Acc = da.GetData("F" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)
                    dt_Acc = da.GetData("W" & "%", Now)
                Case "D6", "E6", "A9"
                    dt_Acc = da.GetData("W" & "%", Now)
                Case "12"
                    dt_Acc = da.GetData("CMA2" & "%", Now)
                Case "OA"
                    dt_Acc = da.GetData("LCR" & "%", Now)
                Case "N0"
                    dt_Acc = da.GetData("JSK" & "%", Now)
                Case "27"
                    dt_Acc = da.GetData("JSM" & "%", Now)
                Case "M8", "M6", "M7", "P5"
                    dt_Acc = da.GetData("SSD" & "%", Now)
                Case "L4"
                    dt_Acc = da.GetData("SHC" & "%", Now)
                Case "25"
                    dt_Acc = da.GetData("SRG" & "%", Now)
                Case "B0"
                    dt_Acc = da.GetData("CMK" & "%", Now)
                Case "81"
                    dt_Acc = da.GetData("ULK" & "%", Now)
                Case "97"
                    dt_Acc = da.GetData("JSG" & "%", Now)
                Case "E7"
                    dt_Acc = da.GetData("R" & "%", Now)
                Case "90"
                    dt_Acc = da.GetData("STS" & "%", Now)
                Case "B3"
                    dt_Acc = da.GetData("GAMD" & "%", Now)
                Case "98"
                    dt_Acc = da.GetData("PV5" & "%", Now)
                Case "O1"
                    'ADD BY YGY 20140929    ↓↓↓↓↓↓
                    dt_Acc = da.GetData("AMD0" & "%", Now)
                    'ADD BY YGY 20140929    ↑↑↑↑↑↑
            End Select
            If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)

            dt_Acc = New DS_KatOut.kh_accumulate_priceDataTable
            Select Case strSeriesKata
                Case "CKH2", "CKLB2"
                    dt_Acc = da.GetData("CKL2" & "%", Now)
                Case "P1100", "P8100", "P4100"
                    dt_Acc = da.GetData("P*100" & "%", Now)
                Case "PVSE2", "PVSE4"
                    dt_Acc = da.GetData("PVSE*" & "%", Now)
                Case "MMD302"
                    dt_Acc = da.GetData("MMD3*2" & "%", Now)
                Case "MMD402"
                    dt_Acc = da.GetData("MMD4*2" & "%", Now)
                Case "MMD502"
                    dt_Acc = da.GetData("MMD5*2" & "%", Now)
                Case "RGIB", "RGCD", "RGCM", "RGID", "RGIM"
                    dt_Acc = da.GetData("RG*" & "%", Now)
                Case "N3S0"
                    dt_Acc = da.GetData("N4S0" & "%", Now)
                Case "N4S0"
                    dt_Acc = da.GetData("N3S0" & "%", Now)
                Case "MDV-L"
                    dt_Acc = da.GetData("MDV" & "%", Now)
                Case "W4GB4", "W4GZ4"
                    dt_Acc = da.GetData("W4G4" & "%", Now)
                Case "W3GA2", "W4GA2", "W4GB2", "W3GB2", "W3GZ2"
                    dt_Acc = da.GetData("W4G2" & "%", Now)
                Case "3SA1", "4SA1", "4SB1"
                    dt_Acc = da.GetData("4S1" & "%", Now)
                Case "AM4F0"
                    dt_Acc = da.GetData("A4F0" & "%", Now)
                    'ADD BY YGY 20140929    ↓↓↓↓↓↓
                    If dt_Acc.Rows.Count > 0 Then
                        dt_accPrice.Merge(dt_Acc)
                    End If
                    dt_Acc = da.GetData("4F0" & "%", Now)
                    If dt_Acc.Rows.Count > 0 Then
                        dt_accPrice.Merge(dt_Acc)
                    End If
                    dt_Acc = da.GetData("M4F0" & "%", Now)
                    'ADD BY YGY 20140929    ↑↑↑↑↑↑
                Case "A4F0"
                    dt_Acc = da.GetData("4F0" & "%", Now)
                Case "4SA0", "4SA1"
                    dt_Acc = da.GetData("M4SA" & "%", Now)
                Case "3PA1", "3PA2"
                    dt_Acc = da.GetData("M3PA" & "%", Now)
                Case "M3MA0", "M3MB0", "M3PA1", "M3PA2", "M3PB1", "M3PB2", "M4F0", "M4F1", "M4F2", _
                    "M4F3", "M4F4", "M4F5", "M4F6", "M4F7", "M4L2", "M4L3", "M4SA0", "M4SB0"
                    'ADD BY YGY 20140929    ↓↓↓↓↓↓
                    dt_Acc = da.GetData(strSeriesKata.Substring(1) & "%", Now)
                    'ADD BY YGY 20140929    ↑↑↑↑↑↑
            End Select
        End Using
        If dt_Acc.Rows.Count > 0 Then dt_accPrice.Merge(dt_Acc)

        GetCacheTable.Tables.Add(dt_Option.Copy)
        GetCacheTable.Tables.Add(dt_ElePattern.Copy)
        GetCacheTable.Tables.Add(dt_VolStd.Copy)
        GetCacheTable.Tables.Add(dt_fullPrice.Copy)
        GetCacheTable.Tables.Add(dt_accPrice.Copy)
        GetCacheTable.Tables.Add(dt_screPrice.Copy)
    End Function

    ''' <summary>
    ''' 組合せ出力可否の判断
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="HTValue"></param>
    ''' <param name="Now_Pos"></param>
    ''' <param name="strMsgCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteCheck(objKtbnStrc As KHKtbnStrc, HTValue As Hashtable, _
                                        ByRef Now_Pos As Integer, ByRef strMsgCd As String) As Boolean

        ExecuteCheck = False
        Try
            Dim str() As String = Nothing
            ' 形番組合せ処理が実行可能なオプションが選択されているかチェックする
            Select Case objKtbnStrc.strcSelection.strSpecNo
                Case "09"    '旧"O"
                    If Not HTValue("6") Is Nothing Then
                        str = HTValue("6").split(",")
                        For I1 = 1 To str.Length - 1
                            If Len(Trim(str(I1))) <> 0 Then
                                strMsgCd = "DEER432"
                                Now_Pos = 6
                                Exit Function
                            End If
                        Next I1
                    End If
                    'ADD BY YGY 20140929    ↓↓↓↓↓↓
                    'ミックスマニホールドを指定した場合は組合せ出力不可
                    If Not HTValue("1") Is Nothing Then
                        For I1 = 1 To CType(HTValue("1"), ArrayList).Count - 1
                            If Trim(CType(HTValue("1"), ArrayList).Item(I1)) = "8" Then
                                strMsgCd = "DEER432"
                                Now_Pos = 1
                                Exit Function
                            End If
                        Next
                    End If
                    If Not HTValue("3") Is Nothing Then
                        For I1 = 1 To CType(HTValue("3"), ArrayList).Count - 1
                            If Trim(CType(HTValue("3"), ArrayList).Item(I1)) = "HX" Then
                                strMsgCd = "DEER432"
                                Now_Pos = 3
                                Exit Function
                            End If
                        Next
                    End If
                    'ADD BY YGY 20140929    ↑↑↑↑↑↑
                Case "17"    'GAMD0
                    'ADD BY YGY 20140929    ↓↓↓↓↓↓
                    'ミックスマニホールドを指定した場合は組合せ出力不可
                    If Not HTValue("1") Is Nothing Then
                        For I1 = 1 To CType(HTValue("1"), ArrayList).Count - 1
                            If Trim(CType(HTValue("1"), ArrayList).Item(I1)) = "X" Then
                                strMsgCd = "DEER432"
                                Now_Pos = 1
                                Exit Function
                            End If
                        Next
                    End If
                    'ADD BY YGY 20140929    ↑↑↑↑↑↑
                Case "00"    '旧の「Z」
                    ' ロッド先端特注を指定した場合は組合せ出力不可
                    Select Case Trim(objKtbnStrc.strcSelection.strSeriesKataban)
                        Case "SSD"
                            Select Case Trim(objKtbnStrc.strcSelection.strKeyKataban)
                                Case String.Empty
                                    If Not HTValue("22") Is Nothing Then
                                        str = HTValue("22").split(",")
                                        For I1 = 1 To str.Length - 1
                                            If Len(Trim(str(I1))) <> 0 Then
                                                strMsgCd = "DEER433"
                                                Now_Pos = 22
                                                Exit Function
                                            End If
                                        Next
                                    End If
                                Case "K"
                                    If Not HTValue("20") Is Nothing Then
                                        str = HTValue("20").split(",")
                                        For I1 = 1 To str.Length - 1
                                            If Len(Trim(str(I1))) <> 0 Then
                                                strMsgCd = "DEER433"
                                                Now_Pos = 20
                                                Exit Function
                                            End If
                                        Next
                                    End If
                            End Select
                        Case "CMK2"
                            If Not HTValue("17") Is Nothing Then
                                str = HTValue("17").split(",")
                                For I1 = 1 To str.Length - 1
                                    If Len(Trim(str(I1))) <> 0 Then
                                        strMsgCd = "DEER433"
                                        Now_Pos = 17
                                        Exit Function
                                    End If
                                Next
                            End If
                        Case "SCM"
                            Select Case Trim(objKtbnStrc.strcSelection.strKeyKataban)
                                Case String.Empty
                                    If Not HTValue("15") Is Nothing Then
                                        str = HTValue("15").split(",")
                                        For I1 = 1 To str.Length - 1
                                            If Len(Trim(str(I1))) <> 0 Then
                                                strMsgCd = "DEER433"
                                                Now_Pos = 15
                                                Exit Function
                                            End If
                                        Next
                                    End If
                                Case "B"
                                    If Not HTValue("19") Is Nothing Then
                                        str = HTValue("19").split(",")
                                        For I1 = 1 To str.Length - 1
                                            If Len(Trim(str(I1))) <> 0 Then
                                                strMsgCd = "DEER433"
                                                Now_Pos = 19
                                                Exit Function
                                            End If
                                        Next
                                    End If
                            End Select
                        Case "SCA2"
                            Select Case Trim(objKtbnStrc.strcSelection.strKeyKataban)
                                Case "", "V"
                                    If Not HTValue("15") Is Nothing Then
                                        str = HTValue("15").split(",")
                                        For I1 = 1 To str.Length - 1
                                            If Len(Trim(str(I1))) <> 0 Then
                                                strMsgCd = "DEER433"
                                                Now_Pos = 15
                                                Exit Function
                                            End If
                                        Next
                                    End If
                                Case "B"
                                    If Not HTValue("19") Is Nothing Then
                                        str = HTValue("19").split(",")
                                        For I1 = 1 To str.Length - 1
                                            If Len(Trim(str(I1))) <> 0 Then
                                                strMsgCd = "DEER433"
                                                Now_Pos = 19
                                                Exit Function
                                            End If
                                        Next
                                    End If
                            End Select
                    End Select
                Case "52", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "70", "71", "72", "89", "90"
                    ' ミックスマニホールドを指定した場合は組合せ出力不可
                    If Not HTValue("1") Is Nothing Then
                        str = HTValue("1").split(",")
                        For I1 = 1 To str.Length - 1
                            If Trim(str(I1)) = "8" Then
                                strMsgCd = "DEER432"
                                Now_Pos = 1
                                Exit Function
                            End If
                        Next
                    End If
                Case "73", "74", "75", "76", "77", "78", "79", "80", "81", "82", "83", "84", "85", "86", "87", "88"
                    ' ミックスマニホールドを指定した場合は組合せ出力不可      M4K**
                    If Not HTValue("1") Is Nothing Then
                        str = HTValue("1").split(",")
                        For I1 = 1 To str.Length - 1
                            If Trim(str(I1)) = "80" Or Trim(str(I1)) = "81" Then
                                strMsgCd = "DEER432"
                                Now_Pos = 1
                                Exit Function
                            End If
                        Next
                    End If
                Case "51"   'B
                    ' ミックスマニホールドを指定した場合は組合せ出力不可
                    If Not HTValue("3") Is Nothing Then
                        str = HTValue("3").split(",")
                        For I1 = 1 To str.Length - 1
                            If Trim(str(I1)) = "8" Then
                                strMsgCd = "DEER432"
                                Now_Pos = 3
                                Exit Function
                            End If
                        Next
                    End If
                    'Case "12", "18", "19", "20", "21", "22", "23", "54", "55", "56", "57", "58", "59", _
                    '    "91", "92", "94", "95", "A1", "A2"
                    '    If objKtbnStrc.strcSelection.strSeriesKataban.StartsWith("VS") Then
                    '        strMsgCd = "DEER432"
                    '        Exit Function
                    '    Else
                    '        ' ミックスマニホールドを指定した場合は組合せ出力不可
                    '        If Not HTValue("1") Is Nothing Then
                    '            str = HTValue("1").split(",")
                    '            For I1 = 1 To str.Length - 1
                    '                If Trim(str(I1)) = "8" Then
                    '                    strMsgCd = "DEER432"
                    '                    Now_Pos = 1
                    '                    Exit Function
                    '                End If
                    '            Next
                    '        End If
                    '    End If
                Case "12", "18", "19", "20", "21", "22", "23", "94", "95"
                    'CHANGED BY YGY 20140929    ↓↓↓↓↓↓
                    '「VS」で始まる機種の場合、ミックスマニホールドを指定したら出力不可
                    If SiyouCheck_StartWithVS(objKtbnStrc.strcSelection.strSeriesKataban, HTValue) = True Then
                        strMsgCd = "DEER432"
                        Now_Pos = 3
                        Exit Function
                    End If
                Case "54", "55", "56", "57", "58", "59", "91", "92", "A1", "A2", "B2", "B3", "B4"
                    'ミックスマニホールドを指定した場合は組合せ出力不可
                    If Not HTValue("1") Is Nothing Then
                        For I1 = 1 To CType(HTValue("1"), ArrayList).Count - 1
                            If Trim(CType(HTValue("1"), ArrayList).Item(I1)) = "8" Then
                                strMsgCd = "DEER432"
                                Now_Pos = 1
                                Exit Function
                            End If
                        Next
                    End If
                    'CHANGED BY YGY 20140929    ↑↑↑↑↑↑
                Case Else
                    Exit Select
            End Select

            ExecuteCheck = True
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' VSで始まる機種が仕様書情報の要否チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SiyouCheck_StartWithVS(ByVal strSeries As String, ByVal HTValue As Hashtable) As Boolean
        Dim blnResult As Boolean = False
        Select Case strSeries
            Case "VSKM"
                If CType(HTValue("1"), ArrayList).Contains("Z") Or _
                   CType(HTValue("2"), ArrayList).Contains("00") Or _
                   CType(HTValue("3"), ArrayList).Contains("Z") Or _
                   CType(HTValue("4"), ArrayList).Contains("CX") Or _
                   CType(HTValue("8"), ArrayList).Contains("Z") Or _
                   CType(HTValue("10"), ArrayList).Contains("Z") Then
                    blnResult = True
                End If
            Case "VSJM"
                If CType(HTValue("1"), ArrayList).Contains("Z") Or _
                   CType(HTValue("2"), ArrayList).Contains("00") Or _
                   CType(HTValue("3"), ArrayList).Contains("Z") Or _
                   CType(HTValue("4"), ArrayList).Contains("CX") Or _
                   CType(HTValue("10"), ArrayList).Contains("Z") Then
                    blnResult = True
                End If
            Case "VSNM"
                If CType(HTValue("1"), ArrayList).Contains("Z") Or _
                   CType(HTValue("2"), ArrayList).Contains("00") Or _
                   CType(HTValue("3"), ArrayList).Contains("CX") Or _
                   CType(HTValue("9"), ArrayList).Contains("Z") Then
                    blnResult = True
                End If
            Case "VSNPM"
                If CType(HTValue("1"), ArrayList).Contains("CX") Or _
                   CType(HTValue("6"), ArrayList).Contains("Z") Then
                    blnResult = True
                End If
            Case "VSXM"
                If CType(HTValue("1"), ArrayList).Contains("Z") Or _
                   CType(HTValue("2"), ArrayList).Contains("00") Or _
                   CType(HTValue("3"), ArrayList).Contains("Z") Or _
                   CType(HTValue("4"), ArrayList).Contains("CX") Or _
                   CType(HTValue("9"), ArrayList).Contains("Z") Then
                    blnResult = True
                End If
            Case "VSZM"
                If CType(HTValue("1"), ArrayList).Contains("Z") Or _
                   CType(HTValue("2"), ArrayList).Contains("00") Or _
                   CType(HTValue("3"), ArrayList).Contains("Z") Or _
                   CType(HTValue("4"), ArrayList).Contains("CX") Or _
                   CType(HTValue("9"), ArrayList).Contains("Z") Then
                    blnResult = True
                End If
            Case "VSJPM"
                If CType(HTValue("1"), ArrayList).Contains("Z") Or _
                   CType(HTValue("2"), ArrayList).Contains("CX") Or _
                   CType(HTValue("9"), ArrayList).Contains("Z") Then
                    blnResult = True
                End If
            Case "VSXPM"
                If CType(HTValue("1"), ArrayList).Contains("Z") Or _
                   CType(HTValue("2"), ArrayList).Contains("CX") Or _
                   CType(HTValue("7"), ArrayList).Contains("Z") Then
                    blnResult = True
                End If
            Case "VSZPM"
                If CType(HTValue("1"), ArrayList).Contains("CX") Or _
                   CType(HTValue("6"), ArrayList).Contains("Z") Then
                    blnResult = True
                End If
            Case Else
                blnResult = False
        End Select

        Return blnResult
    End Function

    ''' <summary>
    ''' 選択したデータより形番を生成する
    ''' </summary>
    ''' <param name="PosIdx"></param>
    ''' <remarks></remarks>
    Public Shared Sub Kataban_Deployment(objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, ByVal PosIdx As Integer, _
                                         ByVal HTValue As Hashtable, ByRef ItemCode() As String, _
                                         ByRef HTResult As ArrayList, ByRef HTItem As ArrayList, ByVal DS_Tab As DataSet, _
                                         ByRef HTCheck As Hashtable, ByRef HTCheckOp As Hashtable, _
                                         ByRef HTOut As ArrayList, ByRef strPath As String, ByRef intAll As Long)
        Dim I1 As Integer
        Dim I2 As Integer
        If HTCheck Is Nothing Then HTCheck = New Hashtable
        If HTCheckOp Is Nothing Then HTCheckOp = New Hashtable

        '300万まで出力する(VB6と同じ)、8時間ぐらい
        If intAll + HTResult.Count >= 3000000 Then Exit Sub

        Try
            '次のオプションを選択できるのデータリストを決める
            Dim strOptionComma As String = String.Empty
            Dim dt_View As New DS_KatOut.DT_OptionDataTable
            Dim dr_View() As DataRow = DS_Tab.Tables("Option").Select("ktbn_strc_seq_no='" & (PosIdx).ToString & "'")
            For inti As Integer = 0 To dr_View.Length - 1
                dt_View.ImportRow(dr_View(inti))
            Next

            'オプションリスト取得
            Dim obj As New KHOptionCtl
            Dim strListOption(,) As String = Nothing
            If ItemCode Is Nothing Then ReDim ItemCode(HTValue.Count)
            objKtbnStrc.strcSelection.strOpSymbol = ItemCode
            Call obj.subOptionList(Nothing, objKtbnStrc, "1", "", "", "ja", PosIdx, strListOption, DS_Tab.Tables("Option"), DS_Tab.Tables("ElePattern"))

            Dim nextOpt As New ArrayList
            If strListOption Is Nothing OrElse UBound(strListOption) = 0 Then
                nextOpt.Add("無記号")
            Else
                For inti As Integer = 0 To UBound(strListOption) - 1
                    Dim str As String = strListOption(inti + 1, 1).ToString.Trim
                    If str = "その他電圧" Or str.ToUpper = "OTHER VOLTAGE" Then
                        Continue For
                    End If
                    nextOpt.Add(IIf(str.Length <= 0, "無記号", str))
                Next
            End If
            If HTValue(PosIdx.ToString) Is Nothing Then
                Exit Sub
            Else
                For inti As Integer = nextOpt.Count - 1 To 0 Step -1
                    Dim flg As Boolean = False
                    If Not HTValue(PosIdx.ToString) Is Nothing Then
                        Dim str() As String = HTValue(PosIdx.ToString).split(",")
                        For intj As Integer = 0 To str.Length - 1
                            If nextOpt(inti) = str(intj) Then
                                flg = True
                                Exit For
                            End If
                        Next
                    End If
                    If Not flg Then
                        nextOpt.RemoveAt(inti)
                    End If
                Next
            End If
            Dim lstOpt As ArrayList = nextOpt
            'If ItemCode Is Nothing Then ReDim ItemCode(HTValue.Count)
            If lstOpt.Count = 0 Then lstOpt.Add(String.Empty)
            If HTCheckOp(PosIdx.ToString) Is Nothing Then
                HTCheckOp.Add(PosIdx.ToString, lstOpt)
            Else
                HTCheckOp(PosIdx.ToString) = lstOpt
            End If
            For I1 = 0 To lstOpt.Count - 1
                ' 選択オプション保存
                If lstOpt.Item(I1) = "無記号" Then
                    ItemCode(PosIdx) = String.Empty
                Else
                    ItemCode(PosIdx) = lstOpt.Item(I1)
                End If
                ' 電圧・口径・ストロークを設定する
                If OptionDefaultSetting(objKtbnStrc, PosIdx, ItemCode) = True Then
                    ' 最終オプションの場合は形番を確認し生成する
                    If HTValue.Count = PosIdx Then
                        ' 最終的な形番をチェックする
                        objKtbnStrc.strcSelection.strOpSymbol = ItemCode
                        If Kataban_Create_Check(objKtbnStrc, ItemCode, HTCheckOp) = True Then
                            ' 形番を生成する
                            Dim FulPartsNo As String = Kataban_Create(objKtbnStrc, HTValue, ItemCode)
                            ' 生成した形番を保存する
                            ' 既に同一の形番が存在する場合は破棄する
                            If HTCheck.ContainsKey(FulPartsNo) Then
                                Exit Sub
                            Else
                                HTCheck.Add(FulPartsNo, "")
                                HTResult.Add(FulPartsNo)
                                ItemCode(0) = objKtbnStrc.strcSelection.strSeriesKataban
                                Dim myItem(ItemCode.Length - 1) As String
                                For inti As Integer = 0 To ItemCode.Length - 1
                                    myItem(inti) = ItemCode(inti)
                                Next
                                HTItem.Add(myItem)

                                '300万まで出力する(VB6と同じ)、8時間ぐらい
                                If intAll + HTResult.Count >= 3000000 Then
                                    Exit Sub
                                End If

                                '3000件ずつ出力する
                                'If HTResult.Count >= 3000 Then
                                If HTResult.Count >= 100 Then
                                    '単価取得
                                    Dim objUnitPrice As New KHUnitPrice
                                    Dim strOutFile As String = String.Empty
                                    For inti As Integer = 0 To HTResult.Count - 1
                                        objKtbnStrc.strcSelection.strFullKataban = HTResult(inti)
                                        objKtbnStrc.strcSelection.strOpSymbol = HTItem(inti)
                                        objKtbnStrc.strcSelection.intListPrice = 0
                                        objKtbnStrc.strcSelection.intRegPrice = 0
                                        objKtbnStrc.strcSelection.intSsPrice = 0
                                        objKtbnStrc.strcSelection.intBsPrice = 0
                                        objKtbnStrc.strcSelection.intGsPrice = 0
                                        objKtbnStrc.strcSelection.intPsPrice = 0
                                        objKtbnStrc.strcSelection.strKatabanCheckDiv = ""
                                        objKtbnStrc.strcSelection.strPlaceCd = ""

                                        Call objUnitPrice.subPriceInfoSet_ForkatOut(objCon, objKtbnStrc, "JPN", "", DS_Tab)

                                        strOutFile &= objKtbnStrc.strcSelection.strFullKataban & ControlChars.Tab

                                        'チェック区分
                                        If HTOut.Contains("kataban_check_div") Then
                                            strOutFile &= objKtbnStrc.strcSelection.strKatabanCheckDiv & ControlChars.Tab
                                        End If
                                        If HTOut.Contains("place_cd") Then
                                            strOutFile &= objKtbnStrc.strcSelection.strPlaceCd & ControlChars.Tab
                                        End If
                                        If HTOut.Contains("disp_name") Then
                                            strOutFile &= objKtbnStrc.strcSelection.strGoodsNm & ControlChars.Tab
                                        End If
                                        If HTOut.Contains("ls_price") Then
                                            strOutFile &= CInt(objKtbnStrc.strcSelection.intListPrice) & ControlChars.Tab
                                        End If
                                        If HTOut.Contains("rg_price") Then
                                            strOutFile &= CInt(objKtbnStrc.strcSelection.intRegPrice) & ControlChars.Tab
                                        End If
                                        If HTOut.Contains("ss_price") Then
                                            strOutFile &= CInt(objKtbnStrc.strcSelection.intSsPrice) & ControlChars.Tab
                                        End If
                                        If HTOut.Contains("bs_price") Then
                                            strOutFile &= CInt(objKtbnStrc.strcSelection.intBsPrice) & ControlChars.Tab
                                        End If
                                        If HTOut.Contains("gs_price") Then
                                            strOutFile &= CInt(objKtbnStrc.strcSelection.intGsPrice) & ControlChars.Tab
                                        End If
                                        If HTOut.Contains("ps_price") Then
                                            strOutFile &= CInt(objKtbnStrc.strcSelection.intPsPrice)
                                        End If

                                        strOutFile &= ControlChars.NewLine
                                    Next
                                    If strOutFile.Length > 0 Then System.IO.File.AppendAllText(strPath, strOutFile, System.Text.Encoding.UTF8)
                                    strOutFile = String.Empty
                                    intAll += HTResult.Count

                                    HTResult.Clear()
                                    HTItem.Clear()
                                    HTCheck.Clear()
                                    HTCheckOp.Clear()
                                End If

                                '300万まで出力する(VB6と同じ)、8時間ぐらい
                                If intAll + HTResult.Count >= 3000000 Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    Else
                        ' 次のオプション選択へ
                        Call Kataban_Deployment(objCon, objKtbnStrc, PosIdx + 1, HTValue, ItemCode, HTResult, HTItem, _
                                                DS_Tab, HTCheck, HTCheckOp, HTOut, strPath, intAll)
                    End If
                End If
                ' 1つも選択可能なオプションが存在しない項目は2回目以降はスキップする
                If I2 > lstOpt.Count - 1 Then
                    Exit For
                End If
            Next I1
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 電圧・口径・ストロークを設定する
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="PosIdx"></param>
    ''' <param name="ItemCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function OptionDefaultSetting(objKtbnStrc As KHKtbnStrc, ByVal PosIdx As Integer, _
                                         ByVal ItemCode() As String)
        Dim DenCoil As String = String.Empty
        Dim DenCaliber As String = String.Empty
        Dim Voltage As String = String.Empty
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
        Dim strPriceNo As String = objKtbnStrc.strcSelection.strPriceNo

        OptionDefaultSetting = False
        Try
            Select Case strPriceNo
                Case "02", "03"
                    Select Case PosIdx
                        Case 4
                            If strSeriesKata.StartsWith("AB") Or strSeriesKata.StartsWith("AG") Then
                                ' コイル情報
                                DenCoil = Trim(ItemCode(PosIdx))
                            End If
                        Case 5
                            If strSeriesKata.StartsWith("GAG") Or strSeriesKata.StartsWith("GAB") Then
                                ' コイル情報
                                DenCoil = Trim(ItemCode(PosIdx))
                            End If
                    End Select
                Case "17"
                    Select Case PosIdx
                        Case 1
                            If strSeriesKata.StartsWith("AD11") Or strSeriesKata.StartsWith("AP11") Then
                                ' 接続口径情報セーブ
                                DenCaliber = Trim(ItemCode(PosIdx))
                            End If
                        Case 3
                            ' コイル情報
                            DenCoil = Trim(ItemCode(PosIdx))
                    End Select
                Case "18", "36", "M3"
                    Select Case PosIdx
                        Case 3
                            ' コイル情報
                            DenCoil = Trim(ItemCode(PosIdx))
                    End Select
                Case "35"
                    Select Case PosIdx
                        Case 4
                            'コイル情報
                            DenCoil = Trim(ItemCode(PosIdx))
                    End Select
                Case "37"
                    Select Case PosIdx
                        Case 1
                            ' 接続口径情報セーブ
                            DenCaliber = Trim(ItemCode(PosIdx))
                    End Select
                Case "38"
                    Select Case PosIdx
                        Case 5
                            ' コイル情報
                            DenCoil = Trim(ItemCode(PosIdx))
                    End Select
            End Select

            ''ガイド区分
            'Select Case objKtbnStrc.strcSelection.strOpElementDiv(PosIdx)
            '    Case "1"
            '        Select Case Trim(ItemCode(PosIdx))
            '            Case "1"
            '                Voltage = "AC100V"
            '            Case "2"
            '                Voltage = "AC200V"
            '            Case "3"
            '                Voltage = "DC24V"
            '            Case "4"
            '                Voltage = "DC12V"
            '            Case "5"
            '                Voltage = "AC110V"
            '            Case "6"
            '                Voltage = "AC220V"
            '            Case Else
            '                Voltage = Trim(ItemCode(PosIdx))
            '        End Select

            '        Dim DenDivision As String = String.Empty
            '        Dim DenMin As Integer = 0
            '        Dim DenMax As Integer = 0

            '        If strPriceNo = "02" Or strPriceNo = "03" Then
            '            Select Case Left(strSeriesKata, 1)
            '                Case "A"
            '                    If Left(strSeriesKata, 3) = "AB4" Then
            '                        If (Trim(DenCoil) = "3A" Or Trim(DenCoil) = "3K") And Left(Voltage, 2) = "DC" Or _
            '                           (Trim(DenCoil) = "5A" Or Trim(DenCoil) = "5K") And Left(Voltage, 2) = "AC" Then
            '                            DenKey = Left(strSeriesKata, 4)
            '                        Else
            '                            DenKey = Left(strSeriesKata, 3)
            '                        End If
            '                    Else
            '                        DenKey = Left(strSeriesKata, 3)
            '                    End If
            '                Case "G"
            '                    If Mid(strSeriesKata, 2, 3) = "AB4" Then
            '                        If (Trim(DenCoil) = "3A" Or Trim(DenCoil) = "3K") And Left(Voltage, 2) = "DC" Or _
            '                           (Trim(DenCoil) = "5A" Or Trim(DenCoil) = "5K") And Left(Voltage, 2) = "AC" Then
            '                            DenKey = Mid(strSeriesKata, 2, 4)
            '                        Else
            '                            DenKey = Mid(strSeriesKata, 2, 3)
            '                        End If
            '                    Else
            '                        DenKey = Mid(strSeriesKata, 2, 3)
            '                    End If
            '            End Select
            '        Else
            '            DenKey = Trim(strSeriesKata)
            '        End If

            '        'If DBDenRead(DenKey, Trim(DenCaliber), Trim(DenCoil), Left(Voltage, 2)) = False Then
            '        '    Exit Function
            '        'End If

            '        'DenDivision = rsden.Fields("区分")
            '        'DenMin = rsden.Fields("下限")
            '        'DenMax = rsden.Fields("上限")

            '        'For I1 = 1 To 50
            '        '    S = "在庫区分" & I1
            '        '    If IsNull(rsden(S)) = True Then
            '        '        DenSDivision(I1) = Space(0)
            '        '    Else
            '        '        DenSDivision(I1) = rsden(S)
            '        '    End If

            '        '    S = "標準" & I1
            '        '    If IsNull(rsden(S)) = True Then
            '        '        DenDDivision(I1) = Space(0)
            '        '    Else
            '        '        DenDDivision(I1) = rsden(S)
            '        '    End If

            '        '    S = "電圧" & I1
            '        '    DenVolt(I1) = rsden(S)

            '        '    If rsden(S) = 0 Then
            '        '        MaxVoltage = I1 - 1

            '        '        Exit For
            '        '    End If
            '        'Next I1
            '    Case "3"
            '        Dim Stroke As Integer = Val(ItemCode(PosIdx))

            '        Dim StrDivision As String = ""             ' データ区分
            '        Dim StrMin As Integer = 0                  ' 最小値
            '        Dim StrMax As Integer = 0                  ' 最大値
            '        Dim StrKeyMax As Integer = 0               ' 単価キーＭＡＸ
            '        Dim StrKeyFore As Integer = 0              ' 直前の単価キー

            '        ' ストロークファイルRead
            '        'If DBStrRead(Trim(strSeriesKata), 0) = False Then
            '        '    If BoreSz = 0 Then
            '        '        Call WriteErrorLog("DEER11")
            '        '        Exit Function
            '        '    End If

            '        '    ' ストロークファイルRead
            '        '    If DBStrRead(Trim(strSeriesKata), CDbl(BoreSz)) = False Then
            '        '        Call WriteErrorLog("RSER10")
            '        '        Exit Function
            '        '    End If
            '        'End If

            '        'StrDivision = rsstr.Fields("データ区分")   ' データ区分
            '        'StrMin = rsstr.Fields("最小値")            ' 最小値
            '        'StrMax = rsstr.Fields("最大値")            ' 最大値
            '        'StrKeyMax = rsstr.Fields("単価キーＭＡＸ")  ' 単価キーＭＡＸ
            '        'StrKeyFore = rsstr.Fields("単価キー前")     ' 直前の単価キー

            '        'For I1 = 1 To 50
            '        '    S = "ストローク" & I1
            '        '    StrStroke(I1) = rsstr.Fields(S)
            '        '    If rsstr(S) = 0 Then
            '        '        MaxStroke = I1 - 1

            '        '        Exit For
            '        '    End If
            '        'Next I1
            '    Case "5"
            '        'BoreSz = Val(ItemCode(PosIdx))
            'End Select

            OptionDefaultSetting = True
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 形番補正用共通関数
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="HTValue"></param>
    ''' <param name="ItemCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Kataban_Create(objKtbnStrc As KHKtbnStrc, ByVal HTValue As Hashtable, _
                                          ByVal ItemCode() As String) As String
        Dim I1 As Integer
        Dim I2 As Integer
        Dim I3 As Integer

        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
        Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban
        Dim strPriceNo As String = objKtbnStrc.strcSelection.strPriceNo
        Kataban_Create = String.Empty
        Try
            ' 形番を補正する
            For I1 = 1 To ItemCode.Length - 1
                If objKtbnStrc.strcSelection.strOpAdditionDiv(I1).ToString >= "2" Then
                    If Left(Trim(strSeriesKata), 4) = "AB21" Then
                        If ItemCode(4) = "00B" Then
                            If Len(Trim(ItemCode(3))) = 0 Then
                                ItemCode(3) = "0"
                            End If
                        End If
                    Else
                        If Len(Trim(ItemCode(I1))) <> 0 Then
                            For I2 = ItemCode.Length - 1 To 1 Step -1
                                If objKtbnStrc.strcSelection.strOpAdditionDiv(I2) >= "1" And _
                                    objKtbnStrc.strcSelection.strOpAdditionDiv(I2) < objKtbnStrc.strcSelection.strOpAdditionDiv(I1) Then
                                    If Len(Trim(ItemCode(I2))) = 0 Then
                                        ItemCode(I2) = ""
                                        Dim MaxLen As Integer = 0
                                        Dim str_val() As String = HTValue(I2.ToString).ToString.Split(",")
                                        For I3 = 1 To str_val.Length - 1
                                            If str_val(I3).ToString.Length > MaxLen Then
                                                MaxLen = Len(str_val(I3))
                                            End If
                                        Next I3
                                        Dim str As String = String.Empty
                                        ItemCode(I2) = str.PadLeft(MaxLen, "0")
                                    Else
                                        Exit For
                                    End If
                                End If
                            Next I2
                        End If
                    End If
                End If
            Next I1

            If strSeriesKata.StartsWith("NAB") Then
                If ItemCode(4).Trim = "" And ItemCode(5) = "B" Then
                    ItemCode(4) = "0"
                End If
            End If

            ' 形番を生成する
            Kataban_Create = Trim(objKtbnStrc.strcSelection.strSeriesKataban & _
                                  IIf(objKtbnStrc.strcSelection.strHyphen = CdCst.HyphenDiv.Necessary, CdCst.Sign.Hypen, String.Empty))

            For I1 = 1 To ItemCode.Length - 1
                Kataban_Create = Trim(Kataban_Create) & Trim(ItemCode(I1)) & IIf(objKtbnStrc.strcSelection.strOpHyphenDiv(I1) = CdCst.HyphenDiv.Necessary, CdCst.Sign.Hypen, String.Empty)
                ' 重複するハイフンを消去する
                '選択したオプションを結合
                'AMD3,4,5 R,X対応 ハイフン削除して結合
                If strKeyKata Is Nothing Then strKeyKata = ""
                If (strSeriesKata.StartsWith("AMD3") And (strKeyKata = "1" Or strKeyKata = "2")) Or _
                    (strSeriesKata.StartsWith("AMD4") And Len(strKeyKata) = 0) Or _
                    (strSeriesKata.StartsWith("AMD5") And Len(strKeyKata) = 0) Or _
                    (strSeriesKata.StartsWith("AMD0") And Len(strKeyKata) = 1) Then            'RM1310067 2013/10/23 追加
                    If ItemCode.Length > 8 AndAlso Len(ItemCode(8).ToString) <> 0 Then
                        If I1 = 5 Then
                            If Kataban_Create.EndsWith("-") Then
                                Kataban_Create = Strings.Left(Kataban_Create, Kataban_Create.Length - 1)
                            End If
                        End If
                    End If
                End If
                Kataban_Create = Kataban_Create.Trim.Replace("--", "-")
            Next I1

            ' 最後がハイフンだったら消去
            Kataban_Create = Kataban_Create.Replace("--", "-")
            If Kataban_Create.EndsWith("-") Then
                Kataban_Create = Left(Kataban_Create, Kataban_Create.Length - 1)
            End If
        Catch ex As Exception
            Kataban_Create = String.Empty
            Call WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 形番チェック
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="ItemCode"></param>
    ''' <param name="HTCheckOp"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Kataban_Create_Check(objKtbnStrc As KHKtbnStrc, ByVal ItemCode() As String, _
                                                Optional ByVal HTCheckOp As Hashtable = Nothing) As Boolean

        Kataban_Create_Check = False

        'OKボタン時のオプションチェック
        Dim strMsgCd As String = String.Empty
        Dim strRodEndOption As String = String.Empty     'ロッド先端特注
        Dim strOptionOther As String = String.Empty      'オプション外指定
        Dim intSeqNo As Integer = -1
        '形番ﾁｪｯｸ
        If Not HTCheckOp Is Nothing Then
            For inti As Integer = 1 To ItemCode.Length - 1
                If Not HTCheckOp(inti.ToString) Is Nothing Then
                    Dim strlst As ArrayList = HTCheckOp(inti.ToString)
                    Dim flg As Boolean = False
                    For intj As Integer = 0 To strlst.Count - 1
                        If strlst(intj) = "無記号" Then
                            flg = True
                            Exit For
                        End If
                        If ItemCode(inti) = strlst(intj) Then
                            flg = True
                            Exit For
                        End If
                    Next
                    If Not flg Then
                        Exit Function
                    End If
                End If
            Next
        End If

        '組合せテスト
        Dim obj As New KHOptionCtl
        If Not obj.fncOtherOptionCheck(objKtbnStrc, intSeqNo, "", strMsgCd) Then Exit Function

        Kataban_Create_Check = True

    End Function


End Class
