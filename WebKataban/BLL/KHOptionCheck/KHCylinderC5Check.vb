Module KHCylinderC5Check

    '********************************************************************************************
    '*【関数名】
    '*  fncCylinderC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ   True    :全てチェック
    '*                                                   False   :チェック区分変更のみは除外   
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・シリーズSTG-K/PCU2/AHB/SSG/JSK2/JSK2-Vを追加
    '*                                          更新日：2009/02/03      更新者：T.Yagyu
    '*  ・RM0811134:SRT3シリーズを追加
    '*  ・RM0811134:SRL3シリーズのC5条件追加
    '*  ・RM0811134:SRG3シリーズを追加
    '*  ・RM0811133:CAC4シリーズを追加 2009/07/27 Y.Miura
    '*  ・RM0811133:UCAC2シリーズを追加 2009/08/01 Y.Miura
    '*  ・RM0808030:MDC2シリーズを追加 2009/09/04 Y.Miura
    '*  ・RM0808030:MSD/MSDGシリーズを追加 2009/10/15 Y.Miura
    '*  ・RM0808030:STKシリーズを追加 2009/10/15 Y.Miura
    '*  ・RM0808030:STR2シリーズを追加 2009/10/15 Y.Miura
    '*  ・RM0808030:SRM3シリーズを追加 2009/10/15 Y.Miura
    '*  ・RM0906034:二次電池対応 2009/07～ Y.Miura
    '*  ・RM1001018:UCAC2シリーズ　スイッチのC5適用 2010/01/18 Y.Miura
    '*  ・RM1001043:二次電池P4*のC5適用をなくす 2010/02～ Y.Miura
    '********************************************************************************************
    Public Function fncCylinderC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                       Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncCylinderC5Check = False

            '機種毎にチェック
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                'RM0811133 2009/07/27 Y.Miura
                'Case "CAC3"
                Case "CAC3", "CAC4"
                    If fncCAC3C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                    'RM0811133 2009/08/01 Y.Miura
                Case "UCAC2", "UCAC2-L2"
                    If fncUCAC2C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "CMK2"
                    If fncCMK2C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "JSC3"
                    If fncJSC3C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                    'RM1302XXX 2013/02/04 Y.Tachi
                Case "JSC4"
                    If fncJSC4C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "JSG", "JSG-V"
                    If fncJSGC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SCA2"
                    If fncSCA2C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SCG", "SCG-D", "SCG-G", "SCG-G2", "SCG-G3", _
                     "SCG-G4", "SCG-M", "SCG-O", "SCG-Q", "SCG-U"
                    If fncSCGC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SCM"
                    If fncSCMC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SCS"
                    If fncSCSC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                    'RM1302XXX 2013/02/04 Y.Tachi
                Case "SCS2"
                    If fncSCS2C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SSD"
                    If fncSSDC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "STG-B", "STG-M", "STG-K"
                    If fncSTGC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "STL-B", "STL-M", "STS-B", "STS-M"
                    If fncSTSC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "LCG", "LCG-Q"
                    If fncLCGC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "LCR", "LCR-Q" 'RM10030086 2010/04/07 Y.Miura 追加
                    If fncLCRC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SSG"
                    If fncSSGC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "JSK2", "JSK2-V"
                    If fncJSK2C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SRL3"
                    If fncSRL3C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SSD2"
                    If fncSSD2C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SRT3" 'RM0811134:SRT3
                    If fncSRT3C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "SRG3" 'RM0811134:SRG3
                    If fncSRG3C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "MDV", "MDV-L", "LFC-KL", "PCU2", "AHB"
                    fncCylinderC5Check = True
                Case "MDC2-L", "MDC2-XL", "MDC2-YL", "MSD-L", "MSD-KL", "MSD-XL", "MSD-YL", "MSDG-L", _
                    "STM", "LCM", "LCM-A", "LCM-P", "LCM-R", "LCW"
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "F3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "F3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "LCT", "LCX", "LCX-Q"
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "LSH"
                    'RM1705004 2017/05/11 削除
                    'If objKtbnStrc.strcSelection.strKeyKataban <> "1" And _
                    '    objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "C" Then
                    '    fncCylinderC5Check = True
                    'End If
                    If objKtbnStrc.strcSelection.strKeyKataban <> "1" And _
                        (objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "F3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "F3PV") Then
                        fncCylinderC5Check = True
                    End If
                Case "CKV2", "CKV2-M"
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "T3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "RCC2", "RCC2-G4", "RCS"
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "RCS2"   'RM1803075_RCS2 追加
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "T3PV" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "F3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "F3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "CAV2", "COVP2", "COVN2"
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T3PV" Then
                        fncCylinderC5Check = True
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "DC24V" Then
                        fncCylinderC5Check = True
                    End If
                Case "MVC"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "BBS-A", "BBS-O", "BBS-OB"
                    fncCylinderC5Check = True
                Case "SMG"
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "NN" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "GN" Then
                        fncCylinderC5Check = True
                    End If

                    If objKtbnStrc.strcSelection.strKeyKataban = "2" Then
                        fncCylinderC5Check = True
                    End If

                    ''微速Fの場合はＣ５
                    'If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F" Then
                    '    fncCylinderC5Check = True
                    'End If

                    ''クリーン仕様P5,P51,P7,P71の場合はＣ５
                    'If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P5" Or _
                    '    objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P51" Or _
                    '    objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P7" Or _
                    '    objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P71" Then
                    '    fncCylinderC5Check = True
                    'End If

                    'RM1001043 2010/02/12 Y.Miura 二次電池C5加算廃止
                    'RM0906034 2009/08/20 Y.Miura　二次電池対応
                    'Case "LCS", "LCS-F", "LCS-Q"
                    'If fncLiIonC5Check(objKtbnStrc, 10, bolJudgeDiv) = True Then
                    '    fncCylinderC5Check = True
                    'End If
                    'RM1001043 2010/02/12 Y.Miura 二次電池C5加算廃止
                    '    'RM0908030 2009/09/04 Y.Miura　二次電池対応
                    'Case "MDC2", "MDC2-L"
                    '    If fncLiIonC5Check(objKtbnStrc, 7, bolJudgeDiv) = True Then
                    '        fncCylinderC5Check = True
                    '    End If
                    'RM1001043 2010/02/12 Y.Miura 二次電池C5加算廃止
                    'RM0908030 2009/09/04 Y.Miura　二次電池対応
                    'Case "SMD2", "SMD2-L"
                    '    If fncLiIonC5Check(objKtbnStrc, 8, bolJudgeDiv) = True Then
                    '        fncCylinderC5Check = True
                    '    End If
                    'RM1001043 2010/02/12 Y.Miura 二次電池C5加算廃止
                    'RM0908030 2009/10/15 Y.Miura　二次電池対応
                    'Case "MSD", "MSD-L", "MSD-K", "MSD-KL", "MSDG-L"
                    '    '二次電池機種のみ 
                    '    If objKtbnStrc.strcSelection.strOpSymbol.Length >= "9" Then
                    '        If fncLiIonC5Check(objKtbnStrc, 8, bolJudgeDiv) = True Then
                    '            fncCylinderC5Check = True
                    '        End If
                    '    End If
                    'RM1001043 2010/02/12 Y.Miura 二次電池C5加算廃止
                    '    'RM0908030 2009/10/15 Y.Miura　二次電池対応
                    'Case "STK"
                    '    If fncLiIonC5Check(objKtbnStrc, 7, bolJudgeDiv) = True Then
                    '        fncCylinderC5Check = True
                    '    End If
                    'RM1001043 2010/02/12 Y.Miura 二次電池C5加算廃止
                    'RM0908030 2009/10/15 Y.Miura　二次電池対応
                    'Case "STR2-B", "STR2-M"
                    '    If fncLiIonC5Check(objKtbnStrc, 9, bolJudgeDiv) = True Then
                    '        fncCylinderC5Check = True
                    '    End If
                    'RM1001043 2010/02/12 Y.Miura 二次電池C5加算廃止
                    '    'RM0908030 2009/10/15 Y.Miura　二次電池対応
                    'Case "SRM3", "SRM3-Q"
                    '    If fncLiIonC5Check(objKtbnStrc, 8, bolJudgeDiv) = True Then
                    '        fncCylinderC5Check = True
                    '    End If
                    'RM0908030 2009/10/19 Y.Miura　二次電池対応
                Case "LHAG"
                    '二次電池機種のみ 
                    If objKtbnStrc.strcSelection.strOpSymbol.Length >= "7" Then
                        If fncLiIonC5Check(objKtbnStrc, 6, bolJudgeDiv) = True Then
                            fncCylinderC5Check = True
                        End If
                    End If
                    'RM1306001 2013/06/06
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "SX" Then
                        fncCylinderC5Check = True
                    End If
                    'RM0908030 2009/10/19 Y.Miura　二次電池対応
                Case "HMD"
                    If objKtbnStrc.strcSelection.strOpSymbol.Length >= "6" Then
                        If fncLiIonC5Check(objKtbnStrc, 5, bolJudgeDiv) = True Then
                            fncCylinderC5Check = True
                        End If
                    End If
                    'RM1001043 2010/02/12 Y.Miura 二次電池C5加算廃止
                    '    'RM0908030 2009/10/19 Y.Miura　二次電池対応
                    'Case "BHG", "HKP", "HLBG", "HLC", "HMF", "CKG", "BHE"
                    '    '二次電池機種のみ 
                    '    If objKtbnStrc.strcSelection.strOpSymbol.Length >= "8" Then
                    '        If fncLiIonC5Check(objKtbnStrc, 7, bolJudgeDiv) = True Then
                    '            fncCylinderC5Check = True
                    '        End If
                    '    End If
                Case "HMF", "CKG", "BHE", "BHA"
                    'P40二次電池機種のみ

                    If objKtbnStrc.strcSelection.strOpSymbol.Length >= "8" Then
                        If fncLiIonC5Check(objKtbnStrc, 7, bolJudgeDiv, True) = True Then
                            fncCylinderC5Check = True
                        End If
                    End If
                    'RM1005033 2010/07/02 Y.Miura ソース整理
                    '2011/03/14 MOD RM1103016(4月VerUP:LCS2シリーズ　追加) START--->

                Case "MCP-W", "MCP-S", "LCS2", "LCS2-Q"
                    'Case "MCP-W", "MCP-S"
                    '2011/03/14 MOD RM1103016(4月VerUP:LCS2シリーズ　追加) <---END
                    fncCylinderC5Check = True
                Case "FCK-L", "FCK-M", "FCK-H"
                    'RM1210067 2013/04/22
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "2" Then
                        fncCylinderC5Check = True
                    End If
                Case "FCD", "FCS", "FCH", "FCD-L", "FCS-L", "FCH-L", "FCD-D", "FCD-DL", "FCD-K", "FCD-KL"
                    'RM1412043 2014/12/11
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim.Trim = "N" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim.Trim = "G" Then
                        fncCylinderC5Check = True
                    End If
                Case "FK"
                    'fncCylinderC5Check = True
                    'RM1306005 2013/06/04 追加
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "SX" Then
                        fncCylinderC5Check = True
                    End If
                Case "HLD"
                    'fncCylinderC5Check = True                                   'Del by Zxjike 2013/10/01
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "2" Then   'Add by Zxjike 2013/10/01
                        fncCylinderC5Check = True
                    End If
                    'RM1210067 2013/02/07 Y.Tachi
                Case "CKL2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "2" Then
                        fncCylinderC5Check = True
                    End If
                Case "CKLB2"
                    'fncCylinderC5Check = True                                   'Del by Zxjike 2013/10/01
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "2" Then   'Add by Zxjike 2013/10/01
                        fncCylinderC5Check = True
                    End If
                    'Case "HCP", "HAP"
                Case "HCP"
                    'fncCylinderC5Check = True                                   'Del by Zxjike 2013/10/01
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "2" Then   'Add by Zxjike 2013/10/01
                        fncCylinderC5Check = True
                    End If
                Case "SFR", "SFRT"
                    'fncCylinderC5Check = True                                   'Del by Zxjike 2013/10/01
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "2" Then   'Add by Zxjike 2013/10/01
                        fncCylinderC5Check = True
                    End If
                Case "USSD", "USSD-L", "USSD-K", "USSD-KL"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "2" Then
                        fncCylinderC5Check = True
                    End If
                    'RM1210067 2013/04/04 ローカル版との差異修正
                Case "HAP"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "2" Then
                        fncCylinderC5Check = True
                    End If
                Case "MRL2", "MRL2-G", "MRL2-W"
                    'RM1306001 2013/06/04 追加 
                    If fncMRL2C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "MRL2-L", "MRL2-GL", "MRL2-WL"
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "MRG2"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "T3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "SRM3"
                    'RM1306001 2013/06/04 追加 
                    If fncSRM3C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "UCA2", "UCA2-B", "UCA2-BL", "UCA2-L"
                    'RM1306001 2013/06/04 追加 
                    If fncUCA2C5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "ULK", "ULK-V", "HCM", "JSM2-V"
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "CMA2", "CMA2-D", "CMA2-E", "CMA2-T", "JSM2", "CMA2-H", "STK"
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "RRC", "MRG2"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "T3PV" Then
                        fncCylinderC5Check = True
                    End If
                Case "NCK"
                    'RM1306001 2013/06/04 追加 
                    If fncNCKC5Check(objKtbnStrc, bolJudgeDiv) = True Then
                        fncCylinderC5Check = True
                    End If
                Case "FJ"
                    'RM1306001 2013/06/04 追加
                    If objKtbnStrc.strcSelection.strOpSymbol(4) = "SX" Then
                        fncCylinderC5Check = True
                    End If
                    If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                        fncCylinderC5Check = True
                    End If
                    'RM1310067 2013/10/23 追加
                Case "HHC", "HHD", "CKT", "CKU", "HLF"
                    fncCylinderC5Check = True
                Case "LHA"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "L" Then
                        fncCylinderC5Check = True
                    End If
                Case "3GA1", "3GA2", _
                     "4GA1", "4GA2", _
                     "3GB1", "3GB2", _
                     "4GB1", "4GB2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "66", "67", "76", "77"
                                fncCylinderC5Check = True
                        End Select
                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Contains("M") Then
                            fncCylinderC5Check = True
                        End If
                    End If
                Case "M3GA1", "M3GA2", _
                     "M4GA1", "M4GA2", _
                     "M3GB1", "M3GB2", _
                     "M4GB1", "M4GB2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" _
                        Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "66", "67", "76", "77", "8"
                                fncCylinderC5Check = True
                        End Select
                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Contains("M") Then
                            fncCylinderC5Check = True
                        End If
                    End If
                Case "MFC" 'RM1610011 K.Ohwaki
                    fncCylinderC5Check = True
                    'RM1710011_形番追加
                Case "UB"
                    fncCylinderC5Check = True
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncCAC3C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダCAC3のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2008/04/21      更新者：T.Sato
    '*   ・受付No.RM0802086対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '*  ・受付No：RM0811133  CAC4新発売
    '*                                      更新日：2009/07/27   更新者：Y.Miura
    '********************************************************************************************
    Private Function fncCAC3C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncCAC3C5Check = False

            '配管ねじ判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                Case "N", "G"
                    'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                    If bolJudgeDiv Then
                        fncCAC3C5Check = True
                    End If
                Case Else
            End Select

            'スイッチ判定
            If bolJudgeDiv Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    Case "T2YDU"
                        fncCAC3C5Check = True
                End Select
            End If

            'スズキ特注
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "R", "S"
                    If objKtbnStrc.strcSelection.strOpSymbol(15).Trim = "S040" Or _
                objKtbnStrc.strcSelection.strOpSymbol(15).Trim = "S050" Then
                        fncCAC3C5Check = False
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncCMK2C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダCMK2のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2008/03/25      更新者：T.Sato
    '*   ・受付No.RM0707057対応  CMK2ロッド先端特注対応
    '*                           ボックス削除に伴いロッド先端パターン判定を修正
    '*                                          更新日：2008/04/21      更新者：T.Sato
    '*   ・受付No.RM0802086対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '*   ・受付No.RM0908030対応  二次電池対応機器
    '*                                          更新日：2009/09/04      更新者：Y.Miura
    '*   ・受付No.RM1001043対応  二次電池対応機器　チェック区分変更 3→2　
    '*                                          更新日：2010/02/22      更新者：Y.Miura
    '********************************************************************************************
    Private Function fncCMK2C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncCMK2C5Check = False

            'キー形番毎に判定
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                'RM0908030 2009/09/04 Y.Miura 二次電池対応機器
                'Case ""
                Case "", "4", "5"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", "S", "SR", "B", "P", _
                             "R", "Q", "M", "C", "Z", _
                             "H", "T", "F", "G2", "G3", _
                             "JG2", "JG3"
                        Case Else
                            fncCMK2C5Check = True
                    End Select

                    '支持形式判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "CC"
                            'Qを含む場合
                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                                fncCMK2C5Check = True
                            End If
                        Case "CC1"
                            'QまたはZを含む場合
                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Or _
                               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Z") >= 0 Then
                                fncCMK2C5Check = True
                            End If
                    End Select

                    '配管ねじ、クッション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "C"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "B", "G2"
                                    fncCMK2C5Check = True
                            End Select
                        Case "GC", "NC", "GN", "NN"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncCMK2C5Check = True
                            End If
                    End Select

                    'スイッチ形番
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                        Case "T3PH", "T3PV"
                            fncCMK2C5Check = True
                        Case Else
                    End Select

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(15), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "J", "L", "M0", "M1"
                                '無し
                            Case "F", "FE"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "G2"
                                        fncCMK2C5Check = True
                                End Select
                            Case "V"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "Q", "G2", "G3"
                                        fncCMK2C5Check = True
                                End Select
                                'RM0908030 2009/09/04 Y.Miura 二次電池対応機器
                            Case "P4", "P40"
                                'RM1001043　二次電池C5加算しない、チェック区分変更しない
                                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                'If bolJudgeDiv Then
                                'fncCMK2C5Check = True
                                'End If
                            Case "P6"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "Z", "G2", "G3"
                                        fncCMK2C5Check = True
                                End Select

                                '配管ねじ、クッション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "C", "GC", "NC"
                                        fncCMK2C5Check = True
                                End Select
                            Case "P7", "P71"
                                'バリエーション判定
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                                    fncCMK2C5Check = True
                                End If
                            Case "A2"
                                fncCMK2C5Check = True
                        End Select
                    Next

                    'ロッド先端パターン判定
                    If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                        fncCMK2C5Check = True
                    End If

                    'オプション(食品製造工程向け商品)
                    If objKtbnStrc.strcSelection.strOpSymbol(16).Trim = "FP1" Then
                        fncCMK2C5Check = True
                    End If

                    'RM1306005 2013/06/04 追加
                    '2013/06/19 変更
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(17).Trim = "SX" Then
                            fncCMK2C5Check = True
                        End If
                    End If

                Case "D", "E"
                    'バリエーション判定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "D" Then
                        fncCMK2C5Check = True
                    End If

                    '配管ねじ、クッション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case ""
                        Case "GC", "NC", "GN", "NN"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncCMK2C5Check = True
                            End If
                        Case Else
                            fncCMK2C5Check = True
                    End Select

                    'スイッチ形番
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        Case "T3PH", "T3PV"
                            fncCMK2C5Check = True
                        Case Else
                    End Select

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "P7", "P71", "A2"
                                fncCMK2C5Check = True
                        End Select
                    Next

                    'ロッド先端パターン判定
                    If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                        fncCMK2C5Check = True
                    End If

                    'オプション(食品製造工程向け商品)
                    If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "FP1" Then
                        fncCMK2C5Check = True
                    End If

                    'RM1306005 2013/06/04 追加
                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "SX" Then
                        fncCMK2C5Check = True
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncJSC3C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダJSC3のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncJSC3C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncJSC3C5Check = False

            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "K", "T1", "T2", "G", "G1", "VK", _
                     "VG", "VG1", "VKG", "VKG1", "KH", _
                     "KT", "KG", "KG1", "KTG1", "TG1", _
                     "NG", "LNG", "NG1", "LNG1", "HG", "LHG", _
                     "HG1", "LHG1", "T2G1", "KT2"
                    fncJSC3C5Check = True
            End Select

            '強磁界SW識別判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "L2"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "", "VK", "VG", "VG1", "VKG", "VKG1", "KH"
                        Case Else
                            fncJSC3C5Check = True
                    End Select
                Case "L2T"
                    fncJSC3C5Check = True
            End Select

            '支持形式判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                Case "TF", "TD", "TE"
                    fncJSC3C5Check = True
            End Select

            '配管ねじ判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Case "N", "G"
                    fncJSC3C5Check = True
            End Select

            'スイッチ判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                Case "E0"
                    fncJSC3C5Check = True
            End Select

            'ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                Case "40", "50", "63"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 600 Then
                        fncJSC3C5Check = True
                    End If
                Case "80"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 700 Then
                        fncJSC3C5Check = True
                    End If
                Case "100"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 800 Then
                        fncJSC3C5Check = True
                    End If
                Case "125", "140", "160"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 800 Then
                        fncJSC3C5Check = True
                    End If
                Case "180"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 900 Then
                        fncJSC3C5Check = True
                    End If
            End Select

            'オプション判定
            If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("A2") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("P12") >= 0 Then
                fncJSC3C5Check = True
            End If

            'ロッド先端特注判定
            If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                fncJSC3C5Check = True
            End If

            'オプション外判定
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                fncJSC3C5Check = True
            End If

            ' T2YDPU,T2YDUスイッチはC5(価格は販促価格)
            If bolJudgeDiv Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                    Case "T2YDPU", "T2YDU"
                        fncJSC3C5Check = True
                End Select
            End If

            'スズキ特注
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "R", "S"
                    If objKtbnStrc.strcSelection.strOpSymbol(16).Trim = "S040" Or _
                objKtbnStrc.strcSelection.strOpSymbol(16).Trim = "S050" Then
                        fncJSC3C5Check = False
                    End If
            End Select


        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncJSC4C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダJSC4のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncJSC4C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncJSC4C5Check = False

            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "H", "T", "LH", "NG", "LNG", "NG1", "LNG1", "HG", "LHG", "HG1", "LHG1", "TG1"
                    fncJSC4C5Check = True
            End Select

            '配管ねじ判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Case "N", "G"
                    fncJSC4C5Check = True
            End Select

            'ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                Case "125", "140", "160"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 800 Then
                        fncJSC4C5Check = True
                    End If
                Case "180"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 900 Then
                        fncJSC4C5Check = True
                    End If
            End Select

            'オプション判定
            If InStr(objKtbnStrc.strcSelection.strOpSymbol(13), "A2") <> 0 Then
                fncJSC4C5Check = True
            End If

            'ロッド先端特注判定
            If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                fncJSC4C5Check = True
            End If

            'オプション外判定
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                fncJSC4C5Check = True
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncJSGC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダJSGのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncJSGC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncJSGC5Check = False

            '配管ねじ判定
            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                fncJSGC5Check = True
            End If

            'ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "40", "50", "63"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 600 Then
                        fncJSGC5Check = True
                    End If
                Case "80"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 700 Then
                        fncJSGC5Check = True
                    End If
                Case "100"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 800 Then
                        fncJSGC5Check = True
                    End If
            End Select

            'スイッチる判定(T2YDPUはチェック「3」。ただし価格は販促価格)
            If bolJudgeDiv Then
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T2YDPU" Or _
                objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T3PH" Or _
                objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T3PV" Then
                    fncJSGC5Check = True
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSCA2C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSCA2のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2008/04/21      更新者：T.Sato
    '*   ・受付No.RM0802086対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '********************************************************************************************
    Private Function fncSCA2C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncSCA2C5Check = False

            'キー形番毎に判定
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "", "2"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "T1", "T2", "PQ2", "PK", "PH", _
                             "PQ2K", "PKH", "RQ2", "RK", "RO", _
                             "RU", "RG", "RG1", "RG2", "RG3", _
                             "RG4", "RQ2K", "RKO", "RKG", "RKG1", _
                             "RKG4", "Q2K", "KH", "KT", "KT1", _
                             "KT2", "KO", "KG", "KG1", "KG4", _
                             "KTG1", "KT1G1", "KT2G1", "TG1", "T1G1", _
                             "T2G1", "T2G4"
                            fncSCA2C5Check = True
                    End Select

                    'スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                            fncSCA2C5Check = True
                        End If
                    End If

                    '配管ねじ判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case ""
                        Case "N", "G"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSCA2C5Check = True
                            End If
                        Case Else
                            fncSCA2C5Check = True
                    End Select

                    '落下防止機構判定
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "HR" Then
                        fncSCA2C5Check = True
                    End If

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(13), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "J", "L"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G2") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G3") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") < 0 Then
                                    fncSCA2C5Check = True
                                End If
                            Case "M"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1") >= 0 Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") >= 0 Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") >= 0 Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") >= 0 Then
                                    fncSCA2C5Check = True
                                End If
                            Case "P6"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                                    fncSCA2C5Check = True
                                End If
                                '2011/03/25 ADD RM1103062(4月VerUP：営業所問合せ対応) START--->
                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                                    fncSCA2C5Check = True
                                End If
                                '2011/03/25 ADD RM1103062(4月VerUP：営業所問合せ対応) <---END

                            Case "P12", "A2"
                                fncSCA2C5Check = True
                        End Select
                    Next

                    'ロッド先端特注判定
                    If objKtbnStrc.strcSelection.strRodEndOption <> "" Then
                        fncSCA2C5Check = True
                    End If

                    '2012/07/27 オプション外判定
                    If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                        fncSCA2C5Check = True
                    End If

                    ' T2YDUスイッチの場合はC5(ただし販促価格)
                    If bolJudgeDiv Then
                        If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "T2YDU" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "T3PV" Then
                            fncSCA2C5Check = True
                        End If
                    End If

                    'オプション(食品製造工程向け商品)
                    If objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "FP1" Then
                        fncSCA2C5Check = True
                    End If

                Case "B", "C"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "BK", "BH", "BT", "BT1", "BT2", _
                             "BO", "BG", "BG1", "BG2", "BG3", _
                             "BG4", "BKH", "BKT", "BKT1", "BKT2", _
                             "BKO", "BKG", "BKG1", "BKG4", "BKTG1", _
                             "BKT1G1", "BKT2G1", "BTG1", "BT1G1", "BT2G1", _
                             "WK", "WH", "WT", "WT1", "WT2", _
                             "WG", "WG1", "WG2", "WG3", "WG4", _
                             "WKH", "WKT", "WKT1", "WKT2", "WKG", _
                             "WKG1", "WKG4", "WKTG1", "WKT1G1", "WTG1", _
                             "WT1G1", "WT2G1"
                            fncSCA2C5Check = True
                    End Select

                    'スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                            fncSCA2C5Check = True
                        End If
                    End If

                    '配管ねじ判定
                    'S1：配管ねじ
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "N", "G"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSCA2C5Check = True
                            End If
                    End Select
                    'S2：配管ねじ
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                        Case "N", "G"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSCA2C5Check = True
                            End If
                    End Select

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "J", "L"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G2") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G3") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") < 0 Then
                                    fncSCA2C5Check = True
                                End If
                            Case "P6", "P12", "A2"
                                fncSCA2C5Check = True
                        End Select
                    Next

                    'ロッド先端特注判定
                    If objKtbnStrc.strcSelection.strRodEndOption <> "" Then
                        fncSCA2C5Check = True
                    End If

                    '2012/07/27 オプション外判定
                    If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                        fncSCA2C5Check = True
                    End If

                    ' T2YDUスイッチの場合はC5(ただし販促価格)
                    If bolJudgeDiv Then
                        If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "T2YDU" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "T2YDU" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "T3PH" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "T3PH" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "T3PV" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "T3PV" Then
                            fncSCA2C5Check = True
                        End If
                    End If

                    'オプション(食品製造工程向け商品)
                    If objKtbnStrc.strcSelection.strOpSymbol(18).Trim = "FP1" Then
                        fncSCA2C5Check = True
                    End If

                Case "D", "E"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "DQ2", "DK", "DH", "DG", "DG1", _
                             "DG2", "DG3", "DG4", "DQ2K", "DKH", _
                             "DKG", "DKG1", "DKG4", "DT", "DT1", "DT2"
                            fncSCA2C5Check = True
                    End Select

                    'スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                            fncSCA2C5Check = True
                        End If
                    End If

                    '配管ねじ判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case ""
                        Case "N", "G"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSCA2C5Check = True
                            End If
                        Case Else
                            fncSCA2C5Check = True
                    End Select

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "J", "L"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G2") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G3") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") < 0 Then
                                    fncSCA2C5Check = True
                                End If
                            Case "M"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1") >= 0 Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") >= 0 Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") >= 0 Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") >= 0 Then
                                    fncSCA2C5Check = True
                                End If
                                '2011/03/25 MOD RM1103062(4月VerUP：営業所問合せ対応) START--->
                            Case "P6", "P12", "A2"
                                'Case "P6"
                                '    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "D" Then
                                '        fncSCA2C5Check = True
                                '    End If
                                'Case "P12", "A2"
                                '2011/03/25 MOD RM1103062(4月VerUP：営業所問合せ対応) <---END
                                fncSCA2C5Check = True
                        End Select
                    Next

                    'ロッド先端特注判定
                    If objKtbnStrc.strcSelection.strRodEndOption <> "" Then
                        fncSCA2C5Check = True
                    End If

                    '2012/07/27 オプション外判定
                    If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                        fncSCA2C5Check = True
                    End If

                    ' T2YDUスイッチの場合はC5(ただし販促価格)
                    If bolJudgeDiv Then
                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "T2YDU" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "T3PV" Then
                            fncSCA2C5Check = True
                        End If
                    End If

                    'オプション(食品製造工程向け商品)
                    If objKtbnStrc.strcSelection.strOpSymbol(13).Trim = "FP1" Then
                        fncSCA2C5Check = True
                    End If
               
                Case "V"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "PV1", "PV2", "PV", "PV1K", "PV2K", _
                             "PVK", "RV1", "RV2", "RV", "RV1K", _
                             "RV2K", "RVK", "RV1G", "RV2G", "RVG", _
                             "RV1G1", "RV2G1", "RVG1", "RV1G4", "RV2G4", _
                             "RVG4", "RV1KG", "RV2KG", "RVKG", "RV1KG1", _
                             "RV2KG1", "RVKG1", "RV1KG4", "RVKG4", "V1K", _
                             "V2K", "VK", "V1G", "V2G", "VG", _
                             "V1G1", "V2G1", "VG1", "V1G4", "V2G4", _
                             "VG4", "V1KG", "V2KG", "VKG", "V1KG1", _
                             "V2KG1", "VKG1", "V1KG4", "V2KG4", "VKG4"
                            fncSCA2C5Check = True
                    End Select

                    'スイッチ判定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                            fncSCA2C5Check = True
                        End If
                    End If

                    '配管ねじ判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case ""
                        Case "N", "G"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSCA2C5Check = True
                            End If
                        Case Else
                            fncSCA2C5Check = True
                    End Select

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(13), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "J", "L"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G2") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G3") < 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") < 0 Then
                                    fncSCA2C5Check = True
                                End If
                            Case "A2"
                                fncSCA2C5Check = True
                        End Select
                    Next

                    'ロッド先端特注判定
                    If objKtbnStrc.strcSelection.strRodEndOption <> "" Then
                        fncSCA2C5Check = True
                    End If

                    '2012/07/27 オプション外判定
                    If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                        fncSCA2C5Check = True
                    End If

                    ' T2YDUスイッチの場合はC5(ただし販促価格)
                    If bolJudgeDiv Then
                        If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "T2YDU" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "T3PH" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "T3PV" Then
                            fncSCA2C5Check = True
                        End If
                    End If
            End Select

            'ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                Case "40", "50", "63"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 600 Then
                        fncSCA2C5Check = True
                    End If

                    'バリエーション「B」を含む場合はS2もチェック
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 600 Then
                            fncSCA2C5Check = True
                        End If
                    End If
                Case "80"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 700 Then
                        fncSCA2C5Check = True
                    End If

                    'バリエーション「B」を含む場合はS2もチェック
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 700 Then
                            fncSCA2C5Check = True
                        End If
                    End If
                Case "100"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 800 Then
                        fncSCA2C5Check = True
                    End If

                    'バリエーション「B」を含む場合はS2もチェック
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 800 Then
                            fncSCA2C5Check = True
                        End If
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSCGC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSCGのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncSCGC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncSCGC5Check = False

            '配管ねじによる判定
            If bolJudgeDiv Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    Case ""
                        '2009/12/09 Y.Miura ローカル版と記述を合わせる
                        'Case "N", "G"
                        '    'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                        '    If bolJudgeDiv Then
                        '        fncSCGC5Check = True
                        '    End If
                    Case Else
                        fncSCGC5Check = True
                End Select
            End If

            'ストロークによる判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "32", "40", "50", "63"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 600 Then
                        fncSCGC5Check = True
                    End If
                Case "80"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 700 Then
                        fncSCGC5Check = True
                    End If
                Case "100"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 800 Then
                        fncSCGC5Check = True
                    End If
            End Select

            'スイッチ判定
            If bolJudgeDiv Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    Case "T2YDU", "T3PH", "T3PV"
                        fncSCGC5Check = True
                End Select
            End If

            'RM1306001 2013/06/04 追加
            '2013/07/04 修正
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "SCG-Q"
                Case Else
                    If objKtbnStrc.strcSelection.strKeyKataban <> "4" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "SX" Then
                            fncSCGC5Check = True
                        End If
                    End If
            End Select
            'RM1001043 2010/02/22 Y.Miura　二次電池のC5チェックをなくす
            'RM0907070 2009/08/21 Y.Miura　二次電池対応
            'P4※二次電池は特注
            'Dim strOpArray() As String
            'Dim intLoopCnt As Integer
            'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
            'For intLoopCnt = 0 To strOpArray.Length - 1
            '    Select Case strOpArray(intLoopCnt).Trim
            '        Case "P4", "P40"
            '            'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
            '            'If bolJudgeDiv Then
            '            fncSCGC5Check = True
            '            'End If
            '    End Select
            'Next

            'オプション(食品製造工程向け商品)
            If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "FP1" Then
                fncSCGC5Check = True
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSCMC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSCMのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2008/04/21      更新者：T.Sato
    '*   ・受付No.RM0802086対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '*   ・受付No：RM1001043対応  二次電池P4*のC5チェックをなくす
    '*                                      更新日：2010/02/22   更新者：Y.Miura
    '********************************************************************************************
    Private Function fncSCMC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncSCMC5Check = False

            'キー形番毎にチェックする
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                'RM0907070 2009/08/21 Y.Miura　二次電池対応
                'Case ""
                Case "", "4", "F"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "H", "T1", "T2", "G", "G1", _
                             "G2", "G3", "G4", "XM", "XT2", _
                             "YM", "YT2", "W4M", "W4H", "W4T", _
                             "W4T1", "W4T2", "W4G", "W4G1", "W4G2", _
                             "W4G3", "W4G4", "W4HG", "W4TG1", "W4T1G1", _
                             "W4T2G1", "W4T2G4", "PM", "PH", "PT2", _
                             "RM", "RO", "RG", "RG1", "RG4", _
                             "RF", "HG", "TG1", "T1G1", "T2G1", _
                             "T2G4"
                            fncSCMC5Check = True
                    End Select

                    '支持形式判定
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "LD" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                            fncSCMC5Check = True
                        End If
                    End If

                    '口径判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "W4", "P", "R"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "80", "100"
                                    fncSCMC5Check = True
                            End Select
                    End Select

                    '配管ねじ判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case ""
                        Case "N", "G"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSCMC5Check = True
                            End If
                        Case Else
                            fncSCMC5Check = True
                    End Select

                    'オプションによる判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(13), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "J", "K", "L"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Q" Then
                                    fncSCMC5Check = True
                                End If
                            Case "M"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Q" Then
                                    fncSCMC5Check = True
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M") >= 0 Then
                                    fncSCMC5Check = True
                                End If
                                'RM1001043 2010/02/22 Y.Miura　二次電池のC5対応をなくす
                                'RM0907070 2009/08/21 Y.Miura　二次電池対応
                            Case "P4", "P40"
                                'fncSCMC5Check = True
                            Case "P5", "P51", "A2"
                                fncSCMC5Check = True
                            Case "P6"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Q" Then
                                    fncSCMC5Check = True
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "LD" Then
                                    fncSCMC5Check = True
                                End If
                            Case "P7", "P71"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("F") >= 0 Then
                                    fncSCMC5Check = True
                                End If
                        End Select
                    Next

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "F"
                            'オプション(食品製造工程向け商品)
                            If objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "FP1" Then
                                fncSCMC5Check = True
                            End If
                        Case Else
                            'ロッド先端パターン判定
                            If objKtbnStrc.strcSelection.strOpSymbol(15).Trim <> "" Then
                                fncSCMC5Check = True
                            End If

                            'RM1306001 2013/06/04 追加
                            '2013/06/19 修正
                            If objKtbnStrc.strcSelection.strKeyKataban <> "4" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(16).Trim = "SX" Then
                                    fncSCMC5Check = True
                                End If
                            End If
                    End Select

                    ' T2YDUスイッチの場合はC5(ただし販促価格)
                    If bolJudgeDiv Then
                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "T2YDU" Then
                            fncSCMC5Check = True
                        End If
                    End If

                Case "B", "G"
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "B", "W"
                        Case Else
                            fncSCMC5Check = True
                    End Select

                    '口径判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "B", "W"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "80", "100"
                                    fncSCMC5Check = True
                            End Select
                    End Select

                    '配管ねじによる判定
                    'S1：配管ねじ
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case ""
                        Case "N", "G"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSCMC5Check = True
                            End If
                        Case Else
                            fncSCMC5Check = True
                    End Select

                    'S2：配管ねじ
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                        Case ""
                        Case "N", "G"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSCMC5Check = True
                            End If
                        Case Else
                            fncSCMC5Check = True
                    End Select

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "P5", "P51", "P7", "P71", "A2"
                                fncSCMC5Check = True
                            Case Else
                        End Select
                    Next

                    ' T2YDUスイッチの場合はC5(ただし販促価格)
                    If bolJudgeDiv Then
                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T2YDU" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(13).Trim = "T2YDU" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T3PH" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(13).Trim = "T3PH" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T3PV" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(13).Trim = "T3PV" Then
                            fncSCMC5Check = True
                        End If
                    End If

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "G"
                            'オプション(食品製造工程向け商品)
                            If objKtbnStrc.strcSelection.strOpSymbol(18).Trim = "FP1" Then
                                fncSCMC5Check = True
                            End If
                        Case Else
                            'RM1306001 2013/06/04 追加
                            If objKtbnStrc.strcSelection.strOpSymbol(20).Trim = "SX" Then
                                fncSCMC5Check = True
                            End If

                            'ロッド先端パターン判定
                            If objKtbnStrc.strcSelection.strOpSymbol(19).Trim <> "" Then
                                fncSCMC5Check = True
                            End If
                    End Select

                Case "D", "H"
                    'バリエーション判定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "D" Then
                        fncSCMC5Check = True
                    End If

                    '配管ねじ判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case ""
                        Case "N", "G"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSCMC5Check = True
                            End If
                        Case Else
                            fncSCMC5Check = True
                    End Select

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "P5", "P51", "P7", "P71", "A2"
                                fncSCMC5Check = True
                        End Select
                    Next

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "H"
                            'オプション(食品製造工程向け商品)
                            If objKtbnStrc.strcSelection.strOpSymbol(13).Trim = "FP1" Then
                                fncSCMC5Check = True
                            End If
                        Case Else

                            'ロッド先端パターン判定
                            If objKtbnStrc.strcSelection.strOpSymbol(14).Trim <> "" Then
                                fncSCMC5Check = True
                            End If

                            'RM1306001 2013/06/04 追加
                            If objKtbnStrc.strcSelection.strOpSymbol(15).Trim = "SX" Then
                                fncSCMC5Check = True
                            End If
                    End Select

                    ' T2YDUスイッチの場合はC5(ただし販促価格)
                    If bolJudgeDiv Then
                        If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "T2YDU" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "T3PH" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "T3PV" Then
                            fncSCMC5Check = True
                        End If
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSCSC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSCSのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncSCSC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncSCSC5Check = False

            'バリエーションによる判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "NP", "LNP", "T1", "LNT2", "NG", _
                     "LNG", "G1", "NG1", "LNG1", "PH", _
                     "LPH", "PT", "PT1", "LNPT2", "HG", _
                     "LHG", "HG1", "LHG1", "TG1", "T1G1", _
                     "ND", "DH", "LDH", "DT", "DT1", _
                     "LNDT2", "DG", "NDG", "LNDG", "DG1", _
                     "NDG1", "LNDG1", "DHG", "LDHG", "DHG1", _
                     "LDHG1", "DTG1", "DT1G1", "LNDT2G1", "NB", _
                     "LNB", "NW", "LNW", "BH", "LBH", _
                     "BT", "BT1", "LNBT2", "BG", "NBG", _
                     "LNBG", "BG1", "NBG1", "LNBG1", "WH", _
                     "LWH", "WT", "WT1", "LNWT2", "WG", _
                     "NWG", "LNWG", "WG1", "NWG1", "LNWG1", _
                     "BHG", "LBHG", "BHG1", "LBHG1", "BTG1", _
                     "BT1G1", "LNBT2G1", "WHG", "LWHG", "WHG1", _
                     "LWHG1", "WTG1", "WT1G1", "LNWT2G1"
                    fncSCSC5Check = True
            End Select

            '配管ねじ、クッション判定
            'S1
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                Case "G", "N"
                    fncSCSC5Check = True
            End Select
            'S2
            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                Case "G", "N"
                    fncSCSC5Check = True
            End Select

            'オプションによる判定
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "U1", "A2"
                        fncSCSC5Check = True
                End Select
            Next

            'ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "125", "140", "160"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 800 Then
                        fncSCSC5Check = True
                    End If
                Case "180"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 900 Then
                        fncSCSC5Check = True
                    End If
                Case "200"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 1000 Then
                        fncSCSC5Check = True
                    End If
                Case "250"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 1200 Then
                        fncSCSC5Check = True
                    End If
            End Select

            'ロッド先端特注判定
            If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                fncSCSC5Check = True
            End If

            'オプション外判定
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                fncSCSC5Check = True
            End If

            'T2YDUPスイッチはC5(価格は販促価格)
            If bolJudgeDiv Then
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T2YDUP" Or _
                   objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "T2YDUP" Then
                    fncSCSC5Check = True
                End If
            End If

            'シリーズ形番判定
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "2"
                    fncSCSC5Check = True
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSCS2C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSCS2のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncSCS2C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncSCS2C5Check = False

            'バリエーションによる判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                '2014/01/20 E.Murata "W"以下追加
                Case "P", "H", "LH", "G", "D", "LND", "B", "W", _
                     "NP", "LNP", "LNT2", "NG", "LNG", "G1", "NG1", _
                     "LNG1", "PH", "LPH", "PT", "LNPT2", "HG", "LHG", _
                     "HG1", "LHG1", "ND", "DH", "LDH", "DT", "LNDT2", _
                     "DG", "NDG", "LNDG", "DG1", "NDG1", "LNDG1", _
                     "DHG", "LDHG", "DHG1", "LDHG1", "NB", "LNB", _
                     "NW", "LNW", "BH", "LBH", "BT", "LNBT2", "BG", _
                     "NBG", "LNBG", "BG1", "NBG1", "LNBG1", "WH", _
                     "LWH", "WT", "LNWT2", "WG", "NWG", "LNWG", "WG1", _
                     "NWG1", "LNWG1", "BHG", "LBHG", "BHG1", "LBHG1", _
                     "WHG", "LWHG", "WHG1", "LWHG1"
                    fncSCS2C5Check = True
            End Select

            '配管ねじ、クッション判定
            'S1
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                Case "G", "N"
                    fncSCS2C5Check = True
            End Select
            'S2
            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                Case "G", "N"
                    fncSCS2C5Check = True
            End Select

            'オプションによる判定
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "U1", "A2", "P6"
                        fncSCS2C5Check = True
                End Select
            Next
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(18), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "U1", "A2", "P6"
                        fncSCS2C5Check = True
                End Select
            Next

            'ストローク判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "125", "140", "160"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 800 Then
                        fncSCS2C5Check = True
                    End If
                Case "180"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 900 Then
                        fncSCS2C5Check = True
                    End If
                Case "200"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 1000 Then
                        fncSCS2C5Check = True
                    End If
                Case "250"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 1200 Then
                        fncSCS2C5Check = True
                    End If
            End Select

            'RM1305007 2013/05/07
            'ロッド先端特注判定
            If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                fncSCS2C5Check = True
            End If

            'オプション外判定
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                fncSCS2C5Check = True
            End If

            '食品製造工程向け商品
            Select Case objKtbnStrc.strcSelection.strOpSymbol(19).Trim
                Case "FP1"
                    fncSCS2C5Check = True
            End Select

            '二次電池
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "4"
                    fncSCS2C5Check = True
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSSDC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSSDのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】                          更新日：2007/10/24   更新者：NII A.Takahashi
    '*   ・ロッド先端特注を選択した場合C5対応とする
    '*                                      更新日：2008/04/21   更新者：T.Sato
    '*   ・受付No.RM0802086対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '********************************************************************************************
    Private Function fncSSDC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncSSDC5Check = False

            'キー形番毎にチェック
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "", "4"
                    'バリエーション①判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "XB", "XBT", "XBT2", "XT", "XT2", "YB", "YBT", "YBT2", "YT", "YT2", _
                            "BQ", "BM", "BMO", "BT", "BTG1", "BT1", "BT1G1", "BT1L", "BG1T1L", "BT2", "BT2G1", _
                            "BO", "BG", "BG1", "BG2", "BG3", "BG4", "WM", "WMO", "WT", "WT1", "WT2", "WO", "MO", _
                            "TG1", "T1", "T1G1", "G1T1L", "T2", "T2G1", "G5"
                            fncSSDC5Check = True
                        Case "B", "W", "T", "O", "G2", "G3"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "125", "140", "160"
                                    fncSSDC5Check = True
                            End Select
                        Case "T1L"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "80", "100"
                                    fncSSDC5Check = True
                            End Select
                        Case "G"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "16", "20", "25", "125", "140", "160"
                                    fncSSDC5Check = True
                            End Select
                        Case "G1"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "16", "20", "125", "140", "160"
                                    fncSSDC5Check = True
                            End Select
                    End Select

                    'バリエーション②(スイッチ付)判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "L4"
                            'バリエーション判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "", "G1"
                                Case Else
                                    fncSSDC5Check = True
                            End Select
                    End Select

                    'バリエーション③(微速)判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "F"
                            'バリエーション判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case ""
                                Case Else
                                    fncSSDC5Check = True
                            End Select

                            'バリエーション②(スイッチ付)判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "L4"
                                    fncSSDC5Check = True
                            End Select

                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "125", "140", "160"
                                    fncSSDC5Check = True
                            End Select
                    End Select

                    '配管ねじ、クッション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "GN", "NN", "GD", "ND"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSSDC5Check = True
                            End If
                    End Select

                    'S1バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        Case "K", "M", "KM"
                            fncSSDC5Check = True
                    End Select

                    'S2バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                        Case "K", "M", "KM"
                            fncSSDC5Check = True
                    End Select

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "M"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "", "X", "Y", "B", "W", _
                                         "M", "T", "O"
                                    Case Else
                                        fncSSDC5Check = True
                                End Select

                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "125", "140", "160"
                                        fncSSDC5Check = True
                                End Select
                            Case "M1"
                                fncSSDC5Check = True
                                'RM0906034 2009/08/05 Y.Miura　二次電池対応
                                'オプション判定
                            Case "P6"
                                'S1バリエーション:S2バリエーション判定
                                Select Case True
                                    Case objKtbnStrc.strcSelection.strOpSymbol(6).IndexOf("K") < 0 And _
                                         objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("K") < 0
                                        'バリエーション判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                            Case ""
                                            Case Else
                                                fncSSDC5Check = True
                                        End Select

                                        'バリエーション(スイッチ付)判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            Case "L4"
                                                fncSSDC5Check = True
                                        End Select

                                        '内径判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            Case "125", "140", "160"
                                                fncSSDC5Check = True
                                        End Select
                                    Case objKtbnStrc.strcSelection.strOpSymbol(6).IndexOf("K") >= 0 Or _
                                         objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("K") >= 0
                                        fncSSDC5Check = True
                                End Select
                            Case "P5", "P51"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case ""
                                    Case Else
                                        fncSSDC5Check = True
                                End Select

                                'バリエーション(スイッチ付)判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "L4"
                                        fncSSDC5Check = True
                                End Select

                                'オプション判定
                                If objKtbnStrc.strcSelection.strOpSymbol(19).IndexOf("A2") >= 0 Then
                                    fncSSDC5Check = True
                                End If
                            Case "P7", "P71"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case ""
                                    Case Else
                                        fncSSDC5Check = True
                                End Select

                                'バリエーション（スイッチ付）判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "L4"
                                        fncSSDC5Check = True
                                End Select

                                'オプション判定
                                If objKtbnStrc.strcSelection.strOpSymbol(19).IndexOf("A2") >= 0 Then
                                    fncSSDC5Check = True
                                End If
                                '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) START--->
                            Case "A2", "R1", "R2"
                                'Case "A2"
                                '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) <---END
                                fncSSDC5Check = True
                            Case "P4", "P40"
                                'RM1002043 2010/02/22 Y.Miura　二次電池C5加算廃止 
                                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                'If bolJudgeDiv Then
                                'fncSSDC5Check = True
                                'End If
                                fncSSDC5Check = True
                        End Select
                    Next

                    'ロッド先端パターン判定
                    If objKtbnStrc.strcSelection.strRodEndOption <> "" Then
                        fncSSDC5Check = True
                    End If

                    'T2YDUスイッチの場合はC5(ただし加算は販促価格)
                    If bolJudgeDiv Then
                        'RM0912039 2009/12/16 Y.Miura スイッチ追加
                        'If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "T2YDU" Or _
                        '   objKtbnStrc.strcSelection.strOpSymbol(16).Trim = "T2YDU" Then
                        '    fncSSDC5Check = True
                        'End If
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            Case "T2YDU", "T2HR3", "T2VR3", "T3PH", "T3PV"
                                fncSSDC5Check = True
                        End Select
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                            Case "T2YDU", "T2HR3", "T2VR3", "T3PH", "T3PV"
                                fncSSDC5Check = True
                        End Select

                    End If

                    'RM1306001 2013/06/04 追加
                    '2013/06/19 修正
                    If objKtbnStrc.strcSelection.strKeyKataban <> "4" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(22).Trim = "SX" Then
                            fncSSDC5Check = True
                        End If
                    End If
                Case "D", "E"
                    'バリエーション①判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "DQ", "DM", "DMO", "DT", "DTG1", "DT1", "DT1G1", "DT1L", "DG1T1L", "DT2", "DT2G1", _
                            "DO", "DG", "DG2", "DG3", "KD", "KDM", "KDMO", "KDT", "KDT1", "KDTG1", "KDT1G1", _
                            "KDT2", "KDT2G1", "KDO", "KDG", "KDG1", "KDG2", "KDG3", "KDG4"
                            fncSSDC5Check = True
                        Case "DG1"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "16", "20", "125", "140", "160"
                                    fncSSDC5Check = True
                            End Select
                    End Select

                    'バリエーション②(スイッチ付)判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "L4"
                            fncSSDC5Check = True
                    End Select

                    'バリエーション③(微速)判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "F"
                            fncSSDC5Check = True
                    End Select

                    '配管ねじ、クッション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "GC", "NC", "GN", "NN"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSSDC5Check = True
                            End If
                    End Select

                    'ストローク判定
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "D", "DG1", "DG4"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16", "20"
                                    'ストローク判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 31 Then
                                        fncSSDC5Check = True
                                    End If
                                Case "25", "32", "40", "50", "63", _
                                     "80", "100"
                                    'ストローク判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= 51 Then
                                        fncSSDC5Check = True
                                    End If
                            End Select
                    End Select

                    'オプションによる判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                                '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) START--->
                            Case "M", "M1", "P6", "P5", "P51", _
                                 "P7", "P71", "A2", "R1", "R2", _
                                 "P4", "P40"
                                'Case "M", "M1", "P6", "P5", "P51", _
                                '     "P7", "P71", "A2"
                                '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) <---END
                                fncSSDC5Check = True
                        End Select
                    Next

                    'ロッド先端パターン判定
                    If objKtbnStrc.strcSelection.strRodEndOption <> "" Then
                        fncSSDC5Check = True
                    End If

                    'T2YDUスイッチの場合はC5(ただし加算は販促価格)
                    If bolJudgeDiv Then
                        'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        'If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "T2YDU" Then
                        '    fncSSDC5Check = True
                        'End If
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            Case "T2YDU", "T2HR3", "T2VR3", "T3PH", "T3PV"
                                fncSSDC5Check = True
                        End Select
                    End If

                    'RM1306001 2013/06/04 追加
                    '2013/06/19　修正
                    If objKtbnStrc.strcSelection.strKeyKataban = "D" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "SX" Then
                            fncSSDC5Check = True
                        End If
                    End If
                Case "K", "P"
                    'バリエーション①判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "KB", "KBM", "KBMO", "KBT", "KBTG1", "KBT1", "KBT1G1", "KBT2", "KBT2G1", "KBO", _
                            "KBU", "KBG", "KBG1", "KBG2", "KBG3", "KBG4", "KW", "KWM", "KWMO", "KWT", "KWT1", _
                            "KWT2", "KWO", "KM", "KMO", "KT", "KTG1", "KT1", "KT1G1", "KT1L", "KT2", "KT2G1", _
                            "KO", "KG5"
                            fncSSDC5Check = True
                        Case "KG"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "16", "20", "25"
                                    fncSSDC5Check = True
                            End Select
                        Case "KG1"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "16", "20"
                                    fncSSDC5Check = True
                            End Select
                    End Select

                    'バリエーション②(スイッチ付)による判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "L4"
                            'バリエーション判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "K", "KG1"
                                Case Else
                                    fncSSDC5Check = True
                            End Select
                    End Select

                    'バリエーション③(微速)判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "F"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "20"
                                    'S1ストローク判定
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 101 Then
                                            fncSSDC5Check = True
                                        End If
                                    End If

                                    'S2ストローク判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 101 Then
                                        fncSSDC5Check = True
                                    End If
                                Case "25", "32", "40", "50"
                                    'S1ストローク判定
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 151 Then
                                            fncSSDC5Check = True
                                        End If
                                    End If

                                    'S2ストローク判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 151 Then
                                        fncSSDC5Check = True
                                    End If
                                Case "63", "80", "100"
                                    'S1ストローク判定
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 201 Then
                                            fncSSDC5Check = True
                                        End If
                                    End If

                                    'S2ストローク判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 201 Then
                                        fncSSDC5Check = True
                                    End If
                            End Select
                    End Select

                    '配管ねじ、クッション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "C"
                            'バリエーション判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "K"
                                Case Else
                                    fncSSDC5Check = True
                            End Select

                            'バリエーション(スイッチ付)判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "L4"
                                    fncSSDC5Check = True
                            End Select

                            'バリエーション(微速)判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "F"
                                    fncSSDC5Check = True
                            End Select
                        Case "GC", "NC", "GN", "NN"
                            'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                            If bolJudgeDiv Then
                                fncSSDC5Check = True
                            End If
                    End Select

                    'S1バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "M"
                            fncSSDC5Check = True
                    End Select

                    'S2バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                        Case "M"
                            fncSSDC5Check = True
                    End Select

                    'S1・S2ストローク判定
                    'バリエーション判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "KG2", "KG3"
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "20"
                                    'S1ストローク判定
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 101 Then
                                            fncSSDC5Check = True
                                        End If
                                    End If
                                    'S2ストローク判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 101 Then
                                        fncSSDC5Check = True
                                    End If
                                Case "25", "32", "40", "50"
                                    'S1ストローク判定
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 151 Then
                                            fncSSDC5Check = True
                                        End If
                                    End If
                                    'S2ストローク判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 151 Then
                                        fncSSDC5Check = True
                                    End If
                                Case "63", "80", "100"
                                    'S1ストローク判定
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 201 Then
                                            fncSSDC5Check = True
                                        End If
                                    End If
                                    'S2ストローク判定
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 201 Then
                                        fncSSDC5Check = True
                                    End If
                            End Select
                    End Select

                    'オプション判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "M"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "K"
                                    Case Else
                                        fncSSDC5Check = True
                                End Select

                                '配管ねじ、クッション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "C", "GC", "NC"
                                        fncSSDC5Check = True
                                End Select

                                '内径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "125", "140", "160"
                                        fncSSDC5Check = True
                                End Select
                            Case "P6"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "K"
                                    Case Else
                                        fncSSDC5Check = True
                                End Select

                                'バリエーション(スイッチ付)判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "L4"
                                        fncSSDC5Check = True
                                End Select

                                '配管ねじ、クッション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "C", "GC", "NC"
                                        fncSSDC5Check = True
                                End Select

                                'オプション判定
                                If objKtbnStrc.strcSelection.strOpSymbol(17).IndexOf("M") >= 0 Then
                                    fncSSDC5Check = True
                                End If
                                'RM0907070 2009/08/20 Y.Miura　二次電池対応
                                'オプション判定
                            Case "P4", "P40"
                                'RM1112XXX 2011/12/22 Y.Tachi　二次電池C5加算対応
                                'RM1001043 2010/02/22 Y.Miura 二次電池C5加算廃止
                                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                'If bolJudgeDiv Then
                                fncSSDC5Check = True
                                'End If
                            Case "P5", "P51"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "K"
                                    Case Else
                                        fncSSDC5Check = True
                                End Select

                                'バリエーション(スイッチ付)判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "L4"
                                        fncSSDC5Check = True
                                End Select

                                '配管ねじ、クッション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "C", "GC", "NC"
                                        fncSSDC5Check = True
                                End Select

                                'オプション判定
                                If objKtbnStrc.strcSelection.strOpSymbol(17).IndexOf("A2") >= 0 Then
                                    fncSSDC5Check = True
                                End If
                            Case "P7", "P71", "P12"
                                'バリエーション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "K"
                                    Case Else
                                        fncSSDC5Check = True
                                End Select

                                'バリエーション(スイッチ付)判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "L4"
                                        fncSSDC5Check = True
                                End Select

                                'バリエーション(微速)判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Case "F"
                                        fncSSDC5Check = True
                                End Select

                                '配管ねじ、クッション判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "C", "GC", "NC"
                                        fncSSDC5Check = True
                                End Select

                                'オプション判定
                                If objKtbnStrc.strcSelection.strOpSymbol(17).IndexOf("A2") >= 0 Then
                                    fncSSDC5Check = True
                                End If
                                '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) START--->
                            Case "A2", "R1", "R2"
                                'Case "A2"
                                '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) <---END
                                fncSSDC5Check = True
                        End Select
                    Next

                    'ロッド先端パターン判定
                    If objKtbnStrc.strcSelection.strRodEndOption <> "" Then
                        fncSSDC5Check = True
                    End If

                    'T2YDUスイッチの場合はC5(ただし加算は販促価格)
                    If bolJudgeDiv Then
                        'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        'If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "T2YDU" Or _
                        '   objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "T2YDU" Then
                        '    fncSSDC5Check = True
                        'End If
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                            Case "T2YDU", "T2HR3", "T2VR3", "T3PH", "T3PV"
                                fncSSDC5Check = True
                        End Select
                    End If

                    'RM1306001 2013/06/04 追加
                    '2013/06/19 修正
                    If objKtbnStrc.strcSelection.strKeyKataban = "K" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(20).Trim = "SX" Then
                            fncSSDC5Check = True
                        End If
                    End If

                    'スズキ特注
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "R", "S"
                            If objKtbnStrc.strcSelection.strOpSymbol(21).Trim = "S040" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(21).Trim = "S050" Then
                                fncSSDC5Check = False
                            End If
                    End Select

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSTGC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSTGのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/06/26      更新者：NII A.Takahashi
    '*   ・オプション「O」を選択した場合はC5対応に修正
    '*　 ・オプション「P6」を選択かつバリエーションもしくは配管ねじ種類で「C」を選択した場合はC5対応に修正
    '*　 ・バリエーション「G5」を選択した場合はC5対応に修正
    '*                                          更新日：2008/04/21      更新者：T.Sato
    '*   ・受付No.RM0802086対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '********************************************************************************************
    Private Function fncSTGC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncSTGC5Check = False

            'バリエーション判定
            'If InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "G5") <> 0 Or _
            '   InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "T2") <> 0 Then
            '    fncSTGC5Check = True
            'End If
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "G5", "T2"
                    fncSTGC5Check = True
                Case "C", "G", "G1"
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "P4" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "P40" Then
                        fncSTGC5Check = True
                    End If
            End Select

            '配管ねじ判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "GC", "NC", "GN", "NN"
                    'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                    If bolJudgeDiv Then
                        fncSTGC5Check = True
                    End If
                Case "C"
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "P4" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "P40" Then
                        fncSTGC5Check = True
                    End If
            End Select

            'RM1001043 2010/02/22 Y.Miura 二次電池C5加算廃止
            ''RM0906034 2009/08/18 Y.Miura　二次電池対応
            ''P4※二次電池は特注
            'Dim strOpArray() As String
            'Dim intLoopCnt As Integer
            'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
            'For intLoopCnt = 0 To strOpArray.Length - 1
            '    Select Case strOpArray(intLoopCnt).Trim
            '        Case "P4", "P40", "P41"
            '            'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
            '            'If bolJudgeDiv Then
            '            fncSTGC5Check = True
            '            'End If
            '    End Select
            'Next

            'オプション判定
            If InStr(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, "O") <> 0 Then
                fncSTGC5Check = True
            End If
            If InStr(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, "P6") <> 0 Then
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "STG-M", "STG-B"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "C" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "C" Then
                            fncSTGC5Check = True
                        End If
                End Select
            End If

            'スイッチ判定
            If bolJudgeDiv Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    Case "T2YDU"
                        fncSTGC5Check = True
                End Select
            End If

            'RM1306005 2013/06/04 追加
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "STG-M", "STG-B"
                    '2013/06/19 修正
                    If objKtbnStrc.strcSelection.strKeyKataban = "" Then
                        If InStr(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, "SX") <> 0 Then
                            fncSTGC5Check = True
                        End If
                    End If
            End Select

            'オプション(食品製造工程向け商品)
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "F"
                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "FP1" Then
                        fncSTGC5Check = True
                    End If
            End Select

            'スズキ特注
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "R", "S"
                    If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "S040" Or _
                objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "S050" Then
                        fncSTGC5Check = False
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSTSC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSTS/STLのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2008/04/21      更新者：T.Sato
    '*   ・受付No.RM0802086対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '********************************************************************************************
    Private Function fncSTSC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncSTSC5Check = False

            'バリエーション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "T1", "T1L", "PQ", "PV1", "PV2", _
                     "PV1S", "PV2S", "PC", "PT2", "PO", _
                     "PG", "PG1", "PG4", "CT", "CT1", _
                     "CT2", "CO", "CG", "CG1", "CG2", _
                     "CG3", "CG4", "CTG1", "CT1G1", "CT2G1", _
                     "TG1", "T1G1", "T1LG1", "T2G1", "PV1O", _
                     "PV2O", "PV1SO", "PV2SO", "PCT2", "PCO", _
                     "PCG", "PCG1", "PCG4", "PCT2G1"
                    fncSTSC5Check = True
            End Select

            'スイッチ識別判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "L5"
                    fncSTSC5Check = True
            End Select

            '口径判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "100"
                    'オプションによる判定
                    If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("O") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("P52") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("P53") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("P72") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("P73") >= 0 Then
                        fncSTSC5Check = True
                    End If
            End Select

            '配管ねじ、クッション判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                Case "C"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                        fncSTSC5Check = True
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "100" Then
                        fncSTSC5Check = True
                    End If
                Case "GC", "NC", "GN", "NN"
                    'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                    If bolJudgeDiv Then
                        fncSTSC5Check = True
                    End If
            End Select

            'オプション判定
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "O"
                        fncSTSC5Check = True
                    Case "P52", "P53"
                        'バリエーション判定
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                            fncSTSC5Check = True
                        End If

                        'スイッチ識別判定
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L5" Then
                            fncSTSC5Check = True
                        End If

                        '配管ねじ、クッション判定
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                            fncSTSC5Check = True
                        End If

                        'オプション判定
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M1") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("F") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("O") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P6") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P72") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P73") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("E") >= 0 Then
                            fncSTSC5Check = True
                        End If
                    Case "M", "M1", "P6", "E"
                        'バリエーション判定
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                                fncSTSC5Check = True
                            End If
                        Else
                            fncSTSC5Check = True
                        End If
                    Case "P72", "P73"
                        'バリエーション判定
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "" Then
                            '配管ねじ、クッション判定
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                                fncSTSC5Check = True
                            End If
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("F") < 0 Then
                                fncSTSC5Check = True
                            Else
                                '配管ねじ、クッション判定
                                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                                    fncSTSC5Check = True
                                End If
                            End If
                        End If
                End Select
            Next

            'T2YDUスイッチの場合C5(ただし価格は販促価格)
            If bolJudgeDiv Then
                If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "T2YDU" Or _
                    objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "T3PH" Or _
                    objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "T3PV" Then
                    fncSTSC5Check = True
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncLCGC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダLCGのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2008/04/21      更新者：T.Sato
    '*   ・受付No.RM0802086対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '********************************************************************************************
    Private Function fncLCGC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncLCGC5Check = False

            ' 配管ねじ判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "N", "G"
                    'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                    If bolJudgeDiv Then
                        fncLCGC5Check = True
                    End If
            End Select

            'スイッチ
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                Case "F3PH", "F3PV", "T3PH", "T3PV"
                    If bolJudgeDiv Then
                        fncLCGC5Check = True
                    End If
            End Select

            'RM1306001 2013/06/04 追加
            If objKtbnStrc.strcSelection.strKeyKataban = "1" Then
                If objKtbnStrc.strcSelection.strOpSymbol(13).Trim = "SX" Then
                    fncLCGC5Check = True
                End If
            End If

            'RM1001043 2010/02/22 Y.Miura 　二次電池C5加算廃止
            ''RM0906034 2009/08/18 Y.Miura　二次電池対応
            ''P4※二次電池は特注
            'Dim strOpArray() As String
            'Dim intLoopCnt As Integer
            'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
            'For intLoopCnt = 0 To strOpArray.Length - 1
            '    Select Case strOpArray(intLoopCnt).Trim
            '        Case "P4", "P40"
            '            'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
            '            'If bolJudgeDiv Then
            '            fncLCGC5Check = True
            '            'End If
            '    End Select
            'Next

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncLCRC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダLCGのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2010/04/07      更新者：Y.Miura
    '********************************************************************************************
    Private Function fncLCRC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncLCRC5Check = False

            'ストッパ
            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                Case "W1", "W2", "W3", "W4", "W5", "W6", "C1", "C2", "C3", "C4"
                    fncLCRC5Check = True
            End Select

            'スイッチ   
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "F3PH" Or _
                objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "F3PV" Or _
                objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PH" Or _
                objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PV" Then
                fncLCRC5Check = True
            End If

            'ストローク調整範囲
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                fncLCRC5Check = True
            End If

            'RM1306001 2013/06/04 追加
            If objKtbnStrc.strcSelection.strKeyKataban = "2" Then
                If objKtbnStrc.strcSelection.strOpSymbol(13).Trim = "SX" Then
                    fncLCRC5Check = True
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSSGC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSSGのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*                                          作成日：2007/05/16      作成者：NII A.Takahashi
    '*【更新履歴】
    '*                                          更新日：2008/04/21      更新者：T.Sato
    '*   ・受付No.RM0802086対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '********************************************************************************************
    Private Function fncSSGC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncSSGC5Check = False

            ' 配管ねじ判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "GN", "NN", "GD", "ND"
                    'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                    If bolJudgeDiv Then
                        fncSSGC5Check = True
                    End If
            End Select

            'スイッチ判定
            If bolJudgeDiv Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    Case "T2YDU", "F3PH", "F3PV", "T3PH", "T3PV"
                        fncSSGC5Check = True
                End Select
            End If

            '二次電池判定
                If objKtbnStrc.strcSelection.strKeyKataban = "4" Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        Case "P4", "P40"
                            fncSSGC5Check = True
                    End Select
                End If
        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncJSK2C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダJSK2のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*                                          作成日：2007/05/25      作成者：NII A.Takahashi
    '********************************************************************************************
    Private Function fncJSK2C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncJSK2C5Check = False

            ' ストローク判定
            If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 700 Then
                fncJSK2C5Check = True
            End If

            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "JSK2-V" Then
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PH" Or _
                objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T3PV" Then
                    fncJSK2C5Check = True
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T3PH" Or _
                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T3PV" Then
                    fncJSK2C5Check = True
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSRL3C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSRL3のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*                                          作成日：2008/01/08      作成者：NII A.Takahashi
    '*【変更履歴】
    '*  ・RM0811134:配管ねじによる判定条件追加
    '********************************************************************************************
    Private Function fncSRL3C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncSRL3C5Check = False

            'RM0912XXX 2009/12/09 Y.Miura ローカル版に合わせる
            If bolJudgeDiv Then
                '配管ねじによる判定
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "N", "G"
                        fncSRL3C5Check = True
                End Select

                'スイッチ判定
                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    Case "M3PH", "M3PV", "T1V", "T1H", "T2H", "T2V", "T3V", "T3H", _
                         "T3PH", "T3PV", "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", _
                         "T0HF", "T0VF", "T0HM", "T0VM", "T0HU", "T0VU", "T2HF", "T2VF", _
                         "T2HM", "T2VM", "T2HU", "T2VU", "T2WHF", "T2WVF", "T2WHM", "T2WVM", _
                         "T2WHU", "T2WVU", "T3HF", "T3VF", "T3PHF", "T3PVF", "T3WHF", "T3WVF", _
                         "T2YDU", "T2YDB", "T2YDG"
                        fncSRL3C5Check = True
                End Select
            End If

            'RM1306001 2013/06/04 追加
            '2013/06/19 修正
            If objKtbnStrc.strcSelection.strKeyKataban = "" Then
                If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "SX" Then
                    fncSRL3C5Check = True
                End If
            End If

            '食品製造工程向け商品
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "F", "H"
                    fncSRL3C5Check = True
            End Select

            'RM1001043 2010/02/22 Y.Miura 二次電池C5加算廃止
            ''RM0907070 2009/08/21 Y.Miura　二次電池対応
            ''P4※二次電池は特注
            'Dim strOpArray() As String
            'Dim intLoopCnt As Integer
            'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
            'For intLoopCnt = 0 To strOpArray.Length - 1
            '    Select Case strOpArray(intLoopCnt).Trim
            '        Case "P4", "P40"
            '            'RM0912XXX 2009/12/09 Y.Miura 二次電池C5加算対応
            '            'If bolJudgeDiv Then
            '            fncSRL3C5Check = True
            '            'End If
            '    End Select
            'Next
        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSSD2C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSSD2のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*                                          作成日：2008/01/08      作成者：NII A.Takahashi
    '********************************************************************************************
    Private Function fncSSD2C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncSSD2C5Check = False

            'C5(ただし加算は販促価格)
            'RM0906034 2009/08/04 Y.Miura
            'RM1001043 2010/02/22 Y.Miura 二次電池のC5加算廃止 
            'P4※二次電池対応はC5
            'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
            'If objKtbnStrc.strcSelection.strFullKataban.IndexOf("P4") >= 0 Then
            '    fncSSD2C5Check = True
            'End If

            If bolJudgeDiv Then
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
                    Case ""
                        'バリエーション①による判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "G5"
                                fncSSD2C5Check = True
                            Case "B", "W", "T1", "O", "G2", "G3"
                                '口径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "125", "140", "160"
                                        fncSSD2C5Check = True
                                End Select
                            Case "T1L"
                                '口径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "80", "100"
                                        fncSSD2C5Check = True
                                End Select
                            Case "G"
                                '口径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16", "125", "140", "160"
                                        fncSSD2C5Check = True
                                End Select
                            Case "G1"
                                '口径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "16", "20", "125", "140", "160"
                                        fncSSD2C5Check = True
                                End Select
                                '2011/01/13 ADD RM1101046(2月VerUP:SSD2シリーズ　チェック区分変更) START--->
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "L4"
                                        fncSSD2C5Check = True
                                End Select
                                '2011/01/13 ADD RM1101046(2月VerUP:SSD2シリーズ　チェック区分変更) <---END
                        End Select

                        'バリエーション③
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F" Then
                            '口径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "125", "140", "160"
                                    fncSSD2C5Check = True
                            End Select

                        End If

                        '配管ねじ、クッション
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "GN", "NN", "GD", "ND"
                                fncSSD2C5Check = True
                        End Select

                        'オプション
                        Dim strOpArray() As String
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt As Integer = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "P5", "P51"
                                    fncSSD2C5Check = True
                            End Select
                        Next

                        'S1
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            Case "T2YDU", "T3PH", "T3PV", "F3PH", "F3PV"
                                fncSSD2C5Check = True
                        End Select
                        'S2
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                            Case "T2YDU", "T3PH", "T3PV", "F3PH", "F3PV"
                                fncSSD2C5Check = True
                        End Select
                        '2010/11/01 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->

                        'RM1306001 2013/06/04 追加
                        If objKtbnStrc.strcSelection.strOpSymbol(22).Trim = "SX" Then
                            fncSSD2C5Check = True
                        End If
                    Case "K"
                        'バリエーション①
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            '2011/01/13 MOD RM1101046(2月VerUP:SSD2シリーズ　チェック区分変更) START--->
                            Case "KG1", "KG2", "KG3", "KG4", "KG5"
                                'Case "KG5"
                                '2011/01/13 MOD RM1101046(2月VerUP:SSD2シリーズ　チェック区分変更) <---END
                                fncSSD2C5Check = True
                        End Select

                        '配管ねじ、クッション
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "GC", "NC", "GN", "NN"
                                fncSSD2C5Check = True
                        End Select

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                            Case "T2YDU", "T3PH", "T3PV"
                                fncSSD2C5Check = True
                        End Select

                        'RM1306001 2013/06/04 追加
                        If objKtbnStrc.strcSelection.strOpSymbol(22).Trim = "SX" Then
                            fncSSD2C5Check = True
                        End If
                        '2010/11/01 DEL RM1011020(12月VerUP:SSD2シリーズ) START--->
                        '    '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END
                        'Case "Q"
                        '    '2010/10/06 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
                        '    '配管ねじ、クッション
                        '    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        '        Case "GN", "NN"
                        '            fncSSD2C5Check = True
                        '    End Select
                        '    '2010/10/06 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END

                        '    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        '        Case "T2YDU", "T3PH", "T3PV"
                        '            fncSSD2C5Check = True
                        '    End Select
                        '2010/11/01 DEL RM1011020(12月VerUP:SSD2シリーズ) <---END
                        '2011/01/13 ADD RM1101046(2月VerUP:SSD2シリーズ　チェック区分変更) START--->
                    Case "D"
                        'バリエーション
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            '2012/12/06 RM1212080 DM追加
                            Case "DG1", "DG4", "DM"
                                fncSSD2C5Check = True
                        End Select
                        '2011/01/13 ADD RM1101046(2月VerUP:SSD2シリーズ　チェック区分変更) <---END
                        'RM1306001 2013/06/04 追加
                        If objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "SX" Then
                            fncSSD2C5Check = True
                        End If
                    Case "4", "L"
                        '2012/01/05 ADD RM1201XXX(12月VerUP:SSD2シリーズ) START--->
                        'オプション
                        Dim strOpArray() As String
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt As Integer = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "P4", "P40"
                                    fncSSD2C5Check = True
                            End Select
                        Next
                        '2012/01/05 ADD RM1201XXX(12月VerUP:SSD2シリーズ) <---END
                    Case "6", "E"
                        '2012/01/05 ADD RM1201XXX(12月VerUP:SSD2シリーズ) START--->
                        'オプション
                        Dim strOpArray() As String
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt As Integer = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "P4", "P40"
                                    fncSSD2C5Check = True
                            End Select
                        Next
                        '2012/01/05 ADD RM1201XXX(12月VerUP:SSD2シリーズ) <---END
                    Case Else
                        '2010/10/06 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
                        '配管ねじ、クッション
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "GN", "NN", "GD", "ND"
                                fncSSD2C5Check = True
                        End Select
                        '2010/10/06 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "T2YDU", "T3PH", "T3PV"
                                fncSSD2C5Check = True
                        End Select
                End Select
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSRT3C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSRT3のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*  RM0811134 SRT3                      作成日：2009/02/03      作成者：T.Yagyu
    '********************************************************************************************
    Private Function fncSRT3C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncSRT3C5Check = False

            'M3PH, M3PVスイッチの場合はC5
            If bolJudgeDiv Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    Case "M3PH", "M3PV"
                        fncSRT3C5Check = True
                End Select
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSRG3C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSRG3のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*  RM0811134 SRG3                      作成日：2009/02/05      作成者：T.Yagyu
    '********************************************************************************************
    Private Function fncSRG3C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncSRG3C5Check = False

            '配管ねじによる判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "N", "G"
                    fncSRG3C5Check = True
            End Select

            If bolJudgeDiv Then
                'M3PH, M3PVスイッチの場合はC5
                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    Case "M3PH", "M3PV"
                        fncSRG3C5Check = True
                End Select
            End If

            'RM1306001 2013/06/04 追加
            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "SX" Then
                fncSRG3C5Check = True
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncUCAC2Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダUCAC2のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2009/08/01      更新者：Y.Miura
    '*   ・受付No.RM0811133対応  チェック区分が『３（Ｃ５）』になる要因がＧネジ、Ｎネジのみの場合
    '*                           販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '********************************************************************************************
    Private Function fncUCAC2C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncUCAC2C5Check = False

            ' 配管ねじ判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "N", "G"
                    'Ｇねじ、Ｎねじの場合はC5(ただし加算は販促価格)
                    If bolJudgeDiv Then
                        fncUCAC2C5Check = True
                    End If
            End Select

            'スイッチT2YDUをC5扱いとする     RM1001018 Y.Miura 追加
            If objKtbnStrc.strcSelection.strOpSymbol.Length > 8 Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(8)
                    Case "T2YDU", "T3PH", "T3PV"
                        If bolJudgeDiv Then
                            fncUCAC2C5Check = True
                        End If
                End Select
            End If

            'スズキ特注
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "R", "S"
                    If objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "S040" Or _
                objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "S050" Then
                        fncUCAC2C5Check = False
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncLiIonC5Check
    '*【処理】
    '*  二次電池のC5チェック
    '*【概要】
    '*  二次電池の場合は販売促進価格を適用して表示のみを『３（Ｃ５）』にする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Integer>       intOptionPos        要素番号
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*  <Boolean>       bolOnlyP40Flg       P40限定フラグ（P4二次電池は対象外）
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2009/09/04      更新者：Y.Miura
    '********************************************************************************************
    Private Function fncLiIonC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByVal intOptionPos As Integer, _
                                   Optional ByVal bolJudgeDiv As Boolean = True, Optional ByVal bolOnlyP40Flg As Boolean = False) As Boolean

        Try

            fncLiIonC5Check = False

            'P4※二次電池は特注
            Dim strOpArray() As String
            Dim intLoopCnt As Integer
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4"
                        If Not bolOnlyP40Flg Then
                            If bolJudgeDiv Then
                                fncLiIonC5Check = True
                            End If
                        End If
                    Case "P40"
                        If bolJudgeDiv Then
                            fncLiIonC5Check = True
                        End If
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncMRL2C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダMRL2のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*                                          作成日：2013/06/04      
    '********************************************************************************************
    Private Function fncMRL2C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncMRL2C5Check = False

            'RM1306001 2013/06/04 追加
            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "SX" Then
                fncMRL2C5Check = True
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSRM3C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダSRM3のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*                                          作成日：2013/06/04      
    '********************************************************************************************
    Private Function fncSRM3C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncSRM3C5Check = False

            'RM1306001 2013/06/04 追加
            '2013/06/19 修正
            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "SX" Then
                    fncSRM3C5Check = True
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncUCA2C5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダUCA2のC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*                                          作成日：2013/06/04      
    '********************************************************************************************
    Private Function fncUCA2C5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncUCA2C5Check = False

            'RM1306001 2013/06/04 追加
            '2013/06/19 修正
            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "SX" Then
                    fncUCA2C5Check = True
                End If
            End If

            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "UCA2-L" Or _
                objKtbnStrc.strcSelection.strSeriesKataban.Trim = "UCA2-BL" Then
                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T3PH" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T3PV" Then
                        fncUCA2C5Check = True
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncNCKC5Check
    '*【処理】
    '*  シリンダC5チェック
    '*【概要】
    '*  シリンダNCKのC5をチェックする
    '*【引数】
    '*  <Object>        objKtbnStrc         引当形番情報
    '*  <Boolean>       bolJudgeDiv         判定フラグ
    '*【戻り値】
    '*  <Boolean>
    '*【作成履歴】
    '*                                          作成日：2013/06/04      
    '********************************************************************************************
    Private Function fncNCKC5Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                    Optional ByVal bolJudgeDiv As Boolean = True) As Boolean

        Try

            fncNCKC5Check = False

            'RM1306001 2013/06/04 追加
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "SX" Then
                fncNCKC5Check = True
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

End Module
