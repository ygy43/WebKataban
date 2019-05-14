'************************************************************************************
'*  ProgramID  ：KHPriceP5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2008/01/09   作成者：NII A.Takahashi
'*
'*  概要       ：スーパーコンパクトシリンダ　ＳＳＤ２
'*
'*【修正履歴】
'*                                      更新日：2008/05/07   更新者：T.Sato
'*  ・受付No：RM0802088対応　バリエーション（'Ｄ','Ｍ','Ｑ','Ｘ','Ｙ'）追加に伴う修正
'* 　　　　　　　　　　　　　特に（'Ｑ'）はボックスが１つ多い点を考慮して修正
'*  ・受付No：RM0906034  二次電池対応機器　SSD2
'*                                      更新日：2009/08/04   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPriceP5

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim intStrokeS1 As Integer = 0      'RM1010017 ADD 
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolC5Flag As Boolean
        Dim intOpAmount As Integer
        Dim intOpAmountBW As Integer
        Dim bolOpP4 As Boolean              'RM0906034 2009/08/04 Y.Miura　二次電池対応

        Dim strVariation As String          'バリエーション
        Dim strSwitchAttached As String     'スイッチ
        Dim strBoreSize As String           '口径
        Dim strCushion As String            '配管ねじ、クッション
        Dim strStroke As String             'ストローク
        Dim strPositionLocking As String    '落下防止位置
        Dim strSwitchModel As String        'スイッチ
        Dim strLeadWireLen As String        'リード線長さ
        Dim strSwitchQty As String          '数
        Dim strLod As String                'ロッド先端

        '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
        Dim strStrokeS1 As String           'ストローク(S1)
        Dim strPositionLockingS1 As String  '落下防止位置(S1)
        Dim strSwitchModelS1 As String      'スイッチ(S1)
        Dim strLeadWireLenS1 As String      'リード線長さ(S1)
        Dim strSwitchQtyS1 As String        '数(S1)
        '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END

        Dim strOption As String             'オプション
        Dim strFP1 As String                '食品製造向け
        Dim strMountingBracket As String    '支持金具
        Dim strAccessory As String          '付属品

        Try



            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            Select Case objKtbnStrc.strcSelection.strKeyKataban
                '2010/11/01 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
                Case "", "K", "L", "4"
                    ''2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
                    'Case ""
                    '2010/11/01 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END
                    strVariation = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         'バリエーション①
                    strSwitchAttached = objKtbnStrc.strcSelection.strOpSymbol(2).Trim    'バリエーション②(スイッチ)
                    strBoreSize = objKtbnStrc.strcSelection.strOpSymbol(4).Trim          '口径
                    strCushion = objKtbnStrc.strcSelection.strOpSymbol(5).Trim           '配管ねじ、クッション
                    strStrokeS1 = objKtbnStrc.strcSelection.strOpSymbol(7).Trim          'Ｓ１：ストローク
                    strPositionLockingS1 = objKtbnStrc.strcSelection.strOpSymbol(8).Trim 'Ｓ１：落下防止位置
                    strSwitchModelS1 = objKtbnStrc.strcSelection.strOpSymbol(9).Trim     'Ｓ１：スイッチ
                    strLeadWireLenS1 = objKtbnStrc.strcSelection.strOpSymbol(10).Trim    'Ｓ１：リード線長さ
                    strSwitchQtyS1 = objKtbnStrc.strcSelection.strOpSymbol(11).Trim      'Ｓ１：数
                    strLod = objKtbnStrc.strcSelection.strOpSymbol(12).Trim              'Ｓ１：ロッド先端
                    strStroke = objKtbnStrc.strcSelection.strOpSymbol(14).Trim           'Ｓ２：ストローク
                    strPositionLocking = objKtbnStrc.strcSelection.strOpSymbol(15).Trim  'Ｓ２：落下防止位置
                    strSwitchModel = objKtbnStrc.strcSelection.strOpSymbol(16).Trim      'Ｓ２：スイッチ
                    strLeadWireLen = objKtbnStrc.strcSelection.strOpSymbol(17).Trim      'Ｓ２：リード線長さ
                    strSwitchQty = objKtbnStrc.strcSelection.strOpSymbol(18).Trim        'Ｓ２：数
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(19).Trim           'オプション
                    strMountingBracket = objKtbnStrc.strcSelection.strOpSymbol(20).Trim  '支持金具
                    strAccessory = objKtbnStrc.strcSelection.strOpSymbol(21).Trim        '付属品
                    strFP1 = ""
                    '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END
                    '2010/11/01 DEL RM1011020(12月VerUP:SSD2シリーズ) START--->
                    'Case "Q"
                    '    strVariation = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        'バリエーション
                    '    strSwitchAttached = objKtbnStrc.strcSelection.strOpSymbol(2).Trim   'スイッチ
                    '    strBoreSize = objKtbnStrc.strcSelection.strOpSymbol(3).Trim         '口径
                    '    strCushion = objKtbnStrc.strcSelection.strOpSymbol(4).Trim          '配管ねじ、クッション
                    '    strStroke = objKtbnStrc.strcSelection.strOpSymbol(5).Trim           'ストローク
                    '    strPositionLocking = objKtbnStrc.strcSelection.strOpSymbol(6).Trim  '落下防止位置
                    '    strSwitchModel = objKtbnStrc.strcSelection.strOpSymbol(7).Trim      'スイッチ
                    '    strLeadWireLen = objKtbnStrc.strcSelection.strOpSymbol(8).Trim      'リード線長さ
                    '    strSwitchQty = objKtbnStrc.strcSelection.strOpSymbol(9).Trim        '数
                    '    strOption = objKtbnStrc.strcSelection.strOpSymbol(10).Trim          'オプション
                    '    strMountingBracket = objKtbnStrc.strcSelection.strOpSymbol(11).Trim '支持金具
                    '    strAccessory = objKtbnStrc.strcSelection.strOpSymbol(12).Trim       '付属品
                    '2010/11/01 DEL RM1011020(12月VerUP:SSD2シリーズ) <---END
                Case "7", "N"
                    strVariation = objKtbnStrc.strcSelection.strOpSymbol(1).Trim         'バリエーション①
                    strSwitchAttached = objKtbnStrc.strcSelection.strOpSymbol(2).Trim    'バリエーション②(スイッチ)
                    strBoreSize = objKtbnStrc.strcSelection.strOpSymbol(4).Trim          '口径
                    strCushion = objKtbnStrc.strcSelection.strOpSymbol(5).Trim           '配管ねじ、クッション
                    strStrokeS1 = objKtbnStrc.strcSelection.strOpSymbol(7).Trim          'Ｓ１：ストローク
                    strPositionLockingS1 = objKtbnStrc.strcSelection.strOpSymbol(8).Trim 'Ｓ１：落下防止位置
                    strSwitchModelS1 = objKtbnStrc.strcSelection.strOpSymbol(9).Trim     'Ｓ１：スイッチ
                    strLeadWireLenS1 = objKtbnStrc.strcSelection.strOpSymbol(10).Trim    'Ｓ１：リード線長さ
                    strSwitchQtyS1 = objKtbnStrc.strcSelection.strOpSymbol(11).Trim      'Ｓ１：数
                    strLod = objKtbnStrc.strcSelection.strOpSymbol(12).Trim              'Ｓ１：ロッド先端
                    strStroke = objKtbnStrc.strcSelection.strOpSymbol(14).Trim           'Ｓ２：ストローク
                    strPositionLocking = objKtbnStrc.strcSelection.strOpSymbol(15).Trim  'Ｓ２：落下防止位置
                    strSwitchModel = objKtbnStrc.strcSelection.strOpSymbol(16).Trim      'Ｓ２：スイッチ
                    strLeadWireLen = objKtbnStrc.strcSelection.strOpSymbol(17).Trim      'Ｓ２：リード線長さ
                    strSwitchQty = objKtbnStrc.strcSelection.strOpSymbol(18).Trim        'Ｓ２：数
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(19).Trim           'オプション
                    strFP1 = objKtbnStrc.strcSelection.strOpSymbol(20).Trim              '食品製造向け
                    strMountingBracket = objKtbnStrc.strcSelection.strOpSymbol(21).Trim  '支持金具
                    strAccessory = objKtbnStrc.strcSelection.strOpSymbol(22).Trim        '付属品
                Case "F"
                    strVariation = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        'バリエーション
                    strSwitchAttached = objKtbnStrc.strcSelection.strOpSymbol(2).Trim   'スイッチ
                    strBoreSize = objKtbnStrc.strcSelection.strOpSymbol(3).Trim         '口径
                    strCushion = objKtbnStrc.strcSelection.strOpSymbol(4).Trim          '配管ねじ、クッション
                    strStroke = objKtbnStrc.strcSelection.strOpSymbol(5).Trim           'ストローク
                    strPositionLocking = ""                                             '落下防止位置
                    strSwitchModel = objKtbnStrc.strcSelection.strOpSymbol(6).Trim      'スイッチ
                    strLeadWireLen = objKtbnStrc.strcSelection.strOpSymbol(7).Trim      'リード線長さ
                    strSwitchQty = objKtbnStrc.strcSelection.strOpSymbol(8).Trim        '数
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(9).Trim           'オプション
                    strFP1 = objKtbnStrc.strcSelection.strOpSymbol(10).Trim             '食品製造向け
                    strMountingBracket = objKtbnStrc.strcSelection.strOpSymbol(11).Trim '支持金具
                    strAccessory = objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                Case Else
                    strVariation = objKtbnStrc.strcSelection.strOpSymbol(1).Trim        'バリエーション
                    strSwitchAttached = objKtbnStrc.strcSelection.strOpSymbol(2).Trim   'スイッチ
                    strBoreSize = objKtbnStrc.strcSelection.strOpSymbol(3).Trim         '口径
                    strCushion = objKtbnStrc.strcSelection.strOpSymbol(4).Trim          '配管ねじ、クッション
                    strStroke = objKtbnStrc.strcSelection.strOpSymbol(5).Trim           'ストローク
                    strPositionLocking = ""                                             '落下防止位置
                    strSwitchModel = objKtbnStrc.strcSelection.strOpSymbol(6).Trim      'スイッチ
                    strLeadWireLen = objKtbnStrc.strcSelection.strOpSymbol(7).Trim      'リード線長さ
                    strSwitchQty = objKtbnStrc.strcSelection.strOpSymbol(8).Trim        '数
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(9).Trim          'オプション
                    strMountingBracket = objKtbnStrc.strcSelection.strOpSymbol(10).Trim '支持金具
                    strAccessory = objKtbnStrc.strcSelection.strOpSymbol(11).Trim       '付属品
                    strFP1 = ""
            End Select

            'RM0906034 2009/08/04 Y.Miura　追加↓↓
            'オプションより二次電池対応か判断する
            bolOpP4 = False
            strOpArray = Split(strOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOpP4 = True
                End Select
            Next
            'RM0906034 2009/08/04 Y.Miura　追加↑↑

            '数量設定
            '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
            intOpAmount = 1
            intOpAmountBW = 1
            '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END

            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "D", "E", "F"
                    intOpAmount = 2
                    '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) START--->
                    '↓2013/09/20 ローカル版との差異修正
                Case "", "4", "7"

                    Select Case Left(strVariation.Trim, 1)
                        Case "B", "W"
                            intOpAmountBW = 2
                    End Select
                    'Case Else
                    '    intOpAmount = 1
                    '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) <---END
            End Select

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            'C5チェック
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "L", "4", "6", "E", "7", "F", "N"
                    bolC5Flag = True
                    '↓RM1306001 2013/06/06 追加
                Case "", "K"
                    If objKtbnStrc.strcSelection.strOpSymbol(22).Trim = "SX" Then
                        bolC5Flag = True
                    End If
                Case "D"
                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "SX" Then
                        bolC5Flag = True
                    End If
            End Select

            'ストローク設定
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(strBoreSize), _
                                                  CInt(strStroke))


            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
                Case "", "4", "7"
                    Select Case Left(strVariation.Trim, 1)
                        Case "B", "W"
                            'ストローク設定(S1)
                            intStrokeS1 = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                                    CInt(strBoreSize), _
                                                                    CInt(IIf(strStrokeS1.Equals(String.Empty), 0, strStrokeS1)))
                            'S1
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       "BASE" & CdCst.Sign.Hypen & strBoreSize & CdCst.Sign.Hypen & intStrokeS1.ToString

                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                    End Select

                    'S2
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               "BASE" & CdCst.Sign.Hypen & strBoreSize & CdCst.Sign.Hypen & intStroke.ToString
                    '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END
                Case "D", "K"
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               "BASE" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strKeyKataban & CdCst.Sign.Hypen & _
                                                               strBoreSize & CdCst.Sign.Hypen & intStroke.ToString
                Case "L", "N"        'RM0906034 2009/08/04 Y.Miura　追加
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               "BASE" & CdCst.Sign.Hypen & "K" & CdCst.Sign.Hypen & _
                                                               strBoreSize & CdCst.Sign.Hypen & intStroke.ToString
                Case "E", "F"
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               "BASE" & CdCst.Sign.Hypen & "D" & CdCst.Sign.Hypen & _
                                                               strBoreSize & CdCst.Sign.Hypen & intStroke.ToString
                Case Else
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               "BASE" & CdCst.Sign.Hypen & strBoreSize & CdCst.Sign.Hypen & intStroke.ToString
            End Select
            decOpAmount(UBound(decOpAmount)) = 1
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If

            'バリエーション加算価格キー
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
                Case "", "4", "7"
                    Select Case strVariation
                        '2010/11/01 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
                        Case "T1", "T1L", "O", "B", "W", "G", "G1", "G4", "G5", "M", "Q"
                            'Case "T1", "T1L", "O", "B", "W", "G", "G1", "G4", "G5"
                            '2010/11/01 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                "-VAL-" & strVariation & CdCst.Sign.Hypen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If

                        Case "G2", "G3"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                "-VAL-" & strVariation & CdCst.Sign.Hypen & strBoreSize & _
                                                                CdCst.Sign.Hypen & intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                    End Select
                    '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END
                    '2010/11/01 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->
                Case "K", "L"
                    Select Case strVariation
                        Case "KU", "KG5"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                "-VAL-" & strVariation & CdCst.Sign.Hypen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case "KG1", "KG4"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                "-VAL-" & Right(strVariation, 2) & CdCst.Sign.Hypen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case "KG2", "KG3"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                "-VAL-" & strVariation & CdCst.Sign.Hypen & strBoreSize & _
                                                                CdCst.Sign.Hypen & intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                    End Select
                Case "D"
                    Select Case strVariation
                        Case "DG1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                "-VAL-" & Right(strVariation, 2) & CdCst.Sign.Hypen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If

                        Case "DG4"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                "-VAL-" & strVariation & CdCst.Sign.Hypen & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case "DM"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                "-VAL-M-" & strBoreSize
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                    End Select
                    '2010/11/01 ADD RM1011020(12月VerUP:SSD2シリーズ) <---END

                    '2010/11/01 DEL RM1011020(12月VerUP:SSD2シリーズ) START--->
                    'Case "M"
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAL-M-" & strBoreSize
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    '    If bolC5Flag = True Then
                    '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    '    End If
                    'Case "Q"
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAL-Q-" & strBoreSize
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    '    If bolC5Flag = True Then
                    '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    '    End If
                    '2010/11/01 DEL RM1011020(12月VerUP:SSD2シリーズ) <---END
            End Select

            '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
            'バリエーション③
            '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
            If (objKtbnStrc.strcSelection.strKeyKataban.Trim = "" OrElse _
                objKtbnStrc.strcSelection.strKeyKataban.Trim = "K" OrElse _
                objKtbnStrc.strcSelection.strKeyKataban.Trim = "L" OrElse _
                objKtbnStrc.strcSelection.strKeyKataban.Trim = "4") _
            AndAlso objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F" Then
                'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" AndAlso _
                'objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F" Then
                '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                '口径判定
                Select Case strBoreSize
                    Case "12", "16", "20"
                        'ストローク(S2)
                        Select Case True
                            Case strStroke <= 15
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-5-15"
                            Case strStroke >= 16 And strStroke <= 30
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-16-30"
                            Case strStroke >= 31
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-31-50"
                        End Select
                    Case "25", "32", "40", "50", "63", "80", "100"
                        'ストローク(S2)
                        Select Case True
                            Case strStroke <= 25
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-5-25"
                            Case strStroke >= 26 And strStroke <= 50
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-26-50"
                            Case strStroke >= 51 And strStroke <= 75
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-51-75"
                            Case strStroke >= 76
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-76-100"
                        End Select
                    Case "125", "140", "160"
                        'ストローク(S2)
                        Select Case True
                            Case strStroke <= 50
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-5-50"
                            Case strStroke >= 51 And strStroke <= 100
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-51-100"
                            Case strStroke >= 101 And strStroke <= 200
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-101-200"
                            Case strStroke >= 201
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                        "-F-" & strBoreSize & "-201-300"
                        End Select

                End Select

                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If

            End If
            '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END

            'スイッチ加算価格キー
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "", "K", "L", "4", "E", "6", "D", "2", "7", "F", "N"
                    If strSwitchAttached <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "SW" & CdCst.Sign.Hypen & strSwitchAttached & CdCst.Sign.Hypen & strBoreSize
                        '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) START--->
                        decOpAmount(UBound(decOpAmount)) = intOpAmountBW
                        'decOpAmount(UBound(decOpAmount)) = 1
                        '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) <---END

                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If
            End Select

            'クッション加算価格キー
            '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) START--->
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "", "K", "L", "4", "E", "D", "7", "F", "N"
                    Select Case strCushion
                        Case "D", "GD", "ND"
                            'If strCushion <> "" Then
                            '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) START--->
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) START--->
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       "OP" & CdCst.Sign.Hypen & Right(strCushion, 1) & CdCst.Sign.Hypen & strBoreSize
                            'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                            '                                           "OP" & CdCst.Sign.Hypen & strCushion & CdCst.Sign.Hypen & strBoreSize
                            '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) START--->
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                            '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->
                        Case "C", "GC", "NC"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim _
                                                                    & CdCst.Sign.Hypen & "K-*C" & CdCst.Sign.Hypen & strBoreSize

                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                            '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) <---END
                    End Select
            End Select

            'スイッチ加算価格キー
            '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "", "K", "L", "4", "7", "N"
                    ''2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                    '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END
                    If strSwitchModelS1 <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "SW" & CdCst.Sign.Hypen & strSwitchModelS1
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchQtyS1)

                        '↓2013/09/20 ローカル版と差異修正
                        If bolOpP4 Then  'P4
                            'スイッチ加算
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       "SW" & CdCst.Sign.Hypen & "P4"
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchQty)
                        End If
                    End If
            End Select
            '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END

            If strSwitchModel <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           "SW" & CdCst.Sign.Hypen & strSwitchModel
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchQty)

                'RM0906034 2009/08/04 Y.Miura　二次電池対応追加↓↓
                If bolOpP4 Then  'P4
                    'スイッチ加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               "SW" & CdCst.Sign.Hypen & "P4"
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchQty)
                End If
                'RM0906034 2009/08/04 Y.Miura　二次電池対応追加↑↑
            End If

            'リード線長さ加算価格キー
            '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "", "K", "L", "4", "7", "N"
                    ''2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                    '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END
                    If strSwitchModelS1 <> "" AndAlso strLeadWireLenS1 <> "" Then

                        '2010/11/17 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
                        Dim strKataban As String = ""
                        'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '2010/11/17 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END

                        Select Case strSwitchModelS1
                            'RM1307003 2013/07/04追加(F2S,F3S)
                            Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T2WH", "T2WV", _
                                 "T3H", "T3V", "T3YH", "T3YV", "T3WH", "T3WV", _
                                 "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", "T2JH", "T2JV", "T3PH", "T3PV", _
                                 "F2H", "F2V", "F3H", "F3V", "F2YH", "F2YV", "F3YH", "F3YV", "F2S", "F3S"
                                strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "SWLW(1)" & CdCst.Sign.Hypen & strLeadWireLenS1
                            Case "T2YD"
                                strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "SWLW(2)" & CdCst.Sign.Hypen & strLeadWireLenS1
                            Case "T2YDT"
                                strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "SWLW(3)" & CdCst.Sign.Hypen & strLeadWireLenS1
                            Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                                strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "SWLW(7)" & CdCst.Sign.Hypen & strLeadWireLenS1
                            Case "V0", "V7"
                                strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "SWLW(8)" & CdCst.Sign.Hypen & strLeadWireLenS1
                        End Select

                        '2010/11/17 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
                        If strKataban.Trim.Length > 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchQtyS1)

                        End If
                        '2010/11/17 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END

                    End If
            End Select
            '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END

            If strSwitchModel <> "" Then
                If strLeadWireLen <> "" Then
                    '2010/11/17 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
                    Dim strKataban As String = ""
                    'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    'ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "", "K"
                            ''2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) START--->
                            'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                            '2010/11/17 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END
                            Select Case strSwitchModel
                                'RM1307003 2013/07/04追加(F2S,F3S)
                                Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T2WH", "T2WV", _
                                     "T3H", "T3V", "T3YH", "T3YV", "T3WH", "T3WV", _
                                     "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", "T2JH", "T2JV", "T3PH", "T3PV", _
                                     "F2H", "F2V", "F3H", "F3V", "F2YH", "F2YV", "F3YH", "F3YV", "F2S", "F3S"
                                    strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "SWLW(1)" & CdCst.Sign.Hypen & strLeadWireLen
                                Case "T2YD"
                                    strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "SWLW(2)" & CdCst.Sign.Hypen & strLeadWireLen
                                Case "T2YDT"
                                    strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "SWLW(3)" & CdCst.Sign.Hypen & strLeadWireLen
                                Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                                    strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "SWLW(7)" & CdCst.Sign.Hypen & strLeadWireLen
                                Case "V0", "V7"
                                    strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "SWLW(8)" & CdCst.Sign.Hypen & strLeadWireLen
                            End Select

                        Case Else
                            Select Case strSwitchModel
                                'RM1307003 2013/07/04追加(F2S,F3S)
                                Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T2WH", "T2WV", _
                                     "T3H", "T3V", "T3YH", "T3YV", "T3WH", "T3WV", _
                                     "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", "T2JH", "T2JV", "T3PH", "T3PV", _
                                     "F2H", "F2V", "F3H", "F3V", "F2YH", "F2YV", "F3YH", "F3YV", "F2S", "F3S"
                                    strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "SWLW(1)" & CdCst.Sign.Hypen & strLeadWireLen
                                Case "T2YD"
                                    strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "SWLW(2)" & CdCst.Sign.Hypen & strLeadWireLen
                                Case "T2YDT"
                                    strKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "SWLW(3)" & CdCst.Sign.Hypen & strLeadWireLen
                            End Select
                    End Select

                    '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) <---END

                    '2010/11/17 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
                    If strKataban.Trim.Length > 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strKataban

                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchQty)
                    End If
                    '2010/11/17 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END
                End If
            End If

            'オプション加算価格キー
            strOpArray = Split(strOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "", "K", "L", "4", "7", "N"
                        ''2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
                        'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "P6"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                            Case "M"
                                Select Case strBoreSize
                                    Case "12", "16", "20", "25"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                   strBoreSize & CdCst.Sign.Hypen & intStroke.ToString

                                        decOpAmount(UBound(decOpAmount)) = intOpAmountBW
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case "32", "40", "50", "63", "80", "100", "125", "140", "160"
                                        Select Case Left(strVariation.Trim, 1)
                                            Case "B", "W"
                                                'S1
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                           "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                           strBoreSize & CdCst.Sign.Hypen & intStrokeS1.ToString

                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If

                                        End Select

                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                   strBoreSize & CdCst.Sign.Hypen & intStroke.ToString

                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "P5", "P51", "P7", "P71"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "OP" & CdCst.Sign.Hypen & Left(strOpArray(intLoopCnt).Trim, 2) & "*" & _
                                                                           CdCst.Sign.Hypen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If

                                '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->
                            Case "M0", "M1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                                '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) <---END
                                '2010/12/10 ADD RM1012055(1月VerUP:SSD2シリーズ) START--->
                            Case "S"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If

                                '2010/12/10 ADD RM1012055(1月VerUP:SSD2シリーズ) <---END
                            Case "P4", "P40"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = intOpAmount

                        End Select

                    Case Else
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "P6"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算不要
                                'If bolC5Flag = True Then
                                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                'End If
                                'RM0906034 2009/08/04 Y.Miura　二次電池対応追加↓↓
                            Case "P4", "P40"
                                Select Case objKtbnStrc.strcSelection.strKeyKataban
                                    Case "E"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "OP-D" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize
                                        decOpAmount(UBound(decOpAmount)) = intOpAmount
                                End Select
                                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算不要
                                'If bolC5Flag = True Then
                                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                'End If
                                'RM0906034 2009/08/04 Y.Miura　二次電池対応追加↑↑
                            Case "M"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           strBoreSize & CdCst.Sign.Hypen & intStroke.ToString
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "M0", "M1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "OP" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intOpAmount
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                End Select
                '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) <---END
            Next

            'FP加算価格キー
            If strFP1 <> "" Then

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                           strFP1 & CdCst.Sign.Hypen & strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If

            End If

            '支持金具加算価格キー
            If strMountingBracket <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           strMountingBracket & CdCst.Sign.Hypen & strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算不要
                'If bolC5Flag = True Then
                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                'End If
            End If

            '付属品加算価格キー
            If strAccessory <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           strAccessory & CdCst.Sign.Hypen & strBoreSize

                'RMXXXXXXX 2009/09/11 Y.Miura 付属品の数量がゼロになる不具合修正
                decOpAmount(UBound(decOpAmount)) = intOpAmount
                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算不要
                'If bolC5Flag = True Then
                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                'End If

            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
