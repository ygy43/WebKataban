Module KHWaterValveCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  流体制御バルブチェック
    '*【概要】
    '*  流体制御バルブをチェックする
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

        Dim strOpArray() As String = Nothing
        Dim intLoopCnt As Integer = Nothing

        Try

            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "GAB412", "GAB452"
                    Dim intOptionPos As Integer = 10
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも230まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim) - 3)) > 231 Then
                            intKtbnStrcSeqNo = intOptionPos
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                     If bolOptionZ = True Then
                        'コイルハウジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "3A", "3M", "3I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                         "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC235V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "DC12V", "DC24V", "DC100V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    Else
                        ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case ""
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                         "AC415V", "AC420V", "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", "AC210V", _
                                         "AC215V", "AC216V", "AC220V", "AC225V", "DC6V", "DC8V", "DC12V", "DC21V", "DC24V", _
                                         "DC28V", "DC30V", "DC33V", "DC45V", "DC48V", "DC50V", "DC70V", "DC85V", "DC90V", _
                                         "DC100V", "DC110V", "DC120V", "DC125V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"

                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V", "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", _
                                         "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", _
                                         "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "DC6V", "DC8V", "DC12V", _
                                         "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC33V", "DC34V", "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", _
                                         "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "DC12V", "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC100V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                         "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                         "AC200V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If

                Case "GAB422"

                    Dim intOptionPos As Integer = 9
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = intOptionPos
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    ' コイルハイジング判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case ""
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2E", "2G"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "DC12V", "DC24V", "DC48V", "DC100V"
                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2H"
                            If bolOptionS = True Then
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Else
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            End If
                        Case "3A"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                     "AC575V", "AC578V", "DC6V", "DC9V", "DC12V", "DC14V", "DC24V", "DC26V", _
                                     "DC30V", "DC36V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V", "DC200V", "DC220V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3M", "3I"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                     "DC6V", "DC9V", "DC12V", "DC14V", "DC24V", "DC26V", _
                                     "DC30V", "DC36V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V", "DC200V", "DC220V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3N", "3J"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                     "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "DC12V", "DC14V", "DC24V", "DC100V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4A"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                     "AC575V", "AC578V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4M"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4N"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                     "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5A", "5M", "5I"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC125V", "AC200V", "AC220V", "AC230V", _
                                     "AC240V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5N", "5J"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC125V", "AC200V", "AC220V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                    End Select

                Case "GAB462"

                    Dim intOptionPos As Integer = 9
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = intOptionPos
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    ' コイルハイジング判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case ""
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2E", "2G"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "DC12V", "DC24V", "DC48V", "DC100V"
                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2H"
                            If bolOptionS = True Then
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Else
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            End If
                        Case "3A"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                     "AC575V", "AC578V", "DC6V", "DC9V", "DC12V", "DC14V", "DC24V", "DC26V", _
                                     "DC30V", "DC36V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V", "DC200V", "DC220V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3M", "3I"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                     "DC6V", "DC9V", "DC12V", "DC14V", "DC24V", "DC26V", _
                                     "DC30V", "DC36V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V", "DC200V", "DC220V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3N", "3J"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                     "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "DC12V", "DC14V", "DC24V", "DC100V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4A"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                     "AC575V", "AC578V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4M"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4N"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                     "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5A", "5M", "5I"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC125V", "AC200V", "AC220V", "AC230V", _
                                     "AC240V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5N", "5J"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC125V", "AC200V"

                                Case Else
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                    End Select

                Case "GAB312", "GAB352"

                    Dim intOptionPos As Integer
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    'RM1004012 2010/04/22 Y.Miura
                    '電圧要素位置をセット
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "GAB312", "GAB352"
                            intOptionPos = 10
                        Case Else
                            intOptionPos = 9
                    End Select

                    ' サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = intOptionPos
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
                    
                    'ドライエア用Ｚ選択時
                    If bolOptionZ = True Then
                        'コイルハウジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "3A", "3K", "3P"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC48V", "DC100V", _
                                         "DC110V", "DC200V", "DC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3H", "3Q"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "DC24V", "DC100V", "DC110V", "DC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                         "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC235V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC12V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    Else
                        ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case ""
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "DC5V", "DC6V", "DC12V", _
                                         "DC14V", "DC24V", "DC25V", "DC28V", "DC48V", _
                                         "DC74V", "DC85V", "DC88V", "DC90V", "DC100V", _
                                         "DC110V", "DC120V", "DC124V", "DC125V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS = True Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V"
                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V", "DC24V"
                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", _
                                         "AC500V", "AC575V", "AC600V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "DC12V", _
                                         "DC13V", "DC14V", "DC17V", "DC21V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", _
                                         "AC500V", "AC575V", "AC600V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select

                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5I", "5M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V", _
                                         "AC220V", "AC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If
                Case "GAG31", "GAG33", "GAG34", "GAG35"
                    Dim intOptionPos As Integer = 10
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = intOptionPos
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    'ドライエア用Ｚ選択時
                    If bolOptionZ = True Then
                        'コイルハウジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "3A", "3M", "3I"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                         "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC235V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC12V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    Else
                        ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case ""
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "DC5V", "DC6V", "DC12V", _
                                         "DC14V", "DC24V", "DC25V", "DC28V", "DC48V", _
                                         "DC74V", "DC85V", "DC88V", "DC90V", "DC100V", _
                                         "DC110V", "DC120V", "DC124V", "DC125V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS = True Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V"
                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V", "DC24V"
                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", _
                                         "AC500V", "AC575V", "AC600V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "DC12V", _
                                         "DC13V", "DC14V", "DC17V", "DC21V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", _
                                         "AC500V", "AC575V", "AC600V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select

                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5I", "5M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V", _
                                         "AC220V", "AC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If

                Case "GAG41", "GAG43", "GAG44", "GAG45"
                    Dim intOptionPos As Integer = 10
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも231まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim) - 3)) > 231 Then
                            intKtbnStrcSeqNo = intOptionPos
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    'ドライエア用Ｚ選択時
                    If bolOptionZ = True Then
                        'コイルハウジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "3A", "3M", "3I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                         "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC235V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "DC12V", "DC24V", "DC100V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    Else
                        ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case ""
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                         "AC415V", "AC420V", "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", "AC210V", _
                                         "AC215V", "AC216V", "AC220V", "AC225V", "DC6V", "DC8V", "DC12V", "DC21V", "DC24V", _
                                         "DC28V", "DC30V", "DC33V", "DC45V", "DC48V", "DC50V", "DC70V", "DC85V", "DC90V", _
                                         "DC100V", "DC110V", "DC120V", "DC125V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"

                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V", "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", _
                                         "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", _
                                         "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "DC6V", "DC8V", "DC12V", _
                                         "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC33V", "DC34V", "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", _
                                         "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "DC12V", "DC14V", "DC16V", _
                                         "DC20V", "DC21V", "DC24V", "DC100V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                         "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                         "AC200V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If

                Case "PVS"
                    Dim bolOptionL As Boolean = False
                    Dim bolOptionL1 As Boolean = False
                    Dim bolOptionL3 As Boolean = False

                    'オプション分解＆抽出
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "L"
                                bolOptionL = True
                            Case "L1"
                                bolOptionL1 = True
                            Case "L3"
                                bolOptionL3 = True
                        End Select
                    Next

                    If bolOptionL = True Or bolOptionL1 = True Then
                        '電圧判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "AC220V", "DC12V", "DC48V"
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8390"
                                fncCheckSelectOption = False
                        End Select
                    End If

                    If bolOptionL3 = True Then
                        '電圧判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "DC12V", "DC48V"
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8390"
                                fncCheckSelectOption = False
                        End Select
                    End If

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "B3", "L3"
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) = "AC" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W8380"
                                    fncCheckSelectOption = False
                                End If
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W8370"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "15A", "20A", "25A", "32A", "40A"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", _
                                         "AC120V", "AC125V", "AC200V", "AC220V", "AC230V", _
                                         "AC240V", "AC250V", "AC380V", "AC400V", "AC440V", _
                                         "DC12V", "DC24V", "DC26V", "DC48V", "DC100V", _
                                         "DC110V", "DC125V", "DC200V", "DC220V", "DC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "AC12V", "AC24V", "AC48V", "AC90V", "AC100V", _
                                         "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", _
                                         "AC200V", "AC210V", "AC220V", "AC230V", "AC235V", _
                                         "AC240V", "AC250V", "AC380V", "AC400V", "AC415V", _
                                         "AC440V", "DC6V", "DC12V", "DC20V", "DC22V", _
                                         "DC24V", "DC26V", "DC30V", "DC48V", "DC95V", _
                                         "DC100V", "DC110V", "DC125V", "DC200V", "DC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            End If
                        Case "50A", "65A", "80A"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "AC12V", "AC24V", "AC48V", "AC100V", "AC110V", _
                                         "AC115V", "AC120V", "AC200V", "AC210V", "AC220V", _
                                         "AC230V", "AC240V", "AC380V", "AC400V", "AC415V", _
                                         "AC440V", "DC6V", "DC12V", "DC24V", "DC48V", _
                                         "DC88V", "DC100V", "DC110V", "DC125V", "DC200V", _
                                         "DC220V", "DC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "AC12V", "AC24V", "AC48V", "AC100V", "AC110V", _
                                         "AC115V", "AC120V", "AC200V", "AC210V", "AC220V", _
                                         "AC230V", "AC240V", "AC380V", "AC400V", "AC415V", _
                                         "AC440V", "DC6V", "DC12V", "DC24V", "DC48V", _
                                         "DC88V", "DC100V", "DC110V", "DC125V", "DC200V", _
                                         "DC220V", "DC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = 6
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            End If
                    End Select
                Case "PKA"
                    Dim bolOptionL As Boolean = False
                    Dim bolOptionL1 As Boolean = False
                    Dim bolOptionL3 As Boolean = False

                    'オプション分解＆抽出
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "L"
                                bolOptionL = True
                            Case "L1"
                                bolOptionL1 = True
                            Case "L3"
                                bolOptionL3 = True
                        End Select
                    Next

                    If bolOptionL = True Or bolOptionL1 = True Then
                        '電圧判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "AC220V", "DC12V", "DC48V"
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8390"
                                fncCheckSelectOption = False
                        End Select
                    End If

                    If bolOptionL3 = True Then
                        '電圧判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "DC12V", "DC48V"
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8390"
                                fncCheckSelectOption = False
                        End Select
                    End If

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "B3", "L3"
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) = "AC" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W8380"
                                    fncCheckSelectOption = False
                                End If
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W8370"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select
                Case "PKS"
                    Dim bolOptionL As Boolean = False
                    Dim bolOptionL1 As Boolean = False
                    Dim bolOptionL3 As Boolean = False

                    'オプション分解＆抽出
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "L"
                                bolOptionL = True
                            Case "L1"
                                bolOptionL1 = True
                            Case "L3"
                                bolOptionL3 = True
                        End Select
                    Next

                    If bolOptionL = True Or bolOptionL1 = True Then
                        '電圧判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "AC220V", "DC12V", "DC48V"
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8390"
                                fncCheckSelectOption = False
                        End Select
                    End If

                    If bolOptionL3 = True Then
                        '電圧判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "DC12V", "DC48V"
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8390"
                                fncCheckSelectOption = False
                        End Select
                    End If

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "B3", "L3"
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = "AC" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8380"
                                    fncCheckSelectOption = False
                                End If
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8370"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select
                Case "PKW"
                    Dim bolOptionL As Boolean = False
                    Dim bolOptionL1 As Boolean = False
                    Dim bolOptionL3 As Boolean = False

                    'オプション分解＆抽出
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "L"
                                bolOptionL = True
                            Case "L1"
                                bolOptionL1 = True
                            Case "L3"
                                bolOptionL3 = True
                        End Select
                    Next

                    If bolOptionL = True Or bolOptionL1 = True Then
                        '電圧判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "AC220V", "DC12V", "DC48V"
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8390"
                                fncCheckSelectOption = False
                        End Select
                    End If

                    If bolOptionL3 = True Then
                        '電圧判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "DC12V", "DC48V"
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W8390"
                                fncCheckSelectOption = False
                        End Select
                    End If

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "B3", "L3"
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) = "AC" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W8380"
                                    fncCheckSelectOption = False
                                End If
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    intKtbnStrcSeqNo = 7
                                    strMessageCd = "W8370"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select

                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "M" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "AC100V" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "AC200V" Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8360"
                            fncCheckSelectOption = False
                        End If
                    End If
                Case "PDV2"
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Length = 0 Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "4A" Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "DC" Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
                Case "PDV3"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "4A" Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = "DC" Then
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "65A" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "2E" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "DC24V" Then
                                intKtbnStrcSeqNo = 4
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                Case "NP13", "NP14", "NVP11"
                    'RM1004012 2010/04/22 Y.Miura
                    '電圧要素位置をセット
                    Dim intOpPosition As Integer
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "NP13", "NP14", "NVP11"
                            intOpPosition = 5
                        Case Else
                            intOpPosition = 6
                    End Select
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                        Case "1", "2", "3"
                        Case "AC100V"
                        Case "AC200V"
                        Case "DC24V"
                        Case Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "2C"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", _
                                             "AC120V", "AC125V", "AC127V", "AC200V", "AC215V", _
                                             "AC220V", "AC230V", "AC380V", "DC12V", "DC24V", _
                                             "DC45V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2CS"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", "DC24V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2G"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", _
                                             "AC120V", "AC125V", "AC127V", "AC200V", "AC215V", _
                                             "AC220V", "AC230V", "DC12V", "DC24V", "DC45V", _
                                              "DC48V", "DC85V", "DC100V", "DC110V", "DC125V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2GS"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", _
                                             "AC125V", "AC127V", "AC200V", "AC215V", "AC220V", _
                                             "AC230V", "DC12V", "DC24V", "DC85V", "DC100V", _
                                             "DC110V", "DC125V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2H"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                             "AC200V", "AC215V", "AC220V", "AC230V", "DC24V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2HS"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                             "AC200V", "AC215V", "AC220V", "AC230V", "DC12V", "DC24V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3T"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", _
                                             "AC120V", "AC125V", "AC127V", "AC200V", "AC215V", _
                                             "AC220V", "AC230V", "AC380V", "DC12V", "DC24V", _
                                             "DC45V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3TS"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", "DC24V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3R"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                             "AC127V", "AC200V", "DC12V", "DC24V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3RS"
                                    ' 製作可能電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOpPosition).Trim
                                        Case "AC100V", "AC110V", "AC200V", "DC24V"
                                            ' 製作可能
                                        Case Else
                                            intKtbnStrcSeqNo = intOpPosition
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select
                    End Select
                Case "CVSE2", "CVSE3"
                    '↓RM1308014 2013/08/05 修正
                    If objKtbnStrc.strcSelection.strSeriesKataban = "CVSE2" Then
                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Or _
                           objKtbnStrc.strcSelection.strKeyKataban.Trim = "1" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "ST" And _
                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "R" Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W8970"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If

                    Dim bolOptionS As Boolean = False
                    'その他オプション分解＆サージキラーチェック
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    If objKtbnStrc.strcSelection.strSeriesKataban = "CVSE3" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "2K"
                                If bolOptionS = False Then
                                    intKtbnStrcSeqNo = 6
                                    strMessageCd = "W8600"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If

                Case "AD11"

                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも230まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) - 3)) > 231 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                        Case "8A", "10A"
                            ' コイルハイジング判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "2C"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                             "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC50V", _
                                             "AC85V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                             "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", _
                                             "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", _
                                             "AC300V", "AC346V", "AC350V", "AC365V", "AC370V", "AC380V", _
                                             "AC400V", "AC415V", "AC420V", "AC430V", "AC433V", "AC440V", _
                                             "AC450V", "AC460V", "AC480V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2E", "2G"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", "AC38V", "AC39V", _
                                             "AC42V", "AC48V", "AC50V", "AC80V", "AC90V", "AC95V", "AC100V", _
                                             "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", "AC220V", _
                                             "AC225V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2H"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", "AC38V", "AC39V", _
                                             "AC42V", "AC48V", "AC50V", "AC80V", "AC90V", "AC95V", "AC100V", _
                                             "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", _
                                             "AC346V", "AC350V", "AC360V", "AC365V", "AC380V", "AC400V", _
                                             "AC415V", "AC420V", "AC433V", "AC440V", "AC450V", "AC460V", _
                                             "AC480V", "AC500V", "AC575V", "AC600V", "DC4V", "DC5V", "DC6V", _
                                             "DC12V", "DC13V", "DC14V", "DC17V", "DC21V", "DC24V", "DC25V", _
                                             "DC26V", "DC28V", "DC30V", "DC42V", "DC48V", "DC50V", _
                                             "DC85V", "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", "DC124V", _
                                             "DC125V", "DC140V", "DC200V", "DC220V", "DC230V", "DC240V", "DC300V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3M", "3I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", "AC38V", "AC39V", _
                                             "AC42V", "AC48V", "AC50V", "AC80V", "AC90V", "AC95V", "AC100V", _
                                             "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", _
                                             "AC346V", "AC350V", "AC360V", "AC365V", "AC380V", "AC400V", "DC4V", "DC5V", "DC6V", _
                                             "DC12V", "DC13V", "DC14V", "DC17V", "DC21V", "DC24V", "DC25V", _
                                             "DC26V", "DC28V", "DC30V", "DC42V", "DC48V", "DC50V", _
                                             "DC85V", "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", "DC124V", _
                                             "DC125V", "DC140V", "DC200V", "DC220V", "DC230V", "DC240V", "DC300V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3N", "3J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC200V", "DC12V", "DC24V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", "AC38V", "AC39V", _
                                             "AC42V", "AC48V", "AC50V", "AC80V", "AC90V", "AC95V", "AC100V", _
                                             "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", _
                                             "AC346V", "AC350V", "AC360V", "AC365V", "AC380V", "AC400V", _
                                             "AC415V", "AC420V", "AC433V", "AC440V", "AC450V", "AC460V", _
                                             "AC480V", "AC500V", "AC575V", "AC600V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4M"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", "AC38V", "AC39V", _
                                             "AC42V", "AC48V", "AC50V", "AC80V", "AC90V", "AC95V", "AC100V", _
                                             "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", _
                                             "AC346V", "AC350V", "AC360V", "AC365V", "AC380V", "AC400V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4N"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC200V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5A", "5M", "5I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC125V", _
                                             "AC200V", "AC220V", "AC240V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5N", "5J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC200V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select

                        Case "15A", "20A", "25A"

                            ' コイルハイジング判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "2C"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                             "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2E", "2G"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "DC6V", "DC8V", "DC12V", "DC21V", "DC24V", "DC28V", "DC30V", _
                                             "DC33V", "DC45V", "DC48V", "DC50V", "DC70V", "DC85V", "DC90V", "DC100V", _
                                             "DC110V", "DC120V", "DC125V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2H"
                                    If bolOptionS = True Then
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                                 "AC210V", "AC215V", "AC216V", "AC220V"
                                            Case Else
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    Else
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                                 "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"
                                            Case Else
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    End If
                                Case "3A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                             "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                             "AC575V", "AC578V", "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                             "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", "DC42V", _
                                             "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", _
                                             "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", _
                                             "DC240V", "DC250V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3M", "3I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                             "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                             "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", "DC42V", _
                                             "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", _
                                             "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", _
                                             "DC240V", "DC250V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3N", "3J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                             "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V",
                                             "DC12V", "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC100V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                             "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                             "AC575V", "AC578V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4M"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4N"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                             "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5A", "5M", "5I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                             "AC140V", "AC200V", "AC220V", "AC240V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5N", "5J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                             "AC140V", "AC200V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select

                        Case Else

                    End Select

                Case "AP11"

                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも230まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) - 3)) > 231 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                        Case "8A", "10A"
                            ' コイルハイジング判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "2C"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC25V", "AC30V", _
                                             "AC38V", "AC42V", "AC39V", "AC48V", "AC50V", _
                                                 "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                                 "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                                 "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                                 "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                                 "AC220V", "AC225V", "AC230V", "AC240V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2E", "2G"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "DC5V", "DC6V", "DC12V", _
                                         "DC14V", "DC24V", "DC25V", "DC28V", "DC48V", _
                                         "DC74V", "DC85V", "DC88V", "DC90V", "DC100V", _
                                         "DC110V", "DC120V", "DC124V", "DC125V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2H"
                                    If bolOptionS = True Then
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                                 "AC105V", "AC208V", "AC210V", "AC216V"
                                            Case Else
                                                intKtbnStrcSeqNo = 10
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    Else
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                                 "AC105V", "AC208V", "AC210V", "AC216V", "DC24V"
                                            Case Else
                                                intKtbnStrcSeqNo = 10
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    End If
                                Case "3A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", "AC38V", "AC39V", _
                                             "AC42V", "AC48V", "AC50V", "AC80V", "AC90V", "AC95V", "AC100V", _
                                             "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", _
                                             "AC346V", "AC350V", "AC360V", "AC365V", "AC380V", "AC400V", _
                                             "AC415V", "AC420V", "AC433V", "AC440V", "AC450V", "AC460V", _
                                             "AC480V", "AC500V", "AC575V", "AC600V", "DC4V", "DC5V", "DC6V", _
                                             "DC12V", "DC13V", "DC14V", "DC17V", "DC21V", "DC24V", "DC25V", _
                                             "DC26V", "DC28V", "DC30V", "DC42V", "DC48V", "DC50V", _
                                             "DC85V", "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", "DC124V", _
                                             "DC125V", "DC140V", "DC200V", "DC220V", "DC230V", "DC240V", "DC300V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3M", "3I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", "AC38V", "AC39V", _
                                             "AC42V", "AC48V", "AC50V", "AC80V", "AC90V", "AC95V", "AC100V", _
                                             "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", _
                                             "AC346V", "AC350V", "AC360V", "AC365V", "AC380V", "AC400V", "DC4V", "DC5V", "DC6V", _
                                             "DC12V", "DC13V", "DC14V", "DC17V", "DC21V", "DC24V", "DC25V", _
                                             "DC26V", "DC28V", "DC30V", "DC42V", "DC48V", "DC50V", _
                                             "DC85V", "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", "DC124V", _
                                             "DC125V", "DC140V", "DC200V", "DC220V", "DC230V", "DC240V", "DC300V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3N", "3J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                             "AC150V", "AC160V", "AC190V", "AC200V", "DC12V", _
                                             "DC16V", "DC14V", "DC20V", "DC21V", "DC24V", "DC100V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", "AC38V", "AC39V", _
                                             "AC42V", "AC48V", "AC50V", "AC80V", "AC90V", "AC95V", "AC100V", _
                                             "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", _
                                             "AC346V", "AC350V", "AC360V", "AC365V", "AC380V", "AC400V", _
                                             "AC415V", "AC420V", "AC433V", "AC440V", "AC450V", "AC460V", _
                                             "AC480V", "AC500V", "AC575V", "AC600V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4M"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", "AC38V", "AC39V", _
                                             "AC42V", "AC48V", "AC50V", "AC80V", "AC90V", "AC95V", "AC100V", _
                                             "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                             "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", _
                                             "AC346V", "AC350V", "AC360V", "AC365V", "AC380V", "AC400V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4N"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC200V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5A", "5M", "5I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC125V", _
                                             "AC200V", "AC220V", "AC240V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5N", "5J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC200V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select

                        Case "15A", "20A", "25A"

                            ' コイルハイジング判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "2C"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                             "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2E", "2G"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "DC6V", "DC8V", "DC12V", "DC21V", "DC24V", "DC28V", "DC30V", _
                                             "DC33V", "DC45V", "DC48V", "DC50V", "DC70V", "DC85V", "DC90V", "DC100V", _
                                             "DC110V", "DC120V", "DC125V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2H"
                                    If bolOptionS = True Then
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                                 "AC210V", "AC215V", "AC216V", "AC220V"
                                            Case Else
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    Else
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                                 "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"
                                            Case Else
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    End If
                                Case "3A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                             "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                             "AC575V", "AC578V", "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                             "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", "DC42V", _
                                             "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", _
                                             "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", _
                                             "DC240V", "DC250V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3M", "3I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                             "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                             "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", "DC42V", _
                                             "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", _
                                             "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", _
                                             "DC240V", "DC250V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3N", "3J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                             "AC117V", "AC120V", "AC125V", "AC127V", _
                                             "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "DC12V", "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC100V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                             "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                             "AC575V", "AC578V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4M"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4N"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                             "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5A", "5M", "5I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                             "AC200V", "AC220V", "AC240V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5N", "5J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                             "AC200V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select

                        Case Else

                    End Select

                Case "ADK11"
                    'RM1004012 2010/05/01 Y.Miura 電圧の位置を5→6に変更
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        Dim bolOptionZ As Boolean = False
                        Dim bolOptionS As Boolean = False
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "Z"
                                    bolOptionZ = True
                                Case "S"
                                    bolOptionS = True
                            End Select
                        Next


                        If bolOptionZ = True Then
                            'コイルハウジング判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "3A", "3M", "3I"
                                    '電圧判定
                                    'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                             "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                             "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                             "DC125V", "DC200V", "DC220V", "DC235V"
                                        Case Else
                                            'intKtbnStrcSeqNo = 9
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3N", "3J"
                                    '電圧判定
                                    'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        Case "DC12V", "DC24V", "DC100V"
                                        Case Else
                                            'intKtbnStrcSeqNo = 9
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5A", "5M", "5I"
                                    '電圧判定
                                    'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                        Case Else
                                            'intKtbnStrcSeqNo = 9
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5N", "5J"
                                    '電圧判定
                                    'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                        Case Else
                                            'intKtbnStrcSeqNo = 9
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select

                            'サージキラー付はAC/DCとも230まで
                            If bolOptionS = True Then
                                If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) - 3)) > 231 Then
                                    intKtbnStrcSeqNo = 6
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                                End If
                            End If

                        Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                                Case "8A", "10A"

                                    'サージキラー付はAC/DCとも230まで
                                    If bolOptionS = True Then
                                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) - 3)) > 231 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                        End If
                                    End If

                                    ' コイルハイジング判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "2C"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                                     "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                                     "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                                     "AC130V", "AC135V", "AC140V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", _
                                                     "AC210V", "AC215V", "AC216V", "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                                     "AC260V", "AC300V", "AC346V", "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                                     "AC415V", "AC420V", "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "2E", "2G"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                                     "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                                     "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                                     "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", "AC210V", _
                                                     "AC215V", "AC216V", "AC220V", "AC225V", "DC6V", "DC8V", "DC12V", "DC21V", "DC24V", _
                                                     "DC28V", "DC30V", "DC33V", "DC45V", "DC48V", "DC50V", "DC70V", "DC85V", "DC90V", _
                                                     "DC100V", "DC110V", "DC120V", "DC125V"
                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "2H"
                                            If bolOptionS Then
                                                ' 電圧判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                    Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                                    Case Else
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W8020"
                                                        fncCheckSelectOption = False
                                                End Select
                                            Else
                                                ' 電圧判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                    Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"

                                                    Case Else
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W8020"
                                                        fncCheckSelectOption = False
                                                End Select
                                            End If
                                        Case "3A"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                                     "AC575V", "AC578V", "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                                     "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", _
                                                     "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", _
                                                     "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", _
                                                     "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "3M", "3I"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "DC6V", "DC8V", "DC12V", _
                                                     "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                                     "DC30V", "DC33V", "DC34V", "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", _
                                                     "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                                     "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "3N", "3J"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                                     "AC117V", "AC120V", "AC125V", "AC127V", _
                                                     "AC130V", "AC135V", "AC150V", "AC160V", _
                                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                                     "DC12V", "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC100V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "4A"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                                     "AC575V", "AC578V"
                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "4M"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "4N"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                                     "AC117V", "AC120V", "AC125V", "AC127V", _
                                                     "AC130V", "AC135V", "AC150V", "AC160V", _
                                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "5A", "5M", "5I"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                                     "AC200V", "AC220V", "AC240V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "5N", "5J"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                                     "AC200V", "AC220V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                    End Select

                                Case "15A", "20A", "25A"

                                    'サージキラー付はAC/DCとも235まで
                                    If bolOptionS = True Then
                                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) - 3)) > 236 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                        End If
                                    End If

                                    ' コイルハイジング判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "2C"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", "AC120V", _
                                                     "AC125V", "AC127V", "AC150V", "AC200V", "AC210V", "AC215V", _
                                                     "AC220V", "AC230V", "AC240V", "AC380V", "AC415V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "2E", "2G"
                                            If Not bolOptionS Then
                                                ' 電圧判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                    Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                                         "AC125V", "AC127V", "AC150V", "AC200V", "AC210V", "AC215V", _
                                                         "AC220V", "DC12V", "DC24V", "DC48V", "DC100V"
                                                    Case Else
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W8020"
                                                        fncCheckSelectOption = False
                                                End Select
                                            Else
                                                ' 電圧判定
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                    Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", "AC120V", _
                                                         "AC125V", "AC127V", "AC150V", "AC200V", "AC210V", "AC215V", _
                                                         "AC220V", "DC12V", "DC24V", "DC100V"
                                                    Case Else
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W8020"
                                                        fncCheckSelectOption = False
                                                End Select
                                            End If

                                        Case "2H"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC100V", "AC110V", "AC200V", "AC220V", "DC24V"
                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "3A"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC24V", "AC25V", "AC42V", "AC48V", "AC100V", "AC105V", _
                                                     "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", "AC200V", "AC208V", _
                                                     "AC210V", "AC215V", "AC220V", "AC230V", "AC240V", "AC265V", "AC380V", _
                                                     "AC400V", "AC415V", "AC420V", "AC440V", "AC460V", _
                                                     "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", "DC42V", "DC45V", "DC48V", _
                                                     "DC50V", "DC59V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", _
                                                     "DC120V", "DC125V", "DC200V", "DC220V", "DC235V"
                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "3M", "3I"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC24V", "AC25V", "AC42V", "AC48V", "AC100V", "AC105V", "AC110V", _
                                                     "AC115V", "AC120V", "AC125V", "AC127V", "AC200V", "AC208V", "AC210V", _
                                                     "AC215V", "AC220V", "AC230V", "AC240V", "AC265V", "AC380V", "AC400V", _
                                                     "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", "DC42V", "DC45V", "DC48V", _
                                                     "DC50V", "DC59V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", _
                                                     "DC120V", "DC125V", "DC200V", "DC220V", "DC235V"
                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "3N", "3J"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", _
                                                     "AC200V", "AC208V", "AC210V", "AC215V", "AC220V", _
                                                     "DC12V", "DC24V", "DC100V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "4A"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC24V", "AC42V", "AC45V", "AC48V", "AC50V", "AC100V", "AC105V", _
                                                     "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", "AC200V", _
                                                     "AC210V", "AC215V", "AC220V", "AC230V", "AC240V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "4M"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC12V", "AC24V", "AC42V", "AC45V", "AC48V", "AC50V", "AC100V", "AC105V", _
                                                     "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", "AC200V", _
                                                     "AC210V", "AC215V", "AC220V", "AC230V", "AC240V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "4N"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                                     "AC127V", "AC200V", "AC210V", "AC215V", "AC220V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "5A", "5M", "5I"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                                     "AC200V", "AC220V", "AC240V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                        Case "5N", "5J"
                                            ' 電圧判定
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                                     "AC200V", "AC220V"

                                                Case Else
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W8020"
                                                    fncCheckSelectOption = False
                                            End Select
                                    End Select

                                Case Else

                            End Select
                        End If
                    End If
                Case "ADK21"
                    '2010/09/10 MOD RM1009006(10月VerUP:ADK,APK21ｼﾘｰｽﾞ) START--->
                    '周波数指定有無のﾁｪｯｸ 2008/08/26
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "AC" And _
                        (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3K" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3H" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "4K" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "4H") And _
                        Right(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) <> "HZ" Then
                        intKtbnStrcSeqNo = 6
                        'If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "AC" And _
                        '    (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3K" Or _
                        '    objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3H" Or _
                        '    objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "4K" Or _
                        '    objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "4H") And _
                        '    Right(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) <> "HZ" Then
                        '    intKtbnStrcSeqNo = 5
                        '2010/09/10 MOD RM1009006(10月VerUP:ADK,APK21ｼﾘｰｽﾞ) <---END

                        strMessageCd = "W8680"
                        fncCheckSelectOption = False
                    End If

                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    'サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    ' コイルハイジング判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "3A"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "AC230V", "AC240V", "AC300V", "AC380V", "AC400V", _
                                     "AC415V", "AC440V", "AC460V", "DC12V", "DC24V", "DC48V", "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3N"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "DC12V", "DC24V", "DC100V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3M"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "AC230V", "AC240V", "AC300V", "AC380V", "AC400V", _
                                     "DC12V", "DC24V", "DC48V", "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4A"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "AC230V", "AC240V", "AC415V", "AC440V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4M"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "AC230V", "AC240V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4N"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5A", "5M", "5N"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC110V", "AC120V", _
                                     "AC200V", "AC220V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                    End Select

                Case "ADK12"

                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    'サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    ' コイルハイジング判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "3A"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC24V", "AC25V", "AC42V", "AC48V", "AC100V", _
                                     "AC105V", "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", _
                                     "AC200V", "AC208V", "AC210V", "AC215V", "AC220V", "AC230V", _
                                     "AC240V", "AC265V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC440V", "AC460V", "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", _
                                     "DC30V", "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", "DC88V", _
                                     "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", _
                                     "DC220V", "DC235V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3N", "3J"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", _
                                     "AC200V", "AC208V", "AC210V", "AC215V", "AC220V", "DC12V", _
                                     "DC24V", "DC100V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3M", "3I"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC24V", "AC25V", "AC42V", "AC48V", "AC100V", _
                                     "AC105V", "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", _
                                     "AC200V", "AC208V", "AC210V", "AC215V", "AC220V", "AC230V", _
                                     "AC240V", "AC265V", "AC380V", "AC400V", "DC6V", "DC12V", _
                                     "DC24V", "DC25V", "DC28V", "DC30V", "DC42V", "DC45V", "DC48V", _
                                     "DC50V", "DC59V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", _
                                     "DC110V", "DC120V", "DC125V", "DC200V", "DC220V", "DC235V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4A", "4M"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC24V", "AC42V", "AC45V", "AC48V", "AC50V", "AC100V", _
                                     "AC105V", "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", _
                                     "AC200V", "AC210V", "AC215V", "AC220V", "AC230V", "AC240V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4N"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", _
                                     "AC200V", "AC210V", "AC215V", "AC220V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5A", "5M", "5I"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC140V", "AC200V", "AC220V", "AC240V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5N", "5J"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC140V", "AC200V", "AC220V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                    End Select

                Case "APK11"

                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                        Case "8A", "10A"
                            ' コイルハイジング判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "2C"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                             "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                             "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                             "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", _
                                             "AC210V", "AC215V", "AC216V", "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                             "AC260V", "AC300V", "AC346V", "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                             "AC415V", "AC420V", "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2E", "2G"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                             "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                             "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                             "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", "AC210V", _
                                             "AC215V", "AC216V", "AC220V", "AC225V", "DC6V", "DC8V", "DC12V", "DC21V", "DC24V", _
                                             "DC28V", "DC30V", "DC33V", "DC45V", "DC48V", "DC50V", "DC70V", "DC85V", "DC90V", _
                                             "DC100V", "DC110V", "DC120V", "DC125V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2H"
                                    If bolOptionS Then
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                            Case Else
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    Else
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                            Case Else
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    End If
                                Case "3A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                             "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                             "AC575V", "AC578V", "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                             "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", _
                                             "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", _
                                             "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", _
                                             "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3M", "3I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "DC6V", "DC8V", "DC12V", _
                                             "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                             "DC30V", "DC33V", "DC34V", "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", _
                                             "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                             "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3N", "3J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                             "AC117V", "AC120V", "AC125V", "AC127V", _
                                             "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "DC12V", "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC100V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                             "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                             "AC575V", "AC578V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4M"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                             "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                             "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                             "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                             "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                             "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4N"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                             "AC117V", "AC120V", "AC125V", "AC127V", _
                                             "AC130V", "AC135V", "AC150V", "AC160V", _
                                             "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5A", "5M", "5I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                             "AC200V", "AC220V", "AC240V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5N", "5J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                             "AC200V", "AV220V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select

                        Case "15A", "20A", "25A"

                            ' コイルハイジング判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "2C"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", "AC120V", _
                                             "AC125V", "AC127V", "AC150V", "AC200V", "AC210V", "AC215V", _
                                             "AC220V", "AC230V", "AC240V", "AC380V", "AC415V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "2E", "2G"
                                    If bolOptionS Then
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                                 "AC125V", "AC127V", "AC150V", "AC200V", "AC210V", "AC215V", _
                                                 "AC220V"
                                            Case Else
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    Else
                                        ' 電圧判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "AC24V", "AC48V", "AC100V", "AC110V", "AC115V", "AC120V", _
                                                 "AC125V", "AC127V", "AC150V", "AC200V", "AC210V", "AC215V", _
                                                 "AC220V"
                                            Case Else
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W8020"
                                                fncCheckSelectOption = False
                                        End Select
                                    End If

                                Case "2H"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC24V", "AC28V", "AC48V", "AC100V", "AC105V", _
                                             "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", "AC200V", _
                                             "AC210V", "AC215V", "AC220V", "AC230V", "AC240V", "AC380V", _
                                             "AC400V", "AC440V", "DC12V", "DC24V", "DC30V", "DC45V", "DC48V", _
                                             "DC85V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", _
                                             "DC200V", "DC220V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3M", "3I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC24V", "AC28V", "AC48V", "AC100V", "AC105V", _
                                             "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", "AC200V", _
                                             "AC210V", "AC215V", "AC220V", "AC230V", "AC240V", "AC380V", _
                                             "AC400V", "DC12V", "DC24V", "DC30V", "DC45V", "DC48V", _
                                             "DC85V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", _
                                             "DC200V", "DC220V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "3N", "3J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                             "AC127V", "AC200V", "AC210V", "AC215V", "AC220V", "DC12V", _
                                             "DC24V", "DC100V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4A"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC24V", "AC28V", "AC48V", "AC100V", "AC105V", _
                                             "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", "AC200V", _
                                             "AC210V", "AC215V", "AC220V", "AC230V", "AC240V", "AC380V", _
                                             "AC400V", "AC440V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4M"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC12V", "AC24V", "AC28V", "AC48V", "AC100V", "AC105V", _
                                             "AC110V", "AC115V", "AC120V", "AC125V", "AC127V", "AC200V", _
                                             "AC210V", "AC215V", "AC220V", "AC230V", "AC240V", "AC380V", _
                                             "AC400V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "4N"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                             "AC127V", "AC200V", "AC210V", "AC215V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5A", "5M", "5I"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", "AC230V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Case "5N", "5J"
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select

                        Case Else

                    End Select

                Case "APK21"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "F"
                        Case Else
                            '周波数指定有無のﾁｪｯｸ 2008/08/26
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = "DC" And _
                                (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3K" Or _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "3H" Or _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "4K" Or _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "4H") And _
                                Right(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) = "HZ" Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W2790"
                                fncCheckSelectOption = False
                            End If
                    End Select

                    '2010/09/10 ADD RM1009006(10月VerUP:ADK,APK21ｼﾘｰｽﾞ) <---END

                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    Dim strVoltage As String = String.Empty
                    Dim intNum As Integer = 5

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "F"
                            intNum = 6
                            strVoltage = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        Case Else
                            intNum = 5
                            strVoltage = objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    End Select

                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    'サージキラー付はAC/DCとも230まで
                    If bolOptionS = True Then
                        If CInt(Mid(strVoltage, 3, Len(strVoltage) - 3)) > 231 Then
                            intKtbnStrcSeqNo = intNum
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    ' コイルハイジング判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "3A"
                            ' 電圧判定
                            Select Case strVoltage
                                Case "AC12V", "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "AC230V", "AC240V", "AC300V", "AC380V", "AC400V", _
                                     "AC415V", "AC440V", "AC460V", "DC12V", "DC24V", "DC48V", "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = intNum
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3N"
                            ' 電圧判定
                            Select Case strVoltage
                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "DC12V", "DC24V", "DC100V"
                                Case Else
                                    intKtbnStrcSeqNo = intNum
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3M"
                            ' 電圧判定
                            Select Case strVoltage
                                Case "AC12V", "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "AC230V", "AC240V", "AC300V", "AC380V", "AC400V", _
                                     "DC12V", "DC24V", "DC48V", "DC100V", "DC110V"
                                Case Else
                                    intKtbnStrcSeqNo = intNum
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4A"
                            ' 電圧判定
                            Select Case strVoltage
                                Case "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "AC230V", "AC240V", "AC415V", "AC440V"
                                Case Else
                                    intKtbnStrcSeqNo = intNum
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4M"
                            ' 電圧判定
                            Select Case strVoltage
                                Case "AC24V", "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V", "AC230V", "AC240V"
                                Case Else
                                    intKtbnStrcSeqNo = intNum
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4N"
                            ' 電圧判定
                            Select Case strVoltage
                                Case "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                     "AC200V", "AC220V"
                                Case Else
                                    intKtbnStrcSeqNo = intNum
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5A", "5M", "5N"
                            ' 電圧判定
                            Select Case strVoltage
                                Case "AC100V", "AC110V", "AC120V", _
                                     "AC200V", "AC220V"
                                Case Else
                                    intKtbnStrcSeqNo = intNum
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                    End Select

                Case "AP12", "AP22", "AD12", "AD22"
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    ' コイルハイジング判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "2C"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2E", "2G"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "DC12V", "DC24V", "DC48V", "DC100V"
                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "2H"
                            If bolOptionS = True Then
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Else
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            End If
                        Case "3A"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                     "AC575V", "AC578V", "DC6V", "DC9V", "DC12V", "DC14V", "DC24V", "DC26V", _
                                     "DC30V", "DC36V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V", "DC200V", "DC220V"

                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3M", "3I"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                     "DC6V", "DC9V", "DC12V", "DC14V", "DC24V", "DC26V", _
                                     "DC30V", "DC36V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V", "DC200V", "DC220V"

                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "3N", "3J"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                     "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "DC12V", "DC14V", "DC24V", "DC100V"

                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4A"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                     "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                     "AC575V", "AC578V"

                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4M"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                     "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                     "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                     "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                     "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "4N"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                     "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5A", "5M", "5I"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC125V", "AC200V", "AC220V", "AC230V", _
                                     "AC240V"

                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        Case "5N", "5J"
                            ' 電圧判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC125V", "AC200V", "AC220V"

                                Case Else
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                    End Select
                Case "AP21", "AD21"
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも230まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) - 3)) > 231 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    'ドライエア用Ｚ選択時
                    If bolOptionZ = True Then
                        'コイルハウジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "3A", "3M", "3I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                         "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC235V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "DC12V", "DC24V", "DC100V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    Else
                        ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "2C"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                         "AC415V", "AC420V", "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", "AC210V", _
                                         "AC215V", "AC216V", "AC220V", "AC225V", "DC6V", "DC8V", "DC12V", "DC21V", "DC24V", _
                                         "DC28V", "DC30V", "DC33V", "DC45V", "DC48V", "DC50V", "DC70V", "DC85V", "DC90V", _
                                         "DC100V", "DC110V", "DC120V", "DC125V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS = True Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V", "DC24V"
                                        Case Else
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V", "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", _
                                         "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", _
                                         "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "DC6V", "DC8V", "DC12V", _
                                         "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC33V", "DC34V", "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", _
                                         "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V",
                                         "DC12V", "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select

                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5I", "5M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                         "AC200V", "AC220V", "AC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If

                Case "AB31"
                    'RM0907070 2009/09/08 Y.Miura　二次電池対応
                    '電圧の位置を要素9番目⇒10番目に変更する
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(10).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = 10
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    If bolOptionZ = True Then
                        'コイルハウジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "3A", "3M", "3I"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                         "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC235V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC12V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    Else
                        ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case ""
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC125V", _
                                         "AC200V", "AC220V", "DC6V", "DC12V", "DC24V", _
                                         "DC48V", "DC100V", "DC110V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "DC5V", "DC6V", "DC12V", _
                                         "DC14V", "DC24V", "DC25V", "DC28V", "DC48V", _
                                         "DC74V", "DC85V", "DC88V", "DC90V", "DC100V", _
                                         "DC110V", "DC120V", "DC124V", "DC125V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS = True Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V"
                                        Case Else
                                            intKtbnStrcSeqNo = 10
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V", "DC24V"
                                        Case Else
                                            intKtbnStrcSeqNo = 10
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", _
                                         "AC500V", "AC575V", "AC600V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3K", "3P"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "DC4V", "DC5V", "DC6V", _
                                         "DC12V", "DC13V", "DC14V", "DC17V", "DC21V", _
                                         "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", _
                                         "DC42V", "DC48V", "DC50V", "DC85V", "DC88V", _
                                         "DC90V", "DC100V", "DC110V", "DC115V", "DC124V", _
                                         "DC125V", "DC140V", "DC200V", "DC220V", "DC230V", _
                                         "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3H", "3Q"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "AC220V", _
                                         "DC24V", "DC100V", "DC110V", "DC115V", "DC124V", _
                                         "DC125V", "DC140V", "DC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3L"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "DC24V", _
                                         "DC100V", "DC110V", "DC115V", "DC124V", "DC125V", _
                                         "DC140V", "DC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "DC12V", _
                                         "DC13V", "DC14V", "DC17V", "DC21V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "AC220V", "DC12V", _
                                         "DC13V", "DC14V", "DC17V", "DC21V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3E", "3F"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", _
                                         "AC500V", "AC575V", "AC600V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4K"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4H"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "AC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4L", "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4E", "4F"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5K", "5P", "5E", "5F", "5M", "5I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V", _
                                         "AC220V", "AC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5H", "5Q"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V", _
                                         "AC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5L", "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If
                Case "AB42"
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                        ' サージキラー付はAC/DCとも236まで
                        If bolOptionS = True Then
                            If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(10).Trim) - 3)) > 237 Then
                                intKtbnStrcSeqNo = 10
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                            End If
                        End If

                        ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case ""
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "DC12V", "DC24V", "DC48V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS = True Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                             "AC210V", "AC215V", "AC216V", "AC220V"
                                        Case Else
                                            intKtbnStrcSeqNo = 10
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", _
                                             "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"
                                        Case Else
                                            intKtbnStrcSeqNo = 10
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V", "DC6V", "DC9V", "DC12V", "DC14V", "DC24V", "DC26V", _
                                         "DC30V", "DC36V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V", "DC200V", "DC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                         "DC6V", "DC9V", "DC12V", "DC14V", "DC24V", "DC26V", _
                                         "DC30V", "DC36V", "DC48V", "DC85V", "DC100V", "DC110V", "DC125V", "DC200V", "DC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                     "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                     "DC12V", "DC14V", "DC24V", "DC100V"

                                Case Else
                                    intKtbnStrcSeqNo = 10
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", _
                                     "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                     "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                Case Else
                                    intKtbnStrcSeqNo = 10
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                            Case "5A", "5M", "5I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                         "AC125V", "AC200V", "AC220V", "AC230V", _
                                         "AC240V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                Case "AC100V", "AC110V", "AC115V", "AC120V", _
                                     "AC125V", "AC200V", "AC220V"

                                Case Else
                                    intKtbnStrcSeqNo = 10
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                        End Select
                Case "AB41"
                    'RM0907070 2009/09/08 Y.Miura　二次電池対応
                    '電圧の位置を要素9番目⇒10番目に変更する
                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも230まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(10).Trim) - 3)) > 231 Then
                            intKtbnStrcSeqNo = 10
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    If bolOptionZ = True Then
                        'コイルハウジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "3A", "3M", "3I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                         "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC235V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC12V", "DC24V", "DC100V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    Else
                       ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case ""
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                         "AC415V", "AC420V", "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", "AC210V", _
                                         "AC215V", "AC216V", "AC220V", "AC225V", "DC6V", "DC8V", "DC12V", "DC21V", "DC24V", _
                                         "DC28V", "DC30V", "DC33V", "DC45V", "DC48V", "DC50V", "DC70V", "DC85V", "DC90V", _
                                         "DC100V", "DC110V", "DC120V", "DC125V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = 10
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"

                                        Case Else
                                            intKtbnStrcSeqNo = 10
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V", "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", _
                                         "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", _
                                         "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "DC6V", "DC8V", "DC12V", _
                                         "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC33V", "DC34V", "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", _
                                         "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "DC12V", "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC100V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                         "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                         "AC200V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If

                    '2010/10/04 MOD RM1010017(11月VerUP:AB41シリーズ) START--->
                    '接続口径判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "03", "04"
                            '低圧大流量の場合、ＤＣ電圧製作不可
                            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "8" Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Case "V", "W"
                                        If Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2).Equals("DC") Then
                                            intKtbnStrcSeqNo = 10
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                        End If

                                End Select
                            End If
                    End Select
                    '2010/10/04 MOD RM1010017(11月VerUP:AB41シリーズ) <---END

                Case "AG31", "AG33", "AG34"
                    'RM0907070 2009/09/08 Y.Miura　二次電池対応
                    '電圧の要素位置の変更9番目⇒10番目に変更する
                    Dim intOptionPos As Integer = 10

                    Dim bolOptionZ As Boolean = False
                    Dim bolOptionS As Boolean = False
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Z"
                                bolOptionZ = True
                            Case "S"
                                bolOptionS = True
                        End Select
                    Next

                    ' サージキラー付はAC/DCとも236まで
                    If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim) - 3)) > 237 Then
                            intKtbnStrcSeqNo = intOptionPos
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                    End If

                    If bolOptionZ = True Then
                        'コイルハウジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "3A", "3M", "3I"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                         "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC235V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "DC12V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                '電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    Else
                        ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case ""
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "DC5V", "DC6V", "DC12V", _
                                         "DC14V", "DC24V", "DC25V", "DC28V", "DC48V", _
                                         "DC74V", "DC85V", "DC88V", "DC90V", "DC100V", _
                                         "DC110V", "DC120V", "DC124V", "DC125V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS = True Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V"
                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC110V", "AC200V", "AC220V", _
                                             "AC105V", "AC208V", "AC210V", "AC216V", "DC24V"
                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", _
                                         "AC500V", "AC575V", "AC600V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "DC12V", _
                                         "DC13V", "DC14V", "DC17V", "DC21V", "DC24V", "DC100V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3K", "3P"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "DC4V", "DC5V", "DC6V", _
                                         "DC12V", "DC13V", "DC14V", "DC17V", "DC21V", _
                                         "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", _
                                         "DC42V", "DC48V", "DC50V", "DC85V", "DC88V", _
                                         "DC90V", "DC100V", "DC110V", "DC115V", "DC124V", _
                                         "DC125V", "DC140V", "DC200V", "DC220V", "DC230V", _
                                         "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3H", "3Q"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "AC220V", _
                                         "DC24V", "DC100V", "DC110V", "DC115V", "DC124V", _
                                         "DC125V", "DC140V", "DC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3E", "3F"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "DC4V", "DC5V", _
                                         "DC6V", "DC12V", "DC13V", "DC14V", "DC17V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC42V", "DC48V", "DC50V", "DC85V", _
                                         "DC88V", "DC90V", "DC100V", "DC110V", "DC115V", _
                                         "DC124V", "DC125V", "DC140V", "DC200V", "DC220V", _
                                         "DC230V", "DC240V", "DC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3L"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "DC24V", _
                                         "DC100V", "DC110V", "DC115V", "DC124V", "DC125V", _
                                         "DC140V", "DC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", _
                                         "AC500V", "AC575V", "AC600V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select

                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V"
                                    Case Else
                                        intKtbnStrcSeqNo = 10
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4K"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4H"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V", "AC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4L", "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", _
                                         "AC150V", "AC160V", "AC190V", "AC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4E", "4F"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC16V", "AC20V", "AC24V", "AC30V", _
                                         "AC38V", "AC39V", "AC42V", "AC48V", "AC50V", _
                                         "AC80V", "AC90V", "AC95V", "AC100V", "AC105V", _
                                         "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", _
                                         "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC216V", _
                                         "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC360V", _
                                         "AC365V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC433V", "AC440V", "AC450V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5K", "5P", "5E", "5F", "5I", "5M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V", _
                                         "AC220V", "AC240V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5H", "5Q"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V", _
                                         "AC220V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5L", "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC125V", "AC200V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If
                Case "AG41", "AG43", "AG44"
                        'RM0907070 2009/09/08 Y.Miura　二次電池対応
                        'AG31は電圧の位置を要素9番目⇒10番目に変更する
                        Dim intOptionPos As Integer = 10

                        Dim bolOptionZ As Boolean = False
                        Dim bolOptionS As Boolean = False
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "Z"
                                    bolOptionZ = True
                                Case "S"
                                    bolOptionS = True
                            End Select
                        Next

                    ' サージキラー付はAC/DCとも230まで
                        If bolOptionS = True Then
                        If CInt(Mid(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, 3, Len(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim) - 3)) > 231 Then
                            intKtbnStrcSeqNo = intOptionPos
                            strMessageCd = "W8020"
                            fncCheckSelectOption = False
                        End If
                        End If

                        If bolOptionZ = True Then
                        'コイルハウジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "3A", "3M", "3I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "DC6V", "DC12V", "DC24V", "DC25V", "DC28V", "DC30V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC59V", "DC85V", _
                                         "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC235V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "DC12V", "DC24V", "DC100V"
                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                '電圧判定
                                'Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim  'RM1004012
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", "AC200V", "AC220V"

                                    Case Else
                                        'intKtbnStrcSeqNo = 9
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    Else
                        ' コイルハイジング判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case ""
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", _
                                         "AC210V", "AC215V", "AC216V", "AC220V", "AC225V", "AC230V", "AC240V", "AC250V", _
                                         "AC260V", "AC300V", "AC346V", "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", _
                                         "AC415V", "AC420V", "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2E", "2G"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", "AC30V", _
                                         "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", "AC90V", "AC95V", _
                                         "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", "AC190V", "AC200V", "AC208V", "AC210V", _
                                         "AC215V", "AC216V", "AC220V", "AC225V", "DC6V", "DC8V", "DC12V", "DC21V", "DC24V", _
                                         "DC28V", "DC30V", "DC33V", "DC45V", "DC48V", "DC50V", "DC70V", "DC85V", "DC90V", _
                                         "DC100V", "DC110V", "DC120V", "DC125V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "2H"
                                If bolOptionS Then
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V"

                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                Else
                                    ' 電圧判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                        Case "AC100V", "AC105V", "AC110V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", "DC24V"

                                        Case Else
                                            intKtbnStrcSeqNo = intOptionPos
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                            Case "3A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V", "DC6V", "DC8V", "DC12V", "DC14V", "DC16V", "DC20V", _
                                         "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", "DC30V", "DC33V", "DC34V", _
                                         "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", "DC74V", "DC85V", "DC88V", _
                                         "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", "DC125V", "DC200V", _
                                         "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3M", "3I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "DC6V", "DC8V", "DC12V", _
                                         "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC25V", "DC26V", "DC28V", _
                                         "DC30V", "DC33V", "DC34V", "DC42V", "DC45V", "DC48V", "DC50V", "DC70V", _
                                         "DC74V", "DC85V", "DC88V", "DC89V", "DC90V", "DC100V", "DC110V", "DC120V", _
                                         "DC125V", "DC200V", "DC220V", "DC230V", "DC235V", "DC240V", "DC250V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "3N", "3J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC215V", "AC216V", "AC220V", _
                                         "DC12V", "DC14V", "DC16V", "DC20V", "DC21V", "DC24V", "DC100V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4A"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V", "AC415V", "AC420V", _
                                         "AC430V", "AC433V", "AC440V", "AC450V", "AC460V", "AC480V", "AC500V", _
                                         "AC575V", "AC578V"
                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4M"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC12V", "AC15V", "AC16V", "AC20V", "AC24V", "AC25V", "AC27V", _
                                         "AC30V", "AC35V", "AC38V", "AC42V", "AC45V", "AC48V", "AC85V", _
                                         "AC90V", "AC95V", "AC100V", "AC105V", "AC110V", "AC115V", "AC117V", _
                                         "AC120V", "AC125V", "AC127V", "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC210V", "AC215V", "AC216V", "AC220V", _
                                         "AC225V", "AC230V", "AC240V", "AC250V", "AC260V", "AC300V", "AC346V", _
                                         "AC350V", "AC365V", "AC370V", "AC380V", "AC400V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "4N"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC105V", "AC110V", "AC115V", _
                                         "AC117V", "AC120V", "AC125V", "AC127V", _
                                         "AC130V", "AC135V", "AC150V", "AC160V", _
                                         "AC190V", "AC200V", "AC208V", "AC215V", "AC216V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5A", "5M", "5I"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                         "AC200V", "AC220V", "AC240V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "5N", "5J"
                                ' 電圧判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim
                                    Case "AC100V", "AC110V", "AC115V", "AC120V", "AC140V", _
                                         "AC200V", "AC220V"

                                    Case Else
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If
                Case "WFK"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "3"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "C" Then
                                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                                If strOpArray.Length <> 2 Then
                                    intKtbnStrcSeqNo = 5
                                    strMessageCd = "W9180"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select

            End Select

        Catch ex As Exception

            fncCheckSelectOption = False

            Throw ex

        End Try

    End Function

End Module
