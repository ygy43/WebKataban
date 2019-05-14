Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class YousoBLL

    Public Structure CompData
        Public strSeriesKataban As String                              'シリーズ形番
        Public strKeyKataban As String                                 'キー形番
        Public strFullKataban As String                                'フル形番
        Public strGoodsNm As String                                    '商品名
        Public strHyphen As String                                     '次ハイフン
        Public strOpSymbol() As String                                 'オプション記号
        Public strElementDiv() As String                               '要素区分
        Public strStructureDiv() As String                             '構成区分
        Public strAdditionDiv() As String                              '付加区分
        Public strHyphenDiv() As String                                'ハイフン区分
        Public strKtbnStrcNm() As String                               '構成名称
        Public strKtbnStrcEle(,) As String                             '構成要素
    End Structure

    ''' <summary>
    ''' 形番構成取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strcCompData"></param>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks>形番構成を読み込み該当するデータを変数に格納する</remarks>
    Public Shared Function fncKatabanStrcSelect(ByVal objCon As SqlConnection, ByRef strcCompData As CompData, _
                                                strLanguage As String) As Boolean
        Dim dt As New DataTable
        Dim dalYousoTmp As New YousoDAL

        fncKatabanStrcSelect = False
        Try
            '配列初期化
            ReDim strcCompData.strElementDiv(0)
            ReDim strcCompData.strStructureDiv(0)
            ReDim strcCompData.strAdditionDiv(0)
            ReDim strcCompData.strHyphenDiv(0)
            ReDim strcCompData.strKtbnStrcNm(0)

            dt = dalYousoTmp.fncKatabanStrcSelect(objCon, strcCompData, strLanguage)

            For Each dr As DataRow In dt.Rows
                ReDim Preserve strcCompData.strElementDiv(UBound(strcCompData.strElementDiv) + 1)
                ReDim Preserve strcCompData.strStructureDiv(UBound(strcCompData.strStructureDiv) + 1)
                ReDim Preserve strcCompData.strAdditionDiv(UBound(strcCompData.strAdditionDiv) + 1)
                ReDim Preserve strcCompData.strHyphenDiv(UBound(strcCompData.strHyphenDiv) + 1)
                ReDim Preserve strcCompData.strKtbnStrcNm(UBound(strcCompData.strKtbnStrcNm) + 1)

                strcCompData.strElementDiv(UBound(strcCompData.strElementDiv)) = dr("element_div")
                strcCompData.strStructureDiv(UBound(strcCompData.strStructureDiv)) = dr("structure_div")
                strcCompData.strAdditionDiv(UBound(strcCompData.strAdditionDiv)) = dr("addition_div")
                strcCompData.strHyphenDiv(UBound(strcCompData.strHyphenDiv)) = dr("hyphen_div")

                If IsDBNull(dr("ktbn_strc_nm")) = True Then
                    strcCompData.strKtbnStrcNm(UBound(strcCompData.strKtbnStrcNm)) = dr("defaultNm")
                Else
                    strcCompData.strKtbnStrcNm(UBound(strcCompData.strKtbnStrcNm)) = dr("ktbn_strc_nm")
                End If
            Next

            fncKatabanStrcSelect = True

        Catch ex As Exception
            fncKatabanStrcSelect = False
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 形番構成要素取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strcCompData"></param>
    ''' <returns></returns>
    ''' <remarks>形番構成要素を読み込み該当するデータを変数に格納する</remarks>
    Public Shared Function subKtbnStrcEleSelect(ByVal objCon As SqlConnection, ByRef strcCompData As CompData) As Boolean
        Dim dt As New DataTable
        Dim dalYousoTmp As New YousoDAL

        Dim strKtbnStrcSeq() As String
        Dim strKtbnStrcValue() As String
        Dim strKtbnPlcaelvl() As Long
        Dim intLoopCnt As Integer

        subKtbnStrcEleSelect = False

        Try
            '配列初期化
            ReDim strKtbnStrcSeq(0)
            ReDim strKtbnStrcValue(0)
            ReDim strKtbnPlcaelvl(0)
            ReDim strcCompData.strKtbnStrcEle(0, 3)

            dt = dalYousoTmp.subKtbnStrcEleSelect(objCon, strcCompData)

            For Each dr As DataRow In dt.Rows
                ReDim Preserve strKtbnStrcSeq(UBound(strKtbnStrcSeq) + 1)
                ReDim Preserve strKtbnStrcValue(UBound(strKtbnStrcValue) + 1)
                ReDim Preserve strKtbnPlcaelvl(UBound(strKtbnPlcaelvl) + 1)
                strKtbnStrcSeq(UBound(strKtbnStrcSeq)) = dr("ktbn_strc_seq_no")
                strKtbnStrcValue(UBound(strKtbnStrcValue)) = dr("option_symbol")
                strKtbnPlcaelvl(UBound(strKtbnPlcaelvl)) = dr("place_lvl")
            Next

            ReDim strcCompData.strKtbnStrcEle(UBound(strKtbnStrcSeq), 3)
            For intLoopCnt = 1 To strKtbnStrcSeq.Length - 1
                strcCompData.strKtbnStrcEle(intLoopCnt, 1) = strKtbnStrcSeq(intLoopCnt)
                strcCompData.strKtbnStrcEle(intLoopCnt, 2) = strKtbnStrcValue(intLoopCnt)
                strcCompData.strKtbnStrcEle(intLoopCnt, 3) = strKtbnPlcaelvl(intLoopCnt)
            Next
            subKtbnStrcEleSelect = True
        Catch ex As Exception
            subKtbnStrcEleSelect = False
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 引当形番構成データのチェック
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <returns></returns>
    ''' <remarks>引当形番構成を読み込み該当するデータがあるかチェックする</remarks>
    Public Shared Function fncSelKtbnStrcCheck(ByVal objCon As SqlConnection, _
                                               strUserId As String, strSessionId As String) As Boolean
        Dim dt As New DataTable
        Dim dalYousoTmp As New YousoDAL
        fncSelKtbnStrcCheck = False

        Try
            dt = dalYousoTmp.fncSelKtbnStrcCheck(objCon, strUserId, strSessionId)
            If dt.Rows.Count > 0 Then fncSelKtbnStrcCheck = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 出荷場所レベルの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <param name="strSymbol"></param>
    ''' <param name="intNo"></param>
    ''' <param name="intPlacelvl">生産レベル</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function subGetPlacelvl(ByVal objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                          ByVal strKeyKataban As String, ByVal strSymbol As String, _
                                          ByVal intNo As Integer, ByRef intPlacelvl As Integer) As Boolean
        Dim dt As New DataTable
        Dim dalYousoTmp As New YousoDAL
        subGetPlacelvl = False
        Try
            dt = dalYousoTmp.subGetPlacelvl(objCon, strSeriesKataban, strKeyKataban, strSymbol, intNo, intPlacelvl)
            For Each dr As DataRow In dt.Rows
                intPlacelvl = dr("place_lvl")
                Exit For
            Next
            subGetPlacelvl = True
        Catch ex As Exception
            subGetPlacelvl = False
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 要素パターン取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intFocusNo"></param>
    ''' <param name="intConSeqNoBr"></param>
    ''' <param name="strConOpSymbol"></param>
    ''' <returns></returns>
    ''' <remarks>引当形番構成を読み込み該当するデータがあるかチェックする</remarks>
    Public Shared Function fncElePtnSelect(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc, _
                                           ByVal intFocusNo As Integer, ByRef intConSeqNoBr As ArrayList, _
                                           ByRef strConOpSymbol As ArrayList) As Boolean
        Dim dt As New DataTable
        Dim dalYousoTmp As New YousoDAL
        fncElePtnSelect = False
        Try
            dt = dalYousoTmp.fncElePtnSelect(objCon, objKtbnStrc, intFocusNo, intConSeqNoBr, strConOpSymbol)

            For Each dr As DataRow In dt.Rows
                intConSeqNoBr.Add(dr("condition_seq_no_br"))
                strConOpSymbol.Add(dr("cond_option_symbol"))
            Next

            fncElePtnSelect = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' '生産国データの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetPlacelvl(ByVal objCon As SqlConnection, ByVal strSeries As String, _
                                          ByVal strKey As String) As DataTable
        Dim dalYousoTmp As New YousoDAL
        fncGetPlacelvl = New DataTable

        Try
            fncGetPlacelvl = dalYousoTmp.fncGetPlacelvl(objCon, strSeries, strKey)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 'オプション外設定ファイルの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetOutofopPlacelvl(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                                 ByVal strSessionId As String) As DataTable
        Dim dalYousoTmp As New YousoDAL
        fncGetOutofopPlacelvl = New DataTable

        Try
            fncGetOutofopPlacelvl = dalYousoTmp.fncGetOutofopPlacelvl(objCon, strUserId, strSessionId)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 全ての生産国名を取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetAllCountryLevel(ByVal objConBase As SqlConnection) As DataTable
        Dim dalYousoTmp As New YousoDAL
        fncGetAllCountryLevel = New DataTable

        Try
            fncGetAllCountryLevel = dalYousoTmp.fncGetAllCountryLevel(objConBase)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' ストローク国名を取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strKey"></param>
    ''' <param name="intBoreSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetStrokeCountry(ByVal objCon As SqlConnection, ByVal strSeries As String, _
                                               ByVal strKey As String, ByVal intBoreSize As Long) As DataTable
        Dim dalYousoTmp As New YousoDAL
        fncGetStrokeCountry = New DataTable

        Try
            fncGetStrokeCountry = dalYousoTmp.fncGetStrokeCountry(objCon, strSeries, strKey, intBoreSize)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 標準ストロークを取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeries"></param>
    ''' <param name="strKey"></param>
    ''' <param name="intBoreSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetStdStroke(ByVal objCon As SqlConnection, ByVal strSeries As String, _
                                           ByVal strKey As String, ByVal intBoreSize As Long, _
                                           ByVal intstroke As Long) As DataTable
        Dim dalYousoTmp As New YousoDAL
        fncGetStdStroke = New DataTable

        Try
            fncGetStdStroke = dalYousoTmp.fncGetStdStroke(objCon, strSeries, strKey, intBoreSize, intstroke)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 形番構成要素を取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <param name="strLanguageCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function subKataStrcEleSel(ByVal objCon As SqlConnection, _
                                             ByVal strSeriesKataban As String, ByVal strKeyKataban As String, _
                                             ByVal strLanguageCd As String) As DataTable
        Dim dalYousoTmp As New YousoDAL
        subKataStrcEleSel = New DataTable

        Try
            subKataStrcEleSel = dalYousoTmp.subKataStrcEleSel(objCon, strSeriesKataban, strKeyKataban, strLanguageCd)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' エレパタンの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function subElePatternSel(ByVal objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                    ByVal strKeyKataban As String) As DataTable

        Dim dalYousoTmp As New YousoDAL
        subElePatternSel = New DataTable

        Try
            subElePatternSel = dalYousoTmp.subElePatternSel(objCon, strSeriesKataban, strKeyKataban)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 引当口径検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function subBoreSizeSelect(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc) As String

        Dim dalYousoTmp As New YousoDAL
        Dim dt As New DataTable
        subBoreSizeSelect = String.Empty

        Try
            dt = dalYousoTmp.subBoreSizeSelect(objCon, objKtbnStrc)

            For Each dr As DataRow In dt.Rows
                subBoreSizeSelect = IIf(IsDBNull(dr("ktbn_strc_seq_no")), CdCst.Sign.Blank, _
                                        objKtbnStrc.strcSelection.strOpSymbol(dr("ktbn_strc_seq_no")))
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' ストロークを取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intBoreSize"></param>
    ''' <param name="intMinStroke"></param>
    ''' <param name="intMaxStroke"></param>
    ''' <param name="intUnitStroke"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetStroke(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByVal intBoreSize As Integer, ByRef intMinStroke As Integer, _
                                   ByRef intMaxStroke As Integer, ByRef intUnitStroke As Integer) As Boolean
        Dim dalYousoTmp As New YousoDAL
        Dim dt As New DataTable
        fncGetStroke = False

        Try
            dt = dalYousoTmp.fncGetStroke(objCon, objKtbnStrc, intBoreSize)

            For Each dr As DataRow In dt.Rows
                If dr("min_stroke") <> 0 And dr("max_stroke") <> 0 Then
                    intMinStroke = dr("min_stroke")
                    intMaxStroke = dr("max_stroke")
                    intUnitStroke = dr("stroke_unit")
                    fncGetStroke = True
                End If
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 生産国の判断ロジック（ストローク範囲）
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strPort"></param>
    ''' <param name="lngStrock"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetStroke_Logic(ByVal objKtbnStrc As KHKtbnStrc, ByVal strPort As String, _
                                              ByVal lngStrock As Long) As Boolean
        fncGetStroke_Logic = True
        Try
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "SCA2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            If objKtbnStrc.strcSelection.strOpSymbol(2) = "L2T" Then
                                Select Case strPort
                                    Case "40", "50", "63"
                                        If lngStrock < 50 Or lngStrock > 600 Then Return False
                                    Case "80"
                                        If lngStrock < 50 Or lngStrock > 700 Then Return False
                                    Case "100"
                                        If lngStrock < 50 Or lngStrock > 800 Then Return False
                                End Select
                            End If

                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                                Case "P", "R", "T", "H", "Q2", "O"
                                    Select Case strPort
                                        Case "40", "50", "63"
                                            If lngStrock < 1 Or lngStrock > 600 Then Return False
                                        Case "80"
                                            If lngStrock < 1 Or lngStrock > 700 Then Return False
                                        Case "100"
                                            If lngStrock < 1 Or lngStrock > 800 Then Return False
                                    End Select
                            End Select
                    End Select
                Case "CMK2"  'CMK2機種 Add by Zxjike 2013/11/21
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            If objKtbnStrc.strcSelection.strOpSymbol(1) = "P" Or _
                                objKtbnStrc.strcSelection.strOpSymbol(1) = "R" Then
                                'If intStrockCount = 1 AndAlso lngStrock < 25 Then Return False '条件変更 ID5086 2013/12/11
                                If lngStrock < 25 Then Return False
                            End If
                            If objKtbnStrc.strcSelection.strOpSymbol(15) = "J" Then
                                If lngStrock < 25 Then Return False
                            End If
                    End Select
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 電圧を変更する
    ''' </summary>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strKeyKataban">キー形番</param>
    ''' <param name="strVoltage">電圧</param>
    ''' <param name="intPos">オプション位置</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncChangeVlt(ByVal strSeriesKataban As String, ByVal strKeyKataban As String, _
                                  ByVal strVoltage As String, ByVal intPos As Integer) As String
        fncChangeVlt = strVoltage
        Try
            '電圧を変更する
            Select Case strSeriesKataban
                Case "NP13", "NP14", "NVP11"
                    If intPos = 4 Then
                        Select Case strVoltage
                            Case "AC100V"
                                fncChangeVlt = "1"
                            Case "AC200V"
                                fncChangeVlt = "2"
                            Case "DC24V"
                                fncChangeVlt = "3"
                        End Select
                    End If
                Case "CVS2"
                    If intPos = 9 Then
                        Select Case strVoltage
                            Case "AC100V"
                                fncChangeVlt = "1"
                            Case "AC200V"
                                fncChangeVlt = "2"
                            Case "DC24V"
                                fncChangeVlt = "3"
                        End Select
                    End If
                Case "CVS2E", "CVS3E"
                    If intPos = 5 Then
                        Select Case strVoltage
                            Case "AC100V"
                                fncChangeVlt = "1"
                            Case "AC200V"
                                fncChangeVlt = "2"
                            Case "DC24V"
                                fncChangeVlt = "3"
                        End Select
                    End If
                Case "CVS3", "CVS31"
                    If intPos = 8 Then
                        Select Case strVoltage
                            Case "AC100V"
                                fncChangeVlt = "1"
                            Case "AC200V"
                                fncChangeVlt = "2"
                            Case "DC24V"
                                fncChangeVlt = "3"
                        End Select
                    End If
                Case "CVSE2", "CVSE3"
                    If intPos = 7 Then
                        Select Case strVoltage
                            Case "AC100V"
                                fncChangeVlt = "1"
                            Case "AC200V"
                                fncChangeVlt = "2"
                            Case "DC24V"
                                fncChangeVlt = "3"
                        End Select
                    End If
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 次の画面を判断
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="objOption"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetNextFormMode(objKtbnStrc As KHKtbnStrc, objOption As KHOptionCtl) As Integer
        GetNextFormMode = 0 '0：単価見積、1：仕様書画面、2：ロッド先端形状オーダーメイド寸法入力画面
        Dim bolRodEndFlag As Boolean = False

        'ロッド先端特注仕様判定
        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
            Case "SCM"
                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "B" Then
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(19).Trim) = 0 Then
                        bolRodEndFlag = False
                    Else
                        bolRodEndFlag = True
                    End If
                ElseIf objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(15).Trim) = 0 Then
                        bolRodEndFlag = False
                    Else
                        bolRodEndFlag = True
                    End If
                End If
        End Select

        Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
            Case ""   'ページ遷移(単価見積画面)
                GetNextFormMode = 0
            Case "00" 'ページ遷移(ロッド先端形状オーダーメイド寸法入力画面)
                If bolRodEndFlag Then
                    GetNextFormMode = 2
                Else
                    GetNextFormMode = 0
                End If
            Case "01", "A1", "A2", "A3", "B1", "02", "03", "04", "05", "06", "07", "08", "10", _
                "11", "13", "14", "15", "16", "96", "B2", "B3", "B4"
                GetNextFormMode = 1
            Case "09"
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" Then 'ページ遷移(単価見積画面)
                    GetNextFormMode = 0
                Else 'ページ遷移(仕様書画面)
                    GetNextFormMode = 1
                End If
            Case "12", "18", "19", "20", "21", "22", "23", "94", "95"
                If objOption.fncVaccumMixCheck(objKtbnStrc) Then 'ページ遷移(仕様書画面)
                    GetNextFormMode = 1
                Else 'ページ遷移(単価見積画面)
                    GetNextFormMode = 0
                End If
            Case "17"
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Then 'ページ遷移(仕様書画面)
                    GetNextFormMode = 1
                Else 'ページ遷移(単価見積画面)
                    GetNextFormMode = 0
                End If
            Case "64", "66", "68", "70", "72"
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then 'ページ遷移(仕様書画面)
                    GetNextFormMode = 1
                Else 'ページ遷移(単価見積画面)
                    GetNextFormMode = 0
                End If
                'RM1805001_4Rシリーズ追加
            Case "52", "60", "61", "62", "63", "65", "67", "69", "71", "S", "T", "U", "A4", "A5", "A6", "A7", "A8" '4Hシリーズ追加
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then 'ページ遷移(仕様書画面)
                    GetNextFormMode = 1
                Else 'ページ遷移(単価見積画面)
                    GetNextFormMode = 0
                End If
            Case "53", "73", "74", "75", "76", "77", "78", "79", "80", "81", _
                    "82", "83", "84", "85", "86", "87", "88", "93"
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "80" Or _
                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "81" Then 'ページ遷移(仕様書画面)
                    GetNextFormMode = 1
                Else 'ページ遷移(単価見積画面)
                    GetNextFormMode = 0
                End If
            Case "89", "90", "98"
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then 'ページ遷移(仕様書画面)
                    GetNextFormMode = 1
                Else 'ページ遷移(単価見積画面)
                    GetNextFormMode = 0
                End If
            Case "51"
                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then 'ページ遷移(仕様書画面)
                    GetNextFormMode = 1
                Else 'ページ遷移(単価見積画面)
                    GetNextFormMode = 0
                End If
            Case "54", "55", "56", "57", "58", "59", "91", "92"
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then 'ページ遷移(仕様書画面)
                    GetNextFormMode = 1
                Else 'ページ遷移(単価見積画面)
                    GetNextFormMode = 0
                End If
            Case Else 'ページ遷移(仕様書画面)
                GetNextFormMode = 1
        End Select
    End Function

End Class
