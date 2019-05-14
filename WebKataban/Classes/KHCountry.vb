Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class KHCountry
    ''' <summary>
    ''' 国別生産品の対象国判定
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <returns>表示対象国コード(昇順)</returns>
    ''' <remarks>
    ''' ログインユーザの国コードにより表示可能な国コードを取得
    ''' </remarks>
    Public Shared Function fncCountryTradeGet(objCon As SqlConnection, ByVal strCountryCd As String) As ArrayList
        fncCountryTradeGet = New ArrayList
        Dim dt As New DataTable
        Dim dalCountryTmp As New CountryDAL
        Try
            dt = dalCountryTmp.fncCountryTradeGet(objCon, strCountryCd)

            If dt.Rows.Count > 0 Then
                '取得した対象国をすべて設定
                For Each dr As DataRow In dt.Rows
                    fncCountryTradeGet.Add(dr("disp_country_cd"))
                Next
            Else
                '検索した国コードのみが対象国
                fncCountryTradeGet.Add(strCountryCd)
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 形番のすべての国コードを取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncCountryKeyGet(objCon As SqlConnection, ByVal strKataban As String) As ArrayList
        fncCountryKeyGet = New ArrayList
        Dim dt As New DataTable
        Dim dalCountryTmp As New CountryDAL
        Try
            dt = dalCountryTmp.fncCountryKeyGet(objCon, strKataban)

            If dt.Rows.Count > 0 Then
                Dim strDateTime As DateTime = Now.Date
                For Each dr As DataRow In dt.Rows
                    '失効日対応
                    If (dr("in_effective_date") <= strDateTime) AndAlso (strDateTime < dr("out_effective_date")) Then
                        fncCountryKeyGet.Add(dr("country_cd"))
                    End If
                Next
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' フル形番の国コードの取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strFullKataban">形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncCountryItmMstChkP(objConBase As SqlConnection, ByVal strFullKataban As String, ByVal strCountry As String) As DataTable
        Dim dt As New DataTable
        Dim dalCountryTmp As New CountryDAL

        Try
            dt = dalCountryTmp.fncCountryItmMstChkP(objConBase, strFullKataban, strCountry)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dt

    End Function

    ''' <summary>
    ''' 生産国名の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="intPlacelvl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetPlacelvlName(objCon As SqlConnection, ByVal intPlacelvl As Long) As ArrayList
        fncGetPlacelvlName = New ArrayList
        Dim dt As New DataTable
        Dim dalCountryTmp As New CountryDAL

        Try
            dt = dalCountryTmp.fncGetPlacelvlName(objCon, intPlacelvl)

            If dt.Rows.Count > 0 Then
                '取得した対象国をすべて設定
                For Each dr As DataRow In dt.Rows
                    If dr("place_lvl") <= intPlacelvl Then
                        fncGetPlacelvlName.Add(dr("disp_seq_no") & "," & dr("place_div"))
                        intPlacelvl = intPlacelvl - dr("place_lvl")
                    End If
                Next
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 生産国の判断ロジック（ストローク範囲）
    ''' </summary>
    ''' <param name="lngPlacelvl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetStroke_Logic(ByVal lngPlacelvl As Long) As String
        fncGetStroke_Logic = String.Empty
        Dim strKey As String = "1,2,4,8,16,32,64,128,256,512,1024"
        Dim strKeylist() As String = strKey.Split(",")

        For intl As Integer = strKeylist.Length - 1 To 0 Step -1
            If CLng(strKeylist(intl)) <= lngPlacelvl Then
                lngPlacelvl -= strKeylist(intl)
                If fncGetStroke_Logic = String.Empty Then
                    fncGetStroke_Logic &= strKeylist(intl)
                Else
                    fncGetStroke_Logic &= "," & strKeylist(intl)
                End If
            End If
        Next
    End Function

    ''' <summary>
    ''' 出荷場所の表示名の取得(英語と日本語だけ)
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetCountryName(objCon As SqlConnection) As DataTable
        Dim dtResult As New DataTable
        Dim dalCountryTmp As New CountryDAL

        Try
            dtResult = dalCountryTmp.fncGetCountryName(objCon)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' 出荷場所の表示名の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetAllCountryName(objCon As SqlConnection) As DataTable
        Dim dtResult As New DataTable
        Dim dalCountryTmp As New CountryDAL

        Try
            dtResult = dalCountryTmp.fncGetAllCountryName(objCon)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' 出荷場所変更情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strPlaceCd">出荷場所コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncPlaceChangeInfo(objCon As SqlConnection, ByVal strKataban As String, _
                                              ByRef strPlaceCd As String, ByRef strEvaluationType As String, _
                                              ByRef strSearchDiv As String) As Boolean
        Dim dt As New DataTable
        Dim dalCountryTmp As New CountryDAL

        fncPlaceChangeInfo = False

        Try
            dt = dalCountryTmp.fncPlaceChangeInfo(objCon, strKataban)

            If dt.Rows.Count > 0 Then
                strPlaceCd = dt.Rows(0)("place_cd")
                strEvaluationType = dt.Rows(0)("evaluation_type")
                strSearchDiv = dt.Rows(0)("search_div")
                fncPlaceChangeInfo = True
            Else
                fncPlaceChangeInfo = False
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 出荷場所変更情報取得処理（GLC在庫品）
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strPlaceCd">出荷場所コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncStockPlaceInfo(objCon As SqlConnection, ByVal strKataban As String, _
                                              ByRef strPlaceCd As String) As DataTable

        Dim dt As New DataTable
        Dim dalCountryTmp As New CountryDAL

        Try
            dt = dalCountryTmp.fncStockPlaceInfo(objCon, strKataban, strPlaceCd)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dt

    End Function

    ''' <summary>
    ''' 中国生産品判断
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetData_Logic_China(ByVal objKtbnStrc As KHKtbnStrc, ByVal strCountryCd As String) As Boolean

        Dim strOpArray() As String
        Dim Option1_Kbn As String
        Dim Option2_Kbn As String

        fncGetData_Logic_China = True
        Try
            With objKtbnStrc.strcSelection
                Select Case .strSeriesKataban
                    Case "CMK2"
                        Select Case .strKeyKataban
                            Case ""
                                If .strOpSymbol(1) = "Q" And _
                                    (.strOpSymbol(2) = "CC" Or _
                                    .strOpSymbol(2) = "CC1") Then
                                    Return False
                                End If
                            Case "4"
                                If .strOpSymbol(1) = "Q" And _
                                    (.strOpSymbol(2) = "CC" Or _
                                    .strOpSymbol(15) = "P40") Then     'ADD BY ID5086 2016/08/11禁則追加
                                    Return False
                                End If

                                'ID5086 2016/08/11解除　ADD BY YGY 20141120中国ユーザ以外は生産不可
                                'If Not strCountryCd.Equals("PRC") Then
                                '    Return False
                                'End If
                        End Select
                    Case "M4F5", "M4F6", "M4F7"  'Add by Zxjike 2014/02/13
                        '要素5「電線接続」が「B」「BL」のとき、要素9「電圧」の「AC110V」「AC200V」「DC110V」選択不可
                        Select Case .strOpSymbol(5)
                            Case "B", "BL"
                                Select Case .strOpSymbol(9)
                                    Case "AC110V", "AC200V", "DC110V"
                                        Return False
                                End Select
                        End Select
                        '要素5「電線接続」が「L」のとき、要素9「電圧」の「DC110V」選択不可
                        Select Case .strOpSymbol(5)
                            Case "L"
                                Select Case .strOpSymbol(9)
                                    Case "DC110V"
                                        Return False
                                End Select
                        End Select
                        '要素6「オプション」が「S」のとき、要素9「電圧」の「DC110V」選択不可
                        If .strOpSymbol(6).ToString.Contains("S") Then
                            Select Case .strOpSymbol(9)
                                Case "DC110V"
                                    Return False
                            End Select
                        End If
                    Case "M4KA2", "M4KB2" 'Add by Zxjike 2014/03/05
                        '要素4「電線接続」が「L」のとき、要素8「電圧」の「DC12V」選択不可
                        If .strOpSymbol(4).ToString.Contains("L") Then
                            Select Case .strOpSymbol(8)
                                Case "DC12V"
                                    Return False
                            End Select
                        End If
                        If .strSeriesKataban = "M4KB2" Then
                            '要素7「連数」が「9、10」のとき、要素2「接続口径」の「08」のみ選択可能
                            If .strOpSymbol(7).ToString.Contains("9") Or _
                                .strOpSymbol(7).ToString.Contains("10") Then
                                If .strOpSymbol(2) <> "08" Then
                                    Return False
                                End If
                            End If
                        End If
                    Case "AB31"
                        '要素1「接続口径」が「01」「1G」「1N」「2G」「2N」の時、要素3「ボディ・シール材質組合せ」の「D」「E」「F」が選択不可
                        Select Case .strOpSymbol(1)
                            Case "01", "1G", "1N", "2G", "2N"
                                Select Case .strOpSymbol(3)
                                    Case "D", "E", "F"
                                        Return False
                                End Select
                        End Select

                        '要素4「コイルハウジング」が「3A」の時、要素10「電圧」の[AC100V], [AC110V], [AC200V], [AC220V]が選択不可
                        Select Case .strOpSymbol(4)
                            Case "3A"
                                Select Case .strOpSymbol(10)
                                    Case "AC100V", "AC110V", "AC200V", "AC220V"
                                        Return False
                                End Select
                        End Select

                    Case "AB41"
                        '要素1「接続口径」が[2G], [2N], [3G], [3N], [04], [4G], [4N]の時、要素3「ボディ・シール材質組合せ」の「D」「E」「F」が選択不可
                        Select Case .strOpSymbol(1)
                            Case "2G", "2N", "3G", "3N", "04", "4G", "4N"
                                Select Case .strOpSymbol(3)
                                    Case "D", "E", "F"
                                        Return False
                                End Select
                        End Select

                        '要素2「オリフィス」が「8」の時、要素3「ﾎﾞﾃﾞｨ･ｼｰﾙ材質組合せ」の「D」「E」「F」が選択不可
                        Select Case .strOpSymbol(2)
                            Case "8"
                                Select Case .strOpSymbol(3)
                                    Case "D", "E", "F"
                                        Return False
                                End Select
                        End Select

                    Case "AB42"
                        '要素1「接続口径」が[2G], [2N], [3G], [3N]の時、要素3「ボディ・シール材質組合せ」の「D」「E」「F」が選択不可
                        Select Case .strOpSymbol(1)
                            Case "2G", "2N", "3G", "3N"
                                Select Case .strOpSymbol(3)
                                    Case "D", "E", "F"
                                        Return False
                                End Select
                        End Select

                    Case "AG31", "AG33", "AG34"
                        '要素1「接続口径」が「01」「1G」「1N」「2G」「2N」の時、要素3「ボディ・シール材質組合せ」の「D」「E」「F」が選択不可
                        Select Case .strOpSymbol(1)
                            Case "01", "1G", "1N", "2G", "2N"
                                Select Case .strOpSymbol(3)
                                    Case "D", "E"
                                        Return False
                                End Select
                        End Select

                        '要素4「コイルハウジング」が「3A」の時、要素10「電圧」の[AC100V], [AC110V], [AC200V], [AC220V]が選択不可
                        Select Case .strOpSymbol(4)
                            Case "3A"
                                Select Case .strOpSymbol(10)
                                    Case "AC100V", "AC110V", "AC200V", "AC220V"
                                        Return False
                                End Select
                        End Select

                    Case "AG41", "AG43", "AG44"
                        '要素1「接続口径」が「2G」「2N」「3G」「3N」の時、要素3「ボディ・シール材質組合せ」の「D」「E」が選択不可
                        Select Case .strOpSymbol(1)
                            Case "2G", "2N", "3G", "3N"
                                Select Case .strOpSymbol(3)
                                    Case "D", "E"
                                        Return False
                                End Select
                        End Select



                        'FRL禁則条件追加  RM1703020  2017/03/14 追加  --------------------------------------------------------->

                    Case "C1000", "C1010", "C1020", "C1030", "C1040", "C1050", "C3000", "C3010", "C3020", "C3030", "C3040", "C3050", _
                         "C4000", "C4010", "C4020", "C4030", "C4040", "C4050", "C8000", "C8010", "C8020", "C8030", "C8040", "C8050"

                        '要素3「オプション」が「L」の時、要素7「配管アダプタセット・アタッチメント(添付)」の「G40P」「G45P」「G49P」「G50P」「G59P」が選択不可

                        '変数の初期化
                        Option1_Kbn = ""
                        Option2_Kbn = ""

                        '要素3の中に「L」があるかどうかチェックし、あった場合は変数に「1」をセットする
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "L"
                                    Option1_Kbn = "1"
                            End Select
                        Next

                        '要素3の中に「L」があった場合のみ以下の処理を実施
                        If Option1_Kbn = "1" Then

                            '要素7の中に「G40P」「G45P」「G49P」「G50P」「G59P」があるかどうかチェックし、あった場合は変数に「1」をセットする
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "G40P", "G45P", "G49P", "G50P", "G59P"
                                        Option2_Kbn = "1"
                                End Select
                            Next

                        End If

                        '各区分に「1」が入っていた場合falseで返す
                        If Option1_Kbn = "1" And Option2_Kbn = "1" Then
                            Return False
                        End If

                    Case "C2000", "C2010", "C2020", "C2030", "C2040", "C2050", "C2500", "C2520", "C2530", "C2550"

                        '要素3「オプション」が「L」の時、要素7「配管アダプタセット・アタッチメント(添付)」の「G40P」「G45P」「G49P」が選択不可

                        '変数の初期化
                        Option1_Kbn = ""
                        Option2_Kbn = ""

                        '要素3の中に「L」があるかどうかチェックし、あった場合は変数に「1」をセットする
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "L"
                                    Option1_Kbn = "1"
                            End Select
                        Next

                        '要素3の中に「L」があった場合のみ以下の処理を実施
                        If Option1_Kbn = "1" Then

                            '要素7の中に「G40P」「G45P」「G49P」があるかどうかチェックし、あった場合は変数に「1」をセットする
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "G40P", "G45P", "G49P"
                                        Option2_Kbn = "1"
                                End Select
                            Next

                        End If

                        '各区分に「1」が入っていた場合falseで返す
                        If Option1_Kbn = "1" And Option2_Kbn = "1" Then
                            Return False
                        End If

                    Case "R1000", "R1100", "R3000", "R3100", "R4000", "R4100", "R8000", "R8100"

                        '要素3「オプション」が「L」の時、要素5「配管アダプタセット・アタッチメント(添付)」の「G40P」「G45P」「G49P」「G50P」「G59P」が選択不可

                        '変数の初期化
                        Option1_Kbn = ""
                        Option2_Kbn = ""

                        '要素3の中に「L」があるかどうかチェックし、あった場合は変数に「1」をセットする
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "L"
                                    Option1_Kbn = "1"
                            End Select
                        Next

                        '要素3の中に「L」があった場合のみ以下の処理を実施
                        If Option1_Kbn = "1" Then

                            '要素5の中に「G40P」「G45P」「G49P」「G50P」「G59P」があるかどうかチェックし、あった場合は変数に「1」をセットする
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "G40P", "G45P", "G49P", "G50P", "G59P"
                                        Option2_Kbn = "1"
                                End Select
                            Next

                        End If

                        '各区分に「1」が入っていた場合falseで返す
                        If Option1_Kbn = "1" And Option2_Kbn = "1" Then
                            Return False
                        End If

                    Case "R2000", "R2100"

                        '要素3「オプション」が「L」の時、要素6「配管アダプタセット・アタッチメント(添付)」の「GP40P」「GP45P」「GP49P」が選択不可

                        '変数の初期化
                        Option1_Kbn = ""
                        Option2_Kbn = ""

                        '要素3の中に「L」があるかどうかチェックし、あった場合は変数に「1」をセットする
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "L"
                                    Option1_Kbn = "1"
                            End Select
                        Next

                        '要素3の中に「L」があった場合のみ以下の処理を実施
                        If Option1_Kbn = "1" Then

                            '要素6の中に「G40P」「G45P」「G49P」があるかどうかチェックし、あった場合は変数に「1」をセットする
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "G40P", "G45P", "G49P"
                                        Option2_Kbn = "1"
                                End Select
                            Next

                        End If

                        '各区分に「1」が入っていた場合falseで返す
                        If Option1_Kbn = "1" And Option2_Kbn = "1" Then
                            Return False
                        End If

                    Case "W1000", "W1100", "W3000", "W3100", "W4000", "W4100"

                        '要素3「オプション」が「L」の時、要素7「配管アダプタセット・アタッチメント(添付)」の「G40P」「G45P」「G49P」「G50P」「G59P」が選択不可

                        '変数の初期化
                        Option1_Kbn = ""
                        Option2_Kbn = ""

                        '要素3の中に「L」があるかどうかチェックし、あった場合は変数に「1」をセットする
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "L"
                                    Option1_Kbn = "1"
                            End Select
                        Next

                        '要素3の中に「L」があった場合のみ以下の処理を実施
                        If Option1_Kbn = "1" Then

                            '要素7の中に「G40P」「G45P」「G49P」「G50P」「G59P」があるかどうかチェックし、あった場合は変数に「1」をセットする
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "G40P", "G45P", "G49P", "G50P", "G59P"
                                        Option2_Kbn = "1"
                                End Select
                            Next

                        End If

                        '各区分に「1」が入っていた場合falseで返す
                        If Option1_Kbn = "1" And Option2_Kbn = "1" Then
                            Return False
                        End If

                    Case "W2000", "W2100"

                        '要素3「オプション」が「L」の時、要素7「配管アダプタセット・アタッチメント(添付)」の「G40P」「G45P」「G49P」が選択不可

                        '変数の初期化
                        Option1_Kbn = ""
                        Option2_Kbn = ""

                        '要素3の中に「L」があるかどうかチェックし、あった場合は変数に「1」をセットする
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "L"
                                    Option1_Kbn = "1"
                            End Select
                        Next

                        '要素3の中に「L」があった場合のみ以下の処理を実施
                        If Option1_Kbn = "1" Then

                            '要素7の中に「G40P」「G45P」「G49P」があるかどうかチェックし、あった場合は変数に「1」をセットする
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "G40P", "G45P", "G49P"
                                        Option2_Kbn = "1"
                                End Select
                            Next

                        End If

                        '各区分に「1」が入っていた場合falseで返す
                        If Option1_Kbn = "1" And Option2_Kbn = "1" Then
                            Return False
                        End If

                    Case "W8000", "W8100"

                        '要素3「オプション」が「L」の時、要素5「配管アダプタセット・アタッチメント(添付)」の「G40P」「G45P」「G49P」「G50P」「G59P」が選択不可

                        '変数の初期化
                        Option1_Kbn = ""
                        Option2_Kbn = ""

                        '要素3の中に「L」があるかどうかチェックし、あった場合は変数に「1」をセットする
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "L"
                                    Option1_Kbn = "1"
                            End Select
                        Next

                        '要素3の中に「L」があった場合のみ以下の処理を実施
                        If Option1_Kbn = "1" Then

                            '要素5の中に「G40P」「G45P」「G49P」「G50P」「G59P」があるかどうかチェックし、あった場合は変数に「1」をセットする
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "G40P", "G45P", "G49P", "G50P", "G59P"
                                        Option2_Kbn = "1"
                                End Select
                            Next

                        End If

                        '各区分に「1」が入っていた場合falseで返す
                        If Option1_Kbn = "1" And Option2_Kbn = "1" Then
                            Return False
                        End If

                        'FRL禁則条件追加  RM1703020  2017/03/14 追加  <---------------------------------------------------------

                    Case "DSC"     'RM1703022　DSC制御追加　2017/7/19
                        Select Case .strKeyKataban
                            Case ""
                                If .strOpSymbol(5) <> "" Then
                                    Return False
                                End If
                            Case "C"
                                If (.strOpSymbol(3) = "3" Or _
                                    .strOpSymbol(3) = "8") And
                                    .strOpSymbol(5) = "L" Then
                                    Return False
                                End If
                        End Select
                        'RM1802***_限定発売対応解除
                        'Case "GRC"     'RM1707052　上海向け限定発売対応

                        '    If Not strCountryCd.Equals("PRC") Then
                        '        Return False
                        '    End If

                    Case "STG-B"    'RM1712***　中国生産品制御追加

                        'RM1807097_禁則削除
                        'If .strKeyKataban = "" Then
                        '    '要素２が12or20の場合のみ要素９のP72,P73生産可
                        '    If .strOpSymbol(2) <> "12" And .strOpSymbol(2) <> "20" Then
                        '        If .strOpSymbol(9) = "P72" Or .strOpSymbol(9) = "P73" Then
                        '            Return False
                        '        End If
                        '    End If
                        'End If

                    Case "PWC"    'RM1804035_PWC生産国禁則追加
                        '上海販売限定（中国限定）
                        If strCountryCd <> "PRC" Then
                            Return False
                        End If

                End Select
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 生産国の判断ロジック(タイ)
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetData_Logic_Thailand(ByVal objKtbnStrc As KHKtbnStrc, ByVal strCountryCd As String) As Boolean
        fncGetData_Logic_Thailand = True
        Try
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "SSD"    'ADD BY YGY 2014/09/08
                    'ASEAN販売限定
                    Dim strAseanCountry As List(Of String) = CdCst.strAseanCode
                    If Not strAseanCountry.Contains(strCountryCd) Then
                        Return False
                    Else
                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case ""
                                '要素9「S1:スイッチ」が「T1V,T3WH,T3WV,T3YH,T8V」の時、要素10「S1:リード線長さ」が
                                '1mのみ生産可能
                                If objKtbnStrc.strcSelection.strOpSymbol(9) = "T1V" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(9) = "T3WH" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(9) = "T3WV" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(9) = "T3YH" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(9) = "T8V" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(10) <> "" Then
                                        Return False
                                    End If
                                End If
                                '要素9「S1:スイッチ」が「T1H,T2YV」の時、要素10「S1:リード線長さ」が
                                '1m,3mのみ生産可能
                                If objKtbnStrc.strcSelection.strOpSymbol(9) = "T1H" OrElse _
                                       objKtbnStrc.strcSelection.strOpSymbol(9) = "T2YV" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(10) <> "" AndAlso _
                                       objKtbnStrc.strcSelection.strOpSymbol(10) <> "3" Then
                                        Return False
                                    End If
                                End If
                                '要素16「S2:スイッチ」が「T1V,T3WH,T3WV,T3YH,T8V」の時、要素17「S2:リード線長さ」が
                                '1mのみ生産可能
                                If objKtbnStrc.strcSelection.strOpSymbol(16) = "T1V" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(16) = "T3WH" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(16) = "T3WV" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(16) = "T3YH" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(16) = "T8V" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(17) <> "" Then
                                        Return False
                                    End If
                                End If
                                '要素16「S2:スイッチ」が「T1H,T2YV」の時、要素17「S2:リード線長さ」が
                                '1m,3mのみ生産可能
                                If objKtbnStrc.strcSelection.strOpSymbol(16) = "T1H" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(16) = "T2YV" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(17) <> "" AndAlso _
                                       objKtbnStrc.strcSelection.strOpSymbol(17) <> "3" Then
                                        Return False
                                    End If
                                End If
                            Case "K"
                                '要素8「S1:スイッチ」が「T1V,T3WH,T3WV,T3YH,T8V」の時、要素9「S1:リード線長さ」が
                                '1mのみ生産可能
                                If objKtbnStrc.strcSelection.strOpSymbol(8) = "T1V" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(8) = "T3WH" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(8) = "T3WV" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(8) = "T3YH" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(8) = "T8V" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(9) <> "" Then
                                        Return False
                                    End If
                                End If
                                '要素8「S1:スイッチ」が「T1H,T2YV」の時、要素9「S1:リード線長さ」が
                                '1m,3mのみ生産可能
                                If objKtbnStrc.strcSelection.strOpSymbol(8) = "T1H" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(8) = "T2YV" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(9) <> "" AndAlso _
                                       objKtbnStrc.strcSelection.strOpSymbol(9) <> "3" Then
                                        Return False
                                    End If
                                End If
                                '要素14「S2:スイッチ」が「T1V,T3WH,T3WV,T3YH,T8V」の時、要素15「S2:リード線長さ」が
                                '1mのみ生産可能
                                If objKtbnStrc.strcSelection.strOpSymbol(14) = "T1V" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(14) = "T3WH" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(14) = "T3WV" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(14) = "T3YH" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(14) = "T8V" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(15) <> "" Then
                                        Return False
                                    End If
                                End If
                                '要素14「S2:スイッチ」が「T1H,T2YV」の時、要素15「S2:リード線長さ」が
                                '1m,3mのみ生産可能
                                If objKtbnStrc.strcSelection.strOpSymbol(14) = "T1H" OrElse _
                                   objKtbnStrc.strcSelection.strOpSymbol(14) = "T2YV" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(15) <> "" AndAlso _
                                       objKtbnStrc.strcSelection.strOpSymbol(15) <> "3" Then
                                        Return False
                                    End If
                                End If
                                'ADD BY YGY 20141202
                                '要素4「口径」が20以下の時、要素7「ストローク」と要素13「ストローク」が最大50
                                '要素4「口径」が25以上の時、要素7「ストローク」と要素13「ストローク」が最大100、最小10
                                Dim strKouKei As String = objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Dim strStroke1 As String = objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Dim strStroke2 As String = objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                Dim intKouKei As Integer = 0
                                Dim intStroke1 As Integer = 0
                                Dim intStroke2 As Integer = 0

                                If Not strKouKei.Equals(String.Empty) Then
                                    intKouKei = Integer.Parse(strKouKei)
                                End If
                                If Not strStroke1.Equals(String.Empty) Then
                                    intStroke1 = Integer.Parse(strStroke1)
                                End If
                                If Not strStroke2.Equals(String.Empty) Then
                                    intStroke2 = Integer.Parse(strStroke2)
                                End If

                                If intKouKei <= 20 AndAlso intKouKei <> 0 Then
                                    '要素4「口径」が20以下の時、要素7「ストローク」と要素13「ストローク」が最大50
                                    If intStroke1 > 50 OrElse intStroke2 > 50 Then
                                        Return False
                                    End If
                                ElseIf intKouKei >= 25 Then
                                    '要素4「口径」が25以上の時、要素7「ストローク」と要素13「ストローク」が最大100、最小10
                                    If intStroke1 > 100 OrElse intStroke2 > 100 OrElse _
                                       (intStroke1 < 10 AndAlso intStroke1 > 0) OrElse _
                                       (intStroke2 < 10 AndAlso intStroke2 > 0) Then
                                        Return False
                                    End If
                                End If
                        End Select
                    End If
                Case "CMK2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            '要素6
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6)
                                Case "T3PH"

                                    '要素6「S1:スイッチ」が「T3PH」の時、要素7「S1:リード線長さ」が
                                    '3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If

                                Case "T1H"

                                    '要素6「S1:スイッチ」が「T1H」の時、要素7「S1:リード線長さ」が
                                    '1mと5mのみ生産可能 RM1701034修正
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(7) <> "5" Then
                                        Return False
                                    End If

                                Case "T3WH", "T3YH"

                                    '要素6「S1:スイッチ」が「T3WH」と「T3YH」の時、要素7「S1:リード線長さ」が
                                    '1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" Then
                                        Return False
                                    End If

                                Case "T2YV"

                                    '要素6「S1:スイッチ」が「T2YV」の時、要素7「S1:リード線長さ」が
                                    '1m,3mのみ生産可能   2017/03/22 追加
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If

                            End Select

                            '要素12
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12)
                                Case "T3PH"

                                    '要素12「S2:スイッチ」が「T3PH」の時、要素13「S2:リード線長さ」が
                                    '3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(13) <> "3" Then
                                        Return False
                                    End If

                                Case "T1H"

                                    '要素12「S2:スイッチ」が「T1H」の時、要素14「S2:リード線長さ」が
                                    '1mと3mのみ生産可能 RM1701034修正
                                    If objKtbnStrc.strcSelection.strOpSymbol(13) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(13) <> "5" Then
                                        Return False
                                    End If

                                Case "T3WH", "T3YH"

                                    '要素12「S2:スイッチ」が「T3WH」「T3YH」の時、要素14「S2:リード線長さ」が
                                    '1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(13) <> "" Then
                                        Return False
                                    End If

                                Case "T2YV"

                                    '要素12「S1:スイッチ」が「T2YV」の時、要素13「S1:リード線長さ」が
                                    '1m,3mのみ生産可能   2017/03/22 追加
                                    If objKtbnStrc.strcSelection.strOpSymbol(13) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(13) <> "3" Then
                                        Return False
                                    End If

                            End Select

                        Case "D"
                            '要素7「S1:スイッチ」が「T3PH」の時、要素8「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(7) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(8) <> "3" Then
                                    Return False
                                End If
                            End If

                            '禁則追加  2017/02/13 追加  -------------------------------------------------------------------------------

                            '要素7「S1:スイッチ」が「T1H」の時、要素8「S1:リード線長さ」が
                            '1m・5mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(7) = "T1H" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(8) <> "" And objKtbnStrc.strcSelection.strOpSymbol(8) <> "5" Then
                                    Return False
                                End If
                            End If
                            '要素7「S1:スイッチ」が「T3WH」または「T3YH」の時、要素8「S1:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(7) = "T3WH" Or objKtbnStrc.strcSelection.strOpSymbol(7) = "T3YH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(8) <> "" Then
                                    Return False
                                End If
                            End If

                            '禁則追加  2017/02/13 追加  -------------------------------------------------------------------------------


                            '禁則追加  2017/03/22 追加  -------------------------------------------------------------------------------

                            '要素7「S1:スイッチ」が「T2YV」の時、要素8「S1:リード線長さ」が
                            '1m,3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(7) = "T2YV" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(8) <> "" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(8) <> "3" Then
                                    Return False
                                End If
                            End If

                            '禁則追加  2017/03/22 追加  -------------------------------------------------------------------------------


                            '禁則追加  2017/03/22 追加  -------------------------------------------------------------------------------

                        Case "5"

                            '要素6
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6)
                                Case "T3PH"

                                    '要素6「S1:スイッチ」が「T3PH」の時、要素7「S1:リード線長さ」が
                                    '3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If

                                Case "T3PV"

                                    '要素6「S1:スイッチ」が「T3PV」の時、要素7「S1:リード線長さ」が
                                    '1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" Then
                                        Return False
                                    End If

                                Case "T1H"

                                    '要素6「S1:スイッチ」が「T1H」の時、要素7「S1:リード線長さ」が
                                    '1mと5mのみ生産可能 RM1701034修正
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(7) <> "5" Then
                                        Return False
                                    End If

                                Case "T3WH", "T3YH"

                                    '要素6「S1:スイッチ」が「T3WH」と「T3YH」の時、要素7「S1:リード線長さ」が
                                    '1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" Then
                                        Return False
                                    End If

                                Case "T2YV"

                                    '要素6「S1:スイッチ」が「T2YV」の時、要素7「S1:リード線長さ」が
                                    '1m,3mのみ生産可能   2017/03/22 追加
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If

                            End Select

                            '要素12
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12)
                                Case "T3PH"

                                    '要素12「S2:スイッチ」が「T3PH」の時、要素13「S2:リード線長さ」が
                                    '3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(13) <> "3" Then
                                        Return False
                                    End If

                                Case "T3PV"

                                    '要素12「S1:スイッチ」が「T3PV」の時、要素13「S1:リード線長さ」が
                                    '1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(13) <> "" Then
                                        Return False
                                    End If

                                Case "T1H"

                                    '要素12「S2:スイッチ」が「T1H」の時、要素14「S2:リード線長さ」が
                                    '1mと3mのみ生産可能 RM1701034修正
                                    If objKtbnStrc.strcSelection.strOpSymbol(13) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(13) <> "5" Then
                                        Return False
                                    End If

                                Case "T3WH", "T3YH"

                                    '要素12「S2:スイッチ」が「T3WH」「T3YH」の時、要素14「S2:リード線長さ」が
                                    '1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(13) <> "" Then
                                        Return False
                                    End If

                                Case "T2YV"

                                    '要素12「S1:スイッチ」が「T2YV」の時、要素13「S1:リード線長さ」が
                                    '1m,3mのみ生産可能   2017/03/22 追加
                                    If objKtbnStrc.strcSelection.strOpSymbol(13) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(13) <> "3" Then
                                        Return False
                                    End If

                            End Select

                            '禁則追加  2017/03/22 追加  -------------------------------------------------------------------------------

                    End Select
                Case "SCG"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            '要素7「S1:スイッチ」が「T3PH」の時、要素8「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(7) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(8) <> "3" Then
                                    Return False
                                End If
                            End If
                    End Select
                Case "SCA2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            '要素10「S1:スイッチ」が「T3PH」の時、要素11「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "3" Then
                                    Return False
                                End If
                            End If

                            '禁則追加  2017/03/22 追加  ------------------------------------------------------------------------------->

                            '要素10「S1:スイッチ」が「T1H」の時、要素11「S1:リード線長さ」が
                            '1m・5mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T1H" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "" And objKtbnStrc.strcSelection.strOpSymbol(11) <> "5" Then
                                    Return False
                                End If
                            End If

                            '要素10「S1:スイッチ」が「T3WH」または「T3YH」の時、要素11「S1:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T3WH" Or objKtbnStrc.strcSelection.strOpSymbol(10) = "T3YH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "" Then
                                    Return False
                                End If
                            End If

                            '要素10「S1:スイッチ」が「T2YV」の時、要素11「S1:リード線長さ」が
                            '1m,3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T2YV" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(11) <> "3" Then
                                    Return False
                                End If
                            End If

                            '禁則追加  2017/03/22 追加  <-------------------------------------------------------------------------------


                            '禁則追加  2017/03/22 追加  ------------------------------------------------------------------------------->

                        Case "2"

                            '要素10「S1:スイッチ」が「T1H」の時、要素11「S1:リード線長さ」が
                            '1m・5mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T1H" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "" And objKtbnStrc.strcSelection.strOpSymbol(11) <> "5" Then
                                    Return False
                                End If
                            End If

                            '要素10「S1:スイッチ」が「T3WH」または「T3YH」の時、要素11「S1:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T3WH" Or objKtbnStrc.strcSelection.strOpSymbol(10) = "T3YH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "" Then
                                    Return False
                                End If
                            End If

                            '要素10「S1:スイッチ」が「T2YV」の時、要素11「S1:リード線長さ」が
                            '1m,3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T2YV" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(11) <> "3" Then
                                    Return False
                                End If
                            End If

                            '要素10「S1:スイッチ」が「T3PH」の時、要素11「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "3" Then
                                    Return False
                                End If
                            End If

                            '要素10「S1:スイッチ」が「T3PV」の時、要素11「S1:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T3PV" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "" Then
                                    Return False
                                End If
                            End If


                            '禁則追加  2017/03/22 追加  <-------------------------------------------------------------------------------

                        Case "B"
                            '要素8「S1:スイッチ」が「T3PH」の時、要素9「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(8) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(9) <> "3" Then
                                    Return False
                                End If
                            End If
                            '要素14「S2:スイッチ」が「T3PH」の時、要素15「S2:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(14) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(15) <> "3" Then
                                    Return False
                                End If
                            End If
                        Case "D"
                            '要素9「S1:スイッチ」が「T3PH」の時、要素10「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(9) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(10) <> "3" Then
                                    Return False
                                End If
                            End If
                        Case "V"
                            '要素10「S1:スイッチ」が「T3PH」の時、要素11「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(10) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(11) <> "3" Then
                                    Return False
                                End If
                            End If
                    End Select


                    '禁則追加  2017/03/22 追加  <-------------------------------------------------------------------------------



                    '禁則追加  2017/03/22 追加  ------------------------------------------------------------------------------->

                Case "SCS2"

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "", "F"

                            '要素14「S1:スイッチ」が「T1H」の時、要素15「S1:リード線長さ」が
                            '1m・5mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(14) = "T1H" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(15) <> "" And objKtbnStrc.strcSelection.strOpSymbol(15) <> "5" Then
                                    Return False
                                End If
                            End If

                            '要素14「S1:スイッチ」が「T3WH」または「T3YH」の時、要素15「S1:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(14) = "T3WH" Or objKtbnStrc.strcSelection.strOpSymbol(14) = "T3YH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(15) <> "" Then
                                    Return False
                                End If
                            End If

                            '要素14「S1:スイッチ」が「T2YV」の時、要素15「S1:リード線長さ」が
                            '1m,3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(14) = "T2YV" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(15) <> "" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(15) <> "3" Then
                                    Return False
                                End If
                            End If

                    End Select

                    '禁則追加  2017/03/22 追加  <-------------------------------------------------------------------------------


                Case "SSD2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            '要素9「S1:スイッチ」が「T3PH」の時、要素10「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(9) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(10) <> "3" Then
                                    Return False
                                End If
                            End If
                            '要素16「S2:スイッチ」が「T3PH」の時、要素17「S2:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) <> "3" Then
                                    Return False
                                End If
                            End If

                            '禁則事項を追加 2017/01/17 追加 RM1701034  -------------------------------->

                            '要素16「S2:スイッチ」が「T1H」の時、要素17「S2:リード線長さ」が
                            '3mは生産不可
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T1H" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "3" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T2YV」の時、要素17「S2:リード線長さ」が
                            '5mは生産不可
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T2YV" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "5" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T3PH」の時、要素17「S2:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) <> "3" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T3PV」の時、要素17「S2:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3PV" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "3" OrElse objKtbnStrc.strcSelection.strOpSymbol(17) = "5" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T3WH」の時、要素17「S2:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3WH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "3" OrElse objKtbnStrc.strcSelection.strOpSymbol(17) = "5" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T3YH」の時、要素17「S2:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3YH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "3" OrElse objKtbnStrc.strcSelection.strOpSymbol(17) = "5" Then
                                    Return False
                                End If
                            End If

                            '禁則事項を追加 2017/01/17 追加 RM1701034  <--------------------------------

                        Case "2"
                            '要素6「S1:スイッチ」が「T3PH」の時、要素7「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(6) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                    Return False
                                End If
                            End If
                        Case "D"
                            '要素6「S1:スイッチ」が「T3PH」の時、要素7「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(6) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                    Return False
                                End If
                            End If
                        Case "K"
                            '要素16「S1:スイッチ」が「T3PH」の時、要素17「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) <> "3" Then
                                    Return False
                                End If
                            End If

                            '禁則事項を追加  2017/03/23 追加  >---------------------------------------

                        Case "7"

                            '要素16「S2:スイッチ」が「T1H」の時、要素17「S2:リード線長さ」が
                            '3mは生産不可
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T1H" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "3" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T2YV」の時、要素17「S2:リード線長さ」が
                            '5mは生産不可
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T2YV" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "5" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T3PH」の時、要素17「S2:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) <> "3" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T3PV」の時、要素17「S2:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3PV" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "3" OrElse objKtbnStrc.strcSelection.strOpSymbol(17) = "5" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T3WH」の時、要素17「S2:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3WH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "3" OrElse objKtbnStrc.strcSelection.strOpSymbol(17) = "5" Then
                                    Return False
                                End If
                            End If

                            '要素16「S2:スイッチ」が「T3YH」の時、要素17「S2:リード線長さ」が
                            '1mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(16) = "T3YH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(17) = "3" OrElse objKtbnStrc.strcSelection.strOpSymbol(17) = "5" Then
                                    Return False
                                End If
                            End If

                            '禁則事項を追加  2017/03/23 追加  ---------------------------------------<



                    End Select

                Case "LCR"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "2"
                            '要素5「S1:スイッチ」が「T3PH」の時、要素6「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(5) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(6) <> "3" Then
                                    Return False
                                End If
                            End If
                    End Select
                Case "RCC2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            '要素5「S1:スイッチ」が「T3PH」の時、要素6「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(5) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(6) <> "3" Then
                                    Return False
                                End If
                            End If
                    End Select
                Case "RCC2-G4"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            '要素5「S1:スイッチ」が「T3PH」の時、要素6「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(5) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(6) <> "3" Then
                                    Return False
                                End If
                            End If
                    End Select
                Case "SCG-G4"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            '要素7「S1:スイッチ」が「T3PH」の時、要素8「S1:リード線長さ」が
                            '3mのみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(7) = "T3PH" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(8) <> "3" Then
                                    Return False
                                End If
                            End If
                    End Select

                    'RM1807132_タイ生産禁則変更（スイッチからT1H,T1V,T3WV,T8V削除）
                Case "JSC3"
                    'ADD BY YGY 20141006
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "1"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(10)
                                Case "H0", "H0Y", "T3WH", "T3YH"
                                    '要素10「スイッチ」が「H0,H0Y,T3WH,T3YH」の時、
                                    '要素11「リード線長さ」が1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(11) <> "" Then
                                        Return False
                                    End If
                                Case "T2YV"
                                    '要素10「スイッチ」が「T2YV」の時、
                                    '要素11「リード線長さ」が1mと3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(11) <> "" AndAlso _
                                       objKtbnStrc.strcSelection.strOpSymbol(11) <> "3" Then
                                        Return False
                                    End If
                            End Select
                    End Select
                Case "STG-B"
                    'ADD BY YGY 20150306
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            '要素1
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                                Case "G4"
                                    '要素1「バリエーション」が「G4」の時、
                                    '要素6「スイッチ形番」が「T2YD,T2YDT,T2YDU」のみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(6) <> "T2YD" AndAlso _
                                       objKtbnStrc.strcSelection.strOpSymbol(6) <> "T2YDT" AndAlso _
                                       objKtbnStrc.strcSelection.strOpSymbol(6) <> "T2YDU" Then
                                        Return False
                                    End If
                            End Select

                            '要素6
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6)
                                Case "T2YV"
                                    '要素6「スイッチ形番」が「T2YV」の時、
                                    '要素「7」が1mと3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso _
                                       objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If
                                Case "T3PV", "T3WH"
                                    '要素6「スイッチ形番」が「T3PV」「T3WH」の時、
                                    '要素「7」が1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" Then
                                        Return False
                                    End If

                                    '禁則事項の追加  2017/03/24 追加   >-----------------------------------------------

                                Case "T1H"
                                    '要素6「スイッチ形番」が「T1H」の時、
                                    '要素「7」が1mと5mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso _
                                       objKtbnStrc.strcSelection.strOpSymbol(7) <> "5" Then
                                        Return False
                                    End If

                                Case "T3PH"
                                    '要素6「スイッチ形番」が「T3PH」の時、
                                    '要素「7」が3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If

                                Case "T3YH"
                                    '要素6「スイッチ形番」が「T3YH」の時、
                                    '要素「7」が1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" Then
                                        Return False
                                    End If

                                    '禁則事項の追加  2017/03/24 追加   <-----------------------------------------------

                            End Select

                            '要素7
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7)
                                Case ""
                                    '要素7「リード線」が「」の時、
                                    '要素6「スイッチ形番」が「T3PH」を生産不可
                                    If objKtbnStrc.strcSelection.strOpSymbol(6) = "T3PH" Then
                                        Return False
                                    End If
                                Case "3"
                                    '要素7「リード線」が「3」の時、
                                    '要素6「スイッチ形番」が「T1H,T2YDT,T3YH」を生産不可
                                    If objKtbnStrc.strcSelection.strOpSymbol(6) = "T1H" OrElse _
                                       objKtbnStrc.strcSelection.strOpSymbol(6) = "T2YDT" OrElse _
                                       objKtbnStrc.strcSelection.strOpSymbol(6) = "T3YH" Then
                                        Return False
                                    End If
                                Case "5"
                                    '要素7「リード線」が「5」の時、
                                    '要素6「スイッチ形番」が「T1H,T2YDT,T3YH,T2YV,T3PH」を生産不可
                                    '要素6「スイッチ形番」の「T1H」は生産可能なので変更  2017/03/24 
                                    If objKtbnStrc.strcSelection.strOpSymbol(6) = "T2YDT" OrElse _
                                       objKtbnStrc.strcSelection.strOpSymbol(6) = "T3YH" OrElse _
                                       objKtbnStrc.strcSelection.strOpSymbol(6) = "T2YV" OrElse _
                                       objKtbnStrc.strcSelection.strOpSymbol(6) = "T3PH" Then
                                        Return False
                                    End If
                            End Select

                    End Select
                Case "STG-M"

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6)
                                Case "T2YV"
                                    '要素6「スイッチ形番」が「T2YV」の時、
                                    '要素「7」が1mと3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso
                                        objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If

                                Case "T3PH"
                                    '要素6「スイッチ形番」が「T3PH」の時、
                                    '要素「7」が3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If

                                Case "T3PV", "T3WH", "T3YH"
                                    '要素6「スイッチ形番」が「T3PV」「T3WH」「T3YH」の時、
                                    '要素「7」が1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" Then
                                        Return False
                                    End If

                                    '2017/03/24 追加 禁則追加のため >---------------------------------------------

                                Case "T1H"
                                    '要素6「スイッチ形番」が「T1H」の時、
                                    '要素「7」が1mと5mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso
                                        objKtbnStrc.strcSelection.strOpSymbol(7) <> "5" Then
                                        Return False
                                    End If

                                    '2017/03/24 追加 禁則追加のため <---------------------------------------------

                            End Select

                            '2017/03/24 追加 禁則追加のため >---------------------------------------------

                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)

                                Case "G4"

                                    '要素1が「G4」の時、
                                    '要素「6」が「T2YD」「T2YDT」「T2YDU」のみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(6) <> "T2YD" AndAlso
                                        objKtbnStrc.strcSelection.strOpSymbol(6) <> "T2YDT" AndAlso
                                        objKtbnStrc.strcSelection.strOpSymbol(6) <> "T2YDU" Then
                                        Return False
                                    End If

                            End Select

                            '2017/03/24 追加 禁則追加のため <---------------------------------------------



                            '2017/03/23 追加 禁則追加のため >---------------------------------------------

                        Case "F"

                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6)
                                Case "T1H"
                                    '要素6「スイッチ形番」が「T1H」の時、
                                    '要素「7」が1mと5mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso
                                        objKtbnStrc.strcSelection.strOpSymbol(7) <> "5" Then
                                        Return False
                                    End If

                                Case "T2YV"
                                    '要素6「スイッチ形番」が「T2YV」の時、
                                    '要素「7」が1mと3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" AndAlso
                                        objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If

                                Case "T3PH"
                                    '要素6「スイッチ形番」が「T3PH」の時、
                                    '要素「7」が3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "3" Then
                                        Return False
                                    End If

                                Case "T3PV", "T3WH", "T3YH"
                                    '要素6「スイッチ形番」が「T3PV」「T3WH」「T3YH」の時、
                                    '要素「7」が1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> "" Then
                                        Return False
                                    End If

                            End Select

                            '2017/03/23 追加 禁則追加のため ---------------------------------------------<

                    End Select


                Case "PV5", "PV5G"
                    'ASEAN販売限定
                    Dim strAseanCountry As List(Of String) = CdCst.strAseanCode
                    If Not strAseanCountry.Contains(strCountryCd) Then
                        Return False
                    End If
                Case "CAC4"

                    Select Case objKtbnStrc.strcSelection.strKeyKataban

                        Case ""
                            'スイッチ形番の場合
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(9)

                                Case "T1H", "T2YDT", "T3YH"

                                    '"T1H", "T2YDT", "T3YH"の場合は「リード線長さ」が1mだけOK
                                    If objKtbnStrc.strcSelection.strOpSymbol(10) <> "" Then

                                        Return False

                                    End If

                                Case "T3PH"

                                    '"T3PH"の場合は「リード線長さ」が3mだけOK
                                    If objKtbnStrc.strcSelection.strOpSymbol(10) <> "3" Then

                                        Return False

                                    End If

                                Case "T3PV", "T3WH"

                                    '"T3PV"、"T3WH"の場合は「リード線長さ」が1mだけOK
                                    If objKtbnStrc.strcSelection.strOpSymbol(10) <> "" Then

                                        Return False

                                    End If

                            End Select

                    End Select

                Case "SRL3"

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7)

                                Case "M3V", "T3WH", "T3YH"
                                    '要素7「スイッチ形番」が「M3V」「T3WH」「T3YH」の時、
                                    '要素「8」が1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(8) <> "" Then
                                        Return False
                                    End If

                                Case "M3WV", "T2YV"
                                    '要素7「スイッチ形番」が「M3WV」「T2YV」の時、
                                    '要素「8」が1mと3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(8) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(8) <> "3" Then
                                        Return False
                                    End If

                            End Select


                            '2017/03/23 追加 禁則追加のため >---------------------------------------------

                        Case "F"

                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7)

                                Case "M3V", "T3WH", "T3YH"
                                    '要素7「スイッチ形番」が「M3V」「T3WH」「T3YH」の時、
                                    '要素「8」が1mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(8) <> "" Then
                                        Return False
                                    End If

                                Case "M3WV", "T2YV"
                                    '要素7「スイッチ形番」が「M3WV」「T2YV」の時、
                                    '要素「8」が1mと3mのみ生産可能
                                    If objKtbnStrc.strcSelection.strOpSymbol(8) <> "" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(8) <> "3" Then
                                        Return False
                                    End If

                            End Select

                            '2017/03/23 追加 禁則追加のため ---------------------------------------------<


                    End Select


                    '2017/02/21 追加 禁則追加のため --------------------------------------------->

                Case "AB31", "AB41", "AG31", "AG33", "AG34", "AG41", "AG43", "AG44"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            'コイルハウジングが「3A」・電圧が「DC24V」の場合のみ生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(4) = "3A" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(10) = "DC24V" Then
                                Else
                                    Return False
                                End If
                            End If
                    End Select

                    '2017/02/21 追加 禁則追加のため <---------------------------------------------

                Case "STR2-M"    'RM1802***　タイ生産品制御追加

                    'RM1806042_生産国禁則追加
                    'ASEAN販売限定
                    Dim strAseanCountry As List(Of String) = CdCst.strAseanCode
                    If Not strAseanCountry.Contains(strCountryCd) Then
                        Return False
                    End If

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            'チューブ内径が「16、20、25、32」のみストローク「60～100」生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(2) = "6" Or objKtbnStrc.strcSelection.strOpSymbol(2) = "10" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(3) = "60" _
                                Or objKtbnStrc.strcSelection.strOpSymbol(3) = "70" _
                                Or objKtbnStrc.strcSelection.strOpSymbol(3) = "80" _
                                Or objKtbnStrc.strcSelection.strOpSymbol(3) = "90" _
                                Or objKtbnStrc.strcSelection.strOpSymbol(3) = "100" Then
                                    Return False
                                End If
                            End If
                    End Select

                Case "SCPD3", "SCPD3-L"

                    'RM1806042_生産国禁則追加
                    'ASEAN販売限定
                    Dim strAseanCountry As List(Of String) = CdCst.strAseanCode
                    If Not strAseanCountry.Contains(strCountryCd) Then
                        Return False
                    End If

                Case "PWC"    'RM1804035_PWC生産国禁則追加
                    'ASEAN販売限定
                    Dim strAseanCountry As List(Of String) = CdCst.strAseanCode
                    If Not strAseanCountry.Contains(strCountryCd) Then
                        Return False
                    End If

            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'RM1801038_インドネシア禁則対応追加
    ''' <summary>
    ''' 生産国の判断ロジック(インドネシア)
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetData_Logic_Indonesia(ByVal objKtbnStrc As KHKtbnStrc, ByVal strCountryCd As String) As Boolean
        fncGetData_Logic_Indonesia = True
        Try
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "SSD2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            '要素1で「G4」選択時、要素4が25,32,40,50,63で生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(1) = "G4" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(4) <> "25" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(4) <> "32" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(4) <> "40" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(4) <> "50" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(4) <> "63" Then
                                    Return False
                                End If
                            End If

                            '要素5で「D」選択時、要素4が40,50,63で生産可能
                            If objKtbnStrc.strcSelection.strOpSymbol(5) = "D" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(4) <> "40" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(4) <> "50" AndAlso _
                                   objKtbnStrc.strcSelection.strOpSymbol(4) <> "63" Then
                                    Return False
                                End If
                            End If
                    End Select
                Case "SCA2"     'RM1807***_バリエーション、ストローク禁則対応追加

                    '要素1「バリエーション」が「Q2,T」、要素4「口径」が「40,50,63」、要素7「ストローク」は600まで生産可能
                    '要素1「バリエーション」が「Q2,T」、要素4「口径」が「80」、要素7「ストローク」は700まで生産可能
                    '要素1「バリエーション」が「Q2,T」、要素4「口径」が「100」、要素7「ストローク」は800まで生産可能
                    Dim strvariation As String = objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    Dim strKouKei As String = objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Dim strStroke As String = objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    Dim intKouKei As Integer = 0
                    Dim intStroke As Integer = 0

                    If Not strKouKei.Equals(String.Empty) Then
                        intKouKei = Integer.Parse(strKouKei)
                    End If
                    If Not strStroke.Equals(String.Empty) Then
                        intStroke = Integer.Parse(strStroke)
                    End If

                    'バリエーションがＱ２、Ｔ
                    Select Case strvariation
                        Case "Q2", "T"
                            Select Case intKouKei
                                Case 40, 50, 63
                                    '要素4「口径」が40,50,63の時、要素7「ストローク」が最大600
                                    If intStroke > 600 Then
                                        Return False
                                    End If
                                Case 80
                                    '要素4「口径」が80の時、要素7「ストローク」が最大700
                                    If intStroke > 700 Then
                                        Return False
                                    End If
                                Case 100
                                    '要素4「口径」が100の時、要素7「ストローク」が最大800
                                    If intStroke > 800 Then
                                        Return False
                                    End If
                            End Select
                    End Select
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Class
