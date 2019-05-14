Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class KHStdDlv

#Region " Definition "

    Private strStdTntMain As String
    Private strStdTntMainTel As String
    Private strStdTntSub As String
    Private strStdTntSubTel As String

#End Region

    ''' <summary>
    ''' 標準納期取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strLangCd">言語コード</param>
    ''' <param name="strStdDlvDt">標準納期</param>
    ''' <param name="strQuantity">適用個数</param>
    ''' <remarks>引当てた形番より標準納期を取得する</remarks>
    Public Sub subStdDlvDtInfo(objCon As SqlConnection, ByVal strKataban As String, _
                               ByVal strLangCd As String, ByRef strStdDlvDt As String, _
                               ByRef strQuantity As String)
        Dim strMsgVale(0) As String
        Dim intStdDlvDt As Integer
        Dim intQuantity As Integer

        Try

            strStdDlvDt = ""
            strQuantity = ""

            '標準納期取得
            intStdDlvDt = Me.fncStdDlvDtGet(objCon, strKataban, intQuantity)

            '戻り値判定
            Select Case intStdDlvDt
                Case Is >= 0
                    '標準納期算出正常
                    Select Case intStdDlvDt
                        Case 90
                            '在庫対応
                            strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "I0080")
                            strQuantity = intQuantity.ToString
                        Case 91
                            '即日対応
                            strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "I0090")
                            strQuantity = intQuantity.ToString
                        Case 92
                            'AM I/Pのみ即日対応
                            strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "I0100")
                            strQuantity = intQuantity.ToString
                        Case 97, 98
                            '納期工場へ問い合わせ
                            strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "I0110")
                            strQuantity = ""
                        Case Else
                            'その他
                            strMsgVale(0) = intStdDlvDt.ToString
                            strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "I0120", strMsgVale)
                            strQuantity = intQuantity.ToString
                    End Select
                Case -1
                    '形番未入力エラー
                    strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "I0130")
                    strQuantity = ""
                Case -2
                    '形番分解エラー
                    strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "I0140")
                    strQuantity = ""
                Case -3
                    '標準納期算出エラー
                    strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "I0140")
                    strQuantity = ""
                Case -99
                    'その他エラー
                    strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "E9999")
                    strQuantity = ""
            End Select

        Catch ex As Exception
            'その他エラー
            strStdDlvDt = ClsCommon.fncGetMsg(strLangCd, "E9999")
            strQuantity = ""
        End Try

    End Sub

    ''' <summary>
    ''' 標準納期取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="intQuantity">適用個数</param>
    ''' <returns></returns>
    ''' <remarks>引当てた形番より標準納期を取得する</remarks>
    Public Function fncStdDlvDtGet(objCon As SqlConnection, ByVal strKataban As String, _
                                    ByRef intQuantity As Integer) As Integer

        Dim strSeriesKataban As String = Nothing
        Dim strErrorMessage As String = Nothing

        Dim strOptionSymbol(24) As String
        Dim strKtbnStrcNm(24) As String
        Dim intStdDate As Integer
        Dim intFixQty As Integer

        Try
            '形番分解
            If Me.fncKtbnResolution(objCon, strKataban, strSeriesKataban, strOptionSymbol, strKtbnStrcNm) Then
                '機種チェック
                If Not Me.fncModelChecker(objCon, strSeriesKataban, strOptionSymbol, strErrorMessage) Then
                    'システムエラー
                    fncStdDlvDtGet = -99
                    intQuantity = 0
                End If
                '標準納期取得
                If Me.fncStandardDate(objCon, strSeriesKataban, strOptionSymbol, intStdDate, intFixQty) Then
                    If strErrorMessage.Trim = "" Then
                        '標準納期設定
                        fncStdDlvDtGet = intStdDate
                        intQuantity = intFixQty
                    Else
                        '納期問合せ
                        fncStdDlvDtGet = 97
                        intQuantity = 0
                    End If
                Else
                    '標準納期計算エラー
                    fncStdDlvDtGet = -3
                    intQuantity = 0
                End If
            Else
                '分解エラー
                fncStdDlvDtGet = -2
                intQuantity = 0
            End If

            If fncStandardDateEx(objCon, strKataban, intStdDate, intFixQty) Then
                '例外標準納期コード設定
                fncStdDlvDtGet = intStdDate
                intQuantity = intFixQty
                Exit Function
            End If

        Catch ex As Exception
            'システムエラー
            fncStdDlvDtGet = -99
            intQuantity = 0
        End Try
    End Function

    ''' <summary>
    ''' 納期取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="intExCode">例外標準納期コード</param>
    ''' <returns>True:例外標準納期コード取得</returns>
    ''' <remarks></remarks>
    Private Function fncStandardDateEx(objCon As SqlConnection, ByVal strKataban As String, _
                                    ByRef intExDate As Integer, ByRef intExCode As Integer) As Boolean

        Dim dt As New DataTable
        Dim dalStdDlv As New StdDlvDAL
        Dim bolReturn As Boolean = False
        Try
            dt = dalStdDlv.fncStandardDateEx(objCon, strKataban)

            If dt.Rows.Count > 0 Then
                intExDate = CInt(dt.Rows(0).Item("ExceptionCode"))
                intExCode = CInt(dt.Rows(0).Item("shipment_qty"))
                bolReturn = True
            End If

        Catch ex As Exception
            '例外処理
            bolReturn = False
        End Try

        Return bolReturn

    End Function

    ''' <summary>
    ''' 形番分解
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strOptionSymbol">オプション記号</param>
    ''' <param name="strKtbnStrcNm">形番構成名称</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncKtbnResolution(objCon As SqlConnection, ByVal strKataban As String, _
                                       ByRef strSeriesKataban As String, _
                                       ByRef strOptionSymbol() As String, _
                                       ByRef strKtbnStrcNm() As String) As Boolean

        Dim intLoopCnt As Integer
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer

        Dim strHyphen(24) As String

        Dim intItemNo As Integer
        Dim intItemNoBackUp As Integer
        Dim intBlankLength As Integer
        Dim strModelWork As String
        Dim strSymbol As String
        Dim strSymbol2 As String
        Dim intLenSymbol As Integer
        Dim intPositionCounter As Integer
        Dim dalStdDlv As New StdDlvDAL
        fncKtbnResolution = False

        Try
            'シリーズテーブル検索
            If Not Me.fncSeriesTableSelect(objCon, strKataban, strSeriesKataban) Then
                Exit Function
            Else
                fncKtbnResolution = True
            End If

            '品名テーブル検索
            If Not Me.fncItemNameTableSelect(objCon, strSeriesKataban, strKtbnStrcNm, strHyphen) Then
                Exit Function
            End If

            'ハイフンテーブルの読み込み終わり
            strModelWork = strKataban
            intPositionCounter = 1

            Dim dt_Symbol As DataTable = dalStdDlv.fncGetAllSymbol(objCon, strSeriesKataban)
            For intItemNo = 1 To 24
                If strHyphen(intItemNo) = CdCst.Sign.Hypen Then
                    strSymbol = fncGetSymbol(strModelWork, strSeriesKataban, intItemNo, True, intBlankLength, dt_Symbol)
                Else
                    strSymbol = fncGetSymbol(strModelWork, strSeriesKataban, intItemNo, False, intBlankLength, dt_Symbol)
                End If

                intItemNoBackUp = intItemNo
                intLenSymbol = Len(strSymbol)

                If intLenSymbol = 0 Then
                    fncKtbnResolution = False
                    Exit For
                ElseIf Left(strSymbol, 1) = "!" Then
                    For intLoopCnt = intItemNo To intItemNo + intLenSymbol - 1
                        strOptionSymbol(intLoopCnt) = "!"
                    Next
                    intItemNo = intItemNo + intLenSymbol - 1
                Else
                    strOptionSymbol(intItemNo) = strSymbol

                    If intBlankLength > 0 Then
                        For intLoopCnt = intItemNo + intBlankLength To 24
                            If strHyphen(intLoopCnt) = CdCst.Sign.Hypen Then
                                strSymbol2 = fncGetSymbol(strModelWork, strSeriesKataban, intLoopCnt, True, intBlankLength, dt_Symbol)
                            Else
                                strSymbol2 = fncGetSymbol(strModelWork, strSeriesKataban, intLoopCnt, False, intBlankLength, dt_Symbol)
                            End If

                            If Len(strSymbol2) = 0 Then
                                Exit For
                            ElseIf Left(strSymbol2, 1) = "!" Then
                                strOptionSymbol(intLoopCnt) = "!"
                            Else
                                If Len(strSymbol2) > intLenSymbol Then
                                    strOptionSymbol(intLoopCnt) = strSymbol2
                                    strOptionSymbol(intItemNo) = "!"
                                    intItemNo = intLoopCnt
                                    strSymbol = strSymbol2
                                End If
                            End If
                        Next
                    End If
                End If

                If intItemNoBackUp > 1 Then
                    If strHyphen(intItemNoBackUp - 1) = "-" Then
                        If Mid(strKataban, intPositionCounter - 1, 1) <> "-" Then
                            fncKtbnResolution = False
                        End If
                    End If
                End If

                If Left(strModelWork, 2) = "DC" And strSymbol = "D" Then strSymbol = "!"

                If Left(strSymbol, 1) <> "!" Then
                    strModelWork = Mid(strModelWork, Len(strSymbol) + 1, 50)
                    intPositionCounter = intPositionCounter + Len(strSymbol)
                    If strHyphen(intItemNoBackUp) = CdCst.Sign.Hypen Then
                        If Left(strModelWork, 1) = CdCst.Sign.Hypen Then
                            strModelWork = Mid(strModelWork, 2, 50)
                            intPositionCounter = intPositionCounter + 1
                        Else
                            If strModelWork <> "" Then
                                fncKtbnResolution = False
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    If strHyphen(intItemNoBackUp) = CdCst.Sign.Hypen Then
                        If Left(strModelWork, 1) = CdCst.Sign.Hypen Then
                            strModelWork = Mid(strModelWork, 2, 50)
                            intPositionCounter = intPositionCounter + 1
                        End If
                    End If
                End If

                If strModelWork.Trim = "" Then
                    Exit For
                End If
            Next

            If strModelWork.Trim = "" Then
                For intLoopCnt1 = intItemNo + 1 To 24
                    intBlankLength = fncGetBlankCounter(strSeriesKataban, intLoopCnt1, dt_Symbol)
                    If intBlankLength = 0 Then
                        fncKtbnResolution = False
                        Exit For
                    ElseIf intBlankLength = -1 Then
                        Exit For
                    Else
                        For intLoopCnt2 = intLoopCnt1 To intLoopCnt1 + intBlankLength - 1
                            strOptionSymbol(intLoopCnt2) = "!"
                        Next
                        intLoopCnt1 = intLoopCnt2 - 1
                    End If
                Next
            Else
                fncKtbnResolution = False
            End If

        Catch ex As Exception
            fncKtbnResolution = False
        End Try

    End Function

    ''' <summary>
    ''' 形番分解
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSeriesTableSelect(objCon As SqlConnection, ByVal strKataban As String, _
                                          ByRef strSeriesKataban As String) As Boolean
        Dim dt As New DataTable
        Dim dalStdDlv As New StdDlvDAL

        Try
            dt = dalStdDlv.fncSeriesTableSelect(objCon, strKataban)

            If dt.Rows.Count > 0 Then

                strSeriesKataban = dt.Rows(0)("Series")

                strStdTntMain = IIf(IsDBNull(dt.Rows(0)("Tnt_Main")), "", dt.Rows(0)("Tnt_Main"))
                strStdTntMainTel = IIf(IsDBNull(dt.Rows(0)("Tnt_Main_Tel")), "", dt.Rows(0)("Tnt_Main_Tel"))
                strStdTntSub = IIf(IsDBNull(dt.Rows(0)("Tnt_Sub")), "", dt.Rows(0)("Tnt_Sub"))
                strStdTntSubTel = IIf(IsDBNull(dt.Rows(0)("Tnt_Sub_Tel")), "", dt.Rows(0)("Tnt_Sub_Tel"))

                fncSeriesTableSelect = True
            Else
                strStdTntMain = ""
                strStdTntMainTel = ""
                strStdTntSub = ""
                strStdTntSubTel = ""

                fncSeriesTableSelect = False
            End If

        Catch ex As Exception
            fncSeriesTableSelect = False
        End Try
    End Function

    ''' <summary>
    ''' 形番分解
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strKtbnStrcNm">形番構成名称</param>
    ''' <param name="strHyphen">ハイフン</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncItemNameTableSelect(objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                            ByRef strKtbnStrcNm() As String, _
                                            ByRef strHyphen() As String) As Boolean
        Dim dt As New DataTable
        Dim dalStdDlv As New StdDlvDAL
        Dim intLoopCnt As Integer

        Try
            dt = dalStdDlv.fncItemNameTableSelect(objCon, strSeriesKataban)

            If dt.Rows.Count > 0 Then
                For intLoopCnt = 1 To 23
                    strHyphen(intLoopCnt) = IIf(IsDBNull(dt.Rows(0)("Hyphen" & intLoopCnt.ToString)), "", dt.Rows(0)("Hyphen" & intLoopCnt.ToString))
                Next
                For intLoopCnt = 1 To 24
                    strKtbnStrcNm(intLoopCnt) = IIf(IsDBNull(dt.Rows(0)("Name" & intLoopCnt.ToString)), "", dt.Rows(0)("Name" & intLoopCnt.ToString))
                Next
                fncItemNameTableSelect = True
            Else
                fncItemNameTableSelect = False
            End If

        Catch ex As Exception
            fncItemNameTableSelect = False
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strKataban"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="intItemNo"></param>
    ''' <param name="bolHyphen"></param>
    ''' <param name="intBlankLength"></param>
    ''' <param name="dt_Symbol"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetSymbol(ByVal strKataban As String, ByVal strSeriesKataban As String, _
                                  ByVal intItemNo As Integer, ByVal bolHyphen As Boolean, _
                                  ByRef intBlankLength As Integer, ByVal dt_Symbol As DataTable) As String
        Dim bolMatchFlag As Boolean = False
        Dim intLoopCnt As Integer
        Dim strSymbol As String

        Try
            fncGetSymbol = ""
            intBlankLength = 0

            Dim dr() As DataRow = dt_Symbol.Select("ItemNo='" & intItemNo & "'", "Symbol")
            For inti As Integer = 0 To dr.Length - 1
                If bolHyphen Then
                    If Len(dr(inti)("Symbol").ToString) < Len(strKataban) Then
                        strSymbol = dr(inti)("Symbol").ToString & CdCst.Sign.Hypen
                    Else
                        strSymbol = dr(inti)("Symbol").ToString
                    End If
                Else
                    strSymbol = dr(inti)("Symbol").ToString
                End If

                If Left(strSymbol, 1) = "!" Then
                    intBlankLength = Len(dr(inti)("Symbol").ToString)
                ElseIf Left(strSymbol, 1) = "#" Then
                    For intLoopCnt = 1 To Len(strKataban)
                        If IsNumeric(Mid(strKataban, intLoopCnt, 1)) Or Mid(strKataban, intLoopCnt, 1) = CdCst.Sign.Dot Then
                            fncGetSymbol &= Mid(strKataban, intLoopCnt, 1)
                            bolMatchFlag = True
                        Else
                            Exit For
                        End If
                    Next
                Else
                    If Left(strKataban, Len(strSymbol)) = strSymbol Then
                        If Len(dr(inti)("Symbol").ToString) > Len(fncGetSymbol) Then
                            fncGetSymbol = dr(inti)("Symbol").ToString
                        End If
                        bolMatchFlag = True
                    End If
                End If
            Next
            If Not bolMatchFlag Then
                If intBlankLength > 0 Then
                    fncGetSymbol = Right("!!!!!!!!!!!!!!!!!!!!!!!!", intBlankLength)
                Else
                    fncGetSymbol = ""
                End If
            End If
        Catch ex As Exception
            fncGetSymbol = ""
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="intItemNo"></param>
    ''' <param name="dt_Symbol"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetBlankCounter(ByVal strSeriesKataban As String, _
                                     ByVal intItemNo As Integer, ByVal dt_Symbol As DataTable) As Integer
        Try
            fncGetBlankCounter = -1

            Dim dr() As DataRow = dt_Symbol.Select("ItemNo='" & intItemNo & "'")
            For inti As Integer = 0 To dr.Length - 1
                fncGetBlankCounter = 0
                If Left(dr(inti)("Symbol").ToString, 1) = "!" Then
                    fncGetBlankCounter = Len(dr(inti)("Symbol").ToString)
                    Exit For
                End If
            Next
        Catch ex As Exception
            fncGetBlankCounter = -1
        End Try

    End Function

    ''' <summary>
    ''' 機種チェック
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strOptionSymbol">オプション記号</param>
    ''' <param name="strErrorMessage">エラーメッセージ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncModelChecker(objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                    ByVal strOptionSymbol() As String, ByRef strErrorMessage As String) As Boolean

        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        Dim sbSql As New StringBuilder
        Dim bolJudgeFlag As Boolean
        fncModelChecker = False

        Try
            strErrorMessage = ""
            'SQL文生成
            sbSql.Append(" SELECT  Check1       , Symbol1      , ")
            sbSql.Append("         Check2       , Symbol2      , ")
            sbSql.Append("         Check3       , Symbol3      , ")
            sbSql.Append("         Check4       , Symbol4      , ")
            sbSql.Append("         Check5       , Symbol5      , ")
            sbSql.Append("         Check6       , Symbol6      , ")
            sbSql.Append("         Check7       , Symbol7      , ")
            sbSql.Append("         Check8       , Symbol8      , ")
            sbSql.Append("         Check9       , Symbol9      , ")
            sbSql.Append("         Check10      , Symbol10     , ")
            sbSql.Append("         Check11      , Symbol11     , ")
            sbSql.Append("         Check12      , Symbol12     , ")
            sbSql.Append("         Check13      , Symbol13     , ")
            sbSql.Append("         Check14      , Symbol14     , ")
            sbSql.Append("         Check15      , Symbol15     , ")
            sbSql.Append("         Check16      , Symbol16     , ")
            sbSql.Append("         Check17      , Symbol17     , ")
            sbSql.Append("         Check18      , Symbol18     , ")
            sbSql.Append("         Check19      , Symbol19     , ")
            sbSql.Append("         Check20      , Symbol20     , ")
            sbSql.Append("         Check21      , Symbol21     , ")
            sbSql.Append("         Check22      , Symbol22     , ")
            sbSql.Append("         Check23      , Symbol23     , ")
            sbSql.Append("         Check24      , Symbol24     , ")
            sbSql.Append("         ErrorMessage ")
            sbSql.Append(" FROM    CheckTable ")
            sbSql.Append(" WHERE   Series = @Series ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@Series", SqlDbType.VarChar, 30).Value = strSeriesKataban
            End With

            objRdr = objCmd.ExecuteReader

            While objRdr.Read()
                bolJudgeFlag = True

                If Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check1"))) And _
                   Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check1"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check1"))), _
                                         IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol1"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol1"))), _
                                         IIf(IsNothing(strOptionSymbol(1)), "", strOptionSymbol(1))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check2"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check2"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check2"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol2"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol2"))), _
                                             IIf(IsNothing(strOptionSymbol(2)), "", strOptionSymbol(2))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check3"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check3"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check3"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol3"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol3"))), _
                                             IIf(IsNothing(strOptionSymbol(3)), "", strOptionSymbol(3))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check4"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check4"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check4"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol4"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol4"))), _
                                             IIf(IsNothing(strOptionSymbol(4)), "", strOptionSymbol(4))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check5"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check5"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check5"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol5"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol5"))), _
                                             IIf(IsNothing(strOptionSymbol(5)), "", strOptionSymbol(5))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check6"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check6"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check6"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol6"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol6"))), _
                                             IIf(IsNothing(strOptionSymbol(6)), "", strOptionSymbol(6))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check7"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check7"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check7"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol7"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol7"))), _
                                             IIf(IsNothing(strOptionSymbol(7)), "", strOptionSymbol(7))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check8"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check8"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check8"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol8"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol8"))), _
                                             IIf(IsNothing(strOptionSymbol(8)), "", strOptionSymbol(8))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check9"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check9"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check9"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol9"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol9"))), _
                                             IIf(IsNothing(strOptionSymbol(9)), "", strOptionSymbol(9))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check10"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check10"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check10"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol10"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol10"))), _
                                             IIf(IsNothing(strOptionSymbol(10)), "", strOptionSymbol(10))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check11"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check11"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check11"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol11"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol11"))), _
                                             IIf(IsNothing(strOptionSymbol(11)), "", strOptionSymbol(11))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check12"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check12"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check12"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol12"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol12"))), _
                                             IIf(IsNothing(strOptionSymbol(12)), "", strOptionSymbol(12))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check13"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check13"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check13"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol13"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol13"))), _
                                             IIf(IsNothing(strOptionSymbol(13)), "", strOptionSymbol(13))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check14"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check14"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check14"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol14"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol14"))), _
                                             IIf(IsNothing(strOptionSymbol(14)), "", strOptionSymbol(14))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check15"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check15"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check15"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol15"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol15"))), _
                                             IIf(IsNothing(strOptionSymbol(15)), "", strOptionSymbol(15))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check16"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check16"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check16"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol16"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol16"))), _
                                             IIf(IsNothing(strOptionSymbol(16)), "", strOptionSymbol(16))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check17"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check17"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check17"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol17"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol17"))), _
                                             IIf(IsNothing(strOptionSymbol(17)), "", strOptionSymbol(17))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check18"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check18"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check18"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol18"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol18"))), _
                                             IIf(IsNothing(strOptionSymbol(18)), "", strOptionSymbol(18))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check19"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check19"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check19"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol19"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol19"))), _
                                             IIf(IsNothing(strOptionSymbol(19)), "", strOptionSymbol(19))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check20"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check20"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check20"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol20"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol20"))), _
                                             IIf(IsNothing(strOptionSymbol(20)), "", strOptionSymbol(20))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check21"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check21"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check21"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol21"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol21"))), _
                                             IIf(IsNothing(strOptionSymbol(21)), "", strOptionSymbol(21))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check22"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check22"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check22"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol22"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol22"))), _
                                             IIf(IsNothing(strOptionSymbol(22)), "", strOptionSymbol(22))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check23"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check23"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check23"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol23"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol23"))), _
                                             IIf(IsNothing(strOptionSymbol(23)), "", strOptionSymbol(23))) Then
                    bolJudgeFlag = False
                ElseIf Not IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check24"))) And _
                       Not Me.fncItemChecker(IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Check24"))), "", objRdr.GetValue(objRdr.GetOrdinal("Check24"))), _
                                             IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("Symbol24"))), "", objRdr.GetValue(objRdr.GetOrdinal("Symbol24"))), _
                                             IIf(IsNothing(strOptionSymbol(24)), "", strOptionSymbol(24))) Then
                    bolJudgeFlag = False
                End If

                If bolJudgeFlag Then
                    strErrorMessage = IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("ErrorMessage"))), "", objRdr.GetValue(objRdr.GetOrdinal("ErrorMessage")))
                    fncModelChecker = True
                    Exit Try
                End If
            End While

        Catch ex As Exception
            fncModelChecker = False
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 機種チェック
    ''' </summary>
    ''' <param name="strCheck">チェック記号</param>
    ''' <param name="strSymbol">オプション記号</param>
    ''' <param name="strOptionSymbol">オプション記号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncItemChecker(ByVal strCheck As String, ByVal strSymbol As String, _
                                    ByVal strOptionSymbol As String) As Boolean
        Dim strWkSymbol As String
        fncItemChecker = False
        Try
            If strCheck.Trim = "" Or strSymbol.Trim = "" Or strOptionSymbol.Trim = "" Then Exit Function

            If Left(strSymbol, 1) = "!" Then
                strWkSymbol = "!"
            Else
                strWkSymbol = strSymbol
            End If

            Select Case strCheck
                Case "＝"
                    If strOptionSymbol = strWkSymbol Then
                        fncItemChecker = True
                    End If
                Case "≠"
                    If strOptionSymbol <> strWkSymbol Then
                        fncItemChecker = True
                    End If
                Case "＜"
                    If Val(strOptionSymbol) < Val(strWkSymbol) Then
                        fncItemChecker = True
                    End If
                Case "＞"
                    If Val(strOptionSymbol) > Val(strWkSymbol) Then
                        fncItemChecker = True
                    End If
                Case "≦"
                    If Val(strOptionSymbol) <= Val(strWkSymbol) Then
                        fncItemChecker = True
                    End If
                Case "≧"
                    If Val(strOptionSymbol) >= Val(strWkSymbol) Then
                        fncItemChecker = True
                    End If
            End Select
        Catch ex As Exception
            fncItemChecker = False
        End Try
    End Function

    ''' <summary>
    ''' 納期情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strOptionSymbol">オプション記号</param>
    ''' <param name="intStdDate">納期</param>
    ''' <param name="intFixQty">適用個数</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncStandardDate(objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                     ByVal strOptionSymbol() As String, _
                                     ByRef intStdDate As Integer, ByRef intFixQty As Integer) As Boolean
        Dim intLoopCnt As Integer
        Dim strStdDay(24) As String
        Dim intOffSet As Integer
        Dim intExpDay As Integer
        Dim intExpQty As Integer
        fncStandardDate = True

        Try
            If Me.fncStandardDay(objCon, strSeriesKataban, strOptionSymbol, strStdDay) Then
                intOffSet = 0
                intStdDate = 0

                For intLoopCnt = 1 To 24
                    If Left(strStdDay(intLoopCnt), 1) = "+" And _
                       intOffSet < Val(Mid(strStdDay(intLoopCnt), 2, 10)) Then
                        intOffSet = Val(Mid(strStdDay(intLoopCnt), 2, 10))
                    Else
                        If intStdDate < Val(strStdDay(intLoopCnt)) Then
                            intStdDate = Val(strStdDay(intLoopCnt))
                        End If
                    End If
                Next

                intStdDate = intStdDate + intOffSet
                intFixQty = Me.fncStandardQuantity(objCon, strSeriesKataban)

                If Me.fncExpDate(objCon, strSeriesKataban, strOptionSymbol, intExpDay, intExpQty) Then
                    intStdDate = intExpDay
                    intFixQty = intExpQty
                End If
            Else
                fncStandardDate = False
            End If
        Catch ex As Exception
            fncStandardDate = False
        End Try

    End Function

    ''' <summary>
    ''' 納期取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strOptionSymbol">オプション記号</param>
    ''' <param name="strStdDay">納期</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncStandardDay(objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                    ByVal strOptionSymbol() As String, ByVal strStdDay() As String) As Boolean
        fncStandardDay = True
        Try
            Dim dt_StandardDate As New DS_Tanka.StandardDateDataTable
            Dim dr() As DataRow = Nothing
            Using da As New DS_TankaTableAdapters.StandardDateTableAdapter
                da.Fill(dt_StandardDate, strSeriesKataban)
            End Using
            If dt_StandardDate.Rows.Count <= 0 Then Exit Function

            For intLoopCnt As Integer = 1 To 24
                If Not strOptionSymbol(intLoopCnt) Is Nothing Then
                    If strOptionSymbol(intLoopCnt) = "!" Then
                        dr = dt_StandardDate.Select("ItemNo='" & intLoopCnt & "' AND Symbol='!'")
                    Else
                        dr = dt_StandardDate.Select("ItemNo='" & intLoopCnt & "' AND Symbol='" & strOptionSymbol(intLoopCnt) & "'")
                    End If

                    If Not dr Is Nothing AndAlso dr.Length = 0 Then
                        If IsNumeric(strOptionSymbol(intLoopCnt)) Then
                            dr = dt_StandardDate.Select("ItemNo='" & intLoopCnt & "' AND Symbol='#'")
                        End If
                        If strOptionSymbol(intLoopCnt) = "!" Then
                            dr = dt_StandardDate.Select("ItemNo='" & intLoopCnt & "' AND Symbol LIKE '!%'")
                        End If
                    End If

                    If Not dr Is Nothing AndAlso dr.Length > 0 Then
                        strStdDay(intLoopCnt) = dr(0).Item("Date")
                    Else
                        If strOptionSymbol(intLoopCnt) = "!" Then
                            strStdDay(intLoopCnt) = "0"
                        Else
                            fncStandardDay = False
                        End If
                    End If

                    If strStdDay(intLoopCnt).Trim = "" Then
                        fncStandardDay = False
                    End If
                Else
                    strStdDay(intLoopCnt) = "0"
                End If
            Next
        Catch ex As Exception
            fncStandardDay = False
        End Try
    End Function

    ''' <summary>
    ''' 適用個数取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncStandardQuantity(objCon As SqlConnection, ByVal strSeriesKataban As String) As Integer
        Dim dt As New DataTable
        Dim dalStdDlv As New StdDlvDAL

        fncStandardQuantity = 0
        Try
            dt = dalStdDlv.fncStandardQuantity(objCon, strSeriesKataban)

            If dt.Rows.Count > 0 Then
                fncStandardQuantity = dt.Rows(0)("Quantity")
            End If
        Catch ex As Exception
            fncStandardQuantity = 0
        End Try

    End Function

    ''' <summary>
    ''' 納期情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strOptionSymbol">オプション記号</param>
    ''' <param name="intExpDay">納期</param>
    ''' <param name="intExpQty">適用個数</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncExpDate(objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                ByVal strOptionSymbol() As String, _
                                ByRef intExpDay As Integer, ByRef intExpQty As Integer) As Boolean
        Dim intLoopCnt As Integer
        Dim strSymbol(24) As String
        Dim bolJudgeFlag As Boolean
        Dim dt As New DataTable
        Dim dalStdDlv As New StdDlvDAL

        fncExpDate = False

        Try
            intExpDay = 0
            intExpQty = 0

            dt = dalStdDlv.fncExpDate(objCon, strSeriesKataban)

            For Each dr As DataRow In dt.Rows
                bolJudgeFlag = True

                If strOptionSymbol(10) = "D" And strOptionSymbol(20) = "DC24V" Then strOptionSymbol(10) = "!"

                For intLoopCnt = 1 To 24
                    strSymbol(intLoopCnt) = IIf(IsDBNull(dr("Symbol" & intLoopCnt.ToString)), "", dr("Symbol" & intLoopCnt.ToString))
                Next

                For intLoopCnt = 1 To 24
                    If Left(strSymbol(intLoopCnt), 1) = ">" Or _
                       Left(strSymbol(intLoopCnt), 1) = "<" Or _
                       Left(strSymbol(intLoopCnt), 1) = "=" Then
                        If Not fncEval(strOptionSymbol(intLoopCnt) & strSymbol(intLoopCnt)) Then
                            bolJudgeFlag = False
                            Exit For
                        End If
                    ElseIf strSymbol(intLoopCnt) <> strOptionSymbol(intLoopCnt) And _
                           strSymbol(intLoopCnt) <> "" Then
                        bolJudgeFlag = False
                        Exit For
                    End If
                Next

                If bolJudgeFlag Then
                    fncExpDate = True
                    If dr("Date") > intExpDay Then
                        intExpDay = dr("Date")
                        intExpQty = dr("Quantity")
                    End If
                End If
            Next

        Catch ex As Exception
            fncExpDate = False
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="St"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncEval(ByVal St As String) As Object
        Dim SSt As String = Nothing         'Stの一文字を入れる変数
        Dim Bit As String = Nothing         'Stの中のビット計算を比べるための変数
        Dim Hen(1000) As String             '変数の配列
        Dim Sik(1000) As Long               '式の配列
        Dim HeP As Integer                  '変数への次入れるべきポイント
        Dim SiP As Integer                  '式への次入れるべきポイント
        Dim KaN As Integer                  'その点が括弧で何重にくるまれているか
        Dim Kans(255) As Boolean            'その点が関数の中かどうか
        Dim Dub As Boolean                  'その点が "" でくるまれているか
        Dim SKd As Boolean                  '式に入れたのが最後か数に入れたのが最後か

        Dim AA As Object = Nothing          '予備用の変数
        Dim BB As Object = Nothing          '予備用の変数
        Dim CC As Object = Nothing          '予備用の変数

        Dim Errlog As Integer

        Try

            fncEval = ""

            SKd = True
Line4:
            If St Like "*[!0-9,! ]*" Then
Line3:
                If St Like "-*" Then
                    St = "0" + St
                    GoTo Line3
                ElseIf St Like "(*)" Then
                    If InStr(InStr(1, St, ")") + 1, St, "(") = 0 Then
                        St = Mid(St, 2, Len(St) - 2)
                        GoTo Line3
                    End If
                ElseIf St Like "&H*" Then
                    St = CDbl(St)
                    GoTo Line4
                End If
            Else
                fncEval = CDbl(St)
                Exit Function
            End If
            'Stの振り分け
            Dim Fa, ISS
            For Fa = 1 To Len(St)
                'SStにStのFa番目の文字を入れる
                SSt = Mid(St, Fa, 1)

                If SSt = """" Then
                    'Bit演算子クリアー
                    Bit = "$$$$"
                    '　""でくるまれているかを反転
                    Dub = Not Dub
                    If SKd Then
                        SKd = False
                    End If
                    ' 次の " を探す
                    ISS = InStr(Fa + 1, St, """")
                    '変数に入れる
                    Hen(HeP) = Mid(St, Fa, ISS - Fa + 1)
                    Dub = Not Dub
                    Fa = ISS
                Else
                    If SSt = "(" Then
                        KaN = KaN + 1
                        If Not SKd Then
                            Kans(KaN) = True
                            Bit = "$$$$"
                            AA = InStr(Fa + 1, St, ")")
                            BB = InStr(Fa + 1, St, "(")
                            If AA < BB Or BB = 0 Then
                                Hen(HeP) = Hen(HeP) + Mid(St, Fa, AA - Fa)
                                Fa = AA - 1
                            ElseIf BB < AA Then
                                Hen(HeP) = Hen(HeP) + Mid(St, Fa, BB - Fa)
                                Fa = BB - 1
                            Else
                                Hen(HeP) = Hen(HeP) + "("
                                SKd = False
                            End If
                        End If
                    ElseIf SSt = ")" Then
                        KaN = KaN - 1
                        If Kans(KaN + 1) Then
                            Bit = ")"
                            Hen(HeP) = Hen(HeP) + ")"
                            SKd = False
                        End If
                    ElseIf KaN = 0 Or (Not Kans(KaN)) Then
                        '括弧の外側の時の処理
                        If ("0" <= SSt And "9" >= SSt) Or SSt = "." Then
                            Hen(HeP) = Hen(HeP) + SSt
                            Bit = "0"
                            SKd = False
                        ElseIf ("A" <= SSt And "Z" >= SSt) Or _
                        ("a" <= SSt And "z" >= SSt) Then
                            Hen(HeP) = Hen(HeP) + StrConv(SSt, vbUpperCase)
                            Bit = StrConv(SSt, vbUpperCase) + Bit
                            SKd = False
                            If Bit Like "DNA[),0-9, ]" Then
                                Hen(HeP) = Left(Hen(HeP), Len(Hen(HeP)) - 3)
                                HeP = HeP + 1
                                Sik(SiP) = &HB0 + (127 - KaN) * 256 'And
                                SiP = SiP + 1
                                SKd = True
                            ElseIf Bit Like "RO[),0-9, ]?" Then
                                Hen(HeP) = Left(Hen(HeP), Len(Hen(HeP)) - 2)
                                HeP = HeP + 1
                                Sik(SiP) = &HC0 + (127 - KaN) * 256 'Or
                                SiP = SiP + 1
                                SKd = True
                            ElseIf Bit Like "ROX[),0-9, ]" Then
                                Hen(HeP) = Left(Hen(HeP), Len(Hen(HeP)) - 3)
                                HeP = HeP + 1
                                Sik(SiP) = &HD0 + (127 - KaN) * 256 'Xor
                                SiP = SiP + 1
                                SKd = True
                            ElseIf Bit Like "DOM[),0-9, ]" Then
                                Hen(HeP) = Left(Hen(HeP), Len(Hen(HeP)) - 3)
                                HeP = HeP + 1
                                Sik(SiP) = &H60 + (127 - KaN) * 256 'Mod
                                SiP = SiP + 1
                                SKd = True
                            End If
                        ElseIf (SSt >= "#" And SSt <= "&") Or SSt = "@" Or SSt = "!" Then
                            Hen(HeP) = Hen(HeP) + SSt
                            Bit = "0"
                            SKd = False
                        ElseIf SSt = " " Then
                            Bit = " "
                        Else
                            If SSt = "+" Or SSt = ";" Then
                                Sik(SiP) = &H70 + (127 - KaN) * 256 '+
                                SiP = SiP + 1
                                If Not SKd Then
                                    HeP = HeP + 1
                                End If
                                SKd = True
                            ElseIf SSt = "-" Then
                                Sik(SiP) = &H71 + (127 - KaN) * 256 '-
                                SiP = SiP + 1
                                If Not SKd Then
                                    HeP = HeP + 1
                                Else
                                    Hen(HeP) = Hen(HeP) + "-"
                                End If
                                SKd = True
                            ElseIf SSt = "*" Then
                                Sik(SiP) = &H40 + (127 - KaN) * 256 '*
                                SiP = SiP + 1
                                If Not SKd Then
                                    HeP = HeP + 1
                                End If
                                SKd = True
                            ElseIf SSt = "/" Then
                                Sik(SiP) = &H41 + (127 - KaN) * 256 '/
                                SiP = SiP + 1
                                If Not SKd Then
                                    HeP = HeP + 1
                                End If
                                SKd = True
                            ElseIf SSt = "\" Then
                                Sik(SiP) = &H50 + (127 - KaN) * 256 '\
                                SiP = SiP + 1
                                If Not SKd Then
                                    HeP = HeP + 1
                                End If
                                SKd = True
                            ElseIf SSt = "^" Then
                                Sik(SiP) = &H20 + (127 - KaN) * 256 '^
                                SiP = SiP + 1
                                If Not SKd Then
                                    HeP = HeP + 1
                                End If
                                SKd = True
                            ElseIf SSt = "<" Then
                                If Mid(St, Fa + 1, 1) = ">" Then
                                    Sik(SiP) = &H91 + (127 - KaN) * 256 '<>
                                    Fa = Fa + 1
                                ElseIf Mid(St, Fa + 1, 1) = "=" Then
                                    Sik(SiP) = &H93 + (127 - KaN) * 256 '<=
                                    Fa = Fa + 1
                                Else
                                    Sik(SiP) = &H92 + (127 - KaN) * 256 '<
                                End If
                                SiP = SiP + 1
                                If Not SKd Then
                                    HeP = HeP + 1
                                End If
                                SKd = True
                            ElseIf SSt = ">" Then
                                If Mid(St, Fa + 1, 1) = "<" Then
                                    Sik(SiP) = &H91 + (127 - KaN) * 256 '<>
                                    Fa = Fa + 1
                                ElseIf Mid(St, Fa + 1, 1) = "=" Then
                                    Sik(SiP) = &H95 + (127 - KaN) * 256 '>=
                                    Fa = Fa + 1
                                Else
                                    Sik(SiP) = &H94 + (127 - KaN) * 256 '>
                                End If
                                SiP = SiP + 1
                                If Not SKd Then
                                    HeP = HeP + 1
                                End If
                                SKd = True
                            ElseIf SSt = "=" Then
                                If Mid(St, Fa + 1, 1) = "<" Then
                                    Sik(SiP) = &H93 + (127 - KaN) * 256 '<=
                                    Fa = Fa + 1
                                ElseIf Mid(St, Fa + 1, 1) = ">" Then
                                    Sik(SiP) = &H95 + (127 - KaN) * 256 '>=
                                    Fa = Fa + 1
                                Else
                                    Sik(SiP) = &H90 + (127 - KaN) * 256 '=
                                End If
                                SiP = SiP + 1
                                If Not SKd Then
                                    HeP = HeP + 1
                                End If
                                SKd = True
                            End If
                        End If
                    Else
                        'カッコ内なので無条件
                        Hen(HeP) = Hen(HeP) + SSt
                    End If
                End If
            Next
            If HeP <> SiP Then
                Errlog = 3
            End If
            If HeP = 0 Then
                '文字だけか数字だけ(これ以上分けられない)
                If Hen(0) Like "*[!0-9,!.]*" Then
                    If Hen(0) Like """*""" Then
                        '文字列
                        Hen(0) = Mid(Hen(0), 2, Len(Hen(0)) - 2)
                        AA = Split(Hen(0), """""")
                        fncEval = Join(AA, """")
                    ElseIf Hen(0) Like "(*)" Then
                        '()をはずす
                        fncEval = fncEval(Mid(Hen(0), 2, Len(Hen(0)) - 2))
                    ElseIf Hen(0) Like "-*" Then
                        fncEval = -fncEval(Right(St, Len(St) - 1))
                    ElseIf StrConv(St, vbUpperCase) Like "NOT*" Then
                        fncEval = Not fncEval(Right(St, Len(St) - 3))
                    Else
                        '変数か関数
                        fncEval = fncFfc(UCase(St))
                    End If
                ElseIf St = "" Then
                    fncEval = ""
                Else
                    'ただの数字
                    fncEval = CDbl(Hen(0))
                End If
            Else
                '式になっている
                Dim Jni(100) As Integer
                Dim SSk(100) As Object
                Dim Boo As Boolean
                Jni(0) = 0
                For Fa = 1 To SiP - 1
                    Jni(Fa) = Fa
                Next Fa
                For Fa = 1 To SiP - 1
                    If Sik(Jni(Fa)) \ 16 < Sik(Jni(Fa - 1)) \ 16 Then
                        CC = Jni(Fa - 1)
                        Jni(Fa - 1) = Jni(Fa)
                        Jni(Fa) = CC
                        Fa = Fa - 1
                    End If
                Next
                For Fa = 0 To HeP
                    If Hen(Fa) Like "*[!0-9]*" Then
                        SSk(Fa) = fncEval(Hen(Fa))
                    Else
                        SSk(Fa) = CDbl(Hen(Fa))
                    End If
                Next

                Dim Fb
                For Fa = 0 To SiP - 1
                    AA = Jni(Fa)
                    For Fb = AA To 0 Step -1
                        If VarType(SSk(Fb)) <> 11 Then
                            BB = Fb
                            GoTo Line1
                        End If
                    Next
Line1:
                    For Fb = AA + 1 To SiP
                        If VarType(SSk(Fb)) <> 11 Then
                            CC = Fb
                            GoTo Line2
                        End If
                    Next
Line2:
                    Select Case Sik(AA) Mod 256
                        Case &H20
                            SSk(BB) = SSk(BB) ^ SSk(CC)
                            SSk(CC) = Boo
                        Case &H40
                            SSk(BB) = SSk(BB) * SSk(CC)
                            SSk(CC) = Boo
                        Case &H41
                            SSk(BB) = SSk(BB) / SSk(CC)
                            SSk(CC) = Boo
                        Case &H50
                            SSk(BB) = SSk(BB) \ SSk(CC)
                            SSk(CC) = Boo
                        Case &H60
                            SSk(BB) = SSk(BB) Mod SSk(CC)
                            SSk(CC) = Boo
                        Case &H70
                            If VarType(SSk(BB)) = vbString Or VarType(SSk(CC)) = vbString Then
                                SSk(BB) = fncSStr(SSk(BB)) & fncSStr(SSk(CC))
                            Else
                                SSk(BB) = SSk(BB) + SSk(CC)
                            End If
                            SSk(CC) = Boo
                        Case &H71
                            SSk(BB) = SSk(BB) - SSk(CC)
                            SSk(CC) = Boo
                        Case &H90
                            SSk(BB) = SSk(BB) = SSk(CC)
                            SSk(CC) = Boo
                        Case &H91
                            SSk(BB) = SSk(BB) <> SSk(CC)
                            SSk(CC) = Boo
                        Case &H92
                            SSk(BB) = SSk(BB) < SSk(CC)
                            SSk(CC) = Boo
                        Case &H93
                            SSk(BB) = SSk(BB) <= SSk(CC)
                            SSk(CC) = Boo
                        Case &H94
                            SSk(BB) = SSk(BB) > SSk(CC)
                            SSk(CC) = Boo
                        Case &H95
                            SSk(BB) = SSk(BB) >= SSk(CC)
                            SSk(CC) = Boo
                        Case &HB0
                            SSk(BB) = SSk(BB) And SSk(CC)
                            SSk(CC) = Boo
                        Case &HC0
                            SSk(BB) = SSk(BB) Or SSk(CC)
                            SSk(CC) = Boo
                        Case &HD0
                            SSk(BB) = SSk(BB) Xor SSk(CC)
                            SSk(CC) = Boo
                    End Select
                Next
                fncEval = SSk(BB)
            End If
        Catch ex As Exception
            fncEval = ""
        End Try

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="objVarName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSStr(ByVal objVarName As Object) As String
        Dim intVarType As Integer
        Try
            intVarType = VarType(objVarName)

            If Not (intVarType = 7 Or intVarType = 8) Then
                fncSStr = Str(objVarName)
            Else
                fncSStr = objVarName
            End If
        Catch ex As Exception
            fncSStr = ""
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="St"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncFfc(ByVal St As String)

        Dim AA As String = Nothing
        Dim DD(7) As Object
        Dim Kai As Integer
        Dim Mk As Integer
        Dim Co As Byte = Nothing

        Dim BB As String = Nothing
        Dim Fa As Integer
        Dim CC As String = Nothing
        Dim X As Integer
        Dim Inkey As String = Nothing
        Dim StX As String = Nothing
        Dim MM As String = Nothing
        Dim StY As String = Nothing
        Dim Errlog As Integer

        Dim RD As Double = 3.14159265358979 / 180
        Dim NRD As Double = 180 / 3.14159265358979
        Dim HnValu(200) As Object
        Dim HnList As String = Nothing
        Dim HhValu(50) As Object
        Dim HhList As String = Nothing

        Try

            fncFfc = ""

            If St Like "*(*)" Then
                AA = Left(St, InStr(1, St, "(") - 1)
                BB = Mid(St, Len(AA) + 2, Len(St) - Len(AA) - 2)

                For Fa = 1 To Len(BB)
                    CC = Mid(BB, Fa, 1)
                    If CC = "(" Then
                        Kai = Kai + 1
                    ElseIf CC = ")" Then
                        Kai = Kai - 1
                    ElseIf Kai = 0 And CC = "," Then
                        DD(Co) = Mid(BB, Mk + 1, Fa - Mk - 1)
                        Mk = Fa : Co = Co + 1
                    End If
                Next

                DD(Co) = Mid(BB, Mk + 1, Fa - Mk - 1)
                AA = StrConv(AA, vbUpperCase)

                Select Case AA
                    Case "ABS"
                        fncFfc = Math.Abs(fncEval(DD(0)))
                    Case "AKCNV$"
                        fncFfc = StrConv(fncEval(DD(0)), vbWide)
                    Case "ASC"
                        fncFfc = Asc(fncEval(DD(0)))
                    Case "ATN"
                        fncFfc = Math.Atan(fncEval(DD(0)))
                    Case "ATND"
                        fncFfc = Math.Atan(fncEval(DD(0)) * NRD)
                    Case "ASN"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Atan(X / Math.Sqrt(-AA * AA + 1)) + 2 * Math.Atan(1)
                    Case "ASND"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Atan(X / Math.Sqrt(-AA * AA + 1) * NRD) + 2 * Math.Atan(NRD)
                    Case "ASC"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Atan(X / Math.Sqrt(AA * AA - 1)) + Math.Sign(AA - 1) * 2 * Math.Atan(1)
                    Case "ASCD"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Atan(X / Math.Sqrt(AA * AA - 1) * NRD) + Math.Sign(AA - 1) * 2 * Math.Atan(NRD)
                    Case "ACS"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Atan(X / Math.Sqrt(AA * AA - 1)) + (Math.Sign(CDec(AA)) - 1) * 2 * Math.Atan(1)
                    Case "ACSD"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Atan(X / Math.Sqrt(AA * AA - 1) * NRD) + (Math.Sign(CDec(AA)) - 1) * 2 * Math.Atan(NRD)
                    Case "ACT"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Atan(X) + 2 * Math.Atan(1)
                    Case "ACTD"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Atan(X * NRD) + 2 * Math.Atan(NRD)
                    Case "CDBL"
                        fncFfc = CDbl(fncEval(DD(0)))
                    Case "CEIL"
                        fncFfc = -Int(-fncEval(DD(0)))
                    Case "CHR$"
                        fncFfc = Chr(fncEval(DD(0)))
                    Case "CINT"
                        fncFfc = CInt(fncEval(DD(0)))
                    Case "COT"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Cos(AA) / Math.Sign(CDec(AA))
                    Case "COTD"
                        AA = fncEval(DD(0)) * NRD
                        fncFfc = Math.Cos(AA) / Math.Sign(CDec(AA))
                    Case "COS"
                        fncFfc = Math.Cos(fncEval(DD(0)))
                    Case "COSD"
                        fncFfc = Math.Cos(fncEval(DD(0)) * NRD)
                    Case "CSC"
                        fncFfc = 1 / Math.Sign(fncEval(DD(0)))
                    Case "CSCD"
                        fncFfc = 1 / Math.Sign(fncEval(DD(0)) * NRD)
                    Case "CSNG"
                        fncFfc = CSng(fncEval(DD(0)))
                    Case "CTN"
                        fncFfc = 1 / Math.Tan(fncEval(DD(0)))
                    Case "CTND"
                        fncFfc = 1 / Math.Tan(fncEval(DD(0)) * NRD)
                    Case "ENVIRON$"
                        fncFfc = Environ(fncEval(DD(0)))
                    Case "EOF"
                        fncFfc = EOF(fncEval(DD(0)))
                    Case "EXP"
                        fncFfc = Math.Exp(fncEval(DD(0)))
                    Case "EVAL"
                        fncFfc = fncEval(fncEval(DD(0)))
                    Case "FIX"
                        fncFfc = Fix(fncEval(DD(0)))
                    Case "FP"
                        AA = fncEval(DD(0))
                        fncFfc = AA - CInt(AA)
                    Case "HEX$"
                        fncFfc = Hex(fncEval(DD(0)))
                    Case "HOUR"
                        fncFfc = Hour(fncEval(DD(0)))
                    Case "HSIN"
                        AA = fncEval(DD(0))
                        fncFfc = (Math.Exp(AA) - Math.Exp(-AA)) / 2
                    Case "HCOS"
                        AA = fncEval(DD(0))
                        fncFfc = (Math.Exp(AA) + Math.Exp(-AA)) / 2
                    Case "HTAN"
                        AA = fncEval(DD(0))
                        fncFfc = (Math.Exp(AA) - Math.Exp(-AA)) / (Math.Exp(AA) + Math.Exp(-AA))
                    Case "HSEC"
                        AA = fncEval(DD(0))
                        fncFfc = 2 / (Math.Exp(AA) - Math.Exp(-AA))
                    Case "HCSE"
                        AA = fncEval(DD(0))
                        fncFfc = 2 / (Math.Exp(AA) + Math.Exp(-AA))
                    Case "HCTN"
                        AA = fncEval(DD(0))
                        fncFfc = (Math.Exp(AA) + Math.Exp(-AA)) / (Math.Exp(AA) - Math.Exp(-AA))
                    Case "HASN"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Log(AA + Math.Sqrt(-AA * AA + 1))
                    Case "HACS"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Log(AA + Math.Sqrt(AA * AA - 1))
                    Case "HATN"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Log((1 + AA) / (1 - AA)) / 2
                    Case "HASC"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Log((Math.Sqrt(-AA * AA + 1) + 1) / AA)
                    Case "HACSE"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Log((Math.Sqrt(AA) * Math.Sqrt(AA * AA + 1) + 1) / AA)
                    Case "HATN"
                        AA = fncEval(DD(0))
                        fncFfc = Math.Log((X + 1) / (X - 1)) / 2
                    Case "IIF"
                        fncFfc = IIf(fncEval(DD(0)), fncEval(DD(1)), fncEval(DD(2)))
                    Case "INPUT$"
Line1:
                        If VarType(DD(1)) = vbEmpty Then
                            If Len(Inkey) >= DD(0) Then
                                fncFfc = Left(Inkey, fncEval(DD(0)))
                                Inkey = ""
                            Else
                                GoTo Line1
                            End If
                        Else
                            Errlog = 5
                        End If
                    Case "INSTR"
                        If VarType(DD(2)) = vbEmpty Then
                            fncFfc = InStr(1, fncEval(DD(1)), fncEval(DD(2)))
                        Else
                            fncFfc = InStr(fncEval(DD(0)), fncEval(DD(1)), fncEval(DD(2)))
                        End If
                    Case "INT"
                        fncFfc = Int(fncEval(DD(0)))
                    Case "INP"
                        fncFfc = Int(fncEval(DD(0)))
                    Case "IP"
                        fncFfc = CInt(fncEval(DD(0)))
                    Case "JIS$"
                        fncFfc = Hex(Asc(fncEval(DD(0))))
                    Case "KACNV$"
                        fncFfc = StrConv(fncEval(DD(0)), vbNarrow)
                    Case "KINSTR"
                        If VarType(DD(2)) = vbEmpty Then
                            fncFfc = InStr(1, fncEval(DD(1)), fncEval(DD(2)))
                        Else
                            fncFfc = InStr(fncEval(DD(0)), fncEval(DD(1)), fncEval(DD(2)))
                        End If
                    Case "KLEN"
                        fncFfc = Len(fncEval(DD(0)))
                    Case "KMID$"
                        If VarType(DD(2)) = vbEmpty Then
                            fncFfc = Mid(fncEval(DD(0)), fncEval(DD(1)))
                        Else
                            fncFfc = Mid(fncEval(DD(0)), fncEval(DD(1)), fncEval(DD(2)))
                        End If
                    Case "KNJ$"
                        fncFfc = Chr("&H" & fncEval(DD(0)))
                    Case "LCASE$"
                        fncFfc = LCase(fncEval(DD(0)))
                    Case "LEFT$"
                        fncFfc = Left(fncEval(DD(0)), fncEval(DD(1)))
                    Case "LEN"
                        fncFfc = Len(CStr(fncEval(DD(0))))
                    Case "LOC"
                        fncFfc = Loc(fncEval(DD(0)))
                    Case "LOF"
                        fncFfc = LOF(fncEval(DD(0)))
                    Case "LOG"
                        fncFfc = Math.Log(fncEval(DD(0)))
                    Case "LOG" To "LOG99"
                        fncFfc = Math.Log(fncEval(DD(0))) / Math.Log(CDbl(Mid(AA, 4)))
                    Case "LTRIM$"
                        fncFfc = LTrim(fncEval(DD(0)))
                    Case "MAX"
                        AA = fncEval(DD(0))
                        BB = fncEval(DD(1))
                        fncFfc = IIf(AA > BB, AA, BB)
                    Case "MID$"
                        If VarType(DD(2)) = vbEmpty Then
                            fncFfc = Mid(fncEval(DD(0)), fncEval(DD(1)))
                        Else
                            fncFfc = Mid(fncEval(DD(0)), fncEval(DD(1)), fncEval(DD(2)))
                        End If
                    Case "MIN"
                        AA = fncEval(DD(0))
                        BB = fncEval(DD(1))
                        fncFfc = IIf(AA < BB, AA, BB)
                    Case "MINUTE"
                        fncFfc = Minute(fncEval(DD(0)))
                    Case "OCT$"
                        fncFfc = Oct(fncEval(DD(0)))
                    Case "POS"
                        fncFfc = StX
                    Case "RIGHT$"
                        fncFfc = Right(fncEval(DD(0)), fncEval(DD(1)))
                    Case "RND"
                        fncFfc = Rnd(fncEval(DD(0)))
                    Case "ROUND"
                        fncFfc = Math.Round(fncEval(DD(0)), fncEval(DD(1)))
                    Case "RTRIM$"
                        fncFfc = RTrim(fncEval(DD(0)))
                    Case "REMINDER"
                        AA = fncEval(DD(0))
                        BB = fncEval(DD(1))
                        fncFfc = AA - BB * CInt(AA / BB)
                    Case "SEC"
                        fncFfc = 1 / Math.Cos(fncEval(DD(0)))
                    Case "SECD"
                        fncFfc = 1 / Math.Cos(fncEval(DD(0)) * NRD)
                    Case "SGN"
                        fncFfc = Math.Sign(fncEval(DD(0)))
                    Case "SIN"
                        fncFfc = Math.Sin(fncEval(DD(0)))
                    Case "SIND"
                        fncFfc = Math.Sin(fncEval(DD(0)) * RD)
                    Case "SPACE$"
                        fncFfc = Space(fncEval(DD(0)))
                    Case "SPC"
                        fncFfc = Space(fncEval(DD(0)))
                    Case "SQR"
                        fncFfc = Math.Sqrt(fncEval(DD(0)))
                    Case "STR$"
                        fncFfc = Str(fncEval(DD(0)))
                    Case "SSTR$"
                        fncFfc = fncSStr(fncEval(DD(0)))
                    Case "STRING$"
                        fncFfc = Space(0).PadRight(fncEval(DD(0)), fncEval(DD(1)))
                    Case "TAB"
                        fncFfc = Space(0).PadRight(fncEval(DD(0)), vbTab)
                    Case "TAN"
                        fncFfc = Math.Tan(fncEval(DD(0)))
                    Case "TAND"
                        fncFfc = Math.Tan(fncEval(DD(0)) * RD)
                    Case "TRIM"
                        fncFfc = Trim(fncEval(DD(0)))
                    Case "TRUNCATE"
                        AA = fncEval(DD(1))
                        fncFfc = CInt(fncEval(DD(0)) * 10 ^ AA) / 10 ^ AA
                    Case "UCASE"
                        fncFfc = UCase(fncEval(DD(0)))
                    Case "VAL"
                        fncFfc = Val(fncEval(DD(0)))
                    Case Else
                        '変数である
                        MM = InStr(1, HhList, Left(AA & "        ", 7) & " ")
                        If MM Then
                            Select Case Co
                                Case 0
                                    fncFfc = HhValu(MM \ 8)(fncEval(DD(0)))
                                Case 1
                                    fncFfc = HhValu(MM \ 8)(fncEval(DD(0)), fncEval(DD(1)))
                                Case 2
                                    fncFfc = HhValu(MM \ 8)(fncEval(DD(0)), fncEval(DD(1)), fncEval(DD(2)))
                                Case 3
                                    fncFfc = HhValu(MM \ 8)(fncEval(DD(0)), fncEval(DD(1)), fncEval(DD(2)), fncEval(DD(3)))
                                Case 4
                                    fncFfc = HhValu(MM \ 8)(fncEval(DD(0)), fncEval(DD(1)), fncEval(DD(2)), fncEval(DD(3)), fncEval(DD(4)))
                            End Select
                        Else
                            If St Like "*$" Then
                                fncFfc = ""
                            Else
                                fncFfc = 0
                            End If
                        End If
                End Select
            Else
                '変数か関数の引数のないもの
                Select Case St
                    Case "CSRLIN"
                        fncFfc = StY
                    Case "ERR"
                        fncFfc = Err()
                    Case "COMMAND$"
                        fncFfc = Command()
                    Case "DATE$"
                        fncFfc = Format(Now(), "yy/MM/dd")
                    Case "INKEY$"
                        fncFfc = Inkey
                        Inkey = ""
                    Case "TIME$"
                        fncFfc = Format(Now(), "hh:mm:ss")
                    Case "RND"
                        fncFfc = Rnd()
                    Case Else
                        '変数である
                        MM = InStr(1, HnList, Left(St & "        ", 7) & " ")
                        If MM Then
                            fncFfc = HnValu(MM \ 8)
                        Else
                            If St Like "*$" Then
                                fncFfc = ""
                            Else
                                fncFfc = 0
                            End If
                        End If
                End Select
            End If
        Catch ex As Exception
            fncFfc = ""
        End Try
    End Function
End Class
