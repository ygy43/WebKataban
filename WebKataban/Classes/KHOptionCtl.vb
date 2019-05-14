Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class KHOptionCtl

#Region " Definition "

    Private strStructureDiv As String                                   '構成区分(形番構成)
    Private strMsgCd As String                                          'メッセージコード

    '形番構成要素
    Private Structure KtbnStrcEle
        Public strOptionSymbol As String                                'オプション記号
        Public strOptionNm As String                                    'オプション名称
        Public bolOptionFlag As Boolean                                 '可否フラグ
    End Structure
    Private strcKtbnStrcEle() As KtbnStrcEle

    '要素パターン
    Private Structure ElePattern
        Public strOptionSymbol As String                                'オプション記号
        Public strConditionCd As String                                 '条件コード
        Public intConditionSeqNo As Integer                             '条件順序
        Public intConditionSeqNoBr As Integer                           '条件順序枝番
        Public strCondOptionSymbol As String                            '条件オプション記号
        Public bolCondFlag As Boolean                                   '可否フラグ
    End Structure
    Private strcElePattern() As ElePattern

#End Region

    ''' <summary>
    ''' 形番分解
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <param name="strSelectLang"></param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <param name="strOptions">オプション</param>
    ''' <returns></returns>
    ''' <remarks>複数選択項目の形番をオプション毎に分解する</remarks>
    Public Function fncOptionResolution(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                        strUserId As String, strSessionId As String, strSelectLang As String, _
                                        ByVal intKtbnStrcSeqNo As Integer, _
                                        ByVal strOptions As String) As String()
        Dim strOption As String
        Dim strAryWkOption() As String
        Dim strAryRetOption() As String = Nothing
        Dim strAryOption() As String = Nothing
        Dim strAryOptionIndex() As Integer = Nothing
        Dim strOptionSymbol() As String = Nothing
        Dim strOptionIndex() As Integer = Nothing
        Dim strListOption(,) As String = Nothing
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Dim intLoopCnt3 As Integer
        Dim intMaxIndex As Integer
        fncOptionResolution = Nothing

        Try
            '配列定義
            ReDim strAryRetOption(0)
            ReDim strAryOption(0)
            ReDim strAryOptionIndex(0)
            ReDim strOptionSymbol(0)
            ReDim strOptionIndex(0)

            'カンマが存在する場合は消去する
            strOption = strOptions.Replace(CdCst.Sign.Delimiter.Comma, "")

            ''形番構成要素取得
            'Call Me.subKatabanStrcEleSelect(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
            '                                objKtbnStrc.strcSelection.strKeyKataban, _
            '                                intKtbnStrcSeqNo, CdCst.LanguageCd.DefaultLang)

            'オプションリスト取得
            Call Me.subOptionList(objCon, objKtbnStrc, "1", strUserId, strSessionId, strSelectLang, intKtbnStrcSeqNo, strListOption)

            'オプションを並べ替える(長さ順)
            For intLoopCnt1 = 1 To UBound(strListOption)
                ReDim Preserve strOptionSymbol(intLoopCnt1)
                ReDim Preserve strOptionIndex(intLoopCnt1)
                For intLoopCnt2 = 1 To intLoopCnt1
                    If intLoopCnt1 = intLoopCnt2 Then
                        strOptionSymbol(intLoopCnt1) = strListOption(intLoopCnt1, 1)
                        strOptionIndex(intLoopCnt1) = intLoopCnt1
                    Else
                        If strListOption(intLoopCnt1, 1).Trim.Length > strOptionSymbol(intLoopCnt2).Trim.Length Then
                            For intLoopCnt3 = intLoopCnt1 To intLoopCnt2 Step -1
                                If intLoopCnt3 = intLoopCnt2 Then
                                    strOptionSymbol(intLoopCnt3) = strListOption(intLoopCnt1, 1)
                                    strOptionIndex(intLoopCnt3) = intLoopCnt1
                                Else
                                    strOptionSymbol(intLoopCnt3) = strOptionSymbol(intLoopCnt3 - 1)
                                    strOptionIndex(intLoopCnt3) = strOptionIndex(intLoopCnt3 - 1)
                                End If
                            Next
                            Exit For
                        End If
                    End If
                Next
            Next

            'オプションが存在するかチェックする
            For intLoopCnt1 = 1 To strOptionSymbol.Length - 1
                If strOptionSymbol(intLoopCnt1).Trim <> "" Then
                    If strOption.IndexOf(strOptionSymbol(intLoopCnt1)) >= 0 Then
                        ReDim Preserve strAryOption(UBound(strAryOption) + 1)
                        ReDim Preserve strAryOptionIndex(UBound(strAryOptionIndex) + 1)
                        strAryOption(UBound(strAryOption)) = strOptionSymbol(intLoopCnt1)
                        strAryOptionIndex(UBound(strAryOption)) = strOptionIndex(intLoopCnt1)
                        strOption = strOption.Replace(strOptionSymbol(intLoopCnt1), CdCst.Sign.Delimiter.Comma)
                    End If
                End If
            Next

            For intLoopCnt1 = 1 To strAryOption.Length - 1
                intMaxIndex = 1
                For intLoopCnt2 = 1 To strAryOption.Length - 1
                    If strAryOptionIndex(intMaxIndex) > strAryOptionIndex(intLoopCnt2) Then
                        intMaxIndex = intLoopCnt2
                    End If
                Next
                ReDim Preserve strAryRetOption(UBound(strAryRetOption) + 1)
                strAryRetOption(UBound(strAryRetOption)) = strAryOption(intMaxIndex)
                strAryOptionIndex(intMaxIndex) = 9999
            Next

            '不要なオプションが選択されなかったかチェックする
            strAryWkOption = strOption.Split(CdCst.Sign.Delimiter.Comma)
            For intLoopCnt1 = 0 To strAryWkOption.Length - 1
                If strAryWkOption(intLoopCnt1).Trim <> "" Then
                    ReDim Preserve strAryRetOption(UBound(strAryRetOption) + 1)
                    strAryRetOption(UBound(strAryRetOption)) = strAryWkOption(intLoopCnt1).Trim
                End If
            Next
            '戻り値設定
            fncOptionResolution = strAryRetOption
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 形番構成要素取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strKeyKataban">キー形番</param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <param name="strLanguageCd">言語コード</param>
    ''' <remarks></remarks>
    Private Sub subKatabanStrcEleSelect(ByVal objCon As SqlConnection, ByVal strSeriesKataban As String, ByVal strKeyKataban As String, _
                                        ByVal intKtbnStrcSeqNo As Integer, ByVal strLanguageCd As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing

        Try
            '配列初期化
            ReDim Me.strcKtbnStrcEle(0)
            'SQL Query生成
            sbSql.Append(" SELECT  b.element_div, ")
            sbSql.Append("         b.structure_div, ")
            sbSql.Append("         c.option_symbol, ")
            sbSql.Append("         d.option_nm as default_option_nm, ")
            sbSql.Append("         e.option_nm ")
            sbSql.Append(" FROM    kh_series_kataban a ")
            sbSql.Append(" INNER JOIN  kh_kataban_strc b ")
            sbSql.Append(" ON      a.series_kataban         = b.series_kataban ")
            sbSql.Append(" AND     a.key_kataban            = b.key_kataban ")
            sbSql.Append(" AND     b.series_kataban         = @SeriesKataban ")
            sbSql.Append(" AND     b.key_kataban            = @KeyKataban ")
            sbSql.Append(" AND     b.ktbn_strc_seq_no       = @KtbnStrcSeqNo ")
            sbSql.Append(" AND     b.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     b.out_effective_date     > @StandardDate ")
            sbSql.Append(" INNER JOIN  kh_kataban_strc_ele c ")
            sbSql.Append(" ON      b.series_kataban         = c.series_kataban ")
            sbSql.Append(" AND     b.key_kataban            = c.key_kataban ")
            sbSql.Append(" AND     b.ktbn_strc_seq_no       = c.ktbn_strc_seq_no ")
            sbSql.Append(" AND     c.series_kataban         = @SeriesKataban ")
            sbSql.Append(" AND     c.key_kataban            = @KeyKataban ")
            sbSql.Append(" AND     c.ktbn_strc_seq_no       = @KtbnStrcSeqNo ")
            sbSql.Append(" AND     c.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     c.out_effective_date     > @StandardDate ")
            sbSql.Append(" INNER JOIN  kh_option_nm_mst d ")
            sbSql.Append(" ON      c.series_kataban         = d.series_kataban ")
            sbSql.Append(" AND     c.key_kataban            = d.key_kataban ")
            sbSql.Append(" AND     c.ktbn_strc_seq_no       = d.ktbn_strc_seq_no ")
            sbSql.Append(" AND     c.option_symbol          = d.option_symbol ")
            sbSql.Append(" AND     d.series_kataban         = @SeriesKataban ")
            sbSql.Append(" AND     d.key_kataban            = @KeyKataban ")
            sbSql.Append(" AND     d.ktbn_strc_seq_no       = @KtbnStrcSeqNo ")
            sbSql.Append(" AND     d.language_cd            = @DefaultLangCd ")
            sbSql.Append(" AND     d.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     d.out_effective_date     > @StandardDate ")
            sbSql.Append(" LEFT  JOIN  kh_option_nm_mst e ")
            sbSql.Append(" ON      c.series_kataban         = e.series_kataban ")
            sbSql.Append(" AND     c.key_kataban            = e.key_kataban ")
            sbSql.Append(" AND     c.ktbn_strc_seq_no       = e.ktbn_strc_seq_no ")
            sbSql.Append(" AND     c.option_symbol          = e.option_symbol ")
            sbSql.Append(" AND     e.series_kataban         = @SeriesKataban ")
            sbSql.Append(" AND     e.key_kataban            = @KeyKataban ")
            sbSql.Append(" AND     e.ktbn_strc_seq_no       = @KtbnStrcSeqNo ")
            sbSql.Append(" AND     e.language_cd            = @LanguageCd ")
            sbSql.Append(" AND     e.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     e.out_effective_date     > @StandardDate ")
            sbSql.Append(" WHERE   a.series_kataban         = @SeriesKataban ")
            sbSql.Append(" AND     a.key_kataban            = @KeyKataban ")
            sbSql.Append(" AND     a.in_effective_date     <= @StandardDate ")
            sbSql.Append(" AND     a.out_effective_date     > @StandardDate ")
            sbSql.Append(" ORDER BY  c.disp_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
                .Parameters.Add("@KtbnStrcSeqNo", SqlDbType.Int).Value = intKtbnStrcSeqNo
                .Parameters.Add("@DefaultLangCd", SqlDbType.NVarChar, 150).Value = CdCst.LanguageCd.DefaultLang
                .Parameters.Add("@LanguageCd", SqlDbType.NVarChar, 150).Value = strLanguageCd
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            End With
            objRdr = objCmd.ExecuteReader
            While objRdr.Read()
                '配列再定義
                ReDim Preserve Me.strcKtbnStrcEle(UBound(Me.strcKtbnStrcEle) + 1)
                '構成区分
                Me.strStructureDiv = objRdr.GetValue(objRdr.GetOrdinal("structure_div"))
                With Me.strcKtbnStrcEle(UBound(Me.strcKtbnStrcEle))
                    .strOptionSymbol = objRdr.GetValue(objRdr.GetOrdinal("option_symbol"))
                    .strOptionNm = IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("option_nm"))), objRdr.GetValue(objRdr.GetOrdinal("default_option_nm")), objRdr.GetValue(objRdr.GetOrdinal("option_nm")))
                    .bolOptionFlag = True
                End With
            End While
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
            objCmd = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' オプションリスト取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strDivision">処理区分</param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strLanguageCd">言語コード</param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <param name="strArrayOption">オプション記号＆オプション名称</param>
    ''' <param name="dt_KataStrcEleSel"></param>
    ''' <param name="dt_ElePatternSel"></param>
    ''' <remarks></remarks>
    Public Sub subOptionList(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                             ByVal strDivision As String, ByVal strUserId As String, _
                             ByVal strSessionId As String, ByVal strLanguageCd As String, _
                             ByVal intKtbnStrcSeqNo As Integer, ByRef strArrayOption(,) As String, _
                             Optional ByVal dt_KataStrcEleSel As DataTable = Nothing, _
                             Optional ByVal dt_ElePatternSel As DataTable = Nothing)
        Dim strOpSymbol() As String
        Dim strOpNm() As String
        Dim intLoopCnt As Integer
        Dim intArrayIdx As Integer = 0

        Try
            If dt_KataStrcEleSel Is Nothing Then
                '形番構成要素取得
                Call Me.subKatabanStrcEleSelect(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                                                objKtbnStrc.strcSelection.strKeyKataban, intKtbnStrcSeqNo, strLanguageCd)
            Else     'データテーブルから取得する
                '配列初期化
                ReDim Me.strcKtbnStrcEle(0)
                Dim dr() As DataRow = dt_KataStrcEleSel.Select("ktbn_strc_seq_no='" & intKtbnStrcSeqNo & "'")

                For inti As Integer = 0 To dr.Length - 1
                    '配列再定義
                    ReDim Preserve Me.strcKtbnStrcEle(UBound(Me.strcKtbnStrcEle) + 1)
                    '構成区分
                    Me.strStructureDiv = dr(inti)("structure_div").ToString
                    With Me.strcKtbnStrcEle(UBound(Me.strcKtbnStrcEle))
                        .strOptionSymbol = dr(inti)("option_symbol").ToString
                        .strOptionNm = IIf(IsDBNull(dr(inti)("option_nm")), dr(inti)("default_option_nm"), dr(inti)("option_nm"))
                        .bolOptionFlag = True
                    End With
                Next
            End If

            If dt_ElePatternSel Is Nothing Then
                'リスト生成
                Call Me.subOptionListCreate(objCon, objKtbnStrc, strDivision, strUserId, strSessionId, _
                                            objKtbnStrc.strcSelection.strSeriesKataban, _
                                            objKtbnStrc.strcSelection.strKeyKataban, _
                                            intKtbnStrcSeqNo, dt_ElePatternSel)
            Else
                '配列初期化
                ReDim Me.strcElePattern(0)
                Dim dr() As DataRow = dt_ElePatternSel.Select("ktbn_strc_seq_no='" & intKtbnStrcSeqNo & "'", _
                                      "serach_seq_no, option_symbol, condition_seq_no, condition_seq_no_br")
                For inti As Integer = 0 To dr.Length - 1
                    ReDim Preserve Me.strcElePattern(UBound(Me.strcElePattern) + 1)
                    With Me.strcElePattern(UBound(Me.strcElePattern))
                        .strOptionSymbol = dr(inti)("option_symbol").ToString
                        .strConditionCd = dr(inti)("condition_cd").ToString
                        .intConditionSeqNo = dr(inti)("condition_seq_no")
                        .intConditionSeqNoBr = dr(inti)("condition_seq_no_br")
                        .strCondOptionSymbol = dr(inti)("cond_option_symbol").ToString
                        .bolCondFlag = True
                    End With
                Next
                'オプション判定
                Call Me.subOptionJudgment(objKtbnStrc, strDivision, strUserId, strSessionId, intKtbnStrcSeqNo)
            End If

            '配列定義
            ReDim strOpSymbol(UBound(Me.strcKtbnStrcEle))
            ReDim strOpNm(UBound(Me.strcKtbnStrcEle))
            'Trueのオプションのみ抽出
            For intLoopCnt = 1 To UBound(Me.strcKtbnStrcEle)
                If Me.strcKtbnStrcEle(intLoopCnt).bolOptionFlag = True Then
                    intArrayIdx = intArrayIdx + 1
                    strOpSymbol(intArrayIdx) = Me.strcKtbnStrcEle(intLoopCnt).strOptionSymbol
                    strOpNm(intArrayIdx) = Me.strcKtbnStrcEle(intLoopCnt).strOptionNm
                End If
            Next
            '配列定義
            ReDim strArrayOption(intArrayIdx, 2)
            '戻り値設定
            For intLoopCnt = 1 To intArrayIdx
                strArrayOption(intLoopCnt, 1) = strOpSymbol(intLoopCnt)
                strArrayOption(intLoopCnt, 2) = strOpNm(intLoopCnt)
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objKtbnStrc = Nothing
        End Try

    End Sub

    ''' <summary>
    ''' オプションリスト生成
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strDivision">処理区分</param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strKeyKataban">キー形番</param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <param name="dt_ElePatternSel"></param>
    ''' <remarks>形番構成要素テーブルからデータを取得する</remarks>
    Private Sub subOptionListCreate(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                    ByVal strDivision As String, ByVal strUserId As String, _
                                    ByVal strSessionId As String, ByVal strSeriesKataban As String, _
                                    ByVal strKeyKataban As String, ByVal intKtbnStrcSeqNo As Integer, _
                                    Optional ByVal dt_ElePatternSel As DataTable = Nothing)
        Try
            '要素パターン取得
            If dt_ElePatternSel Is Nothing Then
                Call Me.subElePatternSelect(objCon, strSeriesKataban, strKeyKataban, intKtbnStrcSeqNo)
            End If
            'オプション判定
            Call Me.subOptionJudgment(objKtbnStrc, strDivision, strUserId, strSessionId, intKtbnStrcSeqNo)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 要素パターン取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strKeyKataban">キー形番</param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <remarks></remarks>
    Private Sub subElePatternSelect(ByVal objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                    ByVal strKeyKataban As String, ByVal intKtbnStrcSeqNo As Integer)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing

        Try
            '配列初期化
            ReDim Me.strcElePattern(0)
            'SQL Query生成
            sbSql.Append(" SELECT  '1' as serach_seq_no, ")
            sbSql.Append("         option_symbol, ")
            sbSql.Append("         condition_cd, ")
            sbSql.Append("         condition_seq_no, ")
            sbSql.Append("         condition_seq_no_br, ")
            sbSql.Append("         cond_option_symbol ")
            sbSql.Append(" FROM    kh_ele_pattern ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     ktbn_strc_seq_no    = @KtbnStrcSeqNo ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     option_symbol       = '" & CdCst.ElePattern.Plural & "' ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT  '2' as serach_seq_no, ")
            sbSql.Append("         option_symbol, ")
            sbSql.Append("         condition_cd, ")
            sbSql.Append("         condition_seq_no, ")
            sbSql.Append("         condition_seq_no_br, ")
            sbSql.Append("         cond_option_symbol ")
            sbSql.Append(" FROM    kh_ele_pattern ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     ktbn_strc_seq_no    = @KtbnStrcSeqNo ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     option_symbol       = '" & CdCst.ElePattern.All & "' ")
            sbSql.Append(" UNION ")
            sbSql.Append(" SELECT  '3' as serach_seq_no, ")
            sbSql.Append("         option_symbol, ")
            sbSql.Append("         condition_cd, ")
            sbSql.Append("         condition_seq_no, ")
            sbSql.Append("         condition_seq_no_br, ")
            sbSql.Append("         cond_option_symbol ")
            sbSql.Append(" FROM    kh_ele_pattern ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     ktbn_strc_seq_no    = @KtbnStrcSeqNo ")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     option_symbol  Not In ('" & CdCst.ElePattern.All & "','" & CdCst.ElePattern.Plural & "') ")
            sbSql.Append(" ORDER BY  serach_seq_no, option_symbol, condition_seq_no, condition_seq_no_br ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
                .Parameters.Add("@KtbnStrcSeqNo", SqlDbType.Int).Value = intKtbnStrcSeqNo
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            End With

            'DBオープン
            objRdr = objCmd.ExecuteReader
            While objRdr.Read()
                ReDim Preserve Me.strcElePattern(UBound(Me.strcElePattern) + 1)
                With Me.strcElePattern(UBound(Me.strcElePattern))
                    .strOptionSymbol = objRdr.GetValue(objRdr.GetOrdinal("option_symbol"))
                    .strConditionCd = objRdr.GetValue(objRdr.GetOrdinal("condition_cd"))
                    .intConditionSeqNo = objRdr.GetValue(objRdr.GetOrdinal("condition_seq_no"))
                    .intConditionSeqNoBr = objRdr.GetValue(objRdr.GetOrdinal("condition_seq_no_br"))
                    .strCondOptionSymbol = objRdr.GetValue(objRdr.GetOrdinal("cond_option_symbol"))
                    .bolCondFlag = True
                End With
            End While
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
            objCmd = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' オプション判定
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strDivision">処理区分</param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <remarks></remarks>
    Public Sub subOptionJudgment(objKtbnStrc As KHKtbnStrc, _
                                  ByVal strDivision As String, ByVal strUserId As String, _
                                  ByVal strSessionId As String, ByVal intKtbnStrcSeqNo As Integer)
        Dim intStructureDiv As Integer
        Dim bolSelectCond As Boolean
        Dim bolSkipCond As Boolean
        Dim bolPluralCond As Boolean
        Dim strKeyOptionSymbol As String
        Dim strKeyConditionCd As String
        Dim intKeyConditionSeqNo As Integer
        Dim intKeyConditionSeqNoBr As Integer
        Dim intElePatternStaPos As Integer
        Dim intElePatternEndPos As Integer
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Dim intLoopCnt3 As Integer
        Dim intLoopCnt4 As Integer
        Dim intNowPos As Integer
        Dim strCondSign As String = Nothing
        Dim intStaPos As Integer
        Dim intEndPos As Integer
        Dim intKtbnStrcEleCnt As Integer
        Dim intWStaPos As Integer
        Dim intWEndPos As Integer
        Dim bolPluralChkFlg As Boolean = False
        Dim strAryOption() As String

        Try
            '判定区分設定(形番構成区分よりチェックパターンを設定する)
            intStructureDiv = CInt(Me.strStructureDiv)
            '複数選択条件設定
            If intStructureDiv >= CInt(CdCst.KtbnStructureDiv.PluralCond) Then
                bolPluralCond = True
                intStructureDiv = intStructureDiv - CInt(CdCst.KtbnStructureDiv.PluralCond)
            End If
            'Skip条件設定
            If intStructureDiv >= CInt(CdCst.KtbnStructureDiv.SkipCond) Then
                bolSkipCond = True
                intStructureDiv = intStructureDiv - CInt(CdCst.KtbnStructureDiv.SkipCond)
            End If
            '選択条件設定
            If intStructureDiv >= CInt(CdCst.KtbnStructureDiv.SelectCond) Then
                bolSelectCond = True
            End If

            'Skip条件判定
            If bolSkipCond = True Then
                'オプション数分チェックする
                For intKtbnStrcEleCnt = 1 To Me.strcKtbnStrcEle.Length - 1
                    '該当する要素を検索
                    For intLoopCnt1 = 1 To Me.strcElePattern.Length - 1
                        'Skip要素(*)がある場合
                        If Me.strcElePattern(intLoopCnt1).strOptionSymbol = CdCst.ElePattern.All Then
                            'Skip要素全体の開始・終了位置を設定する
                            '開始位置設定
                            intElePatternStaPos = intLoopCnt1
                            For intLoopCnt2 = intLoopCnt1 To Me.strcElePattern.Length - 1
                                If Me.strcElePattern(intLoopCnt2).strOptionSymbol = CdCst.ElePattern.All Then
                                    '終了位置設定
                                    intElePatternEndPos = intLoopCnt2
                                End If
                            Next

                            '現在位置設定
                            intNowPos = intElePatternStaPos
                            '要素チェック
                            Do Until intNowPos > intElePatternEndPos
                                'ブレイクキー設定
                                strKeyOptionSymbol = Me.strcElePattern(intNowPos).strOptionSymbol
                                strKeyConditionCd = Me.strcElePattern(intNowPos).strConditionCd
                                intKeyConditionSeqNo = Me.strcElePattern(intNowPos).intConditionSeqNo
                                intKeyConditionSeqNoBr = Me.strcElePattern(intNowPos).intConditionSeqNoBr

                                'IN・OUT区分を設定する
                                Select Case Left(Me.strcElePattern(intNowPos).strConditionCd, 1)
                                    Case CdCst.JudgeDiv.InSign
                                        strCondSign = CdCst.JudgeDiv.InSign
                                    Case CdCst.JudgeDiv.OutSign
                                        strCondSign = CdCst.JudgeDiv.OutSign
                                End Select

                                'Skip要素の中のグループの開始・終了位置を設定する
                                '開始位置設定
                                intStaPos = intNowPos
                                Do Until intNowPos > intElePatternEndPos
                                    If strKeyOptionSymbol = Me.strcElePattern(intNowPos).strOptionSymbol And _
                                       strKeyConditionCd = Me.strcElePattern(intNowPos).strConditionCd And _
                                       intKeyConditionSeqNo = Me.strcElePattern(intNowPos).intConditionSeqNo And _
                                       intKeyConditionSeqNoBr = Me.strcElePattern(intNowPos).intConditionSeqNoBr Then
                                        '現在位置設定
                                        intNowPos = intNowPos + 1
                                    Else
                                        Exit Do
                                    End If
                                Loop
                                '終了位置設定
                                intEndPos = intNowPos - 1

                                '各要素パターン毎にTrue・Falseの設定を行う
                                For intLoopCnt2 = intStaPos To intEndPos
                                    'EQ・NEにより設定
                                    If Mid(Me.strcElePattern(intLoopCnt2).strConditionCd, 4, 2) = CdCst.JudgeDiv.Equal Then
                                        bolPluralChkFlg = False
                                    Else
                                        bolPluralChkFlg = True
                                    End If

                                    '選択した複数のオプションをカンマで分割する(複数選択要素の場合の対応)
                                    strAryOption = objKtbnStrc.strcSelection.strOpSymbol(Mid(Me.strcElePattern(intLoopCnt2).strConditionCd, 2, 2)).Split(CdCst.Sign.Delimiter.Comma)
                                    For intLoopCnt3 = 0 To strAryOption.Length - 1
                                        If Me.strcElePattern(intLoopCnt2).strCondOptionSymbol = strAryOption(intLoopCnt3) Then
                                            If Mid(Me.strcElePattern(intLoopCnt2).strConditionCd, 4, 2) = CdCst.JudgeDiv.Equal Then
                                                bolPluralChkFlg = True
                                            Else
                                                bolPluralChkFlg = False
                                            End If
                                        End If
                                    Next

                                    If bolPluralChkFlg = False Then
                                        Me.strcElePattern(intLoopCnt2).bolCondFlag = False
                                    End If
                                Next
                                'グループのオプションの可否判定(各オプションの状況をグループ全体に反映する)
                                For intLoopCnt2 = intStaPos To intEndPos
                                    If Mid(Me.strcElePattern(intLoopCnt2).strConditionCd, 4, 2) = CdCst.JudgeDiv.Equal Then
                                        'EQの場合は一つでもTrueが存在する場合はグループ全体をTrueに書き換える
                                        If Me.strcElePattern(intLoopCnt2).bolCondFlag = True Then
                                            For intLoopCnt3 = intStaPos To intEndPos
                                                Me.strcElePattern(intLoopCnt3).bolCondFlag = True
                                            Next
                                            Exit For
                                        End If
                                    Else
                                        'NEの場合は一つでもFalseが存在する場合はグループ全体をFalseに書き換える
                                        If Me.strcElePattern(intLoopCnt2).bolCondFlag = False Then
                                            For intLoopCnt3 = intStaPos To intEndPos
                                                Me.strcElePattern(intLoopCnt3).bolCondFlag = False
                                            Next
                                            Exit For
                                        End If
                                    End If
                                Next
                            Loop

                            'And条件(*)
                            For intLoopCnt2 = intElePatternStaPos To intElePatternEndPos
                                'And条件の範囲を指定
                                If Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.CondAnd Then
                                    '開始位置設定
                                    intStaPos = intElePatternStaPos
                                    For intLoopCnt3 = intLoopCnt2 To intElePatternStaPos Step -1
                                        If Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) <> CdCst.JudgeDiv.CondAnd Then
                                            For intLoopCnt4 = intLoopCnt3 To intElePatternStaPos Step -1
                                                If Me.strcElePattern(intLoopCnt3).strConditionCd = Me.strcElePattern(intLoopCnt4).strConditionCd And _
                                                   Me.strcElePattern(intLoopCnt3).intConditionSeqNo = Me.strcElePattern(intLoopCnt4).intConditionSeqNo And _
                                                   Me.strcElePattern(intLoopCnt3).intConditionSeqNoBr = Me.strcElePattern(intLoopCnt4).intConditionSeqNoBr Then
                                                    intStaPos = intLoopCnt4
                                                End If
                                            Next
                                            Exit For
                                        End If
                                    Next
                                    '終了位置設定
                                    intEndPos = intElePatternEndPos
                                    For intLoopCnt3 = intLoopCnt2 To intElePatternEndPos
                                        If Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) <> CdCst.JudgeDiv.CondAnd Then
                                            intEndPos = intLoopCnt3 - 1
                                            Exit For
                                        End If
                                    Next
                                    For intLoopCnt3 = intStaPos To intEndPos
                                        'And条件の範囲内にFalseがあった場合は全てFalseに書き換え
                                        If Me.strcElePattern(intLoopCnt3).bolCondFlag = False Then
                                            For intLoopCnt4 = intStaPos To intEndPos
                                                Me.strcElePattern(intLoopCnt4).bolCondFlag = False
                                            Next
                                        End If
                                    Next
                                End If
                            Next

                            'Or条件(+)
                            For intLoopCnt2 = intElePatternStaPos To intElePatternEndPos
                                'Or条件の範囲を指定
                                If Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.CondOr Or _
                                   Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.InSign Or _
                                   Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.OutSign Then
                                    '開始位置設定
                                    intStaPos = intElePatternStaPos
                                    '終了位置設定
                                    intEndPos = intElePatternEndPos
                                    For intLoopCnt3 = intLoopCnt2 + 1 To intElePatternEndPos
                                        If Me.strcElePattern(intLoopCnt3).strConditionCd <> Me.strcElePattern(intLoopCnt2 + 1).strConditionCd Or _
                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNo <> Me.strcElePattern(intLoopCnt2 + 1).intConditionSeqNo Or _
                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNoBr <> Me.strcElePattern(intLoopCnt2 + 1).intConditionSeqNoBr Then
                                            If Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) = CdCst.JudgeDiv.CondOr Or _
                                               Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) = CdCst.JudgeDiv.InSign Or _
                                               Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) = CdCst.JudgeDiv.OutSign Then
                                                intEndPos = intLoopCnt3 - 1
                                                Exit For
                                            End If
                                        End If
                                    Next

                                    For intLoopCnt3 = intStaPos To intEndPos
                                        'Or条件の範囲内にTrueがあった場合は全てTrueに書き換え
                                        If Me.strcElePattern(intLoopCnt3).bolCondFlag = True Then
                                            For intLoopCnt4 = intStaPos To intEndPos
                                                Me.strcElePattern(intLoopCnt4).bolCondFlag = True
                                            Next
                                        End If
                                    Next
                                End If
                            Next

                            'オプションの可否判定
                            For intLoopCnt2 = intElePatternStaPos To intElePatternEndPos
                                If strCondSign = CdCst.JudgeDiv.InSign Then
                                    '選択条件にFalseがある場合はFalseに変更
                                    If Me.strcElePattern(intLoopCnt2).bolCondFlag = False Then
                                        Me.strcKtbnStrcEle(intKtbnStrcEleCnt).bolOptionFlag = False
                                        Exit For
                                    End If
                                Else
                                    '選択条件にTrueがある場合はTrueに変更
                                    If Me.strcElePattern(intLoopCnt2).bolCondFlag = True Then
                                        Me.strcKtbnStrcEle(intKtbnStrcEleCnt).bolOptionFlag = False
                                        Exit For
                                    End If
                                End If
                            Next

                            '現在の要素終了
                            Exit For
                        End If
                    Next
                Next
            End If

            '選択条件判定
            If bolSelectCond Then
                'オプション数分チェックする
                For intKtbnStrcEleCnt = 1 To Me.strcKtbnStrcEle.Length - 1
                    '該当する要素を検索
                    For intLoopCnt1 = 1 To Me.strcElePattern.Length - 1
                        '要素がある場合
                        If Me.strcKtbnStrcEle(intKtbnStrcEleCnt).strOptionSymbol = Me.strcElePattern(intLoopCnt1).strOptionSymbol Then
                            '選択要素全体の開始・終了位置を設定する
                            '開始位置設定
                            intElePatternStaPos = intLoopCnt1
                            For intLoopCnt2 = intLoopCnt1 To Me.strcElePattern.Length - 1
                                If Me.strcKtbnStrcEle(intKtbnStrcEleCnt).strOptionSymbol = Me.strcElePattern(intLoopCnt2).strOptionSymbol Then
                                    '終了位置設定
                                    intElePatternEndPos = intLoopCnt2
                                End If
                            Next

                            '現在位置設定
                            intNowPos = intElePatternStaPos
                            '要素チェック
                            Do Until intNowPos > intElePatternEndPos
                                'ブレイクキー設定
                                strKeyOptionSymbol = Me.strcElePattern(intNowPos).strOptionSymbol
                                strKeyConditionCd = Me.strcElePattern(intNowPos).strConditionCd
                                intKeyConditionSeqNo = Me.strcElePattern(intNowPos).intConditionSeqNo
                                intKeyConditionSeqNoBr = Me.strcElePattern(intNowPos).intConditionSeqNoBr

                                'IN・OUT区分を設定する
                                Select Case Left(Me.strcElePattern(intNowPos).strConditionCd, 1)
                                    Case CdCst.JudgeDiv.InSign
                                        strCondSign = CdCst.JudgeDiv.InSign
                                    Case CdCst.JudgeDiv.OutSign
                                        strCondSign = CdCst.JudgeDiv.OutSign
                                End Select

                                '選択要素の中のグループの開始・終了位置を設定する
                                '開始位置設定
                                intStaPos = intNowPos
                                Do Until intNowPos > intElePatternEndPos
                                    If strKeyOptionSymbol = Me.strcElePattern(intNowPos).strOptionSymbol And _
                                       strKeyConditionCd = Me.strcElePattern(intNowPos).strConditionCd And _
                                       intKeyConditionSeqNo = Me.strcElePattern(intNowPos).intConditionSeqNo And _
                                       intKeyConditionSeqNoBr = Me.strcElePattern(intNowPos).intConditionSeqNoBr Then
                                        '現在位置設定
                                        intNowPos = intNowPos + 1
                                    Else
                                        Exit Do
                                    End If
                                Loop
                                '終了位置設定
                                intEndPos = intNowPos - 1

                                '各要素パターン毎にTrue・Falseの設定を行う
                                For intLoopCnt2 = intStaPos To intEndPos
                                    'EQ・NEにより設定
                                    If Mid(Me.strcElePattern(intLoopCnt2).strConditionCd, 4, 2) = CdCst.JudgeDiv.Equal Then
                                        bolPluralChkFlg = False
                                    Else
                                        bolPluralChkFlg = True
                                    End If

                                    '選択した複数のオプションをカンマで分割する
                                    strAryOption = objKtbnStrc.strcSelection.strOpSymbol(Mid(Me.strcElePattern(intLoopCnt2).strConditionCd, 2, 2)).Split(CdCst.Sign.Delimiter.Comma)
                                    For intLoopCnt3 = 0 To strAryOption.Length - 1
                                        If Me.strcElePattern(intLoopCnt2).strCondOptionSymbol = strAryOption(intLoopCnt3) Then
                                            If Mid(Me.strcElePattern(intLoopCnt2).strConditionCd, 4, 2) = CdCst.JudgeDiv.Equal Then
                                                bolPluralChkFlg = True
                                            Else
                                                bolPluralChkFlg = False
                                            End If
                                        End If
                                    Next

                                    If bolPluralChkFlg = False Then
                                        Me.strcElePattern(intLoopCnt2).bolCondFlag = False
                                    End If
                                Next
                                'グループのオプションの可否判定(各オプションの状況をグループ全体に反映する)
                                For intLoopCnt2 = intStaPos To intEndPos
                                    If Mid(Me.strcElePattern(intLoopCnt2).strConditionCd, 4, 2) = CdCst.JudgeDiv.Equal Then
                                        'EQの場合は一つでもTrueが存在する場合はグループ全体をTrueに書き換える
                                        If Me.strcElePattern(intLoopCnt2).bolCondFlag = True Then
                                            For intLoopCnt3 = intStaPos To intEndPos
                                                Me.strcElePattern(intLoopCnt3).bolCondFlag = True
                                            Next
                                            Exit For
                                        End If
                                    Else
                                        'NEの場合は一つでもFalseが存在する場合はグループ全体をFalseに書き換える
                                        If Me.strcElePattern(intLoopCnt2).bolCondFlag = False Then
                                            For intLoopCnt3 = intStaPos To intEndPos
                                                Me.strcElePattern(intLoopCnt3).bolCondFlag = False
                                            Next
                                            Exit For
                                        End If
                                    End If
                                Next
                            Loop
                            '括弧の中の条件(*)
                            intStaPos = intElePatternStaPos
                            intNowPos = intElePatternEndPos
                            intEndPos = intElePatternEndPos
                            '条件を後ろから検索する
                            Do Until intNowPos < intElePatternStaPos
                                '括弧"("が見つかった場合
                                If Right(Me.strcElePattern(intNowPos).strConditionCd.Trim, 1) = "(" Then
                                    '括弧の範囲設定
                                    '開始位置設定
                                    intStaPos = intElePatternStaPos
                                    For intLoopCnt2 = intNowPos To intElePatternStaPos Step -1
                                        If Me.strcElePattern(intLoopCnt2).strConditionCd = Me.strcElePattern(intNowPos).strConditionCd And _
                                           Me.strcElePattern(intLoopCnt2).intConditionSeqNo = Me.strcElePattern(intNowPos).intConditionSeqNo And _
                                           Me.strcElePattern(intLoopCnt2).intConditionSeqNoBr = Me.strcElePattern(intNowPos).intConditionSeqNoBr Then
                                            intStaPos = intLoopCnt2
                                        End If
                                    Next
                                    '終了位置設定
                                    intEndPos = intElePatternEndPos
                                    For intLoopCnt2 = intNowPos + 1 To intElePatternEndPos
                                        If Right(Me.strcElePattern(intLoopCnt2).strConditionCd.Trim, 1) = "(" Then
                                            intEndPos = intLoopCnt2 - 1
                                            Exit For
                                        End If
                                        If Right(Me.strcElePattern(intLoopCnt2).strConditionCd.Trim, 1) = ")" Then
                                            intEndPos = intLoopCnt2
                                            Exit For
                                        End If
                                    Next
                                    '括弧の範囲内を検索
                                    For intLoopCnt2 = intStaPos To intEndPos
                                        If Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.CondAnd Then
                                            '開始位置設定
                                            intWStaPos = intStaPos
                                            For intLoopCnt3 = intLoopCnt2 To intStaPos Step -1
                                                If Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) <> CdCst.JudgeDiv.CondAnd Then
                                                    For intLoopCnt4 = intLoopCnt3 To intElePatternStaPos Step -1
                                                        If Me.strcElePattern(intLoopCnt3).strConditionCd = Me.strcElePattern(intLoopCnt4).strConditionCd And _
                                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNo = Me.strcElePattern(intLoopCnt4).intConditionSeqNo And _
                                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNoBr = Me.strcElePattern(intLoopCnt4).intConditionSeqNoBr Then
                                                            intWStaPos = intLoopCnt4
                                                        End If
                                                    Next
                                                    Exit For
                                                End If
                                            Next
                                            '終了位置設定
                                            intWEndPos = intEndPos
                                            For intLoopCnt3 = intLoopCnt2 To intEndPos
                                                If Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) <> CdCst.JudgeDiv.CondAnd Then
                                                    intWEndPos = intLoopCnt3 - 1
                                                    Exit For
                                                End If
                                            Next
                                            For intLoopCnt3 = intWStaPos To intWEndPos
                                                'And条件の範囲内にFalseがあった場合は全てFalseに書き換え
                                                If Me.strcElePattern(intLoopCnt3).bolCondFlag = False Then
                                                    For intLoopCnt4 = intWStaPos To intWEndPos
                                                        Me.strcElePattern(intLoopCnt4).bolCondFlag = False
                                                    Next
                                                End If
                                            Next
                                        End If
                                    Next
                                    intEndPos = intStaPos - 1
                                    intNowPos = intStaPos - 1
                                Else
                                    intNowPos = intNowPos - 1
                                End If
                            Loop

                            '括弧の中の条件(+)
                            intStaPos = intElePatternStaPos
                            intNowPos = intElePatternEndPos
                            intEndPos = intElePatternEndPos
                            '条件を後ろから検索する
                            Do Until intNowPos < intElePatternStaPos
                                '括弧"("が見つかった場合
                                If Right(Me.strcElePattern(intNowPos).strConditionCd.Trim, 1) = "(" Then
                                    '括弧の範囲設定
                                    '開始位置設定
                                    intStaPos = intElePatternStaPos
                                    For intLoopCnt2 = intNowPos To intElePatternStaPos Step -1
                                        If Me.strcElePattern(intLoopCnt2).strConditionCd = Me.strcElePattern(intNowPos).strConditionCd And _
                                           Me.strcElePattern(intLoopCnt2).intConditionSeqNo = Me.strcElePattern(intNowPos).intConditionSeqNo And _
                                           Me.strcElePattern(intLoopCnt2).intConditionSeqNoBr = Me.strcElePattern(intNowPos).intConditionSeqNoBr Then
                                            intStaPos = intLoopCnt2
                                        End If
                                    Next
                                    '終了位置設定
                                    intEndPos = intElePatternEndPos
                                    For intLoopCnt2 = intNowPos + 1 To intElePatternEndPos
                                        If Right(Me.strcElePattern(intLoopCnt2).strConditionCd.Trim, 1) = "(" Then
                                            intEndPos = intLoopCnt2 - 1
                                            Exit For
                                        End If
                                        If Right(Me.strcElePattern(intLoopCnt2).strConditionCd.Trim, 1) = ")" Then
                                            intEndPos = intLoopCnt2
                                            Exit For
                                        End If
                                    Next
                                    '括弧の範囲内を検索
                                    For intLoopCnt2 = intStaPos To intEndPos
                                        'Or条件の範囲を指定
                                        If Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.CondOr Or _
                                           Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.InSign Or _
                                           Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.OutSign Then
                                            '開始位置設定
                                            intWStaPos = intStaPos
                                            '終了位置設定
                                            intWEndPos = intEndPos
                                            For intLoopCnt3 = intLoopCnt2 + 1 To intEndPos
                                                If Me.strcElePattern(intLoopCnt3).strConditionCd <> Me.strcElePattern(intLoopCnt2 + 1).strConditionCd Or _
                                                   Me.strcElePattern(intLoopCnt3).intConditionSeqNo <> Me.strcElePattern(intLoopCnt2 + 1).intConditionSeqNo Or _
                                                   Me.strcElePattern(intLoopCnt3).intConditionSeqNoBr <> Me.strcElePattern(intLoopCnt2 + 1).intConditionSeqNoBr Then
                                                    If Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) = CdCst.JudgeDiv.CondOr Or _
                                                       Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) = CdCst.JudgeDiv.InSign Or _
                                                       Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) = CdCst.JudgeDiv.OutSign Then
                                                        intWEndPos = intLoopCnt3 - 1
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                            For intLoopCnt3 = intWStaPos To intWEndPos
                                                'Or条件の範囲内にTrueがあった場合は全てTrueに書き換え
                                                If Me.strcElePattern(intLoopCnt3).bolCondFlag = True Then
                                                    For intLoopCnt4 = intWStaPos To intWEndPos
                                                        Me.strcElePattern(intLoopCnt4).bolCondFlag = True
                                                    Next
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    Next
                                    intEndPos = intStaPos - 1
                                    intNowPos = intStaPos - 1
                                Else
                                    intNowPos = intNowPos - 1
                                End If
                            Loop

                            '括弧の中の条件を書き換え
                            intStaPos = intElePatternStaPos
                            intNowPos = intElePatternEndPos
                            intEndPos = intElePatternEndPos
                            '条件を後ろから検索する
                            Do Until intNowPos < intElePatternStaPos
                                '括弧"("が見つかった場合
                                If Right(Me.strcElePattern(intNowPos).strConditionCd.Trim, 1) = "(" Then
                                    '開始位置設定
                                    For intLoopCnt2 = intNowPos To intElePatternStaPos Step -1
                                        If Me.strcElePattern(intLoopCnt2).strConditionCd = Me.strcElePattern(intNowPos).strConditionCd And _
                                           Me.strcElePattern(intLoopCnt2).intConditionSeqNo = Me.strcElePattern(intNowPos).intConditionSeqNo And _
                                           Me.strcElePattern(intLoopCnt2).intConditionSeqNoBr = Me.strcElePattern(intNowPos).intConditionSeqNoBr Then
                                            intStaPos = intLoopCnt2
                                        End If
                                    Next
                                    '終了位置設定
                                    For intLoopCnt2 = intNowPos + 1 To intEndPos
                                        If Right(Me.strcElePattern(intLoopCnt2).strConditionCd.Trim, 1) = "(" Then
                                            intEndPos = intLoopCnt2 - 1
                                            Exit For
                                        End If
                                        If Right(Me.strcElePattern(intLoopCnt2).strConditionCd.Trim, 1) = ")" Then
                                            intEndPos = intLoopCnt2
                                            Exit For
                                        End If
                                    Next
                                    '開始位置から終了位置まで検索
                                    For intLoopCnt2 = intStaPos To intEndPos
                                        '条件を開始位置の値(頭5桁)に書き換え
                                        Me.strcElePattern(intLoopCnt2).strConditionCd = Left(Me.strcElePattern(intStaPos).strConditionCd.Trim, 5)
                                        Me.strcElePattern(intLoopCnt2).intConditionSeqNo = Me.strcElePattern(intStaPos).intConditionSeqNo
                                        Me.strcElePattern(intLoopCnt2).intConditionSeqNoBr = Me.strcElePattern(intStaPos).intConditionSeqNoBr
                                    Next

                                    intEndPos = intStaPos - 1
                                    intNowPos = intStaPos - 1
                                Else
                                    intNowPos = intNowPos - 1
                                End If
                            Loop

                            'And条件(*)
                            For intLoopCnt2 = intElePatternStaPos To intElePatternEndPos
                                'And条件の範囲を指定
                                If Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.CondAnd Then
                                    '開始位置設定
                                    intStaPos = intElePatternStaPos
                                    For intLoopCnt3 = intLoopCnt2 To intElePatternStaPos Step -1
                                        If Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) <> CdCst.JudgeDiv.CondAnd Then
                                            For intLoopCnt4 = intLoopCnt3 To intElePatternStaPos Step -1
                                                If Me.strcElePattern(intLoopCnt3).strConditionCd = Me.strcElePattern(intLoopCnt4).strConditionCd And _
                                                   Me.strcElePattern(intLoopCnt3).intConditionSeqNo = Me.strcElePattern(intLoopCnt4).intConditionSeqNo And _
                                                   Me.strcElePattern(intLoopCnt3).intConditionSeqNoBr = Me.strcElePattern(intLoopCnt4).intConditionSeqNoBr Then
                                                    intStaPos = intLoopCnt4
                                                End If
                                            Next
                                            Exit For
                                        End If
                                    Next
                                    '終了位置設定
                                    intEndPos = intElePatternEndPos
                                    For intLoopCnt3 = intLoopCnt2 To intElePatternEndPos
                                        If Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) <> CdCst.JudgeDiv.CondAnd Then
                                            intEndPos = intLoopCnt3 - 1
                                            Exit For
                                        End If
                                    Next
                                    For intLoopCnt3 = intStaPos To intEndPos
                                        'And条件の範囲内にFalseがあった場合は全てFalseに書き換え
                                        If Me.strcElePattern(intLoopCnt3).bolCondFlag = False Then
                                            For intLoopCnt4 = intStaPos To intEndPos
                                                Me.strcElePattern(intLoopCnt4).bolCondFlag = False
                                            Next
                                        End If
                                    Next
                                End If
                            Next

                            'Or条件(+)
                            For intLoopCnt2 = intElePatternStaPos To intElePatternEndPos
                                'Or条件の範囲を指定
                                If Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.CondOr Or _
                                   Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.InSign Or _
                                   Left(Me.strcElePattern(intLoopCnt2).strConditionCd, 1) = CdCst.JudgeDiv.OutSign Then
                                    '開始位置設定
                                    intStaPos = intElePatternStaPos
                                    '終了位置設定
                                    intEndPos = intElePatternEndPos
                                    For intLoopCnt3 = intLoopCnt2 + 1 To intElePatternEndPos
                                        If Me.strcElePattern(intLoopCnt3).strConditionCd <> Me.strcElePattern(intLoopCnt2 + 1).strConditionCd Or _
                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNo <> Me.strcElePattern(intLoopCnt2 + 1).intConditionSeqNo Or _
                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNoBr <> Me.strcElePattern(intLoopCnt2 + 1).intConditionSeqNoBr Then
                                            If Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) = CdCst.JudgeDiv.CondOr Or _
                                               Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) = CdCst.JudgeDiv.InSign Or _
                                               Left(Me.strcElePattern(intLoopCnt3).strConditionCd, 1) = CdCst.JudgeDiv.OutSign Then
                                                intEndPos = intLoopCnt3 - 1
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    For intLoopCnt3 = intStaPos To intEndPos
                                        'Or条件の範囲内にTrueがあった場合は全てTrueに書き換え
                                        If Me.strcElePattern(intLoopCnt3).bolCondFlag = True Then
                                            For intLoopCnt4 = intStaPos To intEndPos
                                                Me.strcElePattern(intLoopCnt4).bolCondFlag = True
                                            Next
                                            Exit For
                                        End If
                                    Next
                                End If
                            Next
                            'オプションの可否判定
                            For intLoopCnt2 = intElePatternStaPos To intElePatternEndPos
                                If strCondSign = CdCst.JudgeDiv.InSign Then
                                    '選択条件にFalseがある場合はFalseに変更
                                    If Me.strcElePattern(intLoopCnt2).bolCondFlag = False Then
                                        Me.strcKtbnStrcEle(intKtbnStrcEleCnt).bolOptionFlag = False
                                        Exit For
                                    End If
                                Else
                                    '選択条件にTrueがある場合はTrueに変更
                                    If Me.strcElePattern(intLoopCnt2).bolCondFlag = True Then
                                        Me.strcKtbnStrcEle(intKtbnStrcEleCnt).bolOptionFlag = False
                                        Exit For
                                    End If
                                End If
                            Next
                            '現在の要素終了
                            Exit For
                        End If
                    Next
                Next
            End If

            '複数選択条件判定
            If bolPluralCond = True Then
                If strDivision = "1" Then
                    '処理無し
                Else
                    'オプションチェック
                    '選択した複数のオプションをカンマで分割する
                    strAryOption = objKtbnStrc.strcSelection.strOpSymbol(intKtbnStrcSeqNo).Split(CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt1 = 0 To strAryOption.Length - 1
                        For intLoopCnt2 = 1 To UBound(Me.strcElePattern)
                            '複数選択項目の場合
                            If Me.strcElePattern(intLoopCnt2).strOptionSymbol = CdCst.ElePattern.Plural Then
                                If strAryOption(intLoopCnt1) = Me.strcElePattern(intLoopCnt2).strCondOptionSymbol Then
                                    'グループ設定
                                    '開始位置設定
                                    For intLoopCnt3 = intLoopCnt2 To 1 Step -1
                                        If Me.strcElePattern(intLoopCnt3).strOptionSymbol = Me.strcElePattern(intLoopCnt2).strOptionSymbol And _
                                           Me.strcElePattern(intLoopCnt3).strConditionCd = Me.strcElePattern(intLoopCnt2).strConditionCd And _
                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNo = Me.strcElePattern(intLoopCnt2).intConditionSeqNo And _
                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNoBr = Me.strcElePattern(intLoopCnt2).intConditionSeqNoBr Then
                                            '開始位置設定
                                            intStaPos = intLoopCnt3
                                        End If
                                    Next
                                    '終了位置設定
                                    For intLoopCnt3 = intLoopCnt2 To UBound(Me.strcElePattern)
                                        If Me.strcElePattern(intLoopCnt3).strOptionSymbol = Me.strcElePattern(intLoopCnt2).strOptionSymbol And _
                                           Me.strcElePattern(intLoopCnt3).strConditionCd = Me.strcElePattern(intLoopCnt2).strConditionCd And _
                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNo = Me.strcElePattern(intLoopCnt2).intConditionSeqNo And _
                                           Me.strcElePattern(intLoopCnt3).intConditionSeqNoBr = Me.strcElePattern(intLoopCnt2).intConditionSeqNoBr Then
                                            '終了位置設定
                                            intEndPos = intLoopCnt3
                                        End If
                                    Next
                                    '同一グループ内で他のオプションを選択していないかチェックする
                                    For intLoopCnt3 = intStaPos To intEndPos
                                        For intLoopCnt4 = 0 To strAryOption.Length - 1
                                            '自分自身でない場合
                                            If strAryOption(intLoopCnt1) <> strAryOption(intLoopCnt4) Then
                                                If strAryOption(intLoopCnt4) = Me.strcElePattern(intLoopCnt3).strCondOptionSymbol Then
                                                    For intKtbnStrcEleCnt = 1 To Me.strcKtbnStrcEle.Length - 1
                                                        If strAryOption(intLoopCnt1) = Me.strcKtbnStrcEle(intKtbnStrcEleCnt).strOptionSymbol Then
                                                            Me.strcKtbnStrcEle(intKtbnStrcEleCnt).bolOptionFlag = False
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                            End If
                        Next
                    Next
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objKtbnStrc = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 選択されたオプションをチェックする
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strDivision">処理区分</param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <param name="strOptionSymbol">オプション記号</param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strMessageCd">メッセージコード</param>
    ''' <param name="dt_KataStrcEleSel"></param>
    ''' <param name="dt_ElePatternSel"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncOptionCheck(ByVal objCon As SqlConnection, ByVal strDivision As String, ByVal strUserId As String, _
                                   ByVal strSessionId As String, ByVal intKtbnStrcSeqNo As Integer, _
                                   ByVal strOptionSymbol As String, ByVal objKtbnStrc As KHKtbnStrc, _
                                   Optional ByRef strMessageCd As String = Nothing, _
                                   Optional ByVal dt_KataStrcEleSel As DataTable = Nothing, _
                                   Optional ByVal dt_ElePatternSel As DataTable = Nothing) As Boolean
        Dim strArrayOption(,) As String = Nothing
        Dim intLoopCnt As Integer
        Dim bolFindFlg As Boolean = False
        Dim intBoreSize As Integer
        fncOptionCheck = False
        Try
            'オプションリスト取得 
            Call Me.subOptionList(objCon, objKtbnStrc, strDivision, strUserId, strSessionId, CdCst.LanguageCd.DefaultLang, _
                                  intKtbnStrcSeqNo, strArrayOption, dt_KataStrcEleSel, dt_ElePatternSel)

            'オプションリストが1件も無く、オプションが無い場合はTrue
            If UBound(strArrayOption) = 0 Then
                If strOptionSymbol.Trim.Length <> 0 Then
                Else
                    bolFindFlg = True
                End If
                '組み合わせ出力時は下の1行を有効にする
                'bolFindFlg = True
            Else
                '引当情報取得
                'Call objKtbnStrc.subSelKtbnInfoGet(strUserId, strSessionId)
                'オプションリストに選択したオプションが存在するかチェックする
                For intLoopCnt = 1 To UBound(strArrayOption)
                    If strOptionSymbol = strArrayOption(intLoopCnt, 1) Then
                        '「その他電圧」の場合は選択不可の為False
                        Select Case strOptionSymbol
                            Case CdCst.OtherVoltage.Japanese, CdCst.OtherVoltage.English
                            Case Else
                                '電圧の場合
                                If objKtbnStrc.strcSelection.strOpElementDiv(intKtbnStrcSeqNo) = CdCst.ElementDiv.Voltage Then
                                    'AC/DCの場合および規定の電圧の場合はTrue(/50Hz等を除外する)
                                    Select Case Left(strOptionSymbol, 2)
                                        Case CdCst.PowerSupply.Div1, CdCst.PowerSupply.Div2
                                            'AC/DCの場合
                                            bolFindFlg = True
                                        Case Else
                                            '規定の電圧の場合
                                            Select Case strOptionSymbol
                                                Case CdCst.PowerSupply.AC100V, CdCst.PowerSupply.AC200V, _
                                                     CdCst.PowerSupply.DC24V, CdCst.PowerSupply.DC12V, _
                                                     CdCst.PowerSupply.AC110V, CdCst.PowerSupply.AC220V
                                                    bolFindFlg = True
                                            End Select
                                    End Select
                                Else
                                    bolFindFlg = True
                                End If
                        End Select
                        Exit For
                    End If
                Next
                'Falseの場合
                If bolFindFlg = False Then
                    '電圧及びストロークの場合のチェック
                    Select Case objKtbnStrc.strcSelection.strOpElementDiv(intKtbnStrcSeqNo)
                        Case CdCst.ElementDiv.Voltage
                            'その他電圧が存在する場合のみ入力された電圧をチェックする
                            For intLoopCnt = 1 To UBound(strArrayOption)
                                'Select Case strArrayOption(intLoopCnt, 1).ToUpper
                                '    Case CdCst.OtherVoltage.Japanese, CdCst.OtherVoltage.English
                                If strArrayOption(intLoopCnt, 1).ToUpper.StartsWith(CdCst.OtherVoltage.English) Then
                                    '電圧オプションの場合
                                    '電圧の妥当性チェック
                                    Dim bolVoltage As Boolean = False
                                    Select Case strOptionSymbol
                                        Case CdCst.PowerSupply.AC100V, CdCst.PowerSupply.AC110V, _
                                             CdCst.PowerSupply.AC200V, CdCst.PowerSupply.AC220V, _
                                             CdCst.PowerSupply.DC12V, CdCst.PowerSupply.DC24V
                                            bolVoltage = True
                                        Case Else
                                            Select Case Left(strOptionSymbol, 2)
                                                Case CdCst.PowerSupply.Div1, CdCst.PowerSupply.Div2
                                                    If strOptionSymbol.IndexOf("V") >= 0 Then
                                                        If IsNumeric(Mid(strOptionSymbol, 3, strOptionSymbol.IndexOf("V") - 2)) Then
                                                            bolVoltage = True
                                                        End If
                                                    End If
                                            End Select
                                    End Select
                                    If bolVoltage = True Then
                                        '電圧チェック
                                        If Me.fncVoltageCheck(objCon, objKtbnStrc, strOptionSymbol) = True Then
                                            bolFindFlg = True
                                        Else
                                            Me.strMsgCd = "W0180" 'メッセージコードセット
                                        End If
                                    Else
                                        Me.strMsgCd = "W0180" 'メッセージコードセット
                                    End If
                                    Exit For
                                End If
                                'End Select
                            Next
                        Case CdCst.ElementDiv.Stroke 'ストロークオプションの場合
                            '口径設定
                            intBoreSize = 0
                            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOpElementDiv.Length - 1
                                If objKtbnStrc.strcSelection.strOpElementDiv(intLoopCnt) = CdCst.ElementDiv.Port Then
                                    If IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim) Then
                                        intBoreSize = CInt(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim)
                                    Else
                                        If objKtbnStrc.strcSelection.strSeriesKataban = "ESM" Then
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                                Case "ST"
                                                    intBoreSize = 2
                                                Case "B"
                                                    intBoreSize = 1
                                                Case Else
                                            End Select
                                        End If
                                    End If
                                    Exit For
                                End If
                            Next
                            If intBoreSize <> 0 Then
                                'ストロークチェック
                                If Me.fncStrokeCheck(objCon, objKtbnStrc.strcSelection.strSeriesKataban, _
                                                     objKtbnStrc.strcSelection.strKeyKataban, _
                                                     intBoreSize, IIf(IsNumeric(strOptionSymbol), strOptionSymbol, 0), _
                                                     objKtbnStrc.strcSelection.strMadeCountry) = True Then
                                    bolFindFlg = True
                                    'ストローク(小数点)チェック
                                    If InStr(1, (strOptionSymbol), ".") <> 0 Then
                                        'メッセージコードセット
                                        Me.strMsgCd = "W0170"
                                        bolFindFlg = False
                                    End If
                                Else
                                    Me.strMsgCd = "W0170" 'メッセージコードセット
                                End If
                            Else
                                Me.strMsgCd = "W0170" 'メッセージコードセット
                            End If
                        Case Else
                            bolFindFlg = False
                            Me.strMsgCd = "W2520"
                    End Select
                End If
            End If
            '戻り値設定
            If bolFindFlg Then
                fncOptionCheck = True
            Else
                fncOptionCheck = False
                'メッセージコードセット
                strMessageCd = Me.strMsgCd
            End If
        Catch ex As Exception
            strMessageCd = ex.Message
        Finally
            objKtbnStrc = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 電圧チェック
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc">引当情報</param>
    ''' <param name="strVoltage">電圧</param>
    ''' <returns></returns>
    ''' <remarks>電圧の可否をチェックする</remarks>
    Private Function fncVoltageCheck(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc, _
                                     ByVal strVoltage As String) As Boolean
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        Dim intVoltage As Integer
        Dim strVoltageDiv As String = Nothing
        Dim strSeriesKataban As String = Nothing
        Dim strKeyKataban As String = Nothing
        Dim strPortSize As String = Nothing
        Dim strCoil As String = Nothing

        fncVoltageCheck = False
        Try
            '電圧検索情報取得
            Call subVoltageSearchInfoGet(objKtbnStrc, strVoltage, intVoltage, strVoltageDiv, _
                                         strSeriesKataban, strKeyKataban, strPortSize, strCoil)
            'SQL Query生成
            sbSql.Append(" SELECT  a.min_voltage, ")
            sbSql.Append("         a.max_voltage, ")
            sbSql.Append("         b.std_voltage ")
            sbSql.Append(" FROM    kh_voltage  a ")
            sbSql.Append(" LEFT JOIN  kh_std_voltage_mst  b ")
            sbSql.Append(" ON      a.series_kataban      = b.series_kataban ")
            sbSql.Append(" AND     a.key_kataban         = b.key_kataban ")
            sbSql.Append(" AND     a.port_size           = b.port_size ")
            sbSql.Append(" AND     a.coil                = b.coil ")
            sbSql.Append(" AND     a.voltage_div         = b.voltage_div ")
            sbSql.Append(" WHERE   a.series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     a.key_kataban         = @KeyKataban ")
            If Not strPortSize Is Nothing Then sbSql.Append(" AND     a.port_size       = @PortSize ")
            If Not strCoil Is Nothing Then sbSql.Append(" AND     a.coil                = @Coil ")
            sbSql.Append(" AND     a.voltage_div         = @VoltageDiv ")
            sbSql.Append(" AND     a.in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     a.out_effective_date  > @StandardDate ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
                If Not strPortSize Is Nothing Then
                    .Parameters.Add("@PortSize", SqlDbType.VarChar, 4).Value = strPortSize
                End If
                If Not strCoil Is Nothing Then
                    .Parameters.Add("@Coil", SqlDbType.VarChar, 2).Value = strCoil
                End If
                If strVoltageDiv = CdCst.PowerSupply.Div1 Then
                    .Parameters.Add("@VoltageDiv", SqlDbType.VarChar, 1).Value = CdCst.PowerSupply.AC
                Else
                    .Parameters.Add("@VoltageDiv", SqlDbType.VarChar, 1).Value = CdCst.PowerSupply.DC
                End If
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
            End With

            objRdr = objCmd.ExecuteReader
            While objRdr.Read()
                If objRdr.GetValue(objRdr.GetOrdinal("min_voltage")) = 0 And _
                   objRdr.GetValue(objRdr.GetOrdinal("max_voltage")) = 0 Then
                    '電圧指定の場合
                    '標準電圧に存在するかチェック
                    If intVoltage = objRdr.GetValue(objRdr.GetOrdinal("std_voltage")) Then
                        fncVoltageCheck = True
                        Exit While
                    End If
                Else
                    '範囲指定の場合
                    'Min<=電圧<=Maxであるかチェック
                    If intVoltage >= objRdr.GetValue(objRdr.GetOrdinal("min_voltage")) And _
                       intVoltage <= objRdr.GetValue(objRdr.GetOrdinal("max_voltage")) Then
                        fncVoltageCheck = True
                        Exit While
                    End If
                End If
            End While
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 電圧検索情報取得
    ''' </summary>
    ''' <param name="objKtbnStrc">引当情報</param>
    ''' <param name="strVoltage">電圧</param>
    ''' <param name="intVoltage">電圧</param>
    ''' <param name="strVoltageDiv">電圧区分</param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strKeyKataban">キー形番</param>
    ''' <param name="strPortSize">口径</param>
    ''' <param name="strCoil">コイル</param>
    ''' <remarks></remarks>
    Public Sub subVoltageSearchInfoGet(ByVal objKtbnStrc As KHKtbnStrc, ByVal strVoltage As String, _
                                       ByRef intVoltage As Integer, ByRef strVoltageDiv As String, _
                                       ByRef strSeriesKataban As String, ByRef strKeyKataban As String, _
                                       ByRef strPortSize As String, ByRef strCoil As String)
        Dim intLoopCnt As Integer
        Try
            intVoltage = 0
            strVoltageDiv = Nothing
            strSeriesKataban = Nothing
            strKeyKataban = Nothing
            strPortSize = Nothing
            strCoil = Nothing
            '電圧設定
            Select Case strVoltage
                Case CdCst.PowerSupply.AC100V
                    intVoltage = CInt(Mid(CdCst.PowerSupply.Const1, 3, CdCst.PowerSupply.Const1.IndexOf("V") - 2))
                    strVoltageDiv = CdCst.PowerSupply.Div1
                Case CdCst.PowerSupply.AC200V
                    intVoltage = CInt(Mid(CdCst.PowerSupply.Const2, 3, CdCst.PowerSupply.Const2.IndexOf("V") - 2))
                    strVoltageDiv = CdCst.PowerSupply.Div1
                Case CdCst.PowerSupply.DC24V
                    intVoltage = CInt(Mid(CdCst.PowerSupply.Const3, 3, CdCst.PowerSupply.Const3.IndexOf("V") - 2))
                    strVoltageDiv = CdCst.PowerSupply.Div2
                Case CdCst.PowerSupply.DC12V
                    intVoltage = CInt(Mid(CdCst.PowerSupply.Const4, 3, CdCst.PowerSupply.Const4.IndexOf("V") - 2))
                    strVoltageDiv = CdCst.PowerSupply.Div2
                Case CdCst.PowerSupply.AC110V
                    intVoltage = CInt(Mid(CdCst.PowerSupply.Const5, 3, CdCst.PowerSupply.Const5.IndexOf("V") - 2))
                    strVoltageDiv = CdCst.PowerSupply.Div1
                Case CdCst.PowerSupply.AC220V
                    intVoltage = CInt(Mid(CdCst.PowerSupply.Const6, 3, CdCst.PowerSupply.Const6.IndexOf("V") - 2))
                    strVoltageDiv = CdCst.PowerSupply.Div1
                Case Else
                    intVoltage = CInt(Mid(strVoltage, 3, strVoltage.IndexOf("V") - 2))
                    strVoltageDiv = Left(strVoltage, 2)
            End Select

            '接続口径・コイル設定
            For intLoopCnt = 1 To UBound(objKtbnStrc.strcSelection.strOpElementDiv)
                Select Case objKtbnStrc.strcSelection.strOpElementDiv(intLoopCnt)
                    Case CdCst.ElementDiv.Coil
                        strCoil = objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                    Case CdCst.ElementDiv.VolPort
                        strPortSize = objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                End Select
            Next
            If Not strCoil Is Nothing Then
                Select Case strCoil.Trim
                    Case "0", "00"
                        strCoil = ""
                End Select
            End If

            'シリーズ形番・キー形番
            Select Case objKtbnStrc.strcSelection.strPriceNo.Trim
                Case "02", "03"
                    Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1)
                        Case "A"
                            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AB4" Then
                                If strCoil Is Nothing Then
                                    strSeriesKataban = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3)
                                    strKeyKataban = ""
                                Else
                                    If (strCoil.Trim = "3A" Or strCoil.Trim = "3K") And Left(strVoltage, 2) = "DC" Or _
                                       (strCoil.Trim = "5A" Or strCoil.Trim = "5K") And Left(strVoltage, 2) = "AC" Then
                                        strSeriesKataban = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4)
                                        strKeyKataban = ""
                                    Else
                                        strSeriesKataban = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3)
                                        strKeyKataban = ""
                                    End If
                                End If
                            Else
                                strSeriesKataban = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3)
                                strKeyKataban = ""
                            End If
                        Case "G"
                            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) = "AB4" Then
                                If strCoil Is Nothing Then
                                    strSeriesKataban = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 3)
                                    strKeyKataban = ""
                                Else
                                    If (strCoil.Trim = "3A" Or strCoil.Trim = "3K") And Left(strVoltage, 2) = "DC" Or _
                                       (strCoil.Trim = "5A" Or strCoil.Trim = "5K") And Left(strVoltage, 2) = "AC" Then
                                        strSeriesKataban = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 4)
                                        strKeyKataban = ""
                                    Else
                                        strSeriesKataban = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 3)
                                        strKeyKataban = ""
                                    End If
                                End If
                            Else
                                strSeriesKataban = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 3)
                                strKeyKataban = ""
                            End If
                    End Select
                Case Else
                    strSeriesKataban = objKtbnStrc.strcSelection.strSeriesKataban
                    strKeyKataban = objKtbnStrc.strcSelection.strKeyKataban
            End Select
        Catch ex As Exception
            intVoltage = 0
            strVoltageDiv = ""
            strSeriesKataban = objKtbnStrc.strcSelection.strSeriesKataban
            strKeyKataban = objKtbnStrc.strcSelection.strKeyKataban
            strPortSize = ""
            strCoil = ""
        End Try
    End Sub

    ''' <summary>
    ''' ストロークチェック
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSeriesKataban">シリーズ形番</param>
    ''' <param name="strKeyKataban">キー形番</param>
    ''' <param name="intBoreSize">口径</param>
    ''' <param name="intStroke">ストローク</param>
    ''' <param name="strMadeCountry"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncStrokeCheck(ByVal objCon As SqlConnection, ByVal strSeriesKataban As String, _
                                    ByVal strKeyKataban As String, _
                                    ByVal intBoreSize As Integer, ByVal intStroke As Integer, _
                                    ByVal strMadeCountry As String) As Boolean
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        fncStrokeCheck = False

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  min_stroke, ")
            sbSql.Append("         max_stroke, ")
            sbSql.Append("         stroke_unit ")
            sbSql.Append(" FROM    kh_stroke ")
            sbSql.Append(" WHERE   series_kataban      = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban         = @KeyKataban ")
            sbSql.Append(" AND     bore_size           = @BoreSize")
            sbSql.Append(" AND     in_effective_date  <= @StandardDate ")
            sbSql.Append(" AND     out_effective_date  > @StandardDate ")
            sbSql.Append(" AND     country_cd  = @countrycd ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = strKeyKataban
                .Parameters.Add("@BoreSize", SqlDbType.Int).Value = intBoreSize
                .Parameters.Add("@StandardDate", SqlDbType.DateTime).Value = Now()
                .Parameters.Add("@countrycd", SqlDbType.VarChar, 3).Value = strMadeCountry
            End With

            'DBオープン
            objRdr = objCmd.ExecuteReader
            While objRdr.Read()
                If objRdr.GetValue(objRdr.GetOrdinal("min_stroke")) <= intStroke And _
                   objRdr.GetValue(objRdr.GetOrdinal("max_stroke")) >= intStroke Then
                    fncStrokeCheck = True
                End If
            End While
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 複数選択のオプション順序チェック処理
    ''' </summary>
    ''' <param name="strCommaOption">オプション記号(カンマ区切り)</param>
    ''' <param name="strOptionSymbol">オプション記号</param>
    ''' <param name="strMessageCd">メッセージコード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncOptSeqCheck(ByVal strCommaOption As String, ByVal strOptionSymbol As String, _
                                   Optional ByRef strMessageCd As String = Nothing) As Boolean
        Dim strSelAryOption() As String = Nothing
        Dim intLoopCnt1 As Integer
        fncOptSeqCheck = True

        Try
            'オプションを配列に格納
            strSelAryOption = Split(strCommaOption, CdCst.Sign.Comma)

            If UBound(strSelAryOption) <> 0 Then
                'オプションの順序がオプションリストどおりかどうかチェックする
                For intLoopCnt1 = 0 To UBound(strSelAryOption)
                    If strSelAryOption(intLoopCnt1) = Left(strOptionSymbol.Trim, strSelAryOption(intLoopCnt1).Trim.Length) Then
                        strOptionSymbol = Mid(strOptionSymbol.Trim, strSelAryOption(intLoopCnt1).Trim.Length + 1, strOptionSymbol.Trim.Length)
                    Else
                        fncOptSeqCheck = False
                        strMessageCd = "W8650"
                        Exit Try
                    End If
                Next
            Else
                If strSelAryOption(0).Trim <> strOptionSymbol.Trim Then
                    fncOptSeqCheck = False
                    strMessageCd = "W8650"
                    Exit Try
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            strSelAryOption = Nothing
        End Try

    End Function

    ''' <summary>
    ''' チェック区分取得
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strKatabanCheckDiv"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncKatabanCheckDivGet(objKtbnStrc As KHKtbnStrc, ByVal strKatabanCheckDiv As String) As String
        Try
            'デフォルト設定
            fncKatabanCheckDivGet = strKatabanCheckDiv

            ''引当情報取得
            'objKtbnStrc.subSelKtbnInfoGet(strUserId, strSessionId)

            '機種別にチェック
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "AMD3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1", "2"
                            'オプションが"6"・"7"の時
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "6", "7"
                                    fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End Select

                            '取付 = "R"・"X"のとき、チェック区分 =「3」
                            If objKtbnStrc.strcSelection.strOpSymbol(8).IndexOf("R") >= 0 Or _
                               objKtbnStrc.strcSelection.strOpSymbol(8).IndexOf("X") >= 0 Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If

                            '接続が"12UA"で且つ流体が"Y"の時
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "12UA" And _
                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "Y" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                    End Select
                Case "AMD4"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            'オプションが"6"・"7"の場合はチェック区分「3」
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "6", "7"
                                    fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End Select

                            '取付 = "R"・"X"のとき、チェック区分 =「3」
                            If objKtbnStrc.strcSelection.strOpSymbol(8).IndexOf("R") >= 0 Or _
                               objKtbnStrc.strcSelection.strOpSymbol(8).IndexOf("X") >= 0 Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                    End Select
                Case "AMD5"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            'オプションが"6"・"7"の場合はチェック区分「3」
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "6", "7"
                                    fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End Select

                            '取付 = "R"・"X"のとき、チェック区分 =「3」
                            If objKtbnStrc.strcSelection.strOpSymbol(8).IndexOf("R") >= 0 Or _
                               objKtbnStrc.strcSelection.strOpSymbol(8).IndexOf("X") >= 0 Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If

                            '接続が"25UP","25BUP"且つオプションが"2","3","7"の時、チェック区分「2」
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "25UP", "25BUP"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "2", "3", "7"
                                            fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Standard
                                    End Select
                            End Select
                    End Select
                Case "1137"
                    If objKtbnStrc.strcSelection.strOpSymbol(2).IndexOf("F1") >= 0 Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                    End If
                Case "1144"
                    If objKtbnStrc.strcSelection.strOpSymbol(2).IndexOf("F1J") >= 0 Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                    End If
                Case "2619"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).IndexOf("P94") >= 0 Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                    End If

                Case "F1000", "F2000", "F3000", "F4000", "F6000", _
                     "M1000", "M2000", "M3000", "M4000", "M6000", _
                     "R1000", "R1100", "R2000", "R2100", "R3000", "R3100", "R4000", "R4100", "R6000", "R6100", _
                     "W1000", "W1100", "W2000", "W2100", "W3000", "W3100", "W4000", "W4100", _
                     "MX1000", "MX3000", "MX4000", "MX6000", "MX8000"
                    If objKtbnStrc.strcSelection.strOpSymbol(4).IndexOf("P74") >= 0 Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                    End If
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "W"
                            If objKtbnStrc.strcSelection.strOpSymbol(5).IndexOf("P74") >= 0 Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                            Dim strOpArray() As String
                            Dim intLoopCnt As Integer
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4"
                                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Standard
                                    Case "P40"
                                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                                End Select
                            Next
                        Case Else
                    End Select
                Case "V3010"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "W"
                            Dim strOpArray() As String
                            Dim intLoopCnt As Integer
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4"
                                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Standard
                                    Case "P40"
                                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                                End Select
                            Next
                    End Select
                Case "VSP"
                    'フリーホルダで「F2」を選択した場合はチェック区分を「3」に変更する
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "2"
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "F2" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                        Case "1", "B", "E", "K", "L", _
                             "M", "P", "R", "S", "W", "F"
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "F2" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                    End Select
                Case "VSPG"
                    'フリーホルダで「F2」を選択した場合はチェック区分を「3」に変更する
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1", "B", "E", "K", "L", _
                             "P", "R", "S", "W", "F"
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "F2" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                    End Select
                Case "STL-B", "STL-M", "STS-B", "STS-M"
                    'オプションでP52/P53を選択した場合はチェック区分「3」
                    If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("P52") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("P53") >= 0 Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                    End If
                Case "PV5"
                    '旧ISOバルブでその他電圧を選択した場合はチェック区分「3」
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "5", "6"
                            If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-9") <> 0 Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                    End Select
                Case "SSD"
                    'P5,P51を選択した場合はチェック区分「3」
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(19).IndexOf("P5") >= 0 Then
                            fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                        End If
                    End If
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "K" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(17).IndexOf("P5") >= 0 Then
                            fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                        End If
                    End If
                Case "SCA2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "T3PH" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "T3PV" Then
                            fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                        End If
                    End If
                Case "HLD", "HLC"
                    'スイッチを選択している場合はチェック区分「2」
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) > 0 Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Standard
                    End If
                Case "CXU10", "CXU30"
                    Select Case True
                        Case InStr(objKtbnStrc.strcSelection.strFullKataban, "CXU10-GFAB3") <> 0 Or _
                             InStr(objKtbnStrc.strcSelection.strFullKataban, "CXU30-GFAB4U") <> 0
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "2C" Or _
                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "1" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                        Case InStr(objKtbnStrc.strcSelection.strFullKataban, "CXU10-FAB3") <> 0 Or _
                             InStr(objKtbnStrc.strcSelection.strFullKataban, "CXU30-FAB4U") <> 0
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "2C" Or _
                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "1" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                        Case InStr(objKtbnStrc.strcSelection.strFullKataban, "CXU30-FAD") <> 0
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "2C" Or _
                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "1" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                            'RM0911XXX 2009/11/11 Y.Miura 機種追加
                        Case InStr(objKtbnStrc.strcSelection.strFullKataban, "CXU30-ADK") <> 0
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "2C" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "1" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                            'RM0911XXX 2009/11/11 Y.Miura 機種追加
                        Case InStr(objKtbnStrc.strcSelection.strFullKataban, "CXU10-CHV") <> 0
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "A" Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                            'RM1003086 2010/03/30 Y.Miura 機種追加
                        Case InStr(objKtbnStrc.strcSelection.strFullKataban, "CXU10-EXA") <> 0
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3)
                                Case "2C", "2H"
                                    fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End Select
                        Case InStr(objKtbnStrc.strcSelection.strFullKataban, "CXU10-GEXA") <> 0
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5)
                                Case "2C", "2H"
                                    fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End Select
                    End Select
                Case "PDV3"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "2CS", "2ES", "2HS", "3RS"
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = CdCst.PowerSupply.Const4 Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                    End Select
                Case "AB31", "AB41", "AB42", _
                     "AG31", "AG33", "AG34", "AG41", "AG43", "AG44", _
                     "AD11", "AD12", "AP11", "AP12", _
                     "ADK11", "ADK12", "APK11", _
                     "GAG31", "GAG33", "GAG34", "GAG35", "GAG41", "GAG43", "GAG44", "GAG45", _
                     "GAB312", "GAB352", "GAB412", "GAB422", "GAB452", "GAB462"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            If InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "G") <> 0 Or _
                               InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "N") <> 0 Then
                                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                            End If
                    End Select
                Case "FSL"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        fncKatabanCheckDivGet = "2"
                    End If
                Case "FCS1000", "FCS500"        'RM1001045 2010/02/24 Y.Miura 二次電池対応
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim.Equals("1") Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim.Equals("P40") Then
                            fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                        End If
                    End If
                Case "MRL2", "MRL2-G", "MRL2-W" 'RM1306001 2013/06/06
                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "SX" Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                    End If
                Case "EXA"    'Add by Zxjike 2013/10/01
                    If objKtbnStrc.strcSelection.strKeyKataban = "2" Then Exit Select
                    If (objKtbnStrc.strcSelection.strOpSymbol(1) = "C6" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1) = "C8") And _
                        objKtbnStrc.strcSelection.strOpSymbol(2) = "0" And _
                        objKtbnStrc.strcSelection.strOpSymbol(3) = "2C" And _
                        objKtbnStrc.strcSelection.strOpSymbol(5) = "3" Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Stock
                    End If
                Case "FWD11"
                    If (objKtbnStrc.strcSelection.strOpSymbol(1) = "8" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1) = "10" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1) = "15" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1) = "20" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1) = "25") And _
                        objKtbnStrc.strcSelection.strOpSymbol(2) = "A" And _
                        objKtbnStrc.strcSelection.strOpSymbol(3) = "0" And _
                        objKtbnStrc.strcSelection.strOpSymbol(4) = "2C" And _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Stock
                    End If
                Case "NP13"
                    If (objKtbnStrc.strcSelection.strOpSymbol(1) = "10A" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1) = "15A" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1) = "20A" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1) = "25A") And _
                        objKtbnStrc.strcSelection.strOpSymbol(2) = "1" And _
                        objKtbnStrc.strcSelection.strOpSymbol(3) = "2C" And _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" And _
                        (objKtbnStrc.strcSelection.strOpSymbol(5) = "1" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5) = "2" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5) = "3") Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Stock
                    End If
                Case "FSM2"
                    If objKtbnStrc.strcSelection.strOpSymbol(1) = "N" And _
                        objKtbnStrc.strcSelection.strOpSymbol(2) = "V" And _
                        objKtbnStrc.strcSelection.strOpSymbol(3) = "F" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(7) = "1" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                                Case "100"
                                    If objKtbnStrc.strcSelection.strOpSymbol(5) = "H04" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "" Then
                                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Stock
                                    End If
                                Case "500"
                                    If objKtbnStrc.strcSelection.strOpSymbol(5) = "H06" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "" Then
                                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Stock
                                    End If
                                Case "201"
                                    If objKtbnStrc.strcSelection.strOpSymbol(5) = "S08" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "" And _
                                        objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "" Then
                                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Stock
                                    End If
                                Case "102"
                                    If objKtbnStrc.strcSelection.strOpSymbol(5) = "A15" And _
                                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" And _
                                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "" And _
                                    objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "" And _
                                    objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "" And _
                                    objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "" Then
                                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Stock
                                    End If
                            End Select
                        End If
                    End If
                Case "M3QB1", "M3QE1", "M3QZ1"  'RM1706016　2017/7/19　チェック区分追加
                    '連数が9～20の場合はチェック区分「3」
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim >= 9 _
                        And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 20 Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
                    End If

                    'RM1807033_食品製造工程向けの場合チェック区分変更
                Case "FX1004", "FX1011", "FX1037"
                    If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                        fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Standard
                    End If

            End Select

            'シリンダC5チェック(C5の場合はチェック区分「3」)
            If KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc) = True Then
                fncKatabanCheckDivGet = CdCst.KatabanChackDiv.Special
            End If
        Catch ex As Exception
            fncKatabanCheckDivGet = strKatabanCheckDiv
        End Try

    End Function

    ''' <summary>
    ''' 原価積算No.取得
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strKatabanCheckDiv"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncCostCalcNoGet(objKtbnStrc As KHKtbnStrc, ByVal strKatabanCheckDiv As String) As String
        Try
            'デフォルト設定
            fncCostCalcNoGet = ""

            ''引当情報取得
            'objKtbnStrc.subSelKtbnInfoGet(strUserId, strSessionId)

            If Left(objKtbnStrc.strcSelection.strFullKataban, 3) = "JSG" Then
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "T2YDU" Then
                    fncCostCalcNoGet = CdCst.CostCalcNo.C5
                End If
            End If

            If objKtbnStrc.strcSelection.strFullKataban.EndsWith("-FP1") Then
                If objKtbnStrc.strcSelection.strFullKataban.Contains("4GA") Or _
                    objKtbnStrc.strcSelection.strFullKataban.Contains("4GB") Or _
                    objKtbnStrc.strcSelection.strFullKataban.Contains("3GA") Or _
                    objKtbnStrc.strcSelection.strFullKataban.Contains("3GB") Then
                    Exit Try
                End If
            End If

            If KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc) = True Or _
            Left(objKtbnStrc.strcSelection.strFullKataban, 10) = "CAC3-T2YDU" Then
                fncCostCalcNoGet = CdCst.CostCalcNo.C5
                Exit Try
            End If

            'RM14070XX 2014/07/11 SWのC5対応
            If objKtbnStrc.strcSelection.strFullKataban.StartsWith("SW-") And strKatabanCheckDiv = "3" Then
                If objKtbnStrc.strcSelection.strFullKataban.Trim = "SW-T2YDU" Then
                Else
                    fncCostCalcNoGet = CdCst.CostCalcNo.C5
                End If
                Exit Try
            End If

        Catch ex As Exception
            fncCostCalcNoGet = ""
        End Try

    End Function

    ''' <summary>
    ''' ミックス構成チェック
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncVaccumMixCheck(ByVal objKtbnStrc As KHKtbnStrc) As Boolean
        fncVaccumMixCheck = False

        Try
            '各機種毎にミックス構成が選択されているかチェックする
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "VSKM"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "00" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "CX" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "Z" Then
                        fncVaccumMixCheck = True
                    End If
                Case "VSJM"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "00" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "CX" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "Z" Then
                        fncVaccumMixCheck = True
                    End If
                Case "VSNM"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Z" Or _
                         objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "00" Or _
                         objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CX" Or _
                         objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "Z" Then
                        fncVaccumMixCheck = True
                    End If
                Case "VSNPM"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CX" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "Z" Then
                        fncVaccumMixCheck = True
                    End If
                Case "VSXM"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "00" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "CX" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "Z" Then
                        fncVaccumMixCheck = True
                    End If
                Case "VSZM"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "00" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "CX" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "Z" Then
                        fncVaccumMixCheck = True
                    End If
                Case "VSJPM"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CX" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "Z" Then
                        fncVaccumMixCheck = True
                    End If
                Case "VSXPM"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Z" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CX" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "Z" Then
                        fncVaccumMixCheck = True
                    End If
                Case "VSZPM"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CX" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "Z" Then
                        fncVaccumMixCheck = True
                    End If
                Case "B"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                        fncVaccumMixCheck = True
                    End If
                Case "M4SA0", "M4SB0", "M4HA1", "M4HA2", "M4HA3", "M4JA1", "M4JA2", "M4JA3"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        fncVaccumMixCheck = True
                    End If
                    'RM1805001_4Rシリーズ追加
                Case "M4RD1", "M4RD2", "M4RE1", "M4RE2"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        fncVaccumMixCheck = True
                    End If
                Case "M3KA1", "M4KA1", "M4KA2", "M4KA3", "M4KA4", _
                     "M4KB1", "M4KB2", "M4KB3", "M4KB4"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "80" Then
                                fncVaccumMixCheck = True
                            End If
                        Case "M"
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "81" Then
                                fncVaccumMixCheck = True
                            End If
                    End Select
                Case "M4F0", "M4F1", "M4F2", "M4F3", "M4F4", "M4F5", "M4F6", "M4F7"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        fncVaccumMixCheck = True
                    End If
                Case "M3MA0", "M3MB0", "M3PA1", "M3PA2", "M3PB1", "M3PB2", "M4L2", "M4LB2"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        fncVaccumMixCheck = True
                    End If
                Case "M3QRA1", "M3QRB1", "M3QB1", "M3QE1", "M3QZ1"
                    fncVaccumMixCheck = True
                Case "MV3QRA1", "MV3QRB1"
                    fncVaccumMixCheck = True
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' オプションチェック処理
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <param name="strOptionSymbol">オプション記号</param>
    ''' <param name="strMessageCd">メッセージコード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncOtherOptionCheck(objKtbnStrc As KHKtbnStrc, _
                                        ByRef intKtbnStrcSeqNo As Integer, ByRef strOptionSymbol As String, _
                                        ByRef strMessageCd As String) As Boolean
        Try

            fncOtherOptionCheck = True

            '空圧バルブチェック
            If fncOtherOptionCheck = True Then
                If KHAirValveCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            '流体制御バルブチェック
            If fncOtherOptionCheck = True Then
                If KHWaterValveCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'シリンダチェック
            If fncOtherOptionCheck = True Then
                If KHCylinderCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'シリンダストロークチェック
            If fncOtherOptionCheck = True Then
                If KHCylinderStrokeCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'シリンダチェック(SSD)
            If fncOtherOptionCheck = True Then
                If KHCylinderSSDCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'シリンダチェック(SSD2)
            If fncOtherOptionCheck = True Then
                If KHCylinderSSD2Check.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'シリンダチェック(CMK2)
            If fncOtherOptionCheck = True Then
                If KHCylinderCMK2Check.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'シリンダチェック(SCM)
            If fncOtherOptionCheck = True Then
                If KHCylinderSCMCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'シリンダチェック(SCA2)
            If fncOtherOptionCheck = True Then
                If KHCylinderSCA2Check.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'シリンダチェック(STS)
            If fncOtherOptionCheck = True Then
                If KHCylinderSTSCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'シリンダチェック(JSC3)
            If fncOtherOptionCheck = True Then
                If KHCylinderJSC3Check.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'RM0906034 2009/08/28 Y.Miura　
            'シリンダチェック(FRL)
            If fncOtherOptionCheck = True Then
                If KHCylinderFRLCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'ニューハンドリングシステム＆ハイブリロボチェック
            If fncOtherOptionCheck = True Then
                If KHNewHandleCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'ガス燃焼システムチェック
            If fncOtherOptionCheck = True Then
                If KHGasCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

            'その他チェック
            If fncOtherOptionCheck = True Then
                If KHOtherCheck.fncCheckSelectOption(objKtbnStrc, intKtbnStrcSeqNo, strOptionSymbol, strMessageCd) = False Then
                    fncOtherOptionCheck = False
                End If
            End If

        Catch ex As Exception
            fncOtherOptionCheck = False
        End Try

    End Function

End Class
