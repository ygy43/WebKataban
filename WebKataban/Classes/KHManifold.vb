Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class KHManifold

    Private strUserID As String
    Private strSessionID As String

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="strAUserID"></param>
    ''' <param name="strASessionID"></param>
    ''' <remarks>ユーザーIDとセッションIDを保持</remarks>
    Public Sub New(ByVal strAUserID As String, ByVal strASessionID As String)
        strUserID = strAUserID
        strSessionID = strASessionID
    End Sub

    ''' <summary>
    ''' 引当情報更新
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="decRailLength">取付レール長さ</param>
    ''' <param name="strRailChange">自動更新フラグ</param>
    ''' <remarks></remarks>
    Public Overloads Sub subUpdateSelSpec(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                          ByVal decRailLength As Decimal, Optional ByVal strRailChange As String = "0")
        Dim strKigouDummy(0) As String
        Try
            strKigouDummy(0) = ""
            subUpdateSelSpec(objCon, objKtbnStrc, decRailLength, strKigouDummy, strRailChange)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当情報更新
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="decRailLength">取付レール長さ</param>
    ''' <param name="strKigou">属性記号</param>
    ''' <param name="strRailChange">自動更新フラグ</param>
    ''' <remarks></remarks>
    Public Overloads Sub subUpdateSelSpec(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                          ByVal decRailLength As Decimal,
                                          ByVal strKigou As String(), Optional ByVal strRailChange As String = "0")
        Dim intReturn As Integer
        Try
            '引当仕様書クリア
            intReturn = fncSPSelSpecDel(objCon)
            '引当仕様書構成クリア
            intReturn = fncSPSpecStrcDel(objCon)
            '引当仕様書更新
            intReturn = fncSPSelSpecIns(objCon, objKtbnStrc, decRailLength, strRailChange)
            '引当仕様書構成更新
            intReturn = fncSPSpecStrcIns(objCon, objKtbnStrc, strKigou)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当情報クリア
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteSelSpec(ByVal objCon As SqlConnection)
        Dim intReturn As Integer
        Try
            '引当仕様書クリア
            intReturn = fncSPSelSpecDel(objCon)
            '引当仕様書構成クリア
            intReturn = fncSPSpecStrcDel(objCon)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当情報検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectSelSpec(ByVal objCon As SqlConnection) As DataTable
        Dim objCmd As SqlCommand
        Dim objAdp As SqlDataAdapter
        Dim dtResult As New DataTable
        Dim sbSql As New Text.StringBuilder
        fncSelectSelSpec = Nothing

        Try
            objCmd = objCon.CreateCommand
            sbSql.Append("  SELECT  0   AS " & CdCst.SelSpec.SeqNo & ", ")
            sbSql.Append("          CASE rail_change ")
            sbSql.Append("            WHEN '0'   THEN ")
            sbSql.Append("              '0' ")
            sbSql.Append("            ELSE ")
            sbSql.Append("              CONVERT(varchar,din_rail_length) ")
            sbSql.Append("          END AS " & CdCst.SelSpec.Kataban & ", ")
            sbSql.Append("          ''  AS " & CdCst.SelSpec.CxA & ", ")
            sbSql.Append("          ''  AS " & CdCst.SelSpec.CxB & ", ")
            sbSql.Append("          ''  AS " & CdCst.SelSpec.PosInfo & ", ")
            sbSql.Append("          0   AS " & CdCst.SelSpec.Qty & " ")
            sbSql.Append("  FROM    kh_sel_spec ")
            sbSql.Append("  WHERE   user_id    = @UserId ")
            sbSql.Append("  AND     session_id = @SessionId ")
            sbSql.Append("UNION ")
            sbSql.Append("  SELECT  strc.spec_strc_seq_no AS " & CdCst.SelSpec.SeqNo & ", ")
            sbSql.Append("          strc.option_kataban   AS " & CdCst.SelSpec.Kataban & ", ")
            sbSql.Append("          strc.cxa_kataban      AS " & CdCst.SelSpec.CxA & ", ")
            sbSql.Append("          strc.cxb_kataban      AS " & CdCst.SelSpec.CxB & ", ")
            sbSql.Append("          strc.position_info    AS " & CdCst.SelSpec.PosInfo & ", ")
            sbSql.Append("          strc.quantity         AS " & CdCst.SelSpec.Qty & " ")
            sbSql.Append("  FROM    kh_sel_spec       spec, ")
            sbSql.Append("          kh_sel_spec_strc  strc ")
            sbSql.Append("  WHERE   spec.user_id    = @UserId ")
            sbSql.Append("  AND     spec.session_id = @SessionId ")
            sbSql.Append("  AND     spec.user_id    = strc.user_id ")
            sbSql.Append("  AND     spec.session_id = strc.session_id ")
            sbSql.Append("ORDER BY 1 ")
            objCmd.CommandText = sbSql.ToString
            objCmd.Parameters.Add("@UserId", SqlDbType.VarChar, 10)
            objCmd.Parameters.Add("@SessionId", SqlDbType.NVarChar, 88)
            objCmd.Parameters("@UserId").Value = strUserID
            objCmd.Parameters("@SessionId").Value = strSessionID
            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dtResult)

            fncSelectSelSpec = dtResult
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 引当仕様書テーブルにデータを追加(SP)
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="decRailLength">取付レール長さ</param>
    ''' <param name="strRailChange">自動更新フラグ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSPSelSpecIns(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                     ByVal decRailLength As Decimal, ByVal strRailChange As String) As Integer
        Dim objCmd As SqlCommand
        fncSPSelSpecIns = 0

        Try
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelSpecIns, objCon)
            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionID
                .Parameters.Add("@ModelNo", SqlDbType.Char, 2).Value = Me.fncModelNoGet(objKtbnStrc)
                .Parameters.Add("@WiringSpec", SqlDbType.Char, 1).Value = Me.fncWiringSpecGet(objCon)
                .Parameters.Add("@DinRailLen", SqlDbType.Decimal, 6, 1).Value = decRailLength
                .Parameters.Add("@RailChange", SqlDbType.Char, 1).Value = strRailChange
                .Parameters.Add("@RegPerson", SqlDbType.VarChar, 10).Value = strUserID
                .Parameters.Add("@RegDate", SqlDbType.DateTime).Value = Now()
                .Parameters.Add("@CurPerson", SqlDbType.VarChar, 10).Value = DBNull.Value
                .Parameters.Add("@CurDate", SqlDbType.DateTime).Value = DBNull.Value
            End With
            '実行
            fncSPSelSpecIns = objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 引当仕様書構成テーブルにデータを追加(SP)
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strKigou"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSPSpecStrcIns(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                      ByVal strKigou As String()) As Integer
        Dim objCmd As SqlCommand
        fncSPSpecStrcIns = 0

        Try
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelSpecStrcIns, objCon)
            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10)
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88)
                .Parameters.Add("@SpcStrcSeq", SqlDbType.Int)
                .Parameters.Add("@AttribSym", SqlDbType.Char, 2)
                .Parameters.Add("@OpKataban", SqlDbType.VarChar, 30)
                .Parameters.Add("@CxaKataban", SqlDbType.VarChar, 30)
                .Parameters.Add("@CxbKataban", SqlDbType.VarChar, 30)
                .Parameters.Add("@Position", SqlDbType.VarChar, 50)
                .Parameters.Add("@Quantity", SqlDbType.Int)
                .Parameters.Add("@RegPerson", SqlDbType.VarChar, 10)
                .Parameters.Add("@RegDate", SqlDbType.DateTime)
                .Parameters.Add("@CurPerson", SqlDbType.VarChar, 10)
                .Parameters.Add("@CurDate", SqlDbType.DateTime)
                .Parameters("@UserId").Value = strUserID
                .Parameters("@SessionId").Value = strSessionID
                .Parameters("@RegPerson").Value = strUserID
                .Parameters("@RegDate").Value = Now()
                .Parameters("@CurPerson").Value = DBNull.Value
                .Parameters("@CurDate").Value = DBNull.Value

                For intI As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                    .Parameters("@SpcStrcSeq").Value = intI + 1

                    Dim strAttrib As String = Me.fncAttributeSymbolGet(objCon, objKtbnStrc.strcSelection.strOptionKataban(intI), intI + 1)
                    .Parameters("@AttribSym").Value = strAttrib

                    Select Case strAttrib
                        Case "T8", "TE"  'ｹﾝｻｾｲｾｷｼﾖ(ﾜﾌﾞﾝ)
                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intI)
                                Case CdCst.Manifold.InspReportJp.DummyValue, CdCst.Manifold.InspReportJp.Japanese
                                    .Parameters("@OpKataban").Value = CdCst.Manifold.InspReportJp.SelectValue
                                Case CdCst.Manifold.InspReportEn.DummyValue, CdCst.Manifold.InspReportEn.Japanese
                                    .Parameters("@OpKataban").Value = CdCst.Manifold.InspReportEn.SelectValue
                                Case Else
                                    'ADD BY YGY 20141126
                                    .Parameters("@OpKataban").Value = objKtbnStrc.strcSelection.strOptionKataban(intI)
                            End Select
                        Case Else
                            .Parameters("@OpKataban").Value = objKtbnStrc.strcSelection.strOptionKataban(intI)
                    End Select

                    If intI > objKtbnStrc.strcSelection.strPositionInfo.Count - 1 Then
                        .Parameters("@Position").Value = DBNull.Value
                    Else
                        If objKtbnStrc.strcSelection.strPositionInfo(intI) IsNot Nothing Then
                            .Parameters("@Position").Value = objKtbnStrc.strcSelection.strPositionInfo(intI)
                        Else
                            .Parameters("@Position").Value = DBNull.Value
                        End If
                    End If
                    If intI > objKtbnStrc.strcSelection.intQuantity.Length - 1 Then
                        .Parameters("@Quantity").Value = 0
                    Else
                        .Parameters("@Quantity").Value = CInt(objKtbnStrc.strcSelection.intQuantity(intI))
                    End If

                    If intI > objKtbnStrc.strcSelection.strCXAKataban.Length - 1 Then
                        .Parameters("@CxaKataban").Value = ""
                    Else
                        If objKtbnStrc.strcSelection.strCXAKataban(intI) Is Nothing Then
                            .Parameters("@CxaKataban").Value = ""
                        Else
                            .Parameters("@CxaKataban").Value = objKtbnStrc.strcSelection.strCXAKataban(intI)
                        End If
                    End If
                    If intI > objKtbnStrc.strcSelection.strCXBKataban.Length - 1 Then
                        .Parameters("@CxbKataban").Value = ""
                    Else
                        If objKtbnStrc.strcSelection.strCXBKataban(intI) Is Nothing Then
                            .Parameters("@CxbKataban").Value = ""
                        Else
                            .Parameters("@CxbKataban").Value = objKtbnStrc.strcSelection.strCXBKataban(intI)
                        End If
                    End If

                    '実行
                    fncSPSpecStrcIns = objCmd.ExecuteNonQuery()
                Next
            End With

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
        End Try

    End Function

    ''' <summary>
    ''' 引当仕様書テーブルからデータを削除(SP)
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSPSelSpecDel(ByVal objCon As SqlConnection) As Integer
        Dim objCmd As SqlCommand
        fncSPSelSpecDel = 0

        Try
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelSpecDel, objCon)
            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionID
            End With
            '実行
            fncSPSelSpecDel = objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 引当仕様書構成テーブルからデータを削除(SP)
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSPSpecStrcDel(ByVal objCon As SqlConnection) As Integer
        Dim objCmd As SqlCommand
        fncSPSpecStrcDel = 0

        Try
            objCmd = New SqlCommand(CdCst.DB.SPL.KHSelSpecStrcDel, objCon)
            With objCmd
                .CommandType = CommandType.StoredProcedure
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = strUserID
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = strSessionID
            End With
            '実行
            fncSPSpecStrcDel = objCmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objCmd = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 機種番号取得
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncModelNoGet(objKtbnStrc As KHKtbnStrc) As String
        fncModelNoGet = String.Empty
        Try
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MN3E0"
                    fncModelNoGet = "81"
                Case "MN4E0"
                    fncModelNoGet = "82"
                Case "MN3E00"
                    fncModelNoGet = "83"
                Case "MN4E00"
                    fncModelNoGet = "84"
                Case "MN3EX0"
                    fncModelNoGet = "85"
                Case "MN4EX0"
                    fncModelNoGet = "86"
                Case "MN4KB1"
                    fncModelNoGet = "21"
                Case "MN4KB2"
                    fncModelNoGet = "22"
                Case "M"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1", "2"
                            'バルブ種類判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "3"
                                    'マニホールド形式判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                        Case "D"
                                            fncModelNoGet = "A4"
                                        Case Else
                                            fncModelNoGet = "A1"
                                    End Select
                                Case Else
                                    'マニホールド形式判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                        Case "D"
                                            fncModelNoGet = "A5"
                                        Case Else
                                            fncModelNoGet = "A2"
                                    End Select
                            End Select
                        Case Else
                            'マニホールド形式判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "D"
                                    fncModelNoGet = "A6"
                                Case Else
                                    fncModelNoGet = "A3"
                            End Select
                    End Select
                Case "CMF"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1", "2", "3", "8"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "1"
                                    fncModelNoGet = "B1"
                                Case "2"
                                    fncModelNoGet = "B2"
                                Case Else
                                    fncModelNoGet = "B3"
                            End Select
                        Case Else
                            fncModelNoGet = "B4"
                    End Select
                Case "GMF"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "1"
                            fncModelNoGet = "B5"
                        Case "2"
                            fncModelNoGet = "B6"
                        Case "Z"
                            fncModelNoGet = "B7"
                    End Select
                Case "M3GA1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "M1"
                        Case Else
                            fncModelNoGet = "11"
                    End Select
                Case "M3GA2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "M2"
                        Case Else
                            fncModelNoGet = "12"
                    End Select
                Case "M3GA3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "M3"
                        Case Else
                            fncModelNoGet = "13"
                    End Select
                Case "M4GA1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "M4"
                        Case Else
                            fncModelNoGet = "14"
                    End Select
                Case "M4GA2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "M5"
                        Case Else
                            fncModelNoGet = "15"
                    End Select
                Case "M4GA3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "M6"
                        Case Else
                            fncModelNoGet = "16"
                    End Select
                Case "M3GB1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "M7"
                        Case Else
                            fncModelNoGet = "1A"
                    End Select
                Case "M3GB2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "M8"
                        Case Else
                            fncModelNoGet = "1B"
                    End Select
                Case "M4GB1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "M9"
                        Case Else
                            fncModelNoGet = "17"
                    End Select
                Case "M4GB2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "MA"
                        Case Else
                            fncModelNoGet = "18"
                    End Select
                Case "M4GB3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U", "S", "V"
                            fncModelNoGet = "MB"
                        Case Else
                            fncModelNoGet = "19"
                    End Select
                Case "M4GA4"
                    fncModelNoGet = "1C"
                Case "M4GB4"
                    fncModelNoGet = "1D"
                Case "M3GD1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "MC"
                        Case Else
                            fncModelNoGet = "1E"
                    End Select
                Case "M3GD2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "MD"
                        Case Else
                            fncModelNoGet = "1F"
                    End Select
                Case "M3GD3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "ME"
                        Case Else
                            fncModelNoGet = "1G"
                    End Select
                Case "M4GD1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "MF"
                        Case Else
                            fncModelNoGet = "1H"
                    End Select
                Case "M4GD2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "MG"
                        Case Else
                            fncModelNoGet = "1I"
                    End Select
                Case "M4GD3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "MH"
                        Case Else
                            fncModelNoGet = "1J"
                    End Select
                Case "M3GE1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "MI"
                        Case Else
                            fncModelNoGet = "1K"
                    End Select
                Case "M3GE2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "MJ"
                        Case Else
                            fncModelNoGet = "1L"
                    End Select
                Case "M3GE3"
                    fncModelNoGet = "1M"
                Case "M4GE1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "MK"
                        Case Else
                            fncModelNoGet = "1N"
                    End Select
                Case "M4GE2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "ML"
                        Case Else
                            fncModelNoGet = "1O"
                    End Select
                Case "M4GE3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "MM"
                        Case Else
                            fncModelNoGet = "1P"
                    End Select
                Case "MN3GA1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "N1"
                        Case Else
                            fncModelNoGet = "31"
                    End Select
                Case "MN3GA2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "N2"
                        Case Else
                            fncModelNoGet = "32"
                    End Select
                Case "MN4GA1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "N3"
                        Case Else
                            fncModelNoGet = "33"
                    End Select
                Case "MN4GA2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "N4"
                        Case Else
                            fncModelNoGet = "34"
                    End Select
                Case "MN3GB1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "N7"
                        Case Else
                            fncModelNoGet = "3A"
                    End Select
                Case "MN3GB2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "N8"
                        Case Else
                            fncModelNoGet = "3B"
                    End Select
                Case "MN4GB1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "N9"
                        Case Else
                            fncModelNoGet = "35"
                    End Select
                Case "MN4GB2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NA"
                        Case Else
                            fncModelNoGet = "36"
                    End Select
                Case "MN3GAX12"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "N5"
                        Case Else
                            fncModelNoGet = "37"
                    End Select
                Case "MN4GAX12"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "N6"
                        Case Else
                            fncModelNoGet = "38"
                    End Select
                Case "MN3GBX12"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NB"
                        Case Else
                            fncModelNoGet = "3C"
                    End Select
                Case "MN4GBX12"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NC"
                        Case Else
                            fncModelNoGet = "39"
                    End Select
                Case "MN3GD1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "ND"
                        Case Else
                            fncModelNoGet = "3D"
                    End Select
                Case "MN3GD2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NE"
                        Case Else
                            fncModelNoGet = "3E"
                    End Select
                Case "MN4GD1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NF"
                        Case Else
                            fncModelNoGet = "3F"
                    End Select
                Case "MN4GD2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NG"
                        Case Else
                            fncModelNoGet = "3G"
                    End Select
                Case "MN4GD3"
                    fncModelNoGet = "3H"
                Case "MN3GE1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NJ"
                        Case Else
                            fncModelNoGet = "3I"
                    End Select
                Case "MN3GE2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NK"
                        Case Else
                            fncModelNoGet = "3J"
                    End Select
                Case "MN4GE1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NL"
                        Case Else
                            fncModelNoGet = "3K"
                    End Select
                Case "MN4GE2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NM"
                        Case Else
                            fncModelNoGet = "3L"
                    End Select
                Case "MN4GE3"
                    fncModelNoGet = "3M"
                Case "N"
                    fncModelNoGet = "41"
                Case "M4TB3"
                    fncModelNoGet = "04"
                Case "M4TB4"
                    fncModelNoGet = "05"
                Case "MN3S0"
                    fncModelNoGet = "51"
                Case "MN4S0"
                    fncModelNoGet = "52"
                Case "MT3S0"
                    fncModelNoGet = "53"
                Case "MT4S0"
                    fncModelNoGet = "54"
                Case "LMF0"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "T0U" Then
                                fncModelNoGet = "C1"
                            Else
                                fncModelNoGet = "C2"
                            End If
                        Case Else
                            fncModelNoGet = "C3"
                    End Select
                Case "MNRB500A"
                    fncModelNoGet = "91"
                Case "MNRB500B"
                    fncModelNoGet = "92"
                Case "MNRJB500A"
                    fncModelNoGet = "93"
                Case "MNRJB500B"
                    fncModelNoGet = "94"
                Case "VSKM"
                    fncModelNoGet = "D1"
                Case "VSJM"
                    fncModelNoGet = "D2"
                Case "VSXM"
                    fncModelNoGet = "D3"
                Case "VSZM"
                    fncModelNoGet = "D4"
                Case "VSJPM"
                    fncModelNoGet = "D5"
                Case "VSXPM"
                    fncModelNoGet = "D6"
                Case "VSZPM"
                    fncModelNoGet = "D7"
                Case "MN4TB1"
                    fncModelNoGet = "01"
                Case "MN4TB2"
                    fncModelNoGet = "02"
                Case "MN4TBX12"
                    fncModelNoGet = "03"
                Case "MEVT"
                    fncModelNoGet = "71"
                Case "MW4GB4"
                    fncModelNoGet = "65"
                Case "MW4GZ4"
                    fncModelNoGet = "66"
                Case "MW3GA2"
                    fncModelNoGet = "61"
                Case "MW4GA2"
                    fncModelNoGet = "62"
                Case "MW4GB2"
                    fncModelNoGet = "63"
                Case "MW4GZ2"
                    fncModelNoGet = "64"
                Case "MW3GB2"
                    fncModelNoGet = "67"
                Case "MW3GZ2"
                    fncModelNoGet = "68"
                Case "GAMD0"
                    fncModelNoGet = "E1"
                Case "M3QRA1"
                    fncModelNoGet = "F1"
                Case "M3QRB1"
                    fncModelNoGet = "F2"
                Case "MV3QRA1"
                    fncModelNoGet = "F3"
                Case "MV3QRB1"
                    fncModelNoGet = "F4"
                Case "M3QB1"
                    fncModelNoGet = "F5"
                Case "M3QE1"
                    fncModelNoGet = "F6"
                Case "M3QZ1"
                    fncModelNoGet = "F7"
                Case "MN3Q0"
                    fncModelNoGet = "G1"
                Case "MT3Q0"
                    fncModelNoGet = "G2"
                Case "MN4GDX12" 'RM1303003 2013/03/08 
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NI"
                        Case Else
                            fncModelNoGet = "3N"
                    End Select
                Case "MN4GEX12" 'RM1303003 2013/03/08 
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "U"
                            fncModelNoGet = "NO"
                        Case Else
                            fncModelNoGet = "3O"
                    End Select
                Case "B"
                    fncModelNoGet = "42"
                Case "M4SA0"
                    fncModelNoGet = "A7"
                Case "M4SB0"
                    fncModelNoGet = "A8"
                Case "M3KA1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            fncModelNoGet = "H1"
                        Case "M"
                            fncModelNoGet = "HA"
                    End Select
                Case "M4KA1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            fncModelNoGet = "H2"
                        Case "M"
                            fncModelNoGet = "HB"
                    End Select
                Case "M4KA2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            fncModelNoGet = "H3"
                        Case "M"
                            fncModelNoGet = "HC"
                    End Select
                Case "M4KA3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            fncModelNoGet = "H4"
                        Case "M"
                            fncModelNoGet = "HD"
                    End Select
                Case "M4KA4"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            fncModelNoGet = "H5"
                        Case "M"
                            fncModelNoGet = "HE"
                    End Select
                Case "M4KB1"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            fncModelNoGet = "H6"
                        Case "M"
                            fncModelNoGet = "HF"
                    End Select
                Case "M4KB2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            fncModelNoGet = "H7"
                        Case "M"
                            fncModelNoGet = "HG"
                    End Select
                Case "M4KB3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            fncModelNoGet = "H8"
                        Case "M"
                            fncModelNoGet = "HH"
                    End Select
                Case "M4KB4"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            fncModelNoGet = "H9"
                        Case "M"
                            fncModelNoGet = "HI"
                    End Select
                Case "M4F0"
                    fncModelNoGet = "I1"
                Case "M4F1"
                    fncModelNoGet = "I2"
                Case "M4F2"
                    fncModelNoGet = "I3"
                Case "M4F3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "", "M"
                            fncModelNoGet = "I4"
                        Case "E"
                            fncModelNoGet = "I9"
                        Case "X"
                            fncModelNoGet = "IE"
                    End Select
                Case "M4F4"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "", "M"
                            fncModelNoGet = "I5"
                        Case "E"
                            fncModelNoGet = "IA"
                        Case "X"
                            fncModelNoGet = "IF"
                    End Select
                Case "M4F5"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "", "M"
                            fncModelNoGet = "I6"
                        Case "E"
                            fncModelNoGet = "IB"
                        Case "X"
                            fncModelNoGet = "IG"
                    End Select
                Case "M4F6"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "", "M"
                            fncModelNoGet = "I7"
                        Case "E"
                            fncModelNoGet = "IC"
                        Case "X"
                            fncModelNoGet = "IH"
                    End Select
                Case "M4F7"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "", "M"
                            fncModelNoGet = "I8"
                        Case "E"
                            fncModelNoGet = "ID"
                        Case "X"
                            fncModelNoGet = "II"
                    End Select
                Case "M3MA0"
                    fncModelNoGet = "J1"
                Case "M3MB0"
                    fncModelNoGet = "J2"
                Case "M3PA1"
                    fncModelNoGet = "K1"
                Case "M3PA2"
                    fncModelNoGet = "K2"
                Case "M3PB1"
                    fncModelNoGet = "K3"
                Case "M3PB2"
                    fncModelNoGet = "K4"
                Case "M4L2"
                    fncModelNoGet = "L1"
                Case "M4LB2"
                    fncModelNoGet = "L2"
                Case Else
                    fncModelNoGet = ""
            End Select

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objKtbnStrc = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 配線仕様有無区分取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncWiringSpecGet(objCon As SqlConnection) As String
        Dim objKtbnStrc As New KHKtbnStrc
        Try
            '引当情報取得
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.strUserID, Me.strSessionID)
            Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                Case "", "00", "52", "90"
                    fncWiringSpecGet = ""
                Case Else
                    fncWiringSpecGet = "2"
            End Select
        Catch ex As Exception
            fncWiringSpecGet = ""
        Finally
            objKtbnStrc = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 属性記号設定
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strOptionKataban"></param>
    ''' <param name="intSpecStrcSeqNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncAttributeSymbolGet(objCon As SqlConnection, ByVal strOptionKataban As String, _
                                          ByVal intSpecStrcSeqNo As Integer) As String
        Dim objKtbnStrc As New KHKtbnStrc
        Try
            fncAttributeSymbolGet = ""

            '引当情報取得
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.strUserID, Me.strSessionID)

            Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                Case "01"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 2
                            '電送ブロック
                            fncAttributeSymbolGet = "B1"
                        Case 3 To 10
                            'バルブブロック
                            fncAttributeSymbolGet = "B3"
                        Case 11 To 12
                            'ダミーブロック
                            fncAttributeSymbolGet = "BJ"
                            'Case 12 To 15
                        Case 13 To 16
                            '給排気ブロック
                            fncAttributeSymbolGet = "B5"
                            'Case 16 To 17
                        Case 17 To 18
                            'レギュレータブロック
                            fncAttributeSymbolGet = "BF"
                            'Case 18 To 19
                        Case 19 To 20
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                            'Case 20 To 23
                        Case 21 To 24
                            '付属品
                            Select Case True
                                Case Left(strOptionKataban, 3) = "SLW"
                                    'サイレンサ
                                    fncAttributeSymbolGet = "T1"
                                Case Left(strOptionKataban, 3) = "GWP" Or _
                                     Left(strOptionKataban, 3) = "PG-" Or _
                                     Left(strOptionKataban, 3) = "N4E"
                                    'ブランクプラグ
                                    fncAttributeSymbolGet = "T2"
                                Case Else
                                    'ダミーセット
                                    fncAttributeSymbolGet = "T1"
                            End Select
                            'Case 24 To 27
                        Case 25 To 28
                            '付属品
                            Select Case True
                                'Case InStr(1, strOptionKataban, "検査成績書（和文）") <> 0
                                '    '検査成績書(和文)
                                '    fncAttributeSymbolGet = "T8"
                                'Case InStr(1, strOptionKataban, "検査成績書（英文）") <> 0
                                '    '検査成績書(英文)
                                '    fncAttributeSymbolGet = "TE"
                                Case InStr(1, strOptionKataban, CdCst.Manifold.InspReportJp.SelectValue) <> 0
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case InStr(1, strOptionKataban, CdCst.Manifold.InspReportEn.SelectValue) <> 0
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                                Case InStr(1, strOptionKataban, "CABLE") <> 0
                                    'ケーブル
                                    fncAttributeSymbolGet = "T9"
                                Case InStr(1, strOptionKataban, "PTN2") <> 0
                                    '継手
                                    fncAttributeSymbolGet = "TA"
                                Case InStr(1, strOptionKataban, "CONNECTOR") <> 0
                                    'コネクタ
                                    fncAttributeSymbolGet = "TB"
                                Case InStr(1, strOptionKataban, "SOCKET") <> 0
                                    'ソケット
                                    fncAttributeSymbolGet = "TC"
                                    'Case Else
                                    '    'ダミーセット
                                    '    fncAttributeSymbolGet = "T8"
                            End Select
                            'Case 28
                        Case 29
                            'チューブ抜具不要
                            fncAttributeSymbolGet = "TD"
                            'Case 29
                        Case 30
                            'DINレール長さ
                            fncAttributeSymbolGet = "L1"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "02"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 2
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 3 To 4
                            '給排気ブロック
                            fncAttributeSymbolGet = "B5"
                        Case 5 To 6
                            '給気ブロック
                            fncAttributeSymbolGet = "B6"
                        Case 7 To 8
                            '排気ブロック
                            fncAttributeSymbolGet = "B7"
                        Case 9 To 14
                            'バルブブロック
                            fncAttributeSymbolGet = "B3"
                        Case 15 To 16
                            '仕切ブロック
                            fncAttributeSymbolGet = "BA"
                        Case 17 To 18
                            'サイレンサ
                            fncAttributeSymbolGet = "T1"
                        Case 19 To 20
                            'ブランクプラグ
                            fncAttributeSymbolGet = "T2"
                        Case 21
                            'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.SelectValue Then
                            '    'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.Japanese Then
                            '    '検査成績書(和文)
                            '    fncAttributeSymbolGet = "T8"
                            'Else
                            '    '検査成績書(英文)
                            '    fncAttributeSymbolGet = "TE"
                            'End If

                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                                Case Else
                                    fncAttributeSymbolGet = ""
                            End Select
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "03"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 14
                            '電磁弁
                            If Mid(strOptionKataban.Trim, 5, 1) = "1" Then
                                fncAttributeSymbolGet = "D1"
                            Else
                                fncAttributeSymbolGet = "D2"
                            End If
                        Case 15
                            'マスキングプレート
                            fncAttributeSymbolGet = "D3"
                        Case 16 To 19
                            '給気ブロック
                            Select Case Left(strOptionKataban, 3)
                                Case "GWP"
                                    'ブランクプラグ
                                    fncAttributeSymbolGet = "T2"
                                Case "SLW"
                                    'サイレンサ
                                    fncAttributeSymbolGet = "T1"
                                Case "GWS"
                                    'ワンタッチ継手
                                    fncAttributeSymbolGet = "TA"
                                Case Else
                                    fncAttributeSymbolGet = "T1"
                            End Select
                        Case 20
                            'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.SelectValue Then
                            '    'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.Japanese Then
                            '    '検査成績書(和文)
                            '    fncAttributeSymbolGet = "T8"
                            'Else
                            '    '検査成績書(英文)
                            '    fncAttributeSymbolGet = "TE"
                            'End If
                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                                Case Else
                                    fncAttributeSymbolGet = ""
                            End Select

                        Case 21 To 22
                            'ケーブル
                            fncAttributeSymbolGet = "T9"
                        Case 23
                            'チューブ抜具不要
                            fncAttributeSymbolGet = "TD"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "04"   'RM1803032_スペーサ行数追加対応
                    Select Case intSpecStrcSeqNo
                        Case CdCst.Siyou_04.Valve1 To CdCst.Siyou_04.Valve10
                            '電磁弁
                            If Mid(strOptionKataban.Trim, 5, 1) = "1" Then
                                fncAttributeSymbolGet = "D1"
                            Else
                                fncAttributeSymbolGet = "D2"
                            End If
                        Case CdCst.Siyou_04.MasPlate1 To CdCst.Siyou_04.MasPlate2
                            'マスキングプレート
                            If InStr(1, strOptionKataban, "-MPD") <> 0 Then
                                fncAttributeSymbolGet = "D4"
                            Else
                                fncAttributeSymbolGet = "D3"
                            End If
                        Case CdCst.Siyou_04.Spacer1 To CdCst.Siyou_04.Spacer4
                            'スペーサ
                            If strOptionKataban.ToString.Contains("R-PC-M") Then
                                fncAttributeSymbolGet = "GB"
                            ElseIf strOptionKataban.ToString.Contains("R-IS") Then
                                fncAttributeSymbolGet = "S7"
                            ElseIf strOptionKataban.ToString.Contains("R-R") Then
                                fncAttributeSymbolGet = "S3"
                            Else
                                fncAttributeSymbolGet = "S2"
                            End If
                        Case CdCst.Siyou_04.BlkPlug1 To CdCst.Siyou_04.BlkPlug2
                            'ブランクプラグ
                            fncAttributeSymbolGet = "T2"
                        Case CdCst.Siyou_04.Silencer1 To CdCst.Siyou_04.Silencer2
                            'サイレンサ
                            fncAttributeSymbolGet = "T1"
                        Case CdCst.Siyou_04.ScrPlug
                            'ねじプラグ
                            fncAttributeSymbolGet = "T5"
                        Case 22
                            'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.SelectValue Then
                            '    'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.Japanese Then
                            '    '検査成績書(和文)
                            '    fncAttributeSymbolGet = "T8"
                            'Else
                            '    '検査成績書(英文)
                            '    fncAttributeSymbolGet = "TE"
                            'End If

                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                                Case Else
                                    fncAttributeSymbolGet = ""
                            End Select
                        Case 23 To 24
                            'ケーブル
                            fncAttributeSymbolGet = "T9"
                        Case 25
                            'チューブ抜具不要
                            fncAttributeSymbolGet = "TD"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "05"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            'ベース
                            fncAttributeSymbolGet = "G1"
                        Case 2 To 7
                            '電磁弁形式
                            fncAttributeSymbolGet = "G2"
                        Case 8
                            'A・Bポートプラグ位置
                            fncAttributeSymbolGet = "G3"
                        Case 9
                            'A・Bポートプラグ位置
                            fncAttributeSymbolGet = "G5"
                        Case 10
                            'A・Bポート接続口径(02)
                            fncAttributeSymbolGet = "G6"
                        Case 11
                            'A・Bポート接続口径(03)
                            fncAttributeSymbolGet = "G7"
                        Case 12
                            'A・Bポート接続口径(04)
                            fncAttributeSymbolGet = "G8"
                        Case 13 To 14
                            '給気スペーサ
                            fncAttributeSymbolGet = "G9"
                        Case 15 To 16
                            '排気スペーサ
                            fncAttributeSymbolGet = "GA"
                        Case 17 To 18
                            'パイロットチェック弁
                            fncAttributeSymbolGet = "GB"
                        Case 19 To 22
                            'スペーサー形減圧弁
                            fncAttributeSymbolGet = "GC"
                        Case 23
                            '流路遮蔽版
                            fncAttributeSymbolGet = "GD"
                        Case 24
                            '流路遮蔽版
                            fncAttributeSymbolGet = "GE"
                        Case 25
                            '接続ブロック
                            fncAttributeSymbolGet = "GF"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "06"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            'ベース
                            fncAttributeSymbolGet = "G1"
                        Case 2 To 7
                            '電磁弁形式
                            fncAttributeSymbolGet = "G2"
                        Case 8
                            'A・Bポート接続口径(01)
                            fncAttributeSymbolGet = "GH"
                        Case 9
                            'A・Bポート接続口径(02)
                            fncAttributeSymbolGet = "G6"
                        Case 10
                            'A・Bポート接続口径(C4)
                            fncAttributeSymbolGet = "GI"
                        Case 11
                            'A・Bポート接続口径(C6)
                            fncAttributeSymbolGet = "GJ"
                        Case 12
                            'A・Bポート接続口径(01Z)
                            fncAttributeSymbolGet = "GK"
                        Case 13 To 14
                            '給気スペーサ
                            fncAttributeSymbolGet = "G9"
                        Case 15 To 16
                            '排気スペーサ
                            fncAttributeSymbolGet = "GA"
                        Case 17
                            'パイロットチェック弁
                            fncAttributeSymbolGet = "GB"
                        Case 18
                            '流路遮蔽版
                            fncAttributeSymbolGet = "GD"
                        Case 19
                            '流路遮蔽版
                            fncAttributeSymbolGet = "GE"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "07", "96"     'RM1803032_スペーサ行数追加対応
                    Select Case intSpecStrcSeqNo
                        Case 1
                            '電装ブロック
                            fncAttributeSymbolGet = "B1"
                        Case 2 To 9
                            '電磁弁形式＆マスキングプレート
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "MN3GA1", "MN4GA1", "MN3GB1", "MN4GB1", _
                                     "MN3GA2", "MN4GA2", "MN3GB2", "MN4GB2", _
                                     "MN3GD1", "MN4GD1", "MN3GE1", "MN4GE1", _
                                     "MN3GD2", "MN4GD2", "MN3GE2", "MN4GE2"
                                    If InStr(1, strOptionKataban, "-MP") = 0 Then
                                        '電磁弁
                                        If Mid(strOptionKataban, 6, 1) = "1" Then
                                            fncAttributeSymbolGet = "D1"
                                        Else
                                            fncAttributeSymbolGet = "D2"
                                        End If
                                    Else
                                        'マスキングプレート
                                        If InStr(1, strOptionKataban, "-MPD") = 0 Then
                                            fncAttributeSymbolGet = "D3"
                                        Else
                                            fncAttributeSymbolGet = "D4"
                                        End If
                                    End If
                                Case "MN3GAX12", "MN4GAX12", "MN3GBX12", "MN4GBX12", "MN4GDX12", "MN4GEX12"
                                    If InStr(1, strOptionKataban, "-MP") = 0 Then
                                        '電磁弁
                                        If Mid(strOptionKataban, 5, 1) = "1" Then
                                            '1タイプ(MN*G*1)
                                            If Mid(strOptionKataban, 6, 1) = "1" Then
                                                fncAttributeSymbolGet = "D1"
                                            Else
                                                fncAttributeSymbolGet = "D2"
                                            End If
                                        Else
                                            '2タイプ(MN*G*2)
                                            If Mid(strOptionKataban, 6, 1) = "1" Then
                                                fncAttributeSymbolGet = "D5"
                                            Else
                                                fncAttributeSymbolGet = "D6"
                                            End If
                                        End If
                                    Else
                                        'マスキングプレート
                                        If Mid(strOptionKataban, 5, 1) = "1" Then
                                            '1タイプ(MN*G*1)
                                            If InStr(1, strOptionKataban, "-MPD") = 0 Then
                                                fncAttributeSymbolGet = "D3"
                                            Else
                                                fncAttributeSymbolGet = "D4"
                                            End If
                                        Else
                                            '2タイプ(MN*G*2)
                                            If InStr(1, strOptionKataban, "-MPD") = 0 Then
                                                fncAttributeSymbolGet = "D7"
                                            Else
                                                fncAttributeSymbolGet = "D8"
                                            End If
                                        End If
                                    End If
                            End Select
                        Case 10
                            'ミックスブロック
                            fncAttributeSymbolGet = "BB"
                        Case 11 To 14
                            '個別給気
                            If strOptionKataban.ToString.Contains("R-PC-M") Then
                                fncAttributeSymbolGet = "GB"
                            ElseIf strOptionKataban.ToString.Contains("R-IS") Then
                                fncAttributeSymbolGet = "S7"
                            ElseIf strOptionKataban.ToString.Contains("R-R") Then
                                fncAttributeSymbolGet = "S3"
                            Else
                                fncAttributeSymbolGet = "S2"
                            End If
                        Case 15 To 17
                            '給排気ブロック
                            fncAttributeSymbolGet = "B5"
                        Case 18 To 19
                            '仕切ブロック
                            fncAttributeSymbolGet = "BA"
                        Case 20 To 21
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 22 To 24
                            Select Case Left(strOptionKataban, 3)
                                Case "SLW"
                                    'サイレンサ
                                    fncAttributeSymbolGet = "T1"
                                Case "GWP"
                                    'ブランクプラグ
                                    fncAttributeSymbolGet = "T2"
                                Case Else
                                    'ダミーセット
                                    fncAttributeSymbolGet = "T1"
                            End Select
                        Case 25 To 26
                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    'Case CdCst.Manifold.InspReportJp.Japanese
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    'Case CdCst.Manifold.InspReportEn.Japanese
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                                Case Else
                                    'Select Case True
                                    '    Case InStr(1, strOptionKataban, "CABLE") <> 0
                                    '        'ケーブル
                                    '        fncAttributeSymbolGet = "T9"
                                    '    Case Else
                                    '        'ダミーセット
                                    '        fncAttributeSymbolGet = "T8"
                                    'End Select
                                    fncAttributeSymbolGet = "T9"
                            End Select
                        Case 27
                            'タグ銘板
                            fncAttributeSymbolGet = "T6"
                        Case 28
                            'チューブ抜具不要
                            fncAttributeSymbolGet = "TD"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "08"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            '左側エンドプーレート
                            fncAttributeSymbolGet = "P1"
                        Case 2
                            '右側エンドプーレート
                            fncAttributeSymbolGet = "P2"
                        Case 3 To 12
                            '電磁弁付サブプレート
                            fncAttributeSymbolGet = "P3"
                        Case 13 To 14
                            '中間給気プレート
                            fncAttributeSymbolGet = "P4"
                        Case 15 To 16
                            '中間排気プレート
                            fncAttributeSymbolGet = "P5"
                        Case 17 To 18
                            'サイレンサ
                            fncAttributeSymbolGet = "T1"
                        Case 19 To 20
                            'ブランクプラグ
                            fncAttributeSymbolGet = "T2"
                        Case 21
                            'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.SelectValue Then
                            '    'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.Japanese Then
                            '    '検査成績書(和文)
                            '    fncAttributeSymbolGet = "T8"
                            'Else
                            '    '検査成績書(英文)
                            '    fncAttributeSymbolGet = "TE"
                            'End If
                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                                Case Else
                                    fncAttributeSymbolGet = ""
                            End Select

                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "09"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 2
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 3
                            '配線ブロック
                            fncAttributeSymbolGet = "B1"
                        Case 4 To 8
                            'バルブブロック
                            fncAttributeSymbolGet = "B3"
                        Case 9 To 10
                            'MPV付バルブブロック
                            fncAttributeSymbolGet = "B4"
                        Case 11 To 13
                            'スペーサ形レギュレータ
                            fncAttributeSymbolGet = "S1"
                        Case 14
                            '単独給気スペーサ
                            fncAttributeSymbolGet = "S2"
                        Case 15
                            '単独排気スペーサ
                            fncAttributeSymbolGet = "S3"
                        Case 16 To 17
                            '仕切ブラグ
                            fncAttributeSymbolGet = "C1"
                        Case 18 To 19
                            'サイレンサ
                            fncAttributeSymbolGet = "T1"
                        Case 20 To 21
                            'プラグ
                            fncAttributeSymbolGet = "T4"
                        Case 22
                            'ケーブルクランプ
                            fncAttributeSymbolGet = "T3"
                        Case 23
                            'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.SelectValue Then
                            '    'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.Japanese Then
                            '    '検査成績書(和文)
                            '    fncAttributeSymbolGet = "T8"
                            'Else
                            '    '検査成績書(英文)
                            '    fncAttributeSymbolGet = "TE"
                            'End If
                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                                Case Else
                                    fncAttributeSymbolGet = ""
                            End Select

                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "10"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            '配線ブロック
                            fncAttributeSymbolGet = "B1"
                        Case 2 To 8
                            'バルブブロック
                            fncAttributeSymbolGet = "B3"
                        Case 9 To 10
                            '給排気ブロック
                            fncAttributeSymbolGet = "B5"
                        Case 11 To 12
                            '仕切ブロック
                            fncAttributeSymbolGet = "BA"
                        Case 13 To 14
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 15 To 16
                            'サイレンサ
                            fncAttributeSymbolGet = "T1"
                        Case 17 To 19
                            'ブランクプラグ
                            fncAttributeSymbolGet = "T2"
                        Case 20
                            'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.SelectValue Then
                            '    'If strOptionKataban.Trim = CdCst.Manifold.InspReportJp.Japanese Then
                            '    '検査成績書(和文)
                            '    fncAttributeSymbolGet = "T8"
                            'Else
                            '    '検査成績書(英文)
                            '    fncAttributeSymbolGet = "TE"
                            'End If
                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                            End Select
                        Case 21 To 22
                            'ケーブル
                            fncAttributeSymbolGet = "T9"
                        Case 23
                            'チューブ抜具不要
                            fncAttributeSymbolGet = "TD"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "11"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 2
                            '集中給気ブロック
                            fncAttributeSymbolGet = "BD"
                        Case 3
                            'APS付集中給気ブロック
                            fncAttributeSymbolGet = "BE"
                        Case 4 To 13
                            'レギュレータブロック
                            fncAttributeSymbolGet = "BF"
                        Case 14
                            'MP付サブベース
                            fncAttributeSymbolGet = "BG"
                        Case 15
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 16 To 18
                            'ブランクプラグ
                            fncAttributeSymbolGet = "T2"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "12", "18", "19", "20", "21", "22", "23"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 8
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "VSKM", "VSJM", "VSXM", "VSZM", "VSNM"
                                    '真空エジェクタ
                                    fncAttributeSymbolGet = "H1"
                                Case "VSJPM", "VSXPM", "VSZPM", "VSNPM"
                                    '真空切替ユニット
                                    fncAttributeSymbolGet = "H2"
                            End Select
                        Case 9 To 10
                            'マスキングブロック
                            fncAttributeSymbolGet = "H3"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "13"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            '配線ブロック
                            fncAttributeSymbolGet = "B1"
                        Case 2 To 3
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 4 To 9
                            '電磁弁＆ＭＰ付バルブブロック
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "MN4TB1"
                                    If InStr(1, strOptionKataban, "MPV") = 0 Then
                                        '電磁弁
                                        fncAttributeSymbolGet = "B3"
                                    Else
                                        'MPV
                                        Select Case Mid(strOptionKataban, 5, 1)
                                            Case "1"
                                                fncAttributeSymbolGet = "B4"
                                            Case "2"
                                                fncAttributeSymbolGet = "B9"
                                        End Select
                                    End If
                                Case "MN4TB2"
                                    If InStr(1, strOptionKataban, "MPV") = 0 Then
                                        '電磁弁
                                        fncAttributeSymbolGet = "B3"
                                    Else
                                        'MPV
                                        Select Case Mid(strOptionKataban, 5, 1)
                                            Case "1"
                                                fncAttributeSymbolGet = "B4"
                                            Case "2"
                                                fncAttributeSymbolGet = "B9"
                                        End Select
                                    End If
                                Case "MN4TBX12"
                                    If InStr(1, strOptionKataban, "MPV") = 0 Then
                                        '電磁弁
                                        Select Case Mid(strOptionKataban, 5, 1)
                                            Case "1"
                                                fncAttributeSymbolGet = "B3"
                                            Case "2"
                                                fncAttributeSymbolGet = "B8"
                                        End Select
                                    Else
                                        'MPV
                                        Select Case Mid(strOptionKataban, 5, 1)
                                            Case "1"
                                                fncAttributeSymbolGet = "B4"
                                            Case "2"
                                                fncAttributeSymbolGet = "B9"
                                        End Select
                                    End If
                            End Select
                        Case 10 To 11
                            '給排気ブロック
                            fncAttributeSymbolGet = "B5"
                        Case 12 To 13
                            '給気ブロック
                            fncAttributeSymbolGet = "B6"
                        Case 14 To 15
                            '排気ブロック
                            fncAttributeSymbolGet = "B7"
                        Case 16 To 17
                            '仕切プラグ
                            fncAttributeSymbolGet = "C1"
                        Case 18 To 19
                            Select Case Left(strOptionKataban, 3)
                                Case "SLW"
                                    'サイレンサ
                                    fncAttributeSymbolGet = "T1"
                                Case "GWP"
                                    'ブランクプラグ
                                    fncAttributeSymbolGet = "T2"
                                Case "4T9"
                                    'ケーブルクランプ
                                    fncAttributeSymbolGet = "T3"
                                Case Else
                                    'ダミーセット
                                    fncAttributeSymbolGet = "T1"
                            End Select
                        Case 20 To 21
                            Select Case Left(strOptionKataban, 3)
                                Case "SLW"
                                    'サイレンサ
                                    fncAttributeSymbolGet = "T1"
                                Case "GWP"
                                    'ブランクプラグ
                                    fncAttributeSymbolGet = "T2"
                                Case "4T9"
                                    'ケーブルクランプ
                                    fncAttributeSymbolGet = "T3"
                                Case Else
                                    'ダミーセット
                                    fncAttributeSymbolGet = "T2"
                            End Select
                        Case 22
                            'Case 23
                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    'Case CdCst.Manifold.InspReportJp.Japanese
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    'Case CdCst.Manifold.InspReportEn.Japanese
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                            End Select


                        Case 23 To 24
                            'Case 24 To 25
                            'ケーブル
                            fncAttributeSymbolGet = "T9"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "14"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            'EVT
                            fncAttributeSymbolGet = "F1"
                        Case 2 To 4
                            '電装・給排気ブロック
                            fncAttributeSymbolGet = "E1"
                        Case 5 To 6
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 7 To 8
                            'ブランクプラグ
                            fncAttributeSymbolGet = "T2"
                        Case 9
                            'サイレンサ
                            fncAttributeSymbolGet = "T1"
                        Case 10
                            'マニホールド長さ
                            fncAttributeSymbolGet = "L1"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "15"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 2
                            '入出力ブロック
                            fncAttributeSymbolGet = "BC"
                        Case 3
                            '電装ブロック
                            fncAttributeSymbolGet = "B1"
                        Case 4 To 11
                            If InStr(1, strOptionKataban, "-MP") = 0 Then
                                '電磁弁バルブブロック
                                If Mid(strOptionKataban, 7, 1) = "1" Then
                                    fncAttributeSymbolGet = "D1"
                                Else
                                    fncAttributeSymbolGet = "D2"
                                End If
                            Else
                                'ＭＰ付バルブブロック
                                If InStr(1, strOptionKataban, "-MPS") <> 0 Then
                                    fncAttributeSymbolGet = "D3"
                                ElseIf InStr(1, strOptionKataban, "-MPD") <> 0 Then
                                    fncAttributeSymbolGet = "D4"
                                Else
                                    fncAttributeSymbolGet = "D3"
                                End If
                            End If
                        Case 12 To 15
                            'スペーサ1
                            If InStr(1, strOptionKataban, "-PC") <> 0 Then
                                fncAttributeSymbolGet = "S6"
                            ElseIf InStr(1, strOptionKataban, "-PIS") <> 0 Then
                                fncAttributeSymbolGet = "S7"
                            ElseIf InStr(1, strOptionKataban, "-P") <> 0 Then
                                fncAttributeSymbolGet = "S2"
                            ElseIf InStr(1, strOptionKataban, "-R") <> 0 Then
                                fncAttributeSymbolGet = "S3"
                            End If
                        Case 16 To 17
                            '給排気ブロック
                            fncAttributeSymbolGet = "B5"
                        Case 18 To 19
                            '仕切ブロック
                            fncAttributeSymbolGet = "BA"
                        Case 20 To 21
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 22 To 24
                            Select Case Left(strOptionKataban, 3)
                                Case "SLW"
                                    'サイレンサ
                                    fncAttributeSymbolGet = "T1"
                                Case "GWP"
                                    'ブランクプラグ
                                    fncAttributeSymbolGet = "T2"
                                Case Else
                                    'ダミーセット
                                    fncAttributeSymbolGet = "T1"
                            End Select
                        Case 25
                            '防水プラグ
                            fncAttributeSymbolGet = "T7"
                        Case 26 To 27
                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    'Case CdCst.Manifold.InspReportJp.Japanese
                                    '検査成績書(和文)
                                    fncAttributeSymbolGet = "T8"
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    'Case CdCst.Manifold.InspReportEn.Japanese
                                    '検査成績書(英文)
                                    fncAttributeSymbolGet = "TE"
                                Case Else
                                    'Select Case True
                                    '    Case InStr(1, strOptionKataban, "CABLE") <> 0
                                    '        'ケーブル
                                    '        fncAttributeSymbolGet = "T9"
                                    '    Case Else
                                    '        'ダミーセット
                                    '        fncAttributeSymbolGet = "T8"
                                    'End Select
                                    fncAttributeSymbolGet = "T9"
                            End Select
                        Case 28
                            'ケーブルクランプ
                            fncAttributeSymbolGet = "T3"
                        Case 29
                            'タグ銘板
                            fncAttributeSymbolGet = "T6"
                        Case 30
                            'タグプレート長さ
                            fncAttributeSymbolGet = "L4"
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "16"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 2
                            '入出力ブロック
                            fncAttributeSymbolGet = "BC"
                        Case 3 To 4
                            'エンドブロック
                            fncAttributeSymbolGet = "B2"
                        Case 5
                            '配線ブロック
                            fncAttributeSymbolGet = "B1"
                        Case 6 To 10
                            'バルブブロック
                            fncAttributeSymbolGet = "B3"
                        Case 11 To 12
                            'MPV付バルブブロック
                            fncAttributeSymbolGet = "B4"
                        Case 13
                            '仕切りブロック
                            fncAttributeSymbolGet = "BA"
                        Case 14
                            '単独給気スペーサ
                            fncAttributeSymbolGet = "S2"
                        Case 15
                            '単独排気スペーサ
                            fncAttributeSymbolGet = "S3"
                        Case 16 To 18
                            'スペーサ形レギュレータ
                            fncAttributeSymbolGet = "S1"
                        Case 19 To 20
                            '仕切プラグ
                            fncAttributeSymbolGet = "C1"
                        Case 21 To 23
                            'ブランクプラグ＆サイレンサ
                            'fncAttributeSymbolGet = "T2"
                            Select Case Left(strOptionKataban, 3)
                                Case "SLW"
                                    'サイレンサ
                                    fncAttributeSymbolGet = "T1"
                                Case "GWP"
                                    'ブランクプラグ
                                    fncAttributeSymbolGet = "T2"
                                Case Else
                                    'ダミーセット
                                    fncAttributeSymbolGet = "T1"
                            End Select
                        Case 24
                            'ケーブルクランプ
                            fncAttributeSymbolGet = "T3"
                        Case 25
                            Select Case strOptionKataban.Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue
                                    'Case CdCst.Manifold.InspReportJp.Japanese
                                    fncAttributeSymbolGet = "T8" '検査成績書(和文)
                                Case CdCst.Manifold.InspReportEn.SelectValue
                                    'Case CdCst.Manifold.InspReportEn.Japanese
                                    fncAttributeSymbolGet = "TE" '検査成績書(英文)
                            End Select
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "17"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 5
                            fncAttributeSymbolGet = "BI" '単体バルブ
                        Case 6
                            fncAttributeSymbolGet = "G1" 'ベ－ス
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select

                Case "A1", "A2", "54", "55", "56", "57", "58", "59", "A9", "B1", "B2", "B3", "B4"
                    Select Case intSpecStrcSeqNo
                        Case 1
                            fncAttributeSymbolGet = "B8" '電磁弁ブロック
                        Case 2
                            fncAttributeSymbolGet = "D9"  'マルキングプレート
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "60"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 2
                            fncAttributeSymbolGet = "B8" '電磁弁ブロック
                        Case 3
                            fncAttributeSymbolGet = "D9" 'マルキングプレート
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "51"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 2
                            '電磁弁ブロック
                            fncAttributeSymbolGet = "B8"
                        Case 3
                            'マルキングプレート
                            If InStr(1, strOptionKataban, "-MP") = 0 Then
                                fncAttributeSymbolGet = "B8"
                            Else
                                fncAttributeSymbolGet = "D9"
                            End If
                        Case 4
                            If InStr(1, strOptionKataban, "-MP") = 0 Then
                                fncAttributeSymbolGet = ""
                            Else
                                fncAttributeSymbolGet = "D9"
                            End If
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "63", "64", "65", "66", "67", "68", "69", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "91", "92", "A4", "A5", "A6", "A7", "A8", "98", "U"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 5
                            fncAttributeSymbolGet = "B8" '電磁弁ブロック
                        Case 6
                            fncAttributeSymbolGet = "D9" 'マルキングプレート
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "53", "93", _
                            "S", "T"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 3
                            fncAttributeSymbolGet = "B8" '電磁弁ブロック
                        Case 4
                            fncAttributeSymbolGet = "D9" 'マルキングプレート
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                Case "61", "62"
                    Select Case intSpecStrcSeqNo
                        Case 1 To 4
                            fncAttributeSymbolGet = "B8" '電磁弁ブロック
                        Case 5
                            fncAttributeSymbolGet = "D9" 'マルキングプレート
                        Case Else
                            fncAttributeSymbolGet = ""
                    End Select
                        Case Else
                            fncAttributeSymbolGet = ""
            End Select
        Catch ex As Exception
            fncAttributeSymbolGet = ""
        Finally
            objKtbnStrc = Nothing
        End Try

    End Function

    ''' <summary>
    ''' 項目、項目内容テーブルの項目を設定
    ''' </summary>
    ''' <param name="dtSpecItem"></param>
    ''' <param name="dtContent"></param>
    ''' <remarks></remarks>
    Public Shared Sub subInitTable(ByRef dtSpecItem As DataTable, ByRef dtContent As DataTable)
        Dim objColumn As DataColumn
        Try
            '===========================================
            '項目
            dtSpecItem = New DataTable

            '品名(記号)
            objColumn = New DataColumn
            objColumn.DataType = GetType(String)
            objColumn.ColumnName = CdCst.TblSpecItem.ProdNm
            objColumn.AllowDBNull = False
            dtSpecItem.Columns.Add(objColumn)

            '項目数
            objColumn = New DataColumn
            objColumn.DataType = GetType(Integer)
            objColumn.ColumnName = CdCst.TblSpecItem.ItemCnt
            objColumn.AllowDBNull = False
            dtSpecItem.Columns.Add(objColumn)

            '項目区分
            objColumn = New DataColumn
            objColumn.DataType = GetType(String)
            objColumn.ColumnName = CdCst.TblSpecItem.ItemDiv
            objColumn.AllowDBNull = False
            dtSpecItem.Columns.Add(objColumn)

            '===========================================
            '項目内容
            dtContent = New DataTable
            '品名(形番)
            objColumn = New DataColumn
            objColumn.DataType = GetType(String)
            objColumn.ColumnName = CdCst.TblSpecItem.ProdNm
            objColumn.AllowDBNull = False
            dtContent.Columns.Add(objColumn)
            '項目数
            objColumn = New DataColumn
            objColumn.DataType = GetType(String)
            objColumn.ColumnName = CdCst.TblSpecItem.ItemCnt
            objColumn.AllowDBNull = False
            objColumn.DefaultValue = ""
            dtContent.Columns.Add(objColumn)
            '項目区分
            objColumn = New DataColumn
            objColumn.DataType = GetType(Integer)
            objColumn.ColumnName = CdCst.TblSpecItem.ItemDiv
            objColumn.AllowDBNull = False
            dtContent.Columns.Add(objColumn)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 形番変更
    ''' </summary>
    ''' <param name="dtSpecItem"></param>
    ''' <param name="dtContent"></param>
    ''' <param name="strSpecNo"></param>
    ''' <param name="strSeriesKataban"></param>
    ''' <param name="strOpSymbol"></param>
    ''' <param name="strKeyKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncGetNewKataban(ByVal dtSpecItem As DataTable, ByVal dtContent As DataTable, _
                                ByVal strSpecNo As String, ByVal strSeriesKataban As String, _
                                ByVal strOpSymbol() As String, ByVal strKeyKataban As String, _
                                Optional ByVal strSelLang As String = "") As ArrayList
        fncGetNewKataban = New ArrayList
        Dim strKtbn As String
        Dim intAddCnt As Integer = 0
        Dim strValues As String = ""
        Dim CST_ZERO As String = ""
        Dim CST_BLANK As String = ""

        Try
            Select Case strSpecNo
                Case "64", "66", "68", "70", "72"
                    For idx As Integer = 0 To dtContent.Rows.Count - 1
                        If (Mid(strSeriesKataban, 4, 1) = "6" _
                        Or Mid(strSeriesKataban, 4, 1) = "7") _
                        And Not dtSpecItem.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "MP" Then
                            '形番の頭7桁取得
                            strKtbn = Left(dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm), 7)
                            fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(intAddCnt + 1) & "_" & _
                                strKtbn & strOpSymbol(4) & "_" & CST_BLANK & "_" & CST_BLANK)
                            intAddCnt = intAddCnt + 1
                        Else
                            '形番ﾃｰﾌﾞﾙをそのままｾｯﾄ
                            fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(intAddCnt + 1) & "_" & _
                                dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                            intAddCnt = intAddCnt + 1
                        End If
                    Next
                Case "S"
                    Select Case strSeriesKataban
                        Case "M4HA2"
                            dtContent.Rows(0).Item(CdCst.TblSpecItem.ProdNm) = "4HA219"
                            dtContent.Rows(1).Item(CdCst.TblSpecItem.ProdNm) = "4HA229"
                            dtContent.Rows(2).Item(CdCst.TblSpecItem.ProdNm) = "4HA239"
                        Case "M4HA3"
                            dtContent.Rows(0).Item(CdCst.TblSpecItem.ProdNm) = "4HA319"
                            dtContent.Rows(1).Item(CdCst.TblSpecItem.ProdNm) = "4HA329"
                            dtContent.Rows(2).Item(CdCst.TblSpecItem.ProdNm) = "4HA339"
                    End Select

                    For idx As Integer = 0 To dtContent.Rows.Count - 1
                        If strOpSymbol(1) <> "8" And _
                           InStr(dtSpecItem.Rows(idx).Item(CdCst.TblSpecItem.ProdNm), "MP") Then
                        Else
                            fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                            intAddCnt = intAddCnt + 1
                        End If
                    Next
                Case "T"
                    Select Case strSeriesKataban
                        Case "M4JA2"
                            dtContent.Rows(0).Item(CdCst.TblSpecItem.ProdNm) = "4JA219"
                            dtContent.Rows(1).Item(CdCst.TblSpecItem.ProdNm) = "4JA229"
                            dtContent.Rows(2).Item(CdCst.TblSpecItem.ProdNm) = "4JA239"
                        Case "M4JA3"
                            dtContent.Rows(0).Item(CdCst.TblSpecItem.ProdNm) = "4JA319"
                            dtContent.Rows(1).Item(CdCst.TblSpecItem.ProdNm) = "4JA329"
                            dtContent.Rows(2).Item(CdCst.TblSpecItem.ProdNm) = "4JA339"
                    End Select

                    For idx As Integer = 0 To dtContent.Rows.Count - 1
                        If strOpSymbol(1) <> "8" And _
                           InStr(dtSpecItem.Rows(idx).Item(CdCst.TblSpecItem.ProdNm), "MP") Then
                        Else
                            fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                            intAddCnt = intAddCnt + 1
                        End If
                    Next
                    'RM1805001_4Rシリーズ追加
                Case "U"
                    Select Case strSeriesKataban
                        Case "M4RD2"
                            dtContent.Rows(0).Item(CdCst.TblSpecItem.ProdNm) = "4RD219"
                            dtContent.Rows(1).Item(CdCst.TblSpecItem.ProdNm) = "4RD229"
                            dtContent.Rows(2).Item(CdCst.TblSpecItem.ProdNm) = "4RD239"
                            dtContent.Rows(3).Item(CdCst.TblSpecItem.ProdNm) = "4RD249"
                            dtContent.Rows(4).Item(CdCst.TblSpecItem.ProdNm) = "4RD259"
                        Case "M4RE1"
                            dtContent.Rows(0).Item(CdCst.TblSpecItem.ProdNm) = "4RE119"
                            dtContent.Rows(1).Item(CdCst.TblSpecItem.ProdNm) = "4RE129"
                            dtContent.Rows(2).Item(CdCst.TblSpecItem.ProdNm) = "4RE139"
                            dtContent.Rows(3).Item(CdCst.TblSpecItem.ProdNm) = "4RE149"
                            dtContent.Rows(4).Item(CdCst.TblSpecItem.ProdNm) = "4RE159"
                        Case "M4RE2"
                            dtContent.Rows(0).Item(CdCst.TblSpecItem.ProdNm) = "4RE219"
                            dtContent.Rows(1).Item(CdCst.TblSpecItem.ProdNm) = "4RE229"
                            dtContent.Rows(2).Item(CdCst.TblSpecItem.ProdNm) = "4RE239"
                            dtContent.Rows(3).Item(CdCst.TblSpecItem.ProdNm) = "4RE249"
                            dtContent.Rows(4).Item(CdCst.TblSpecItem.ProdNm) = "4RE259"
                    End Select

                    For idx As Integer = 0 To dtContent.Rows.Count - 1
                        If strOpSymbol(1) <> "8" And _
                           InStr(dtSpecItem.Rows(idx).Item(CdCst.TblSpecItem.ProdNm), "MP") Then
                        Else
                            fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                            intAddCnt = intAddCnt + 1
                        End If
                    Next
                Case "52", "60", "61", "62", "63", "65", "67", "69", "71", "A4", "A5", "A6", "A7", "A8"

                    For idx As Integer = 0 To dtContent.Rows.Count - 1
                        If dtSpecItem.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "MP" Then
                            If (strSeriesKataban = "M4F2" Or strSeriesKataban = "M4F3") _
                            And (strOpSymbol(8) = "C" Or strOpSymbol(8) = "I") Then
                                '設定なし
                            Else
                                fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(intAddCnt + 1) & "_" & _
                                    dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                                intAddCnt = intAddCnt + 1
                            End If
                        Else

                            If Left(dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm), 1) = "A" Then
                                Select Case Trim(strSeriesKataban)
                                    Case "M4F0"
                                        '形番の頭6桁取得
                                        strKtbn = Left(dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm), 6)
                                        CST_ZERO = "-" & strOpSymbol(3)
                                        fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(intAddCnt + 1) & "_" & _
                                            strKtbn & CST_ZERO & "_" & CST_BLANK & "_" & CST_BLANK)
                                        intAddCnt = intAddCnt + 1
                                    Case Else
                                        '形番の頭5桁取得
                                        strKtbn = Left(dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm), 5)
                                        CST_ZERO = "-" & strOpSymbol(3)
                                        fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(intAddCnt + 1) & "_" & _
                                            strKtbn & CST_ZERO & "_" & CST_BLANK & "_" & CST_BLANK)
                                        intAddCnt = intAddCnt + 1
                                End Select
                            Else
                                Select Case Trim(strSeriesKataban)
                                    Case "M4F0"
                                        If strKeyKataban = "M" Then
                                            If strOpSymbol(3) = "06" Then
                                                CST_ZERO = "1-06"
                                            Else
                                                CST_ZERO = "1-M5"
                                            End If
                                        Else
                                            If strOpSymbol(3) = "06" Then
                                                CST_ZERO = "0-06"
                                            Else
                                                CST_ZERO = "0-M5"
                                            End If
                                        End If
                                    Case "M4F1"
                                        If strKeyKataban = "M" Then
                                            If strOpSymbol(3) = "06" Then
                                                CST_ZERO = "1-06"
                                            Else
                                                CST_ZERO = "1-08"
                                            End If
                                        Else
                                            If strOpSymbol(3) = "06" Then
                                                CST_ZERO = "0-06"
                                            Else
                                                CST_ZERO = "0-08"
                                            End If
                                        End If
                                    Case "M4F2"
                                        If strKeyKataban = "M" Then
                                            Select Case strOpSymbol(8).Trim
                                                Case "C"
                                                    CST_ZERO = "8-08"

                                                Case "I"
                                                    CST_ZERO = "8-08"

                                                Case Else
                                                    CST_ZERO = "1-08"
                                            End Select
                                        Else
                                            Select Case strOpSymbol(8).Trim
                                                Case "C"
                                                    CST_ZERO = "9-08"

                                                Case "I"
                                                    CST_ZERO = "9-08"

                                                Case Else
                                                    CST_ZERO = "0-08"
                                            End Select
                                        End If
                                    Case "M4F3"
                                        'RM1312084 2013/12/24
                                        If strKeyKataban <> "X" Then
                                            If strKeyKataban = "M" Then
                                                Select Case strOpSymbol(8).Trim
                                                    Case "C"
                                                        If strOpSymbol(3) = "08" Then
                                                            CST_ZERO = "8-08"
                                                        Else
                                                            CST_ZERO = "8-10"
                                                        End If
                                                    Case "I"
                                                        If strOpSymbol(3) = "08" Then
                                                            CST_ZERO = "8-08"
                                                        Else
                                                            CST_ZERO = "8-10"
                                                        End If
                                                    Case Else
                                                        If strOpSymbol(3) = "08" Then
                                                            CST_ZERO = "1-08"
                                                        Else
                                                            CST_ZERO = "1-10"
                                                        End If
                                                End Select
                                            Else
                                                Select Case strOpSymbol(8).Trim
                                                    Case "C"
                                                        If strOpSymbol(3) = "08" Then
                                                            CST_ZERO = "9-08"
                                                        Else
                                                            CST_ZERO = "9-10"
                                                        End If
                                                    Case "I"
                                                        If strOpSymbol(3) = "08" Then
                                                            CST_ZERO = "9-08"
                                                        Else
                                                            CST_ZERO = "9-10"
                                                        End If
                                                    Case Else
                                                        If strOpSymbol(3) = "08" Then
                                                            CST_ZERO = "0-08"
                                                        Else
                                                            CST_ZERO = "0-10"
                                                        End If
                                                End Select
                                            End If
                                        Else
                                            CST_ZERO = "0EX"
                                        End If
                                        '2013/11/06 修正
                                    Case "M4F4", "M4F5"
                                        'RM1312084 2013/12/24
                                        If strKeyKataban <> "X" Then
                                            If strKeyKataban = "M" Then
                                                CST_ZERO = "8-00"
                                            Else
                                                CST_ZERO = "9-00"
                                            End If
                                        Else
                                            CST_ZERO = "9EX"
                                        End If
                                    Case "M4F6"
                                        'RM1312084 2013/12/24
                                        If strKeyKataban <> "X" Then
                                            If strKeyKataban = "M" Then
                                                CST_ZERO = "8-D00"
                                            Else
                                                CST_ZERO = "9-D00"
                                            End If
                                        Else
                                            CST_ZERO = "9EX"
                                        End If
                                    Case "M4F7"
                                        'RM1312084 2013/12/24
                                        If strKeyKataban <> "X" Then
                                            If strKeyKataban = "M" Then
                                                CST_ZERO = "8-E00"
                                            Else
                                                CST_ZERO = "9-E00"
                                            End If
                                        Else
                                            CST_ZERO = "9EX"
                                        End If
                                    Case Else
                                End Select
                                '形番の頭4桁取得
                                strKtbn = Left(dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm), 4)
                                fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(intAddCnt + 1) & "_" & _
                                    strKtbn & CST_ZERO & "_" & CST_BLANK & "_" & CST_BLANK)
                                intAddCnt = intAddCnt + 1
                            End If
                        End If
                    Next
                Case "51"

                    For idx As Integer = 0 To dtContent.Rows.Count - 1
                        If (dtSpecItem.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "M4") _
                        And ((strOpSymbol(4) = "6") _
                        Or (strOpSymbol(5) = "M0" Or strOpSymbol(5) = "M1" Or strOpSymbol(5) = "M4")) Then

                        Else
                            fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(intAddCnt + 1) & "_" & _
                                dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                            intAddCnt = intAddCnt + 1
                        End If
                    Next
                Case "A1", "A2", "B2", "B3", "B4"
                    Select Case strSeriesKataban
                        Case "M3QRA1"
                            If strOpSymbol(1) = "2" Then
                                dtContent.Rows(0).Item(CdCst.TblSpecItem.ProdNm) = "3QRA129"
                            End If
                        Case "M3QRB1"
                            If strOpSymbol(1) = "2" Then
                                dtContent.Rows(0).Item(CdCst.TblSpecItem.ProdNm) = "3QRB129"
                            End If
                    End Select

                    For idx As Integer = 0 To dtContent.Rows.Count - 1
                        If strOpSymbol(1) <> "8" And _
                           InStr(dtSpecItem.Rows(idx).Item(CdCst.TblSpecItem.ProdNm), "MP") Then
                        Else
                            fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                            intAddCnt = intAddCnt + 1
                        End If
                    Next
                Case "A9", "B1"
                    Select Case strOpSymbol(1)
                        Case "1"
                            If strOpSymbol(8) = "V1" Then
                                For idx As Integer = 0 To dtContent.Rows.Count - 1
                                    If dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119" Then
                                        'dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119+ｾﾝｻ"
                                        If strSelLang.Equals("ja") OrElse strSelLang.Equals(String.Empty) Then
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119+" & CdCst.Senser.ja
                                        Else
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119+" & CdCst.Senser.en
                                        End If
                                        fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                                        intAddCnt = intAddCnt + 1
                                        Exit For
                                    End If
                                Next
                            Else
                                For idx As Integer = 0 To dtContent.Rows.Count - 1
                                    If dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119" Then
                                        fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                                        intAddCnt = intAddCnt + 1
                                        Exit For
                                    Else
                                        fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                                        intAddCnt = intAddCnt + 1
                                        Exit For
                                    End If
                                Next
                            End If
                        Case "2"
                            If strOpSymbol(8) = "V1" Then
                                For idx As Integer = 0 To dtContent.Rows.Count - 1
                                    If dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119" Then
                                        If strSelLang.Equals("ja") OrElse strSelLang.Equals(String.Empty) Then
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA129+" & CdCst.Senser.ja
                                        Else
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA129+" & CdCst.Senser.en
                                        End If
                                        fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                                        intAddCnt = intAddCnt + 1
                                        Exit For
                                    End If
                                Next
                            Else
                                For idx As Integer = 0 To dtContent.Rows.Count - 1
                                    If dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119" Then
                                        dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA129"
                                        fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                                        intAddCnt = intAddCnt + 1
                                        Exit For
                                    Else
                                        dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRB129"
                                        fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                                        intAddCnt = intAddCnt + 1
                                        Exit For
                                    End If
                                Next
                            End If
                        Case "8"
                            If strOpSymbol(8) = "V1" Then
                                For idx As Integer = 0 To dtContent.Rows.Count - 1
                                    If dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119" Then
                                        If strSelLang.Equals("ja") OrElse strSelLang.Equals(String.Empty) Then
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119+" & CdCst.Senser.ja
                                        Else
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA119+" & CdCst.Senser.en
                                        End If
                                    ElseIf dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA129" Then
                                        If strSelLang.Equals("ja") OrElse strSelLang.Equals(String.Empty) Then
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA129+" & CdCst.Senser.ja
                                        Else
                                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) = "3QRA129+" & CdCst.Senser.en
                                        End If
                                    End If
                                Next
                            End If
                            For idx As Integer = 0 To dtContent.Rows.Count - 1
                                fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                                    dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                                intAddCnt = intAddCnt + 1
                            Next
                    End Select
                Case Else
                    For idx As Integer = 0 To dtContent.Rows.Count - 1
                        fncGetNewKataban.Add(CdCst.SpecInfoItem.Kataban & CStr(idx + 1) & "_" & _
                            dtContent.Rows(idx).Item(CdCst.TblSpecItem.ProdNm) & "_" & CST_BLANK & "_" & CST_BLANK)
                        intAddCnt = intAddCnt + 1
                    Next
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function
End Class
