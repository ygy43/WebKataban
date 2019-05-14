Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports System.IO

Public Class TankaBLL

    ''' <summary>
    ''' 単価情報をDBに保存
    ''' </summary>
    ''' <param name="strPriceList"></param>
    ''' <param name="dt_display"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strHostName"></param>
    ''' <param name="processStartTime"></param>
    ''' <remarks></remarks>
    Public Sub subInsertPriceInfoToHistoryTable(ByVal strPriceList(,) As String, ByVal dt_display As DataTable, _
                                              ByVal objKtbnStrc As KHKtbnStrc, ByVal strUserId As String, _
                                              ByVal strHostName As String, ByVal processStartTime As Date)
        'DBに単価情報を保存する
        Dim dt_history As New DS_History.kh_price_historyDataTable
        Dim dr_history As DataRow = dt_history.NewRow
        Try
            '付加情報
            Dim strKeylvl As String = "64,2,1"
            Dim strLevel() As String = strKeylvl.Split(",")
            Dim dr_display() As DataRow = Nothing
            For inti As Integer = 0 To strLevel.Length - 1
                If dt_display Is Nothing Then Exit For
                dr_display = dt_display.Select("strLevel='" & CInt(strLevel(inti)) & "'")
                Select Case CInt(strLevel(inti))
                    Case 64 'EL品情報
                        If dr_display.Length > 0 Then dr_history("ELFlag") = dr_display(0)("strValue").ToString
                    Case 2 '出荷場所
                        If dr_display.Length > 0 Then dr_history("KataPlace") = dr_display(0)("strValue").ToString
                    Case 1 '形番チェック区分
                        If dr_display.Length > 0 Then dr_history("KataCheck") = dr_display(0)("strValue").ToString
                End Select
            Next

            '単価リスト
            For intRow As Integer = 1 To UBound(strPriceList)
                Select Case strPriceList(intRow, 4)
                    Case CdCst.UnitPrice.ListPrice
                        dr_history("ListPrice") = strPriceList(intRow, 2)
                    Case CdCst.UnitPrice.RegPrice
                        dr_history("RegPrice") = strPriceList(intRow, 2)
                    Case CdCst.UnitPrice.SsPrice
                        dr_history("SSPrice") = strPriceList(intRow, 2)
                    Case CdCst.UnitPrice.BsPrice
                        dr_history("BSPrice") = strPriceList(intRow, 2)
                    Case CdCst.UnitPrice.GsPrice
                        dr_history("GSPrice") = strPriceList(intRow, 2)
                    Case CdCst.UnitPrice.PsPrice
                        dr_history("PSPrice") = strPriceList(intRow, 2)
                End Select
            Next
            Dim strEnd As Date = Now
            Dim strTime As TimeSpan = strEnd - processStartTime

            dr_history("UpdateDate") = Now
            dr_history("UpdateComputer") = Right(strHostName.PadRight(10), 10)
            dr_history("UpdateUser") = strUserId
            dr_history("MFNo") = "0"
            dr_history("Kataban_Title") = objKtbnStrc.strcSelection.strGoodsNm
            dr_history("Kataban") = objKtbnStrc.strcSelection.strFullKataban
            dr_history("Runtime") = strTime.Milliseconds
            dt_history.Rows.Add(dr_history)
            Using da As New DS_HistoryTableAdapters.kh_price_historyTableAdapter
                da.Update(dt_history)
            End Using
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' ログファイル出力処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strPriceList">単価リスト</param>
    ''' <param name="dt_display"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strLanguage"></param>
    ''' <param name="strKatabanInfo"></param>
    ''' <remarks>テキスト出力(比較のため、１ヶ月削除保留)</remarks>
    Public Sub subLogFileOutput(objCon As SqlConnection, ByVal strPriceList(,) As String, ByVal dt_display As DataTable, _
                                       strUserId As String, strSessionId As String, strCountryCd As String, _
                                       strLanguage As String, Optional ByVal strKatabanInfo() As String = Nothing)
        Dim strLogFolder As String
        Dim strLogFilePath As String
        Dim strLogFileName As String
        Dim strSystemDatetime As String
        Dim bolDownload As Boolean

        Dim objWriter As StreamWriter
        Dim strOutputText As String

        Try
            bolDownload = False

            ' システム日付設定
            strSystemDatetime = Format(Now(), "yyyyMMdd")
            'Web.configよりログファイル出力フォルダ取得
            strLogFolder = My.Settings.LogFolder
            'Web.configよりログファイル取得
            strLogFileName = My.Settings.LogFileName & strSystemDatetime & CdCst.File.TextExtension
            'ログファイルパス設定
            strLogFilePath = strLogFolder & strLogFileName

            'ディレクトリ存在確認
            If System.IO.Directory.Exists(strLogFolder) = False Then
                '存在しない場合は作成する
                System.IO.Directory.CreateDirectory(strLogFolder)
            End If

            '出力内容取得
            strOutputText = fncLogInfoGet(objCon, strUserId, strSessionId, strCountryCd, _
                                          strLanguage, strPriceList, dt_display, strKatabanInfo)
            'ファイルOpen
            objWriter = New StreamWriter(strLogFilePath, True, System.Text.Encoding.GetEncoding("Shift-Jis"))
            'Select Case strLanguage
            '    Case "ja"
            '        objWriter = New StreamWriter(strLogFilePath, True, System.Text.Encoding.GetEncoding("Shift-Jis"))
            '    Case "zh"
            '        'objWriter = New StreamWriter(strLogFilePath, True, System.Text.Encoding.GetEncoding("Shift-Jis"))
            '        objWriter = New StreamWriter(strLogFilePath, True, System.Text.Encoding.GetEncoding("gb2312"))
            '    Case Else
            '        objWriter = New StreamWriter(strLogFilePath, True, System.Text.Encoding.GetEncoding("utf-8"))
            'End Select
            'objWriter = New StreamWriter(strLogFilePath, True, System.Text.Encoding.Default)
            'ファイル出力
            objWriter.Write(strOutputText)
            'ファイルClose
            objWriter.Close()
            bolDownload = True
        Catch ex As Exception
            'エラー画面に遷移する
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' ログをDBに出力処理
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="objCon"></param>
    ''' <param name="strPriceList">単価リスト</param>
    ''' <param name="dt_display"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strLanguage"></param>
    ''' <remarks>ログを出力する</remarks>
    Public Sub subLogOutput(objConBase As SqlConnection, objCon As SqlConnection, _
                                   ByVal strPriceList(,) As String, ByVal dt_display As DataTable, _
                                   strUserId As String, strSessionId As String, strCountryCd As String, _
                                   strLanguage As String)
        Dim sbScript As New StringBuilder
        Dim bolReturn As Boolean
        Try
            'ログ出力
            bolReturn = fncSearchKatabanLogOutput(objConBase, objCon, strUserId, strSessionId, strCountryCd, _
                                                  strLanguage, strPriceList, dt_display)
        Catch ex As Exception
            'エラー画面に遷移する
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 形番検索ログ出力処理
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strSelectLang"></param>
    ''' <param name="strPriceList">単価リスト</param>
    ''' <param name="dt_display"></param>
    ''' <param name="strKatabanInfo">形番情報</param>
    ''' <returns></returns>
    ''' <remarks>形番の検索状況をログ出力する</remarks>
    Private Function fncSearchKatabanLogOutput(objConBase As SqlConnection, objCon As SqlConnection, _
                                    ByVal strUserId As String, ByVal strSessionId As String, _
                                    ByVal strCountryCd As String, ByVal strSelectLang As String, _
                                    ByVal strPriceList(,) As String, ByVal dt_display As DataTable, _
                                    Optional ByVal strKatabanInfo() As String = Nothing) As Boolean
        Dim objWriter As StreamWriter
        Dim strOutputText As String
        fncSearchKatabanLogOutput = False
        Try
            '出力内容取得
            strOutputText = fncLogInfoGet(objCon, strUserId, strSessionId, strCountryCd, _
                                          strSelectLang, strPriceList, dt_display, strKatabanInfo)
            Dim strSetVal() As String
            strSetVal = strOutputText.Split(CdCst.Sign.Delimiter.Tab)

            '形番検索LOGへ追加
            If Not InsertSearchKatabanLogTbl(objConBase, strSetVal) Then Exit Try
            fncSearchKatabanLogOutput = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objWriter = Nothing
        End Try

    End Function

    ''' <summary>
    ''' 形番検索LOGInsert処理
    ''' </summary>
    ''' <param name="strSetValue">設定値</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InsertSearchKatabanLogTbl(objConBase As SqlConnection, _
                                                      ByVal strSetValue() As String) As Boolean
        'Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objTrans As SqlTransaction
        InsertSearchKatabanLogTbl = False
        Try
            'DB接続文字列の取得
            objCmd = objConBase.CreateCommand()
            ' トランザクションの開始
            objTrans = objConBase.BeginTransaction(IsolationLevel.ReadCommitted, "InsertTran")
            objCmd.Connection = objConBase
            objCmd.Transaction = objTrans

            Dim sbSql As New StringBuilder()
            sbSql.Append(" INSERT INTO kh_SearckKatabanLog " & Environment.NewLine)
            sbSql.Append(" VALUES ( " & Environment.NewLine)
            sbSql.Append(" @LogDateTime, " & Environment.NewLine)          'ログ時間
            sbSql.Append(" @UserID, " & Environment.NewLine)               'ユーザーID
            sbSql.Append(" @CountryCD, " & Environment.NewLine)            '国コード
            sbSql.Append(" @SelectLang, " & Environment.NewLine)           '言語コード
            sbSql.Append(" @FullKataban, " & Environment.NewLine)          'フル形番
            sbSql.Append(" @SpecNo, " & Environment.NewLine)               '仕様書番号
            sbSql.Append(" @CheckKbn, " & Environment.NewLine)             'チェック区分
            sbSql.Append(" @PlaceCd, " & Environment.NewLine)              '出荷場所コード
            sbSql.Append(" @Nouki, " & Environment.NewLine)                '標準納期
            sbSql.Append(" @TekiyoKosu, " & Environment.NewLine)           '適用個数
            sbSql.Append(" @Kosu, " & Environment.NewLine)                 '販売数量個数
            sbSql.Append(" @ELInfo, " & Environment.NewLine)               'EL品情報
            sbSql.Append(" @ListPrice, " & Environment.NewLine)            'ListPrice
            sbSql.Append(" @RegPrice, " & Environment.NewLine)             'RegPrice
            sbSql.Append(" @SsPrice, " & Environment.NewLine)              'SsPrice
            sbSql.Append(" @BsPrice, " & Environment.NewLine)              'BsPrice
            sbSql.Append(" @GsPrice, " & Environment.NewLine)              'GsPrice
            sbSql.Append(" @PsPrice, " & Environment.NewLine)              'PsPrice
            sbSql.Append(" @APrice, " & Environment.NewLine)               'APrice
            sbSql.Append(" @FobPrice) " & Environment.NewLine)             'FobPrice

            Try
                objCmd.CommandText = sbSql.ToString
                objCmd.Parameters.Add("@LogDateTime", SqlDbType.DateTime).Value = strSetValue(0).ToString
                objCmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = strSetValue(1).ToString
                objCmd.Parameters.Add("@CountryCD", SqlDbType.Char).Value = strSetValue(2).ToString
                objCmd.Parameters.Add("@SelectLang", SqlDbType.Char).Value = strSetValue(3).ToString
                objCmd.Parameters.Add("@FullKataban", SqlDbType.VarChar).Value = strSetValue(4).ToString

                objCmd.Parameters.Add("@SpecNo", SqlDbType.Char).Value = strSetValue(5).ToString

                objCmd.Parameters.Add("@CheckKbn", SqlDbType.Char).Value = strSetValue(11).ToString
                objCmd.Parameters.Add("@PlaceCd", SqlDbType.VarChar).Value = strSetValue(10).ToString
                objCmd.Parameters.Add("@Nouki", SqlDbType.VarChar).Value = strSetValue(8).ToString

                objCmd.Parameters.Add("@TekiyoKosu", SqlDbType.Int).Value = IIf(strSetValue(9).Equals(String.Empty), "0", strSetValue(9))

                objCmd.Parameters.Add("@Kosu", SqlDbType.VarChar).Value = strSetValue(7).ToString
                objCmd.Parameters.Add("@ELInfo", SqlDbType.VarChar).Value = strSetValue(6).ToString

                objCmd.Parameters.Add("@ListPrice", SqlDbType.Money).Value = IIf(strSetValue(12).Equals(String.Empty), "0", strSetValue(12))
                objCmd.Parameters.Add("@RegPrice", SqlDbType.Money).Value = IIf(strSetValue(13).Equals(String.Empty), "0", strSetValue(13))
                objCmd.Parameters.Add("@SsPrice", SqlDbType.Money).Value = IIf(strSetValue(14).Equals(String.Empty), "0", strSetValue(14))
                objCmd.Parameters.Add("@BsPrice", SqlDbType.Money).Value = IIf(strSetValue(15).Equals(String.Empty), "0", strSetValue(15))
                objCmd.Parameters.Add("@GsPrice", SqlDbType.Money).Value = IIf(strSetValue(16).Equals(String.Empty), "0", strSetValue(16))
                objCmd.Parameters.Add("@PsPrice", SqlDbType.Money).Value = IIf(strSetValue(17).Equals(String.Empty), "0", strSetValue(17))
                objCmd.Parameters.Add("@APrice", SqlDbType.Money).Value = IIf(strSetValue(18).Equals(String.Empty), "0", strSetValue(18))
                objCmd.Parameters.Add("@FobPrice", SqlDbType.Money).Value = IIf(strSetValue(19).Equals(String.Empty), "0", strSetValue(19))

                ''引数より値をセットする
                'For i As Integer = 0 To strSetValue.Length - 2
                '    Select Case objCmd.Parameters.Item(i).SqlDbType
                '        Case SqlDbType.Int, SqlDbType.Money
                '            If strSetValue(i) = Nothing OrElse strSetValue(i) = "" Then
                '                objCmd.Parameters.Item(i).Value = DBNull.Value
                '            Else
                '                objCmd.Parameters.Item(i).Value = CInt(strSetValue(i))
                '            End If
                '        Case Else
                '            objCmd.Parameters.Item(i).Value = strSetValue(i)
                '    End Select
                'Next
                objCmd.ExecuteNonQuery()
                objTrans.Commit()
            Catch e As Exception
                Try
                    objTrans.Rollback("InsertTran")
                Catch ex As SqlException
                    '何もしない
                End Try
            End Try
            InsertSearchKatabanLogTbl = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objTrans = Nothing
        End Try
    End Function

    ''' <summary>
    ''' ログファイル情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーID</param>
    ''' <param name="strSessionId">セッションID</param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strSelectLang">言語コード</param>
    ''' <param name="strPriceList">単価リスト</param>
    ''' <param name="dt_display">付加情報</param>
    ''' <param name="strKatabanInfo">形番情報</param>
    ''' <returns></returns>
    ''' <remarks>ログファイルの情報を編集し返却する</remarks>
    Private Function fncLogInfoGet(objCon As SqlConnection, ByVal strUserId As String, _
                                  ByVal strSessionId As String, ByVal strCountryCd As String, _
                                  ByVal strSelectLang As String, ByVal strPriceList(,) As String, _
                                  ByVal dt_display As DataTable, _
                                  Optional ByVal strKatabanInfo() As String = Nothing) As String
        Dim objKtbnStrc As New KHKtbnStrc
        Dim sbValue As New StringBuilder

        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Dim strSystemDatetime As String
        Dim strAryStdDlvDt() As String
        Dim strPrice() As String

        Try
            '引当形番情報取得
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, strUserId, strSessionId)

            ' システム日付設定
            strSystemDatetime = Format(Now, "yyyy/MM/dd HH:mm:ss")

            'データ編集
            sbValue.Append(strSystemDatetime & CdCst.Sign.Delimiter.Tab)                            'システム日付
            sbValue.Append(strUserId & CdCst.Sign.Delimiter.Tab)                                    'ユーザーID
            sbValue.Append(strCountryCd & CdCst.Sign.Delimiter.Tab)                                 '国コード
            sbValue.Append(strSelectLang & CdCst.Sign.Delimiter.Tab)                                '言語コード
            If strKatabanInfo Is Nothing Then
                sbValue.Append(objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Delimiter.Tab) 'フル形番
            Else
                sbValue.Append(strKatabanInfo(1) & CdCst.Sign.Delimiter.Tab)                        'フル形番
            End If
            sbValue.Append(objKtbnStrc.strcSelection.strSpecNo & CdCst.Sign.Delimiter.Tab)          '仕様書番号

            '付加情報
            Dim strKeylvl As String = "1024,512,256,128,64,32,16,8,4,2,1"
            Dim strLevel() As String = strKeylvl.Split(",")
            Dim dr_display() As DataRow = Nothing
            For inti As Integer = 0 To strLevel.Length - 1
                If dt_display Is Nothing Then Exit For
                dr_display = dt_display.Select("strLevel='" & CInt(strLevel(inti)) & "'")
                Select Case CInt(strLevel(inti))
                    Case 128 '中国輸出不可
                    Case 64 'EL品情報
                        If dr_display.Length > 0 Then
                            sbValue.Append(dr_display(0)("strValue").ToString & CdCst.Sign.Delimiter.Tab)
                        Else
                            sbValue.Append("" & CdCst.Sign.Delimiter.Tab)
                        End If
                    Case 32 '販売数量単位
                        If dr_display.Length > 0 Then
                            sbValue.Append(dr_display(0)("strValue").ToString & CdCst.Sign.Delimiter.Tab)
                        Else
                            sbValue.Append("" & CdCst.Sign.Delimiter.Tab)
                        End If
                    Case 16 '標準納期
                        If dr_display.Length > 0 Then
                            strAryStdDlvDt = Split(dr_display(0)("strValue").ToString, CdCst.Sign.Delimiter.Pipe)
                            For intLoopCnt2 = 0 To strAryStdDlvDt.Length - 1
                                sbValue.Append(strAryStdDlvDt(intLoopCnt2) & CdCst.Sign.Delimiter.Tab)
                            Next
                        Else
                            For intLoopCnt2 = 0 To 1
                                sbValue.Append("" & CdCst.Sign.Delimiter.Tab)
                            Next
                        End If
                    Case 8 '担当者情報
                        '表示のみ
                    Case 4 '在庫情報
                    Case 2 '出荷場所
                        If dr_display.Length > 0 Then
                            If strKatabanInfo Is Nothing Then
                                sbValue.Append(dr_display(0)("strValue").ToString & CdCst.Sign.Delimiter.Tab)
                            Else
                                sbValue.Append(strKatabanInfo(3) & CdCst.Sign.Delimiter.Tab)
                            End If
                        Else
                            sbValue.Append("" & CdCst.Sign.Delimiter.Tab)
                        End If
                    Case 1 '形番チェック区分
                        If dr_display.Length > 0 Then
                            If strKatabanInfo Is Nothing Then
                                sbValue.Append(dr_display(0)("strValue").ToString & CdCst.Sign.Delimiter.Tab)
                            Else
                                sbValue.Append(strKatabanInfo(2) & CdCst.Sign.Delimiter.Tab)
                            End If
                        Else
                            sbValue.Append("" & CdCst.Sign.Delimiter.Tab)
                        End If
                End Select
            Next

            '単価リスト
            ReDim strPrice(8)
            For intLoopCnt1 = 1 To UBound(strPriceList)
                Select Case strPriceList(intLoopCnt1, 4)
                    Case CdCst.UnitPrice.ListPrice
                        strPrice(1) = strPriceList(intLoopCnt1, 2)
                    Case CdCst.UnitPrice.RegPrice
                        strPrice(2) = strPriceList(intLoopCnt1, 2)
                    Case CdCst.UnitPrice.SsPrice
                        strPrice(3) = strPriceList(intLoopCnt1, 2)
                    Case CdCst.UnitPrice.BsPrice
                        strPrice(4) = strPriceList(intLoopCnt1, 2)
                    Case CdCst.UnitPrice.GsPrice
                        strPrice(5) = strPriceList(intLoopCnt1, 2)
                    Case CdCst.UnitPrice.PsPrice
                        strPrice(6) = strPriceList(intLoopCnt1, 2)
                    Case CdCst.UnitPrice.APrice
                        strPrice(7) = strPriceList(intLoopCnt1, 2)
                    Case CdCst.UnitPrice.FobPrice
                        strPrice(8) = strPriceList(intLoopCnt1, 2)
                End Select
            Next
            For intLoopCnt1 = 1 To strPrice.Length - 1
                If strPrice(intLoopCnt1) = Nothing Then
                    sbValue.Append("" & CdCst.Sign.Delimiter.Tab)
                Else
                    sbValue.Append(strPrice(intLoopCnt1) & CdCst.Sign.Delimiter.Tab)
                End If
            Next
            sbValue.Append(vbCrLf)

            '戻り値設定
            fncLogInfoGet = sbValue.ToString
        Catch ex As Exception
            WriteErrorLog("E001", ex)
            fncLogInfoGet = ""
        Finally
            sbValue = Nothing
            objKtbnStrc = Nothing
        End Try
    End Function

End Class
