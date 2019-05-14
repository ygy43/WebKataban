Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class TypeBLL

    Private Property dllType As New TypeDAL


    ''' <summary>
    ''' 機種選択画面の検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks></remarks>
    Public Function fncSearch(ByVal objCon As SqlConnection, ByVal seriesSearch As KHSeriesSearch, ByVal lstWhereSeries As ArrayList) As DataSet

        Dim strRange As String = String.Empty
        Dim strMinFlag As String = String.Empty
        Dim strKataban As String = Nothing
        Dim strPRCKataban As String = String.Empty
        Dim dsResult As New DataSet

        Try
            With seriesSearch
                If .strSearchDvValue Is Nothing OrElse Len(Trim(.strSearchDvValue)) = 0 OrElse _
                    .strSrsKataValue Is Nothing OrElse Len(Trim(.strSrsKataValue)) = 0 OrElse _
                    .strLangCdValue Is Nothing OrElse Len(Trim(.strLangCdValue)) = 0 Then
                Else
                    If .intRangeValue = 0 Then
                        strRange = Nothing
                    Else
                        strRange = " TOP " & .intRangeValue & " "
                    End If
                    If .strMinKatabanValue Is Nothing Or Len(Trim(.strMinKatabanValue)) = 0 Then
                        strMinFlag = "0"
                    Else
                        strMinFlag = "1"
                    End If

                    If .dsKatabanValue IsNot Nothing Then
                        .dsKatabanValue.Clear()
                        .dsKatabanValue = Nothing
                    End If

                    '条件により検索
                    Select Case .strSearchDvValue
                        Case "0" '機種
                            dsResult = dllType.fncSelectBySeries(objCon, strRange, strMinFlag, .strSrsKataValue, .strCountryCdValue, seriesSearch, lstWhereSeries)
                        Case "1" 'フル形番
                            dsResult = dllType.fncSelectByFullKataban(objCon, strRange, strMinFlag, .strCountryCdValue, seriesSearch, lstWhereSeries)

                            If Not IsDBNull(dsResult.Tables("KatabanTbl")) Then
                                dsResult = subEditData(seriesSearch)
                            End If
                        Case "2" '全て
                            dsResult = dllType.fncSelectByAll(objCon, strRange, strMinFlag, .strSrsKataValue, .strCountryCdValue, seriesSearch, lstWhereSeries)

                            If My.Settings.ShiireSearchMode = 1 Then
                                dsResult = M_KatabanDAL.LoadCommonDB_GetMKatabanByConditionsKatahiki(.strSrsKataValue, seriesSearch, dsResult)
                            End If

                            If Not IsDBNull(dsResult.Tables("KatabanTbl")) Then
                                dsResult = subEditData(seriesSearch)
                            End If

                        Case "3"   '仕入品
                            dsResult = M_KatabanDAL.LoadCommonDB_GetMKatabanByConditionsKatahiki(.strSrsKataValue, seriesSearch, Nothing)

                    End Select
                End If
            End With

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dsResult
    End Function

    ''' <summary>
    ''' 検索データ編集
    ''' </summary>
    ''' <remarks></remarks>
    Private Function subEditData(ByVal seriesSearch As KHSeriesSearch) As DataSet
        Dim strSystem As String
        Dim strParts As String
        Dim strFor As String

        Try
            'コード文字列取得
            strSystem = ClsCommon.fncGetMsg(seriesSearch.strLangCdValue, "I0040")
            strParts = ClsCommon.fncGetMsg(seriesSearch.strLangCdValue, "I0030")
            strFor = ClsCommon.fncGetMsg(seriesSearch.strLangCdValue, "I0050")

            '表示形番(説明)編集
            With seriesSearch.dsKatabanValue.Tables("KatabanTbl")
                For intI As Integer = 0 To seriesSearch.dsKatabanValue.Tables("KatabanTbl").Rows.Count - 1
                    If CStr(.Rows(intI).Item("division")) = CdCst.RetrievalDiv.Full Then
                        If CInt(.Rows(intI).Item("kataban_check_div")) < 4 Then
                            If IsDBNull(.Rows(intI).Item("model_nm")) _
                                OrElse .Rows(intI).Item("model_nm").Equals(String.Empty) Then

                                If IsDBNull(.Rows(intI).Item("parts_nm")) _
                                OrElse .Rows(intI).Item("parts_nm").Equals(String.Empty) Then
                                    .Rows(intI).Item("disp_name") = "(" & strSystem & ")"
                                Else
                                    .Rows(intI).Item("disp_name") = .Rows(intI).Item("parts_nm")
                                End If
                            Else
                                If IsDBNull(.Rows(intI).Item("parts_nm")) _
                                OrElse .Rows(intI).Item("parts_nm").Equals(String.Empty) Then
                                    .Rows(intI).Item("disp_name") = .Rows(intI).Item("model_nm")
                                Else
                                    .Rows(intI).Item("disp_name") = .Rows(intI).Item("model_nm") & "(" & .Rows(intI).Item("parts_nm") & ")"
                                End If
                            End If
                        Else
                            If IsDBNull(.Rows(intI).Item("model_nm")) _
                                OrElse .Rows(intI).Item("model_nm").Equals(String.Empty) Then

                                If IsDBNull(.Rows(intI).Item("parts_nm")) _
                                OrElse .Rows(intI).Item("parts_nm").Equals(String.Empty) Then
                                    .Rows(intI).Item("disp_name") = "(" & strSystem & ")"
                                Else
                                    .Rows(intI).Item("disp_name") = strParts & "(" & .Rows(intI).Item("parts_nm") & ")"
                                End If
                            Else
                                If IsDBNull(.Rows(intI).Item("parts_nm")) _
                                OrElse .Rows(intI).Item("parts_nm").Equals(String.Empty) Then
                                    .Rows(intI).Item("disp_name") = strParts & "(" & .Rows(intI).Item("model_nm") & ")"
                                Else
                                    .Rows(intI).Item("disp_name") = strParts & "(" & .Rows(intI).Item("model_nm") & strFor & "(" & .Rows(intI).Item("parts_nm") & "))"
                                End If
                            End If
                        End If
                    End If
                Next
            End With


        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return seriesSearch.dsKatabanValue
    End Function

    ''' <summary>
    ''' 引当シリーズ形番追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strSrsKataban">シリーズ形番</param>
    ''' <param name="strKeyKataban">キー形番</param>
    ''' <param name="strGoodsNm">商品名</param>
    ''' <remarks></remarks>
    Public Sub subInsertSelSrsKtbnMdl(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                      ByVal strSessionId As String, ByVal strSrsKataban As String, _
                                      ByVal strKeyKataban As String, ByVal strGoodsNm As String, _
                                      ByVal strCurrencyCd As String)
        Try
            dllType.subInsertSelSrsKtbnMdl(objCon, strUserId, strSessionId, strSrsKataban, strKeyKataban, strGoodsNm, strCurrencyCd)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当シリーズ形番追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strFullKataban">フル形番</param>
    ''' <param name="strGoodsNm">商品名</param>
    ''' <remarks></remarks>
    Public Sub subInsertSelSrsKtbnFull(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                    ByVal strSessionId As String, ByVal strFullKataban As String, _
                                    ByVal strGoodsNm As String, ByVal strCurrencyCd As String)
        Try
            dllType.subInsertSelSrsKtbnFull(objCon, strUserId, strSessionId, strFullKataban, strGoodsNm, strCurrencyCd)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' 引当シリーズ形番追加処理(仕入品）
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strFullKataban">フル形番</param>
    ''' <param name="strGoodsNm">商品名</param>
    ''' <remarks></remarks>
    Public Sub subInsertSelSrsKtbnShiire(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                    ByVal strSessionId As String, ByVal strFullKataban As String, _
                                    ByVal strGoodsNm As String, ByVal strCurrencyCd As String)
        Try
            dllType.subInsertSelSrsKtbnShiire(objCon, strUserId, strSessionId, strFullKataban, strGoodsNm, strCurrencyCd)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' 引当シリーズ形番削除処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <remarks></remarks>
    Public Sub subDeleteSelKtbnInfo(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                          ByVal strSessionId As String)
        Try
            dllType.subDeleteSelKtbnInfo(objCon, strUserId, strSessionId)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

End Class
