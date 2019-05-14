Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class LabelBLL
    ''' <summary>
    ''' ラベルデータ取り込み
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strLangCd">プログラムＩＤ</param>
    ''' <param name="strPgmId">言語コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSelectPageAllLabels(ByVal objCon As SqlConnection, ByVal strLangCd As String, ByVal strPgmId As String) As DataTable
        Dim dt As New DataTable
        Dim dtResult As New DataTable

        Try
            '結果テーブルを作成
            Dim dc As New DataColumn("label_content")
            dtResult.Columns.Add(dc)
            dc = New DataColumn("label_div")
            dtResult.Columns.Add(dc)
            dc = New DataColumn("label_seq")
            dtResult.Columns.Add(dc)

            'ラベル情報の取得
            dt = LabelDAL.fncSelectPageAllLabels(objCon, strLangCd, strPgmId)

            '結果テーブルに保存
            If dt.Rows.Count > 0 Then
                Dim drResult As DataRow = Nothing

                For Each dr In dt.Rows
                    drResult = dtResult.NewRow
                    drResult("label_content") = dr("label_content")
                    drResult("label_div") = dr("label_div")
                    drResult("label_seq") = dr("label_seq")
                    dtResult.Rows.Add(drResult)
                Next
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dtResult
    End Function

    ''' <summary>
    ''' ラベルデータの検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strPgmId">プログラムＩＤ</param>
    ''' <param name="strLangCd">言語コード </param>
    ''' <param name="strLblDiv"> ラベル区分       L:画面ラベル  B:ボタンラベル</param>
    ''' <param name="intLblSeq">ラベル番号</param>
    ''' <returns></returns>
    ''' <remarks>ラベルマスタより引数.プログラムID、言語コード、ラベル区分、ラベル番号に該当するラベルを取得する</remarks>
    Public Shared Function fncSelectLabelById(ByVal objCon As SqlConnection, ByVal strPgmId As String, _
                                  ByVal strLangCd As String, ByVal strLblDiv As String, ByVal intLblSeq As Integer) As String
        Dim strResult As String = String.Empty

        Try
            strResult = LabelDAL.fncSelectLabelById(objCon, strPgmId, strLangCd, strLblDiv, intLblSeq)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return strResult
    End Function
End Class
