Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports WebKataban.CdCst

Public Class TankaISOBLL

    ''' <summary>
    ''' 小数点区分取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncDecPointDivSelect(ByVal objCon As SqlConnection, strUserID As String) As String
        Dim strResult As String = String.Empty
        Dim ISOTanka As New TankaISODAL

        Try
            strResult = ISOTanka.fncDecPointDivSelect(objCon, strUserID)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return strResult
    End Function

    ''' <summary>
    ''' 選択情報の取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strSession"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSQL_GetCompData(objCon As SqlConnection, strUserID As String, strSession As String) As DataTable
        Dim dtResult As New DataTable
        Dim ISOTanka As New TankaISODAL

        Try
            dtResult = ISOTanka.fncSQL_GetCompData(objCon, strUserID, strSession)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dtResult
    End Function

End Class
