Imports WebKataban.ClsCommon

Public Class KHPlace
    Public Property dtPlace As DataTable

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        dtPlace = fncCreateTableByColumnNames(New List(Of String) From {"PlaceID", "PlaceName"})
    End Sub

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal lstPlaceIDs As ArrayList, ByVal lstPlaceNames As ArrayList)
        Me.New()
        Try
            If lstPlaceIDs.Count = lstPlaceNames.Count Then
                For inti As Integer = 0 To lstPlaceIDs.Count - 1
                    Dim dr As DataRow = dtPlace.NewRow

                    dr.Item("PlaceID") = lstPlaceIDs(inti)
                    dr.Item("PlaceName") = lstPlaceNames(inti)
                Next
            End If

        Catch ex As Exception

        End Try

    End Sub
End Class
