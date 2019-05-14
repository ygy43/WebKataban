Public Class KHAccPriceModel
#Region "プロパティ"
    Public Property strKatabanCheckDiv As String
    Public Property strPlaceCd As String
    Public Property strCostCalcNo As String
    Public Property intListPrice As Decimal
    Public Property intRegPrice As Decimal
    Public Property intSsPrice As Decimal
    Public Property intBsPrice As Decimal
    Public Property intGsPrice As Decimal
    Public Property intPsPrice As Decimal
    Public Property strCheckDiv As String
    Public Property strCurrency As String
    Public Property strMadeCountry As String
    Public Property strStorageLocation As String
    Public Property strEvaluationType As String

#End Region
    
    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        strKatabanCheckDiv = String.Empty
        strPlaceCd = String.Empty
        strCostCalcNo = Nothing
        intListPrice = 0
        intRegPrice = 0
        intSsPrice = 0
        intBsPrice = 0
        intGsPrice = 0
        intPsPrice = 0
        strCheckDiv = String.Empty
        strCurrency = String.Empty
        strMadeCountry = String.Empty
        strStorageLocation = String.Empty
        strEvaluationType = String.Empty

    End Sub
End Class
