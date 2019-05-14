''' <summary>
''' 全ての情報
''' </summary>
''' <remarks></remarks>
<Serializable()>
Public Class KHInfoModel

#Region "プロパティ"
    Public Property strUserId As String                          'ユーザーＩＤ
    Public Property strSessionId As String                       'セッションＩＤ
    Public Property strDivision As String                        '区分

    Public Property strSeriesKataban As String                   'シリーズ形番
    Public Property strKeyKataban As String                      'キー形番
    Public Property strHyphen As String                          'ハイフン
    Public Property strPriceNo As String                         '価格積上番号
    Public Property strSpecNo As String                          '仕様書番号
    Public Property strFullKataban As String                     'フル形番
    Public Property strGoodsNm As String                         '商品名    
    Public Property strKatabanCheckDiv As String                 '形番チェック区分
    Public Property strPlaceCd As String                         '出荷場所コード
    Public Property strCostCalcNo As String                      '原価積算No.
    Public Property intListPrice As Decimal                      '定価
    Public Property intRegPrice As Decimal                       '登録店価格
    Public Property intSsPrice As Decimal                        'SS店価格
    Public Property intBsPrice As Decimal                        'BS店価格
    Public Property intGsPrice As Decimal                        'GS店価格
    Public Property intPsPrice As Decimal                        'PS店価格
    Public Property decFactor As Decimal                         '掛率
    Public Property intUnitPrice As Decimal                      '単価
    Public Property intAmount As Integer                         '数量
    Public Property strRodEndOption As String                    'ロッド先端特注
    Public Property strOtherOption As String                     'オプション外
    Public Property strPositionOption As String                  '簡易仕様書設置位置
    Public Property strAuthorizationNo As String                   '特価決裁No.メッセージ

    Public Property strOpSymbol As String()                      'オプション記号    
    Public Property strOpElementDiv As String()                  '要素区分
    Public Property strOpStructureDiv As String()                '構成区分
    Public Property strOpAdditionDiv As String()                 '付加区分
    Public Property strOpHyphenDiv As String()                   'ハイフン区分
    Public Property strOpKtbnStrcNm As String()                  '形番構成名称
    Public Property strOpCountryDiv As Long()                    '生産国レベル    'Add by Zxjike 2013/09/04
    Public Property strOpIsoShowFlag As String()                'ISO画面表示形番フラグ

    Public Property strOpKataban As String()                     '形番
    Public Property strOpKatabanCheckDiv As String()             '形番チェック区分
    Public Property strOpPlaceCd As String()                     '出荷場所コード
    Public Property intOpListPrice As Decimal()                  '定価
    Public Property intOpRegPrice As Decimal()                   '登録店価格
    Public Property intOpSsPrice As Decimal()                    'SS店価格
    Public Property intOpBsPrice As Decimal()                    'BS店価格
    Public Property intOpGsPrice As Decimal()                    'GS店価格
    Public Property intOpPsPrice As Decimal()                    'PS店価格
    Public Property decOpamount As Decimal()                     '部品数量    'add by Zxjike 2012/11/20

    Public Property strModelNo As String                         '機種番号
    Public Property strWiringSpec As String                      '配線仕様区分
    Public Property decDinRailLength As Decimal                  'レール長さ

    Public Property strAttributeSymbol As String()               '属性記号
    Public Property strOptionKataban As String()                 'オプション形番
    Public Property strCXAKataban As String()                    'CX-A形番
    Public Property strCXBKataban As String()                    'CX-B形番
    Public Property strPositionInfo As String()                  '設置位置情報
    Public Property intQuantity As Double()                      '使用数

    Public Property strFullManiKataban As String                 'strFullKataban+6桁位置情報

    Public Property strRodEndWFStdVal As String                  'ロッド先端特注WF標準寸法
    Public Property strCurrency As String                        '通貨        'add by Zxjike 2013/05/15
    Public Property strMadeCountry As String                     '生産国      'add by Zxjike 2013/06/07

    Public Property strOutofOpCountryDiv As Long()               'オプション外生産国レベル判定用変数  '2017/04/07  追加
    Public Property strSalesUnit As String                       '販売単位
    Public Property strQuantityPerSalesUnit As String            '入数
    Public Property strSapBaseUnit As String                      'SAP基本数量単位
    Public Property strOrderLot As String                        '発注ロット
    Public Property strStorageLocation As String                 '保管場所
    Public Property strEvaluationType As String                  '評価タイプ
    Public Property strUserKataban As String                     'ユーザー形番

    Public Property dt_vol As DataTable
    Public Property dt_Stroke As DataTable
#End Region


    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal strUserId As String, ByVal strSessionId As String)
        'ユーザーＩＤ
        Me.strUserId = strUserId
        'セッションＩＤ
        Me.strSessionId = strSessionId
        strDivision = ""

        strSeriesKataban = ""
        strKeyKataban = ""
        strHyphen = ""
        strPriceNo = ""
        strSpecNo = ""
        strFullKataban = ""
        strGoodsNm = ""
        strKatabanCheckDiv = ""
        strPlaceCd = ""
        strCostCalcNo = ""
        intListPrice = 0
        intRegPrice = 0
        intSsPrice = 0
        intBsPrice = 0
        intGsPrice = 0
        intPsPrice = 0
        decFactor = 0
        intUnitPrice = 0
        intAmount = 0
        strRodEndOption = ""
        strOtherOption = ""
        strPositionOption = ""

        ReDim strOpSymbol(0)
        ReDim strOpElementDiv(0)
        ReDim strOpStructureDiv(0)
        ReDim strOpAdditionDiv(0)
        ReDim strOpHyphenDiv(0)
        ReDim strOpKtbnStrcNm(0)

        ReDim strOpKataban(0)
        ReDim strOpKatabanCheckDiv(0)
        ReDim strOpPlaceCd(0)
        ReDim intOpListPrice(0)
        ReDim intOpRegPrice(0)
        ReDim intOpSsPrice(0)
        ReDim intOpBsPrice(0)
        ReDim intOpGsPrice(0)
        ReDim intOpPsPrice(0)

        strModelNo = ""
        strWiringSpec = ""
        decDinRailLength = 0

        ReDim strAttributeSymbol(0)
        ReDim strOptionKataban(0)
        ReDim strCXAKataban(0)
        ReDim strCXBKataban(0)
        ReDim strPositionInfo(0)
        ReDim intQuantity(0)
        ReDim decOpamount(0)
        ReDim strOpCountryDiv(0)
        strCurrency = ""
        strMadeCountry = ""
        dt_vol = New DataTable
        dt_Stroke = New DataTable

        ReDim strOutofOpCountryDiv(0)

    End Sub
End Class
