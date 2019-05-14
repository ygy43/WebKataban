Public Class KHSessionInfo

    'ユーザー情報
    <Serializable()> Public Structure UserInfo
        Public UserId As String                                     'ユーザーＩＤ
        Public BaseCd As String                                     '拠点コード   
        Public CountryCd As String                                  '国コード
        Public OfficeCd As String                                   '営業所コード
        Public PersonCd As String                                   '担当者コード
        Public MailAddress As String                                'メールアドレス
        Public LanguageCd As String                                 '言語コード
        Public CurrencyCd As String                                 '通貨コード
        Public EditDiv As String                                    '編集区分
        Public UserClass As String                                  'ユーザー種別
        Public PriceDispLvl As Integer                              '価格表示レベル
        Public AddInformationLvl As Integer                         '付加情報レベル
        Public UseFunctionLvl As Integer                            '利用機能レベル
        Public TnkDispCnt As Integer                                '単価画面表示回数
    End Structure

    '選択情報
    <Serializable()> Public Structure LoginInfo
        Public SessionId As String                                  'セッションＩＤ
        Public SelectLang As String                                 '選択言語
    End Structure

    'フレーム情報
    <Serializable()> Public Structure FrameInfo
        Public TopUrl As String                                     'Top URL
        Public MainUrl As String                                    'Main URL
        Public ButtonUrl As String                                  'Button URL
        Public BottomUrl As String                                  'Bottom URL
        Public Rows As String                                       'Rows
    End Structure

    '受注EDI引数情報
    <Serializable()> Public Structure EdiInfo
        Public KeyInfo As String                                    'Reciveキー情報
    End Structure

    '受注EDI送信情報
    <Serializable()> Public Structure SendEdi
        Public FullKataban As String
        Public CheckKubun As String
        Public PlaceCode As String
        Public PriceTeika As Decimal
        Public PriceTouroku As Decimal
        Public PriceSS As Decimal
        Public PriceBS As Decimal
        Public PriceGS As Decimal
        Public PricePS As Decimal
        Public PriceNet As Decimal
        Public Currency As String
        Public KisyuCode As String
        Public ManifoldSpecData As String
        Public ElKubun As String
        Public SalesUnit As String
        Public SapBaseUnit As String
        Public QuantityPerSalesUnit As String
        Public OrderLot As String
    End Structure

End Class
