Public Module KHCodeConstants

    Public Structure CdCst

        Public Shared strPipe As String = CdCst.Sign.Delimiter.Pipe
        Public Shared strComma As String = CdCst.Sign.Delimiter.Comma                       'カンマ
        Public Shared strTube As String = "ﾁﾕ-ﾌﾞﾇｷｸﾞﾌﾖｳ"                            'チューブ抜具不要
        Public Shared strExcelTmpFileName As String = "Manifold.xlsx"                'Excel一時ファイル名

        'セッション情報
        Public Structure SessionInfo
            Private strDummy As String
            Public Const TimeOut = 20                                               'タイムアウト時間
            'キー情報
            Public Structure Key
                Private strDummy As String
                Public Const UserInfo = "UserInfo"                                  'ユーザー情報
                Public Const LoginInfo = "LoginInfo"                                'ログイン情報
                Public Const LoginError = "LoginError"                              'ログインエラー
                Public Const Exception = "Exception"                                'エラー情報(System.Exception)
                Public Const ErrorCode = "ErrorCode"                                'エラーコード
                Public Const FrameInfo = "FrameInfo"                                'KHFrameのURL情報
                '受注EDI連携
                Public Const EdiInfo = "EdiInfo"                                    '受注EDI受信情報
                Public Const SendEdi = "SendEdi"                                    '受注EDI送信情報
                Public Const HikiateFlg = "HikiateFlg"                              'EDI引当フラグ
            End Structure
        End Structure

        '画面ID
        Public Shared strPageIDs() As String = {"WebUC_Login", "WebUC_Menu",
                                                "WebUC_Type", "WebUC_Youso",
                                                "WebUC_Tanka", "WebUC_RodEndOrder",
                                                "WebUC_RodEnd", "WebUC_OutOfOption",
                                                "WebUC_Stopper", "WebUC_Motor", "WebUC_PriceCopy",
                                                "WebUC_Siyou", "WebUC_ISOTanka",
                                                "WebUC_Error", "WebUC_Master",
                                                "WebUC_KatOut", "WebUC_KatSep", "WebUC_100Test",
                                                "WebUC_PriceDetail", "WebUC_TypeAnonymous"}

        '言語区分
        Public Structure LanguageCd
            Private strDummy As String
            Public Const DefaultLang = "en"                                         'デフォルト言語(英語)
            Public Const SimplifiedChinese = "zh"                                   '日本語
            Public Const TraditionalChinese = "tw"                                  '簡体字
            Public Const Japanese = "ja"                                            '繁体字
            Public Const Korean = "ko"                                              '韓国語
        End Structure

        'EL品判定区分
        Public Structure ELDiv
            Private strDummy As String
            Public Const Yes = "Y"                                                  'EL品
            Public Const No = "N"                                                   'EL品以外
        End Structure

        'ハイフン区分
        Public Structure HyphenDiv
            Private strDummy As String
            Public Const Necessary = "1"                                            '有り
            Public Const Unnecessary = "0"                                          '無し
        End Structure

        '編集区分
        Public Structure EditDiv
            Private strDummy As String
            Public Const Normal = "0"                                               'カンマ編集
            Public Const Other = "1"                                                'ドット編集
        End Structure

        'ラベル
        Public Structure Lbl
            Private strDummy As String
            '区分
            Public Structure Division
                Private strDummy As String
                Public Const Label = "L"                                            'ラベル
                Public Const Button = "B"                                           'ボタンラベル
                Public Const Radio = "R"                                            'ラジオボタンラベル
                Public Const Title = "T"                                            'タイトルラベル
            End Structure
            '名称
            Public Structure Name
                Private strDummy As String
                Public Const Label = "Label"                                        'ラベル
                Public Const Button = "Button"                                      'ボタンラベル
                Public Const RadioButton = "RadioButton"                            'ボタンラベル
                Public Const Title = "Title"                                        'タイトルラベル
                Public Const RadioButtonList = "RadioButtonList"                    'ラジオボタンラベル
            End Structure
        End Structure

        'キャッシュ情報
        Public Structure CacheInfo
            Private strDummy As String
            'キー情報
            Public Structure Key
                Private strDummy As String
                Public Const MenuClass = "MenuClass"                                'メニュー分類
                Public Const MenuContent = "MenuContent"                            'メニュー内容
                Public Const RateMstMnt = "RateMstMnt"                              '掛率情報メンテナンス
                Public Const UserMstMnt1 = "UserMstMnt1"                            'ユーザー情報メンテナンス(固定)
                Public Const UserMstMnt2 = "UserMstMnt2"                            'ユーザー情報メンテナンス(ｽｸﾛｰﾙ)
            End Structure
        End Structure

        '稼動状況
        Public Structure OpeState
            Private strDummy As String
            Public Const Stopping = "0"                                             '停止中
            Public Const Operating = "1"                                            '稼動中
            Public Const Trouble = "E"                                              'トラブル
        End Structure

        'WucSpecInfo 項目プロパティのアイテム
        Public Structure TblSpecItem
            Private strDummy As String
            Public Const ProdNm = "ProdNm"                                          '品名
            Public Const ItemCnt = "ItemCnt"                                        '項目数
            Public Const ItemDiv = "ItemDiv"                                        '項目区分
        End Structure

        'WucSpecInfo 項目内容プロパティのアイテム
        Public Structure TblSpecContent
            Private strDummy As String
            Public Const SeqNo = "SeqNo"                                            '連番
            Public Const ItemNm = "ItemNm"                                          '項目名称
            Public Const ItemValue = "ItemValue"                                    '項目内容値
            Public Const Other1 = "Other1"                                          'その他１
            Public Const Other2 = "Other2"                                          'その他２
        End Structure

        'WucSpecInfo 仕様情報
        Public Structure SpecInfoItem
            Private strDummy As String
            Public Const Kataban = "Kataban"                                        '形番
            Public Const JointCxA = "JointCxA"                                      '継手CXA
            Public Const JointCxB = "JointCxB"                                      '継手CXB
            Public Const UseCnt = "UseCnt"                                          '使用数
        End Structure

        'ユーザー種別
        Public Structure UserClass
            Private strDummy As String
            'Public Const DmGeneralUser = "1"                                        '国内一般ユーザー
            'Public Const OsGeneralUser = "2"                                        '海外一般ユーザー
            Public Const DmAgentRs = "11"                                           '国内代理店(登録店)
            Public Const DmAgentSs = "12"                                           '国内代理店(ＳＳ店)
            Public Const DmAgentBs = "13"                                           '国内代理店(ＢＳ店)
            Public Const DmAgentGs = "14"                                           '国内代理店(ＧＳ店)
            Public Const DmAgentPs = "15"                                           '国内代理店(ＰＳ店)
            Public Const OsAgentCs = "16"                                           '海外代理店(契約店)
            Public Const OsAgentLs = "17"                                           '海外代理店(契約店、E-con) Add by Zxjike 2014/03/10
            Public Const DmSalesOffice = "21"                                       '国内営業所
            'Public Const OsSelComLocEmpBiz = "22"                                   '海外販社(現地採用者(営業))
            'Public Const OsSelComLocEmpRetMnger = "23"                              '海外販社(現地採用者(販管))
            'Public Const OsSelComLocEmp = "24"                                      '海外販社(現地採用者)
            'Public Const OsSelComJpnRep = "25"                                      '海外販社(日本人駐在員)
            'Public Const EngineeringDep = "31"                                      '技術部
            Public Const OsSalesDep = "41"                                          '海外営業部
            Public Const OsSalesDepMnger = "42"                                     '海外営業部(管理者)
            Public Const BizHeadquarters = "45"                                     '営業本部
            Public Const BizHeadquartersMnger = "46"                                '営業本部(管理者)
            Public Const InfoSysForce = "91"                                        '情報システム部
            Public Const InfoSysForceMnger = "95"                                   '情報システム部(管理者)
            Public Const InfoSysForceSysAdmin = "99"                                '情報システム部(システム管理者)
        End Structure

        'DB情報
        Public Structure DB
            Private strDummy As String
            'スキーマ
            Public Structure Scehma
                Private strDummy As String
                Public Const sales = "sales"                                        '形番引当スキーマ
            End Structure
            'ストアドプロシージャ
            Public Structure SPL
                Private strDummy As String
                Public Const KHLoginRec = "KHLoginRec"                                              'ログイン
                Public Const KHLogoutRec = "KHLogoutRec"                                            'ログオフ
                Public Const KHUserPasswdChg = "KHUserPasswdChg"                                    'パスワード変更
                Public Const KHSelSrsKtbnMdlIns = "KHSelSrsKtbnMdlIns"                              '引当シリーズ形番追加(機種)
                Public Const KHSelSrsKtbnFullIns = "KHSelSrsKtbnFullIns"                            '引当シリーズ形番追加(フル形番)
                Public Const KHSelSrsKtbnShiireIns = "KHSelSrsKtbnShiireIns"                        '引当シリーズ形番追加(仕入品)
                Public Const KHSelSrsKtbnPriUpd = "KHSelSrsKtbnPriUpd"                              '引当シリーズ形番更新
                Public Const KHSelSrsKtbnFullKtbnUpd = "KHSelSrsKtbnFullKtbnUpd"                    '引当シリーズ形番更新
                Public Const KHSelSrsKtbnOptionUpd = "KHSelSrsKtbnOptionUpd"                        '引当シリーズ形番更新
                Public Const KHSelSrsKtbnDel = "KHSelSrsKtbnDel"                                    '引当シリーズ形番削除
                Public Const KHSelKtbnStrcIns = "KHSelKtbnStrcIns"                                  '引当形番構成追加
                Public Const KHSelKtbnStrcUpd = "KHSelKtbnStrcUpd"                                  '引当形番構成更新
                Public Const KHSelKtbnStrcDel = "KHSelKtbnStrcDel"                                  '引当形番構成削除
                Public Const KHSelKtbnInfoDel = "KHSelKtbnInfoDel"                                  '引当情報削除
                Public Const KHSelAccPrcStrcIns = "KHSelAccPrcStrcIns"                              '引当積上単価構成追加
                Public Const KHSelAccPrcStrcDel = "KHSelAccPrcStrcDel"                              '引当積上単価構成削除
                Public Const KHErrorContentIns = "KHErrorContentIns"                                'エラー内容追加
                Public Const KHSelSpecIns = "KHSelSpecIns"                                          '引当仕様書追加
                Public Const KHSelSpecDel = "KHSelSpecDel"                                          '引当仕様書削除
                Public Const KHSelSpecStrcIns = "KHSelSpecStrcIns"                                  '引当仕様書構成追加
                Public Const KHSelSpecStrcDel = "KHSelSpecStrcDel"                                  '引当仕様書構成削除
                Public Const KHRateMstMntIns = "KHRateMstMntIns"                                    '掛率情報追加
                Public Const KHRateMstOutEffDtUpd = "KHRateMstOutEffDtUpd"                          '掛率情報更新(失効日)
                Public Const KHRateMstNumAreaUpd = "KHRateMstNumAreaUpd"                            '掛率情報更新(数値ｴﾘｱ)
                Public Const KHRateMstSeqNoUpd = "KHRateMstSeqNoUpd"                                '掛率更新削除(順序)
                Public Const KHRateMstMntDel = "KHRateMstMntDel"                                    '掛率情報削除
                Public Const KHCurrencyExcRateMstMntIns = "KHCurrencyExcRateMstMntIns"              '為替率情報追加
                Public Const KHCurrencyExcRateMstOutEffDtUpd = "KHCurrencyExcRateMstOutEffDtUpd"    '為替率情報更新(失効日)
                Public Const KHCurrencyExcRateMstSeqNoUpd = "KHCurrencyExcRateMstSeqNoUpd"          '為替率情報更新(順序)
                Public Const KHCurrencyExcRateMstDataAreaUpd = "KHCurrencyExcRateMstDataAreaUpd"    '為替率情報更新(ﾃﾞｰﾀｴﾘｱ)
                Public Const KHCurrencyExcRateMstMntDel = "KHCurrencyExcRateMstMntDel"              '為替率情報削除
                Public Const KHCountryItemMstMntIns = "KHCountryItemMstMntIns"                      '国別生産品情報追加
                Public Const KHCountryItemMstOutEffDtUpd = "KHCountryItemMstOutEffDtUpd"            '国別生産品情報更新(失効日)
                Public Const KHCountryItemMstSeqNoUpd = "KHCountryItemMstSeqNoUpd"                  '国別生産品情報更新(順序)
                Public Const KHCountryItemMstDataAreaUpd = "KHCountryItemMstDataAreaUpd"            '国別生産品情報更新(ﾃﾞｰﾀｴﾘｱ)
                Public Const KHCountryItemMstMntDel = "KHCountryItemMstMntDel"                      '国別生産品情報削除
                Public Const KHUserMstMntIns = "KHUserMstMntIns"                                    'ユーザー情報追加
                Public Const KHUserMstOutEffDtUpd = "KHUserMstOutEffDtUpd"                          'ユーザー情報更新(失効日)
                Public Const KHUserMstSeqNoUpd = "KHUserMstSeqNoUpd"                                'ユーザー情報更新(順序)
                Public Const KHUserMstDataAreaUpd = "KHUserMstDataAreaUpd"                          'ユーザー情報更新(ﾃﾞｰﾀｴﾘｱ)
                Public Const KHUserMstMntDel = "KHUserMstMntDel"                                     'ユーザー情報削除
                Public Const KHInfoMstMntIns = "KHInfoMstMntIns"                                    '情報マスタメンテ追加
                Public Const KHInfoMstMntUpd = "KHInfoMstMntUpd"                                    '情報マスタメンテ更新
                Public Const KHInfoMstMntDel = "KHInfoMstMntDel"                                    '情報マスタメンテ削除
                Public Const KHSelRodIns = "KHSelRodIns"                                            '引当仕様書追加
                Public Const KHSelRodDel = "KHSelRodDel"                                            '引当仕様書削除
                Public Const KHModelMstMntUpd = "KHModelMstMntUpd"                                  '機種情報更新
                Public Const MUserIns = "MUserIns"                                                  'Web認証システムユーザーマスタ追加
                Public Const MUserDel = "MUserDel"                                                  'Web認証システムユーザーマスタ削除
                Public Const MUserUpd = "MUserUpd"                                                  'Web認証システムユーザーマスタ更新
                Public Const KHSelOutOfOpIns = "KHSelOutOfOpIns"                                    '引当オプション外特注追加
                Public Const KHSelOutOfOpDel = "KHSelOutOfOpDel"                                    '引当オプション外特注削除
                Public Const KHUpdateHistoryIns = "KHUpdateHistoryIns"                              'マスタ変更履歴追加
            End Structure
        End Structure

        '要素区分
        Public Structure ElementDiv
            Private strDummy As String
            Public Const Voltage = "1"                                              '電圧
            Public Const Stroke = "3"                                               'ストローク
            Public Const Port = "5"                                                 '口径
            Public Const Coil = "6"                                                 'コイル
            Public Const VolPort = "7"                                              '口径(電圧用)
        End Structure

        '要素パターン
        Public Structure ElePattern
            Private strDummy As String
            Public Const All = "*"                                                  '全て
            Public Const Plural = "#"                                               '複数選択
        End Structure

        'オプション判定区分
        Public Structure JudgeDiv
            Private strDummy As String
            Public Const InSign = "I"                                               'IN
            Public Const OutSign = "O"                                              'OUT
            Public Const CondOr = "+"                                               'Or条件
            Public Const CondAnd = "*"                                              'And条件
            Public Const Equal = "EQ"                                               '含む
            Public Const NotEqual = "NE"                                            '含まない
        End Structure

        '構成区分
        Public Structure KtbnStructureDiv
            Private strDummy As String
            Public Const SelectCond = "1"                                           '選択条件
            Public Const SkipCond = "2"                                             'Skip条件
            Public Const PluralCond = "4"                                           '複数選択
        End Structure

        '電源種類
        Public Structure PowerSupply
            Private strDummy As String
            Public Const AC = "1"                                                   'AC電源
            Public Const DC = "2"                                                   'DC電源
            Public Const Div1 = "AC"                                                'AC電源
            Public Const Div2 = "DC"                                                'DC電源
            Public Const AC100V = "1"                                               'AC100V
            Public Const AC200V = "2"                                               'AC200V
            Public Const DC24V = "3"                                                'DC24V
            Public Const DC12V = "4"                                                'DC12V
            Public Const AC110V = "5"                                               'AC110V
            Public Const AC220V = "6"                                               'AC220V
            Public Const Const1 = "AC100V"                                          'AC100V
            Public Const Const2 = "AC200V"                                          'AC200V
            Public Const Const3 = "DC24V"                                           'DC24V
            Public Const Const4 = "DC12V"                                           'DC12V
            Public Const Const5 = "AC110V"                                          'AC110V
            Public Const Const6 = "AC220V"                                          'AC220V
            Public Const Const7 = "AC120V"                                          'AC120V
            Public Const Const8 = "AC240V"                                          'AC240V
        End Structure

        'その他電圧
        Public Structure OtherVoltage
            Private strDummy As String
            Public Const Japanese = "その他電圧"                                     'その他電圧(日本語)
            Public Const English = "OTHER"                                  '選択終了(英語)
        End Structure

        '価格積上区分
        Public Structure PriceAccDiv
            Private strDummy As String
            Public Const Domestic = "0"                                             '国内用(標準)
            Public Const Overseas = "1"                                             '海外用(価格加算無)
            Public Const C5 = "C5"                                                  'C5    
            Public Const DINRail = "DINRail"                                        'DIN Rail
            Public Const Joint = "Joint"                                            '継手
            Public Const Screw = "Screw"                                            'ねじ
            Public Const Open = "Open"                                              'Open Price
        End Structure

        '電圧区分
        Public Structure VoltageDiv
            Private strDummy As String
            Public Const Standard = "1"                                             '標準電圧
            Public Const Options = "2"                                              'オプション
            Public Const Other = "3"                                                'その他電圧
        End Structure

        '国コード
        Public Structure CountryCd
            Private strDummy As String
            Public Const DefaultCountry = "JPN"                                     'デフォルト国(日本)
        End Structure

        '日本出荷場所コード
        Public Shared ShipPlaceJapan As List(Of String) = New List(Of String) From {"P", "S", "K", "C", "JPN", "C11", "P21", "P11", "P51", "P52", "P55", "C51", "C52", "C55", "S51", "S52", "S55", "K51", "K52", "K55", "1001", "1002", "1003", "1004", "1005"}

        '固定メッセージ 
        Public Structure FixedMessage
            Private strDummy As String
            Public Const PriceJPY = "(JPY=日本円です)"                         '単価画面メッセージ
        End Structure

        '通貨コード
        Public Structure CurrencyCd
            Private strDummy As String
            Public Const DefaultCurrency = "JPY"                                    'デフォルト通貨(円)
            Public Const VNMCurrency = ""                                           'デフォルト通貨("")
        End Structure

        '国コード
        Public Structure OfficeCd
            Private strDummy As String
            Public Const Overseas = "II2"                                           'デフォルト国(日本)
        End Structure

        '記号
        Public Structure Sign
            Private strDummy As String
            Public Const Hypen = "-"                                                'ハイフン
            Public Const Colon = ":"                                                'コロン
            Public Const Comma = ","                                                'カンマ
            Public Const Dot = "."                                                  'ドット
            Public Const Blank = ""                                                 'ブランク
            Public Const Equal = "="                                                'イコール
            Public Const Question = "?"                                             'クエスチョン
            Public Const Asterisk = "*"                                             'アスタリスク
            Public Const OtherOpSymbol = "-X"                                       'オプション外
            '区切り文字
            Public Structure Delimiter
                Private strDummy As String
                Public Const Comma = ","                                            'カンマ
                Public Const Pipe = "|"                                             'パイプ
                Public Const Tab = vbTab                                            'タブ
                Public Const CrLf = vbCrLf                                          '改行復帰
            End Structure
            '少数点文字
            Public Structure DecPoint
                Private strDummy As String
                Public Const Comma = ","                                            'カンマ
                Public Const Dot = "."                                              'ドット
            End Structure
        End Structure

        'プログラムＩＤ
        Public Structure PgmId
            Private strDummy As String
            Public Const KHDefault = "Default"
            Public Const KHLogin = "KHLogin"
            Public Const KHMenu = "KHMenu"
            Public Const KHMenu_Head = "KHMenu_Head"
            Public Const KHModelSelection = "KHModelSelection"
            Public Const KHYouso = "KHYouso"
            Public Const KHUnitPrice = "KHUnitPrice"
            Public Const KHRodEndOrder = "KHRodEndOrder"
            Public Const KHRodEnd = "KHRodEnd"
            Public Const KHOutOFOption = "KHOutOFOption"
            Public Const KHPriceCopy = "KHPriceCopy"
            Public Const KHPriceDetail = "KHPriceDetail"
            Public Const KHSiyou = "KHSiyou"
            Public Const KHISOTanka = "KHISOTanka"
            Public Const KHTanka = "KHTanka"
            Public Const KHMaster = "KHMaster"
            Public Const KHUserMaster = "KHUserMstMnt"
            Public Const KHCountryItemMstMnt = "KHCountryItemMstMnt"
            Public Const KHRateMstMnt = "KHRateMstMnt"
            'ファイル出力
            Public Const KHFileOutput = "KHFileOutput"
        End Structure

        '単価情報
        Public Structure UnitPrice
            Private strDummy As String
            Public Const DefaultNmlRate = "1.0000"                                   '掛率(Normal)　'2013/04/25（1.000→1.0000)
            Public Const DefaultOtrRate = "1,0000"                                   '掛率(Other)　 '2013/04/25（1.000→1.0000)
            Public Const DefaultUnitPrice = "0"                                     '単価
            Public Const DefaultNmlRateUnitPrice = "0.00"                            '単価(Normal)　'2013/04/25（0.0→0.00)
            Public Const DefaultOtrRateUnitPrice = "0,00"                            '単価(Other)　 '2013/04/25（0.0→0.00)
            Public Const ListPrice = "ListPrice"                                    '定価
            Public Const RegPrice = "RegPrice"                                      '登録店
            Public Const SsPrice = "SsPrice"                                        'SS店価格
            Public Const BsPrice = "BsPrice"                                        'BS店価格
            Public Const GsPrice = "GsPrice"                                        'GS店価格
            Public Const PsPrice = "PsPrice"                                        'PS店価格
            Public Const APrice = "APrice"                                          'Local価格
            Public Const FobPrice = "FobPrice"                                      'FOB価格
            Public Const Fca2Price = "Fca2Price"                                    'FCA2価格
            Public Const ListPrice_ja = "定価"
            Public Const RegPrice_ja = "登録店"
            Public Const SsPrice_ja = "SS店"
            Public Const BsPrice_ja = "BS店"
            Public Const GsPrice_ja = "GS店"
            Public Const PsPrice_ja = "PS店"
            Public Const ListPrice_prc = "販売価格"                                 '中国生産表示 2008/6/25
            Public Const CostPrice_ja = "仕入価格"                                  '中国生産表示 2008/6/25
            Public Const CostPrice = "CostPrice"                                    '中国生産表示 2008/6/25

            Public Structure C5Rate
                Private strDummy As String
                Public Const ListPrice = 1.61                                       '定価
                Public Const RegPrice = 1.24                                        '登録店
                Public Const SsPrice = 1.16                                         'SS店価格
                Public Const BsPrice = 1.08                                         'BS店価格
                Public Const GsPrice = 1.0                                          'GS店価格
                Public Const PsPrice = 0.92                                         'PS店価格
            End Structure

            'RM1805036_P40加算価格用
            Public Structure P40Rate
                Private strDummy As String
                Public Const ListPrice = 1.2                                        '定価
                Public Const RegPrice = 1.2                                         '登録店
                Public Const SsPrice = 1.2                                          'SS店価格
                Public Const BsPrice = 1.2                                          'BS店価格
                Public Const GsPrice = 1.2                                          'GS店価格
                Public Const PsPrice = 1.2                                          'PS店価格
            End Structure

        End Structure

        '形番チェック区分
        Public Structure KatabanChackDiv
            Private strDummy As String
            Public Const Stock = "1"                                                '在庫品
            Public Const Standard = "2"                                             '標準品
            Public Const Special = "3"                                              '特注品
            Public Const Parts = "4"                                                '部品
        End Structure

        '検索区分
        Public Structure RetrievalDiv
            Private strDummy As String
            Public Const Model = "1"                                                '機種検索
            Public Const Full = "2"                                                 'フル形番検索
            Public Const Shiire = "3"                                                 '仕入品検索
        End Structure

        '原価積算No.
        Public Structure CostCalcNo
            Private strDummy As String
            Public Const C5 = "C5"                                                  '犬山簡易特注
        End Structure

        'ファイル関連
        Public Structure File
            Private strDummy As String
            Public Const TextExtension As String = ".txt"                           'テキストファイル拡張子
            Public Const LogExtension As String = ".log"                            'ログファイル拡張子
            Public Const ExcelExtension As String = ".xls"                          'Excelファイル拡張子
            Public Const CsvExtension As String = ".csv"                            'csvファイル拡張子       RM0911008 2009/11/30 Y.Miura 追加
            Public Const JsonExtension As String = ".json"                          'Jsonファイル拡張子      RM1809***_JSONファイル追加
        End Structure

        'メッセージ
        Public Structure Message
            Private strDummy As String
            'Title
            Public Structure Title
                Private strDummy As String
                Public Const Japanese = "<< 形番引当システム エラー情報 >>"               'Title(日本語)
                Public Const English = "<< Parts-Number Searching System Error Information >>" 'Title(英語)
            End Structure
            'AuthenticationErrTitle
            Public Structure AuthenticationErrTitle
                Private strDummy As String
                Public Const Japanese = "<< 形番引当システム 認証エラー情報 >>"         'Title(日本語)
                Public Const English = "<< Parts-Number Searching System Authentication Error Information >>" 'Title(英語)
            End Structure
            'NotFound
            Public Structure NotFound
                Private strDummy As String
                Public Const Japanese = "該当するメッセージが登録されていません。"         'NotFoundメッセージ(日本語)
                Public Const English = "The corresponding message is not registered. " 'NotFoundメッセージ(英語)
            End Structure
            'System Error
            Public Structure SystemError
                Private strDummy As String
                Public Const Japanese = "アプリケーションでシステムエラーが発生しました。" & "<br />" & _
                                        "システム担当者に連絡して下さい。"                                 'SystemErrorメッセージ(日本語)
                Public Const English = "The system error occurred by the application. " & "<br />" & _
                                       "Please contact the person in charge of the system. "            'SystemErrorメッセージ(英語)
                Public Const MailMsg = "エラーが発生しています。対応して下さい。"                          'Mail用
            End Structure
            'Authentication Error
            Public Structure AuthenticationErr
                Private strDummy As String
                Public Const Japanese = "認証後一定時間が過ぎました。" & "<br />" & _
                                        "再度認証してください。"                                    'AuthenticationErrorメッセージ(日本語)
                Public Const English = "Time is out. Uniformity time passed, after the authentication. " & "<br />" & _
                                       "Authenticate again. "                                       'AuthenticationErrorメッセージ(英語)
            End Structure
            'Authentication Error
            Public Structure AuthenticationErr2
                Private strDummy As String
                Public Const Japanese = "ログインに失敗しました。" & "<br />" & _
                                        "ＣＫＤ株式会社 営業本部 機器販売企画部（TEL:0568-74-1350）までご連絡ください。"                                    'AuthenticationError2メッセージ(日本語)
            End Structure
        End Structure

        Public Structure Manifold
            Private strDummy As String
            Public Structure Necessity
                Private strDummy As String
                Public Const Japanese = "要"
                Public Const English = "Necessity"
            End Structure
            Public Structure UnNecessity
                Private strDummy As String
                Public Const Japanese = "不要"
                Public Const English = "UnNecessity"
            End Structure
            Public Structure InspReportJp
                Private strDummy As String
                Public Const Japanese = "検査成績書（和文）"
                Public Const English = "Test certificate(Japanese)"
                Public Const SelectValue = "InspReportJp"
                Public Const DummyValue = "ｹﾝｻｾｲｾｷｼﾖ(ﾜﾌﾞﾝ)"
            End Structure
            Public Structure InspReportEn
                Private strDummy As String
                Public Const Japanese = "検査成績書（英文）"
                Public Const English = "Test certificate(English)"
                Public Const SelectValue = "InspReportEn"
                Public Const DummyValue = "ｹﾝｻｾｲｾｷｼﾖ(ｴｲﾌﾞﾝ)"
            End Structure
            Public Structure TubeRemover
                Private strDummy As String
                Public Const DummyValue = "ﾁﾕ-ﾌﾞﾇｷｸﾞﾌﾖｳ"
                Public Const Necessity = "1"
                Public Const UnNecessity = "0"
            End Structure
        End Structure

        'Manifold共通 引当情報の項目
        Public Structure SelSpec
            Private strDummy As String
            Public Const SeqNo = "SeqNo"                                            '連番
            Public Const Kataban = "Kataban"                                        '形番
            Public Const CxA = "CxA"                                                'CXA形番
            Public Const CxB = "CxB"                                                'CXB形番
            Public Const PosInfo = "PosInfo"                                        '設置位置情報
            Public Const Qty = "Qty"                                                '使用数
        End Structure

        '形番チェック区分名
        Public Structure KatabanChackDivName
            Private strDummy As String
            Public Const Stock = "標準品"
            Public Const Standard = "オプション品"
            Public Const Special = "オーダーメイド品"
            Public Const Parts = "部品"
        End Structure

        'ロッド先端特注
        Public Structure RodEndCstmOrder
            Private strDummy As String
            Public Const Label = "L"
            Public Const Text = "T"
            Public Const Drop = "D"
            Public Const OtherSize = "Other"
            Public Const EleBoreSize = "5"
            Public Const RdoGroupNm = "RdoGroup"
            Public Const FrmWF = "WF"
            Public Const FrmA = "A"
            Public Const FrmKK = "KK"
            Public Const FrmC = "C"
            Public Const FrmKL = "KL"
            Public Const RodPtnN13 = "N13"
            Public Const RodPtnN15 = "N15"
            Public Const RodPtnN11 = "N11"
            Public Const RodPtnN1 = "N1"
            Public Const RodPtnN12 = "N12"
            Public Const RodPtnN14 = "N14"
            Public Const RodPtnN3 = "N3"
            Public Const RodPtnN31 = "N31"
            Public Const RodPtnN2 = "N2"
            Public Const RodPtnN21 = "N21"
            Public Const RodPtnN13N11 = "N13-N11"
            Public Const RodPtnN11N13 = "N11-N13"
            Public Const ActionOK = "OK"
            Public Const ActionCancel = "Cancel"
        End Structure

        'スタイルシート用
        Public Structure CSSClass
            Private strDummy As String
            Public Const InputText = "InputText"                                    'TextBox(入力可)
            Public Const InputTextCenter2 = "InputTextCenter2"                      'TextBox(入力可)
            Public Const InputTextCenter3 = "InputTextCenter3"                      'TextBox(入力可)
            Public Const InputTextReadOnly = "InputTextReadOnly"                    'TextBox(表示専用)
            Public Const InputTextCenterReadOnly = "InputTextCenterReadOnly"        'TextBox(表示専用)
            Public Const InputTextCenterReadOnly2 = "InputTextCenterReadOnly2"      'TextBox(表示専用)
            Public Const InputTextHypen = "InputTextHypen"                          'TextBoxエラー）
            Public Const InputNoSelect = "InputNoSelect"                            'TextBox(選択不可)
            Public Const InputNoSelectCenter = "InputNoSelectCenter"                'TextBox(選択不可/普通ｶｰｿﾙ)
            Public Const InputNoSelectCenter2 = "InputNoSelectCenter2"              'TextBox(選択不可/普通ｶｰｿﾙ)
            Public Const InputNoSelectLeft = "InputNoSelectLeft"                    'TextBox(選択不可/普通ｶｰｿﾙ)
            Public Const InputNoSelectRight = "InputNoSelectRight"                  'TextBox(選択不可/普通ｶｰｿﾙ)
            Public Const InputSelect = "InputSelect"                                'TextBox
            Public Const inputSelectCenter = "InputSelectCenter"                    'TextBox(普通ｶｰｿﾙ)
            Public Const inputSelectCenter2 = "InputSelectCenter2"                  'TextBox(普通ｶｰｿﾙ)
            Public Const inputSelectLeft = "InputSelectLeft"                        'TextBox(普通ｶｰｿﾙ)
            Public Const InputNumeric = "InputNumeric"                              'TextBox(数値用：入力可)
            Public Const InputNumeric2 = "InputNumeric2"                            'TextBox(数値用：入力可)
            Public Const InputNumericCenter = "InputNumericCenter"                  'TextBox(数値用：入力可/中央)
            Public Const InputNumericReadOnly = "InputNumericReadOnly"              'TextBox(数値用：表示専用)
            Public Const InputNumericCenterReadOnly = "InputNumericCenterReadOnly"  'TextBox(数値/中央/表示専用)
            Public Const InputTextCenterSel = "InputTextCenterSel"                  'TextBox
            Public Const InputPriceList = "InputPriceList"                          'TextBox
            Public Const PriceEstimateText1 = "PriceEstimateText1"                  'TextBox(数値/右/入力可)
            Public Const PriceEstimateText2 = "PriceEstimateText2"                  'TextBox(数値/右/入力不可)
            Public Const PriceEstimateText3 = "PriceEstimateText3"                  'TextBox(文字/中央/入力不可)
            Public Const InputRodEndOrder1 = "InputRodEndOrder1"                    'TextBox(数値/右/入力可)
            Public Const InputRodEndOrder2 = "InputRodEndOrder2"                    'TextBox(文字/左/入力可)
            Public Const InputRodEndOrderReadOnly = "InputRodEndOrderReadOnly"      'TextBox(数値/右/入力不可)
            Public Const Button = "Button"                                          'ボタン
            Public Const SelectButton = "SelectButton"                              'リストの選択ボタン
            Public Const RowNoSelect = "rowNoSelect"                                'TR(選択不可)
            Public Const RowNoSelect2 = "rowNoSelect2"                              'TR(選択不可/普通ｶｰｿﾙ)
            Public Const RowNoSelect3 = "rowNoSelect3"                              'TR(選択不可/普通ｶｰｿﾙ/背景色なし)
            Public Const RowSelect = "rowSelect"                                    'TR(選択可)
            Public Const RowSelect2 = "rowSelect2"                                  'TR(選択可/普通ｶｰｿﾙ)
            Public Const TableList = "tableList"                                    'Table
            Public Const PriceListLabel = "PriceListLabel"                          'TR
            Public Const PriceEstimateLabel2 = "PriceEstimateLabel2"                'Span
            Public Const PriceEstimateLabel3 = "PriceEstimateLabel3"                'Span
            Public Const ListLabel = "ListLabel"                                    'Span(Listのﾍｯﾀﾞｰﾗﾍﾞﾙ)
            Public Const RodListLabel = "RodListLabel"                              'Span(ロッド先端特注Listのヘッダーラベル)
            Public Const FixedMessageLabel1 = "FixedMessageLabel1"                  'Span(国内代理店用メッセージ) RM0911008 2009/11/24 Y.Miura   
            Public Const InputNoSelectTnk1 = "InputNoSelectTnk1"                    'Input
            Public Const InputNoSelectTnk2 = "InputNoSelectTnk2"                    'Input
            Public Const LargeButtonOnBlur = "LargeButtonOnBlur"
            Public Const MenuButtonOnBlur = "MenuButtonOnBlur"                      'メニューボタン
            Public Const DDLBold = "DDLBold"                                        'DropDownList(太字)
            Public Const RodDDL = "RodDDL"                                          'DropDownList(文字/中央/入力不可)
            Public Const RodRadio = "RodRadio"                                      'ラジオボタン(ロッド先端特注用)
        End Structure

        'JavaScript用
        Public Structure JavaScript
            Private strDummy As String
            Public Const OnFocus = "onfocus"                                        'イベント(OnFocus)
            Public Const OnBlur = "onblur"                                          'イベント(OnBlur)
            Public Const OnClick = "onclick"                                        'イベント(OnClick)
            Public Const OnClientClick = "onclientclick"                            'イベント(OnClick)
            Public Const OnChange = "onchange"                                      'イベント(OnChange)
            Public Const OnDblClick = "ondblclick"                                  'イベント(OnDblCkick)
            Public Const OnKeyDown = "onkeydown"                                    'イベント(OnKeyDown)
            Public Const OnKeyUp = "onkeyup"                                        'イベント(OnKeyUp)
            Public Const OnMouseDown = "onmousedown"                                'イベント(OnMouseDown)
            Public Const OnLoad = "onload"                                          'イベント(OnLoad)
            Public Const OnUnLoad = "onunload"                                        'イベント(OnUnload)
            Public Const UnSelectableProperty = "unselectable"                      'Unselectable
            Public Const ReadOnlyProperty = "readOnly"                              'ReadOnly
            Public Const ValueProperty = "value"                                    'Value
        End Structure

        Public Structure MonifoldGrid
            Public Const intGridWidth As Long = 22    '設置位置列の幅
            Public Const intGridHeight As Long = 22   '設置位置列の高さ
        End Structure

        Public Structure FileHeaderInfo
            Public strDummy As String
            Public Const FileOutPut1 As String = ControlChars.Quote & "製品名" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "形番" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "定価" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "登録" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "ＳＳ" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "ＢＳ" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "ＧＳ" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "掛率" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "単価" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "数量" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "金額" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "消費税" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "合計" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "更新日" & ControlChars.Quote

            '項目追加(形番チェック,出荷場所,PS価格)
            Public Const FileOutPut2 As String = ControlChars.Quote & "製品名" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "形番" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "形番チェック" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "出荷場所" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "定価" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "登録" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "ＳＳ" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "ＢＳ" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "ＧＳ" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "ＰＳ" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "掛率" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "単価" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "数量" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "金額" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "消費税" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "合計" & ControlChars.Quote & CdCst.Sign.Comma & _
                                                 ControlChars.Quote & "更新日" & ControlChars.Quote
        End Structure

        ''' <summary>
        ''' セレクト品チェック用データ
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared strSelectM4GCheckData As List(Of String) = New List(Of String)(New String() {"N4G1-EL", "N4G1-ER", "N4G1-Q-6", _
                                                                "N4G1-Q-8", "N4G1-T10", "N4G1-T10W", _
                                                                "N4G2-EL", "N4G2-ER", "N4G2-Q-10", _
                                                                "N4G2-Q-8", "N4G2-T10", "N4G2-T10W", _
                                                                "N4GB110-C4", "N4GB110-C6", "N4GB120-C4", _
                                                                "N4GB120-C6", "N4GB210-C6", "N4GB210-C8", _
                                                                "N4GB220-C4", "N4GB220-C6", "N4GB220-C8", _
                                                                "N4G1R-EL", "N4G1R-ER", "N4G1R-Q-6", _
                                                                "N4G1R-Q-8", "N4G1R-T10", "N4G1R-T10W", _
                                                                "N4G2R-EL", "N4G2R-ER", "N4G2R-Q-10", _
                                                                "N4G2R-Q-8", "N4G2R-T10", "N4G2R-T10W", _
                                                                "N4GB110R-C4", "N4GB110R-C6", "N4GB120R-C4", _
                                                                "N4GB120R-C6", "N4GB210R-C6", "N4GB210R-C8", _
                                                                "N4GB220R-C4", "N4GB220R-C6", "N4GB220R-C8"})

        ''' <summary>
        ''' ASEAN対応国コード
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared strAseanCode As List(Of String) = New List(Of String) From {"THA", "MYS", "SGP", "IDN", "VNM", "IND", "PRC"}        '中国(PRC)追加

        ''' <summary>
        ''' 機種選択画面モード
        ''' </summary>
        Public Shared seriesSelectMode As New Dictionary(Of String, String) From {{"1", "機種選択(ツリー)"}, {"2", "機種選択(検索)"}}

        ''' <summary>
        ''' センサ文字列
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Senser
            Public strDummy As String
            Public Const ja As String = "ｾﾝｻ"
            Public Const en As String = "Senser"
        End Structure

        ''' <summary>
        ''' 商品名文字列（仕入品・修理）
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure GoogsName_Shiire
            Public strDummy As String
            Public Const ja As String = "仕入品・修理"
            Public Const en As String = "Goods on hand"
            Public Const ko As String = "매입품"
            Public Const tw As String = "存貨"
            Public Const zh As String = "购入品"
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義01
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_01
            Private dummy As Integer
            Public Shared Elect1 As Integer = 1
            Public Shared Elect2 As Integer = 2
            Public Shared Wiring As Integer = 3
            Public Shared Valve1 As Integer = 4
            Public Shared Valve2 As Integer = 5
            Public Shared Valve3 As Integer = 6
            Public Shared Valve4 As Integer = 7
            Public Shared Valve5 As Integer = 8
            Public Shared Valve6 As Integer = 9
            Public Shared Valve7 As Integer = 10
            Public Shared Dummy1 As Integer = 11
            Public Shared Dummy2 As Integer = 12
            Public Shared Exhaust1 As Integer = 13
            Public Shared Exhaust2 As Integer = 14
            Public Shared Exhaust3 As Integer = 15
            Public Shared Exhaust4 As Integer = 16
            Public Shared Regulat1 As Integer = 17
            Public Shared Regulat2 As Integer = 18
            Public Shared EndL As Integer = 19
            Public Shared EndR As Integer = 20
            Public Shared Plug1 As Integer = 21
            Public Shared Plug2 As Integer = 22
            Public Shared Plug3 As Integer = 23
            Public Shared Plug4 As Integer = 24
            Public Shared Rail As Integer = 25
            Public Shared Inspect1 As Integer = 26
            Public Shared Inspect2 As Integer = 27
            Public Shared Inspect3 As Integer = 28
            Public Shared Inspect4 As Integer = 29
            Public Shared Tube As Integer = 30
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義02
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_02
            Private dummy As Integer
            Public Shared End1 As Integer = 1           'エンドブロック
            Public Shared End2 As Integer = 2
            Public Shared Exhaust1 As Integer = 3       '給排気ブロック
            Public Shared Exhaust2 As Integer = 4
            Public Shared Exhaust3 As Integer = 5       '給気ブロック
            Public Shared Exhaust4 As Integer = 6
            Public Shared Exhaust5 As Integer = 7       '排気ブロック
            Public Shared Exhaust6 As Integer = 8
            Public Shared Valve1 As Integer = 9         'バルブブロック
            Public Shared Valve2 As Integer = 10
            Public Shared Valve3 As Integer = 11
            Public Shared Valve4 As Integer = 12
            Public Shared Valve5 As Integer = 13
            Public Shared Valve6 As Integer = 14
            Public Shared Partition1 As Integer = 15    '仕切りブロック
            Public Shared Partition2 As Integer = 16
            Public Shared Silencer1 As Integer = 17     'サイレンサ
            Public Shared Silencer2 As Integer = 18
            Public Shared Plug1 As Integer = 19         'ブランクプラグ
            Public Shared Plug2 As Integer = 20
            Public Shared Rail As Integer = 21          '取付レール長さ
            Public Shared Inspect As Integer = 22       '検査成績書
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義03
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_03
            Private dummy As Integer
            Public Shared Elect1 As Integer = 1         '電磁弁
            Public Shared Elect2 As Integer = 2
            Public Shared Elect3 As Integer = 3
            Public Shared Elect4 As Integer = 4
            Public Shared Elect5 As Integer = 5
            Public Shared Elect6 As Integer = 6
            Public Shared Elect7 As Integer = 7
            Public Shared Elect8 As Integer = 8
            Public Shared Elect9 As Integer = 9
            Public Shared Elect10 As Integer = 10
            Public Shared Elect11 As Integer = 11
            Public Shared Elect12 As Integer = 12
            Public Shared Elect13 As Integer = 13
            Public Shared Elect14 As Integer = 14
            Public Shared Masking As Integer = 15       'マスキングプレート
            Public Shared Plug1 As Integer = 16         'ブランクプラグ＆サイレンサ
            Public Shared Plug2 As Integer = 17
            Public Shared Plug3 As Integer = 18
            Public Shared Plug4 As Integer = 19
            Public Shared Rail As Integer = 20          '取付レール長さ
            Public Shared Inspect As Integer = 21
            Public Shared Cable1 As Integer = 22        'ケーブル
            Public Shared Cable2 As Integer = 23
            Public Shared Tube As Integer = 24          'チューブ抜具
        End Structure

        'RM1803032_スペーサ行追加対応
        ''' <summary>
        ''' マニホールド仕様定義04
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_04
            Private dummy As Integer
            Public Shared Valve1 As Integer = 1
            Public Shared Valve2 As Integer = 2
            Public Shared Valve3 As Integer = 3
            Public Shared Valve4 As Integer = 4
            Public Shared Valve5 As Integer = 5
            Public Shared Valve6 As Integer = 6
            Public Shared Valve7 As Integer = 7
            Public Shared Valve8 As Integer = 8
            Public Shared Valve9 As Integer = 9
            Public Shared Valve10 As Integer = 10
            Public Shared MasPlate1 As Integer = 11
            Public Shared MasPlate2 As Integer = 12
            Public Shared Spacer1 As Integer = 13
            Public Shared Spacer2 As Integer = 14
            Public Shared Spacer3 As Integer = 15
            Public Shared Spacer4 As Integer = 16
            Public Shared BlkPlug1 As Integer = 17
            Public Shared BlkPlug2 As Integer = 18
            Public Shared Silencer1 As Integer = 19
            Public Shared Silencer2 As Integer = 20
            Public Shared ScrPlug As Integer = 21
            Public Shared Rail As Integer = 22
            Public Shared TestCert As Integer = 23
            Public Shared Cable1 As Integer = 24
            Public Shared Cable2 As Integer = 25
            Public Shared Tube As Integer = 26
            Public Const Elect1 As Integer = 1         '電磁弁
            Public Const Elect10 As Integer = 10         '電磁弁
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義05
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_05
            Private dummy As Integer
            Public Shared ElType1 As Integer = 1
            Public Shared ElType2 As Integer = 2
            Public Shared ElType3 As Integer = 3
            Public Shared ElType4 As Integer = 4
            Public Shared ElType5 As Integer = 5
            Public Shared ElType6 As Integer = 6
            Public Shared ABPlugR As Integer = 7
            Public Shared ABPlugL As Integer = 8
            Public Shared ABCon02 As Integer = 9
            Public Shared ABCon03 As Integer = 10
            Public Shared ABCon04 As Integer = 11
            Public Shared RepSpace1 As Integer = 12
            Public Shared RepSpace2 As Integer = 13
            Public Shared ExhSpace1 As Integer = 14
            Public Shared ExhSpace2 As Integer = 15
            Public Shared Pilot1 As Integer = 16
            Public Shared Pilot2 As Integer = 17
            Public Shared SpDecomp1 As Integer = 18
            Public Shared SpDecomp2 As Integer = 19
            Public Shared SpDecomp3 As Integer = 20
            Public Shared SpDecomp4 As Integer = 21
            Public Shared ExpCovRep As Integer = 22
            Public Shared ExpCovExh As Integer = 23
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義06
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_06
            Private dummy As Integer
            Public Shared Elect1 As Integer = 1
            Public Shared Elect2 As Integer = 2
            Public Shared Elect3 As Integer = 3
            Public Shared Elect4 As Integer = 4
            Public Shared Elect5 As Integer = 5
            Public Shared Elect6 As Integer = 6
            Public Shared ABCon01 As Integer = 7
            Public Shared ABCon02 As Integer = 8
            Public Shared ABCon04 As Integer = 9
            Public Shared ABCon06 As Integer = 10
            Public Shared ABCon1Z As Integer = 11
            Public Shared RepSpace1 As Integer = 12
            Public Shared RepSpace2 As Integer = 13
            Public Shared ExhSpace1 As Integer = 14
            Public Shared ExhSpace2 As Integer = 15
            Public Shared Pilot As Integer = 16
            Public Shared ExpCovRep As Integer = 17
            Public Shared ExpCovExh As Integer = 18
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義07
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_07
            Private dummy As Integer
            Public Shared Equip As Integer = 1          '電装ブロック
            Public Shared Elect1 As Integer = 2         '電磁弁
            Public Shared Elect2 As Integer = 3
            Public Shared Elect3 As Integer = 4
            Public Shared Elect4 As Integer = 5
            Public Shared Elect5 As Integer = 6
            Public Shared Elect6 As Integer = 7
            Public Shared Elect7 As Integer = 8
            Public Shared Elect8 As Integer = 9
            Public Shared Mix As Integer = 10           'ミックスブロック
            Public Shared Spacer1 As Integer = 11       'スペーサー
            Public Shared Spacer2 As Integer = 12
            Public Shared Spacer3 As Integer = 13
            Public Shared Spacer4 As Integer = 14
            Public Shared Exhaust1 As Integer = 15      '給排気ブロック
            Public Shared Exhaust2 As Integer = 16
            Public Shared Exhaust3 As Integer = 17
            Public Shared Partition1 As Integer = 18    '仕切ブロック
            Public Shared Partition2 As Integer = 19
            Public Shared EndLeft As Integer = 20       'エンドブロック(左)
            Public Shared EndRight As Integer = 21      'エンドブロック(右)
            Public Shared Plug1 As Integer = 22         'ブランクプラグ＆サイレンサ
            Public Shared Plug2 As Integer = 23
            Public Shared Plug3 As Integer = 24
            Public Shared Rail As Integer = 25         '取付レール長さ
            Public Shared Inspect1 As Integer = 26      '検査成績書＆ケーブル
            Public Shared Inspect2 As Integer = 27
            Public Shared Tag As Integer = 28           'タグ銘板
            Public Shared Tube As Integer = 29          'チューブ抜具
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義08
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_08
            Private dummy As Integer
            Public Shared EndP1 As Integer = 1          '左側給気用エンドプレート
            Public Shared EndP2 As Integer = 2          '右側給気用エンドプレート
            Public Shared ElType1 As Integer = 3        '電磁弁付サブブロック
            Public Shared ElType2 As Integer = 4
            Public Shared ElType3 As Integer = 5
            Public Shared ElType4 As Integer = 6
            Public Shared ElType5 As Integer = 7
            Public Shared ElType6 As Integer = 8
            Public Shared ElType7 As Integer = 9
            Public Shared ElType8 As Integer = 10
            Public Shared ElType9 As Integer = 11
            Public Shared ElType10 As Integer = 12
            Public Shared Supply1 As Integer = 13       '中間給気プレート
            Public Shared Supply2 As Integer = 14
            Public Shared Exhaust1 As Integer = 15      '中間排気プレート
            Public Shared Exhaust2 As Integer = 16
            Public Shared Silencer1 As Integer = 17     'サイレンサ
            Public Shared Silencer2 As Integer = 18
            Public Shared Plug1 As Integer = 19         'ブランクプラグ
            Public Shared Plug2 As Integer = 20
            Public Shared Rail As Integer = 21         '取付レール長さ
            Public Shared Inspect1 As Integer = 22      '検査成績書
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義09
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_09
            Private dummy As Integer
            Public Shared Endb1 As Integer = 1          'エンドブロック
            Public Shared Endb2 As Integer = 2
            Public Shared Wiring As Integer = 3         '配線ブロック
            Public Shared ElValve1 As Integer = 4       '電磁弁付バルブブロック
            Public Shared ElValve2 As Integer = 5
            Public Shared ElValve3 As Integer = 6
            Public Shared ElValve4 As Integer = 7
            Public Shared ElValve5 As Integer = 8
            Public Shared MpValve1 As Integer = 9       'マスキングプレート付バルブブロック
            Public Shared MpValve2 As Integer = 10
            Public Shared SpReguP As Integer = 11       'スペーサ形レギュレータ(P)
            Public Shared SpReguA As Integer = 12       'スペーサ形レギュレータ(A)
            Public Shared SpReguB As Integer = 13       'スペーサ形レギュレータ(B)
            Public Shared SupplySp As Integer = 14      '単独給気スペーサ
            Public Shared ExhaustSp As Integer = 15     '単独排気スペーサ
            Public Shared PartitionS As Integer = 16    '仕切りプラグ（給気用）
            Public Shared PartitionE As Integer = 17    '仕切りプラグ（排気用）
            Public Shared Silencer1 As Integer = 18     'サイレンサ(樹脂)
            Public Shared Silencer2 As Integer = 19     'サイレンサ(メタル)
            Public Shared HexPlug1 As Integer = 20      '六角穴付プラグ
            Public Shared HexPlug2 As Integer = 21
            Public Shared Cable As Integer = 22         'ケーブルクランプ
            Public Shared Inspect As Integer = 23       '検査成績書
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義11
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_11
            Private dummy As Integer
            Public Shared EndL As Integer = 1           'エンドブロックL
            Public Shared ChargeAir As Integer = 2      '集中給気ブロック
            Public Shared ChargeAirAPS As Integer = 3   'APS付集中給気ブロック
            Public Shared Regulator1 As Integer = 4     'レギュレータ
            Public Shared Regulator2 As Integer = 5
            Public Shared Regulator3 As Integer = 6
            Public Shared Regulator4 As Integer = 7
            Public Shared Regulator5 As Integer = 8
            Public Shared Regulator6 As Integer = 9
            Public Shared Regulator7 As Integer = 10
            Public Shared Regulator8 As Integer = 11
            Public Shared Regulator9 As Integer = 12
            Public Shared Regulator10 As Integer = 13
            Public Shared Subbase As Integer = 14       'MP付サブベース
            Public Shared EndR As Integer = 15          'エンドブロックR
            Public Shared Plug1 As Integer = 16         'ブランクプラグ
            Public Shared Plug2 As Integer = 17
            Public Shared Plug3 As Integer = 18
            Public Shared Rail As Integer = 19          '取付レール長さ
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義12
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_12
            Private dummy As Integer
            Public Shared Vaccum1 As Integer = 1
            Public Shared Vaccum2 As Integer = 2
            Public Shared Vaccum3 As Integer = 3
            Public Shared Vaccum4 As Integer = 4
            Public Shared Vaccum5 As Integer = 5
            Public Shared Vaccum6 As Integer = 6
            Public Shared Vaccum7 As Integer = 7
            Public Shared Vaccum8 As Integer = 8
            Public Shared Mask1 As Integer = 9
            Public Shared Mask2 As Integer = 10
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義13
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_13
            Private dummy As Integer
            Public Shared Wiring As Integer = 1         '配線ブロック
            Public Shared End1 As Integer = 2           'エンドブロック
            Public Shared End2 As Integer = 3
            Public Shared Valve1 As Integer = 4         'バルブブロック
            Public Shared Valve2 As Integer = 5
            Public Shared Valve3 As Integer = 6
            Public Shared Valve4 As Integer = 7
            Public Shared Valve5 As Integer = 8
            Public Shared Valve6 As Integer = 9
            Public Shared Exhaust1 As Integer = 10      '給排気ブロック
            Public Shared Exhaust2 As Integer = 11
            Public Shared Exhaust3 As Integer = 12      '給気ブロック
            Public Shared Exhaust4 As Integer = 13
            Public Shared Exhaust5 As Integer = 14      '排気ブロック
            Public Shared Exhaust6 As Integer = 15
            Public Shared Partition1 As Integer = 16    '仕切りブロック
            Public Shared Partition2 As Integer = 17
            Public Shared Silencer1 As Integer = 18     'サイレンサ
            Public Shared Silencer2 As Integer = 19
            Public Shared Silencer3 As Integer = 20
            Public Shared Silencer4 As Integer = 21
            Public Shared Rail As Integer = 22          '取付レール長さ
            Public Shared Inspect As Integer = 23       '検査成績書
            Public Shared Cable1 As Integer = 24        'ケーブル
            Public Shared Cable2 As Integer = 25
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義14
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_14
            Private dummy As Integer
            Public Shared Evt As Integer = 1            'EVT
            Public Shared Exhaust1 As Integer = 2       '電装／給排気ブロック
            Public Shared Exhaust2 As Integer = 3
            Public Shared Exhaust3 As Integer = 4
            Public Shared End1 As Integer = 5           'エンドブロック
            Public Shared End2 As Integer = 6
            Public Shared Plug1 As Integer = 7          'ブランクプラグ
            Public Shared Plug2 As Integer = 8
            Public Shared Silencer As Integer = 9       'サイレンサ
            Public Shared Rail As Integer = 10          '取付レール長さ
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義15
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_15
            Private dummy As Integer
            Public Shared InOut1 As Integer = 1
            Public Shared InOut2 As Integer = 2
            Public Shared Elect As Integer = 3
            Public Shared Valve1 As Integer = 4
            Public Shared Valve2 As Integer = 5
            Public Shared Valve3 As Integer = 6
            Public Shared Valve4 As Integer = 7
            Public Shared Valve5 As Integer = 8
            Public Shared Valve6 As Integer = 9
            Public Shared Valve7 As Integer = 10
            Public Shared Valve8 As Integer = 11
            Public Shared Spacer1 As Integer = 12
            Public Shared Spacer2 As Integer = 13
            Public Shared Spacer3 As Integer = 14
            Public Shared Spacer4 As Integer = 15
            Public Shared Exhaust1 As Integer = 16
            Public Shared Exhaust2 As Integer = 17
            Public Shared Partition1 As Integer = 18
            Public Shared Partition2 As Integer = 19
            Public Shared EndR As Integer = 20
            Public Shared EndL As Integer = 21
            Public Shared Plug1 As Integer = 22
            Public Shared Plug2 As Integer = 23
            Public Shared Plug3 As Integer = 24
            Public Shared Rail As Integer = 25
            Public Shared Waterproof As Integer = 26
            Public Shared Inspect1 As Integer = 27
            Public Shared Inspect2 As Integer = 28
            Public Shared Cable As Integer = 29
            Public Shared TagPlaque As Integer = 30
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義16
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_16
            Private dummy As Integer
            Public Shared InOut1 As Integer = 1
            Public Shared InOut2 As Integer = 2
            Public Shared EndL As Integer = 3
            Public Shared EndR As Integer = 4
            Public Shared Elect As Integer = 5
            Public Shared Valve1 As Integer = 6
            Public Shared Valve2 As Integer = 7
            Public Shared Valve3 As Integer = 8
            Public Shared Valve4 As Integer = 9
            Public Shared Valve5 As Integer = 10
            Public Shared MPValve1 As Integer = 11
            Public Shared MPValve2 As Integer = 12
            Public Shared PartitionBlk1 As Integer = 13
            Public Shared Spacer1 As Integer = 14
            Public Shared Spacer2 As Integer = 15
            Public Shared RegulatorP As Integer = 16
            Public Shared RegulatorA As Integer = 17
            Public Shared RegulatorB As Integer = 18
            Public Shared Partition1 As Integer = 19
            Public Shared Partition2 As Integer = 20
            Public Shared Plug1 As Integer = 21
            Public Shared Plug2 As Integer = 22
            Public Shared Plug3 As Integer = 23
            Public Shared Cable As Integer = 24
            Public Shared Inspect As Integer = 25
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義17
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_17
            Private dummy As Integer
            Public Const Unit1 As Integer = 1
            Public Const Unit2 As Integer = 2
            Public Const Unit3 As Integer = 3
            Public Const Unit4 As Integer = 4
            Public Const Unit5 As Integer = 5
            Public Const Base As Integer = 6
        End Structure

        ''' <summary>
        ''' マニホールド仕様定義18
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure Siyou_18
            Private dummy As Integer
            Public Shared Equip As Integer = 1          '電装ブロック
            Public Shared Elect1 As Integer = 2         '電磁弁
            Public Shared Elect2 As Integer = 3
            Public Shared Elect3 As Integer = 4
            Public Shared Elect4 As Integer = 5
            Public Shared Elect5 As Integer = 6
            Public Shared Elect6 As Integer = 7
            Public Shared Elect7 As Integer = 8
            Public Shared Elect8 As Integer = 9
            Public Shared Mix As Integer = 10           'ミックスブロック
            Public Shared Spacer1 As Integer = 11       'スペーサー
            Public Shared Spacer2 As Integer = 12
            Public Shared Spacer3 As Integer = 13
            Public Shared Spacer4 As Integer = 14
            Public Shared Exhaust1 As Integer = 15      '給排気ブロック
            Public Shared Exhaust2 As Integer = 16
            Public Shared Exhaust3 As Integer = 17
            Public Shared Partition1 As Integer = 18    '仕切ブロック
            Public Shared Partition2 As Integer = 19
            Public Shared EndLeft As Integer = 20       'エンドブロック(左)
            Public Shared EndRight As Integer = 21      'エンドブロック(右)
            Public Shared Plug1 As Integer = 22         'ブランクプラグ＆サイレンサ
            Public Shared Plug2 As Integer = 23
            Public Shared Plug3 As Integer = 24
            Public Shared Rail As Integer = 25          '取付レール長さ
            Public Shared Inspect1 As Integer = 26      '検査成績書＆ケーブル
            Public Shared Inspect2 As Integer = 27
            Public Shared Tag As Integer = 28           'タグ銘板
            Public Shared Tube As Integer = 29          'チューブ抜具
        End Structure
    End Structure

    ''' <summary>
    ''' 機種画面の定数
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure CdCstType
        '次ページがこの機種で始まる
        Public Shared strPageFirstKisyu As String = "ListKey"
        '画面上の機種情報を記録
        Public Shared strDTList As String = "DT_List"
        '画面表示件数
        Public Shared intMaxRowCnt As Integer = 15
    End Structure

    ''' <summary>
    ''' 要素画面の定数
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure CdCstYouso
        '1文字の幅
        Public Shared intStrWidth As Integer = 13
        'ハイフン幅
        Public Shared intHypenWidth As Integer = 20
        '要素区分「1(電圧)」の文字数
        Public Shared intVolStrcnt As Integer = 11
        '要素区分「3(ストローク)」の文字数
        Public Shared intStrokeStrcnt As Integer = 5
    End Structure

    ''' <summary>
    ''' 価格表示レベル
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure strcPriceDispLvl
        Private strDummy As String
        Public Const ListPrice As Integer = 1               '定価
        Public Const RegPrice As Integer = 2                '登録店
        Public Const SsPrice As Integer = 4                 'SS
        Public Const BsPrice As Integer = 8                 'BS
        Public Const GsPrice As Integer = 16                'GS
        Public Const PsPrice As Integer = 32                'PS
        Public Const APrice As Integer = 64                 '現地定価
        Public Const FobPrice As Integer = 128              'FOB
        Public Const CostPrice As Integer = 256             '仕入価格
    End Structure

    ''' <summary>
    ''' 付加情報レベル
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure strcAddInformationLvl
        Private strDummy As String
        Public Const KatabanCheckDiv As Integer = 1         '形番チェック区分
        Public Const PlaceCd As Integer = 2                 '出荷場所
        Public Const StockInfo As Integer = 4               '在庫情報
        Public Const PersonInfo As Integer = 8              '担当者情報
        Public Const StdDlvDt As Integer = 16               '標準納期
        Public Const QtyUnit As Integer = 32                '販売数量単位
        Public Const ELInfo As Integer = 64                 'EL品情報
    End Structure

    ''' <summary>
    ''' 利用機能レベル
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure strcUserFunctionLvl
        Private strDummy As String
        Public Const AcosIF As Integer = 1                  'ACOS I/F
        Public Const SapIF As Integer = 2                   'SBO I/F
        Public Const RateMstMnt As Integer = 4              '掛率マスタメンテナンス
        Public Const UserMstMnt As Integer = 8              'ユーザーマスタメンテナンス
        Public Const InfoMstMnt As Integer = 16             '情報マスタメンテナンス
        Public Const CurrencyExcRateMstMnt As Integer = 32  '為替率マスタメンテナンス
        Public Const CountryItemMstMnt As Integer = 64      '為替率マスタメンテナンス
    End Structure

    ''' <summary>
    ''' ファイル出力項目番号
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FileOutputColumns
        ''' <summary>
        ''' 製品名
        ''' </summary>
        ProductName = 1
        ''' <summary>
        ''' 形番
        ''' </summary>
        Kataban = 2
        ''' <summary>
        ''' チェック区分
        ''' </summary>
        CheckKBN = 3
        ''' <summary>
        ''' 出荷場所
        ''' </summary>
        ShipPlace = 4
        ''' <summary>
        ''' 定価
        ''' </summary>
        ListPrice = 5
        ''' <summary>
        ''' 登録店
        ''' </summary>
        RegPrice = 6
        ''' <summary>
        ''' SS
        ''' </summary>
        SsPrice = 7
        ''' <summary>
        ''' BS
        ''' </summary>
        BsPrice = 8
        ''' <summary>
        ''' GS
        ''' </summary>
        GsPrice = 9
        ''' <summary>
        ''' PS
        ''' </summary>
        PsPrice = 10
        ''' <summary>
        ''' 現地定価
        ''' </summary>
        APrice = 11
        ''' <summary>
        ''' 購入価格
        ''' </summary>
        FobPrice = 12
        ''' <summary>
        ''' 掛率
        ''' </summary>
        Rate = 13
        ''' <summary>
        ''' 単価
        ''' </summary>
        UnitPrice = 14
        ''' <summary>
        ''' 数量
        ''' </summary>
        Quantity = 15
        ''' <summary>
        ''' 金額
        ''' </summary>
        Amount = 16
        ''' <summary>
        ''' 消費税
        ''' </summary>
        Tax = 17
        ''' <summary>
        ''' 合計
        ''' </summary>
        Total = 18
        ''' <summary>
        ''' 更新日
        ''' </summary>
        UpdateDate = 19
        ''' 保管場所
        ''' </summary>
        StorageLocation = 20
        ''' 評価タイプ
        ''' </summary>
        EvaluationType = 21
    End Enum

    ''' <summary>
    ''' ファイル出力タイプ
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FileOutputType
        ''' <summary>
        ''' 普通
        ''' </summary>
        Normal = 1
        ''' <summary>
        ''' ISO
        ''' </summary>
        ISO = 2
    End Enum

    ''' <summary>
    ''' シリーズ選択画面タイプ
    ''' </summary>
    Public Enum SeriesSelectPageType
        ''' <summary>
        ''' 検索
        ''' </summary>
        Search
        ''' <summary>
        ''' 一覧
        ''' </summary>
        List
    End Enum

End Module
