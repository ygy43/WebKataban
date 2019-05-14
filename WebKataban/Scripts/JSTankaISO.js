/*******************************************************************************
【関数名】  : ISOTanka_OnKeyup
【概要】    : ISO単価入力欄のEnterキーイベント
【引数】    : strName         <String>    画面名
              intMode         <Integer>   1：掛け率欄　2：単価欄　3：数量欄
【戻り値】  : 無し
*******************************************************************************/
function ISOTanka_OnKeyup(e, intMode, strName) {
    switch (e.keyCode) {
        case 13:
            if (intMode == 1) { document.getElementById(strName + "txt_UnitPrc").focus(); }
            if (intMode == 2) { document.getElementById(strName + "txt_Amount").focus(); }
            if (intMode == 3) {//次のISO単価
                var strID = strName.substr(strName.length - 3, 2);
                if (isNumber(strID)) {
                    strName = strName.substr(0, strName.length - 3) + (parseInt(strID) + 1) + "_";
                } else {
                    strID = strName.substr(strName.length - 2, 1);
                    strName = strName.substr(0, strName.length - 2) + (parseInt(strID) + 1) + "_";
                }
                if (document.getElementById(strName + "txt_Rate")) {
                    document.getElementById(strName + "txt_Rate").focus();
                } else {
                    strName = strName.substr(0, strName.length - 10 - strID.length);
                    if (document.getElementById(strName + "btnOK")) { document.getElementById(strName + "btnOK").focus(); }
                }
            }
            return false;
        case 68:
            //Shift + Dを押下した場合
            if (e.shiftKey == true) {
                var strName = "ctl00_ContentDetail_WebUC_ISOTanka_";
                frmShiftD(strName);
            }
        default:
            break;
    }
}

/*******************************************************************************
【関数名】  : ISOTanka_ChkUnitList
【概要】    : ISO単価のチェック欄（明細反映用）
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function ISOTanka_ChkUnitList(strName, intLoop) {
    var strID;
    if (intLoop == 1) {//すべて
        strID = document.getElementById(strName + "HidSelRowID").value;
    } else {
        var strLastName;
        var strFirstName;
        if (strName.length == 47) {
            strLastName = strName.substr(0, strName.length - 3);
        } else {
            strLastName = strName.substr(0, strName.length - 2);
        }
        strFirstName = strLastName + "1_";
        strLastName = strLastName + (parseInt(intLoop) - 1) + "_";
        strID = document.getElementById(strLastName + "HidSelRowID").value;
    }
    if (strID.length > 0) {
        if (strID.length == 1) { strID = "0" + strID; }
        var intStartID = 2;
        var intj;
        if (document.getElementById(strName + "ChkUnitList").checked) {
            if (intLoop == 1) {//ベースの値を全ての明細へ反映する
                for (intj = intStartID; intj <= intStartID + 20; intj++) {
                    if (intj > 10) {
                        strName = strName.substr(0, strName.length - 3) + parseInt(intj) + "_";
                    } else {
                        strName = strName.substr(0, strName.length - 2) + parseInt(intj) + "_";
                    }
                    if (document.getElementById(strName + "GVPrice")) {
                        fncGridClick(strName, strName + "GVPrice_ctl" + strID, 2, 2);
                        document.getElementById(strName + "ChkUnitList").checked = false;
                    }
                }
            } else {//直前の明細の値を反映する 
                fncGridClick(strName, strName + "GVPrice_ctl" + strID, 2, 2);
                document.getElementById(strFirstName + "ChkUnitList").checked = false;
            }
        }
    }
}
/*******************************************************************************
【関数名】  : f_ISOTanka_File
【概要】    : ISO単価のファイル出力
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function f_ISOTanka_File(strName) {
    //コントロール名
    var strChildName = "";
    //入力情報
    var strInput = "";
    //価格リスト情報
    var strPriceList = "";
    
    //保存した情報をクリア
    document.getElementById(strName + "HidPriceForFile").value = "";
    //価格リストをクリア
    document.getElementById(strName + "HidPriceList").value = "";
    
    for (intj = 1; intj <= 20; intj++) {
        strChildName = strName + "ISODetail" + intj + "_";
        if (document.getElementById(strChildName + "txt_Rate")) {
            
            //"_"で各部品情報を区切する
            if (strInput.length > 0) {
                strInput = strInput + "_";
                strPriceList = strPriceList + "_"; 
             }
            
            //入力情報の取得
            strInput += f_ISOSetSelectInfo(strChildName);
            strPriceList += f_ISOSetPriceList(strChildName);
        }
    }

    //選択情報の保存
    document.getElementById(strName + "HidPriceForFile").value = strInput;
    //価格リストの保存
    document.getElementById(strName + "HidPriceList").value = strPriceList;
}

/*******************************************************************************
【関数名】  : f_SetSelectInfo
【概要】    : 画面から入力した情報を取得して保存
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function f_ISOSetSelectInfo(strName) {
    //掛率、単価、数量、金額、消費税と合計
    var strSelectInfo;
    //掛率    
    strSelectInfo = document.getElementById(strName + "txt_Rate").value + "|";
    //単価
    strSelectInfo = strSelectInfo + document.getElementById(strName + "txt_UnitPrc").value + "|";
    //数量
    strSelectInfo = strSelectInfo + document.getElementById(strName + "txt_Amount").value + "|";
    //金額
    strSelectInfo = strSelectInfo + document.getElementById(strName + "txt_Price").value;
    //消費税と合計
    if (document.getElementById(strName + "txt_Tax")) {
        strSelectInfo = strSelectInfo + "|" + document.getElementById(strName + "txt_Tax").value + "|";
        strSelectInfo = strSelectInfo + document.getElementById(strName + "txt_Total").value;
    } else {
        strSelectInfo += "|" + "|";
    }

    return strSelectInfo;
}

/*******************************************************************************
【関数名】  : f_SetPriceList
【概要】    : 価格リストを取得して保存
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function f_ISOSetPriceList(strName) {
    //価格リスト
    var list = "";

    $("#" + strName + "GVPrice" + " tr").each(function () {
        //if (!this.rowIndex) return;

        var firstColumn = $(this).find("td:first");
        //タイトル
        var title = firstColumn.html();
        //価格
        var price = firstColumn.next().html();
        //価格区分
        var columnKbn = $(this).find("td:last").children(0).val();

        list += columnKbn + ":" + title + ":" + price + "|";

    });

    //価格リストの保存
    return list;

}

/*******************************************************************************
【関数名】  : isNumber
【概要】    : 数字判断
【引数】    : x         <Object>    入力値
【戻り値】  : 無し
*******************************************************************************/
function isNumber(x) {
    if (typeof (x) != 'number' && typeof (x) != 'string')
        return false;
    else
        return (x == parseFloat(x) && isFinite(x));
}

/*******************************************************************************
【関数名】  : fncPrcLstOnchange
【概要】    : 金額を再計算する
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncPrcLstOnchange(strName) {
    var decPrcLst = document.getElementById(strName + "SelUnitValue").value;
    var decCurLst = document.getElementById(strName + "SelCurrValue").value;
    var decRate = 1;  //掛率は1.000に戻す
    var strAmount = "";
    var decUntPrc1 = 0;
    var strParentName;
    if (strName.length == 47) {
        strParentName = strName.substr(0, strName.length - 12);
    } else {
        strParentName = strName.substr(0, strName.length - 11);
    }

    //受注EDI連携は変更不可（非表示のため）
    if (document.getElementById(strParentName + "strHiddenKbn").value == "1") { return true; }

    decUntPrc1 = decPrcLst * decRate;
    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
        //単価①    = 単価ﾘｽﾄ * 掛率(繰り上げ)
        document.getElementById(strName + "txt_UnitPrc").value = fncSetComma(f_JudgeRound("fncRoundUp", String(decUntPrc1), f_DecLen(0, decCurLst), ".", decCurLst));
        //単価②    = 単価ﾘｽﾄ * 掛率(小数第2位四捨五入)
        document.getElementById(strName + "txt_DtlPrc").value = fncSetComma(fncRound(String(decUntPrc1), f_DecLen(1, decCurLst), "."));
        //数量取得
        strAmount = fncRemoveComma(document.getElementById(strName + "txt_Amount").value);
        //掛率
        document.getElementById(strName + "txt_Rate").value = fncSetComma(fncRoundUp(String(decRate), "3", "."));
    } else {
        //単価①    = 単価ﾘｽﾄ * 掛率(繰り上げ)
        document.getElementById(strName + "txt_UnitPrc").value = fncSetDot(f_JudgeRound("fncRoundUp", String(decUntPrc1), f_DecLen(0, decCurLst), ",", decCurLst));
        //単価②    = 単価ﾘｽﾄ * 掛率(小数第2位四捨五入)
        document.getElementById(strName + "txt_DtlPrc").value = fncSetDot(fncRound(String(decUntPrc1), f_DecLen(1, decCurLst), ","));
        //数量取得
        strAmount = fncRemoveDot(document.getElementById(strName + "txt_Amount").value);
        //掛率
        document.getElementById(strName + "txt_Rate").value = fncSetDot(fncRoundUp(String(decRate), "3", ","));
    }
    //数量が入っていたら
    if (!(strAmount == "")) { fncRowCal(strName); }
}

/*******************************************************************************
【関数名】  : fncRateOnchange
【概要】    : 金額を再計算する
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncRateOnchange(strName) {
    var decPrcLst = document.getElementById(strName + "SelUnitValue").value;
    var decCurLst = document.getElementById(strName + "SelCurrValue").value;
    var decRate = document.getElementById(strName + "txt_Rate").value;
    var strAmount = "";
    var decUntPrc1 = 0;
    var strParentName;
    if (strName.length == 47) {
        strParentName = strName.substr(0, strName.length - 12);
    } else {
        strParentName = strName.substr(0, strName.length - 11);
    }

    if (decRate == "") {
        decRate = 0;
        //カンマ・ドット編集
        if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
            document.getElementById(strName + "txt_Rate").value = fncSetComma(fncRoundUp(String(decRate), "3", "."));
        } else {
            document.getElementById(strName + "txt_Rate").value = fncSetDot(fncRoundUp(String(decRate), "3", ","));
        }
    }
    if (isNaN(decRate)) { return false; }
    //カンマとドットを取る(例：小数点区分が"."(ドット)の場合、"1,009.120"→"1009.120")
    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
        decRate = fncRemoveComma(decRate);
    } else {
        decRate = fncRemoveDot(decRate);
    }
    //単価① = 単価ﾘｽﾄ * 掛率(繰り上げ)
    decUntPrc1 = decPrcLst * decRate;
    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
        document.getElementById(strName + "txt_UnitPrc").value = fncSetComma(f_JudgeRound("fncRoundUp", String(decUntPrc1), f_DecLen(0, decCurLst), ".", decCurLst));
        //単価②    = 単価ﾘｽﾄ * 掛率(小数第2位四捨五入)
        document.getElementById(strName + "txt_DtlPrc").value = fncSetComma(fncRound(String(decUntPrc1), f_DecLen(1, decCurLst), "."));
        strAmount = fncRemoveComma(document.getElementById(strName + "txt_Amount").value);
    } else {
        document.getElementById(strName + "txt_UnitPrc").value = fncSetDot(f_JudgeRound("fncRoundUp", String(decUntPrc1), f_DecLen(0, decCurLst), ",", decCurLst));
        //単価②    = 単価ﾘｽﾄ * 掛率(小数第2位四捨五入)
        document.getElementById(strName + "txt_DtlPrc").value = fncSetDot(fncRound(String(decUntPrc1), f_DecLen(1, decCurLst), ","));
        strAmount = fncRemoveDot(document.getElementById(strName + "txt_Amount").value);
    }
    //数量が入っていたら
    if (!(strAmount == "")) { fncRowCal(strName); }
    //項目の「合計」に値が入っている場合、それらを合計したものを「縦合計」の「金額」「合計」にセットする
    fncVertCal(strName);
}

/*******************************************************************************
【関数名】  : fncAmountOnchange
【概要】    : 金額を再計算する
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncAmountOnchange(strName) {
    var decCurLst = document.getElementById(strName + "SelCurrValue").value;
    var aryItemCnt = new Array();
    var strUntPrc = "";
    var strAmount = "";
    var strPrice = "";
    var strTax = "";
    var strTotal = "";
    var strParentName;
    var strChildID = strName.substr(0, strName.length - 2);
    if (strName.length == 47) {
        strParentName = strName.substr(0, strName.length - 12);
    } else {
        strParentName = strName.substr(0, strName.length - 11);
    }

    //項目ごとの数量を配列に格納する
    aryItemCnt = document.getElementById(strParentName + "intItemCnt").value.split("|");
    //変更した項目の数量
    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
        strAmount = fncRemoveComma(document.getElementById(strName + "txt_Amount").value);
    } else {
        strAmount = fncRemoveDot(document.getElementById(strName + "txt_Amount").value);
    }
    if (strAmount == "") {
        strAmount = 0;
        document.getElementById(strName + "txt_Amount").value = strAmount;
    }
    if (isNaN(strAmount)) { return false; }
    //他の項目にも数量を反映する
    if (strName.substr(strName.length - 2, 1) == "1") {
        for (DtlRow = 1; DtlRow < aryItemCnt.length; DtlRow++) {

            //項目単価
            if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
                strUntPrc = fncRemoveComma(document.getElementById(strChildID + DtlRow + "_txt_UnitPrc").value);
            } else {
                strUntPrc = fncRemoveDot(document.getElementById(strChildID + DtlRow + "_txt_UnitPrc").value);
            }

            if (strUntPrc == "") {
                strUntPrc = 0;
                if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
                    document.getElementById(strChildID + DtlRow + "_txt_UnitPrc").value = fncRoundUp("0", f_DecLen(0, decCurLst), ".");
                } else {
                    document.getElementById(strChildID + DtlRow + "_txt_UnitPrc").value = fncRoundUp("0", f_DecLen(0, decCurLst), ",");
                }
            }
            //数量設定
            document.getElementById(strChildID + DtlRow + "_txt_Amount").value = parseInt(strAmount) * aryItemCnt[DtlRow];
            //金額設定
            if (strUntPrc != 0) {
                strPrice = Number(strUntPrc) * parseInt(strAmount) * aryItemCnt[DtlRow];
                strTax = fncRoundDown(Number(strPrice) * 0.08, f_DecLen(0, decCurLst), ".");
                strTotal = Number(strPrice) + Number(strTax);
                //海外代理店は消費税を表示しない
                if (document.getElementById(strParentName + "strHiddenKbn").value == "2") {
                    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
                        document.getElementById(strChildID + DtlRow + "_txt_Price").value = fncSetComma(fncRound(String(strPrice), f_DecLen(0, decCurLst), "."));
                    } else {
                        document.getElementById(strChildID + DtlRow + "_txt_Price").value = fncSetDot(fncRound(String(strPrice), f_DecLen(0, decCurLst), ","));
                    }
                } else {
                    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
                        document.getElementById(strChildID + DtlRow + "_txt_Price").value = fncSetComma(fncRound(String(strPrice), f_DecLen(0, decCurLst), "."));
                        document.getElementById(strChildID + DtlRow + "_txt_Tax").value = fncSetComma(String(strTax), f_DecLen(0, decCurLst), ".");
                        document.getElementById(strChildID + DtlRow + "_txt_Total").value = fncSetComma(fncRound(String(strTotal), f_DecLen(0, decCurLst), "."));
                    } else {
                        document.getElementById(strChildID + DtlRow + "_txt_Price").value = fncSetDot(fncRound(String(strPrice), f_DecLen(0, decCurLst), ","));
                        document.getElementById(strChildID + DtlRow + "_txt_Tax").value = fncSetDot(String(strTax), f_DecLen(0, decCurLst), ".");
                        document.getElementById(strChildID + DtlRow + "_txt_Total").value = fncSetDot(fncRound(String(strTotal), f_DecLen(0, decCurLst), ","));
                    }
                }
            }
        }
    }
    //項目の「合計」に値が入っている場合、それらを合計したものを「縦合計」の「金額」「合計」にセットする
    fncVertCal(strName);
}

/*******************************************************************************
【関数名】  : fncUntPrcOnchange
【概要】    : 金額を再計算する
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncUntPrcOnchange(strName) {
    var strPrcLst = document.getElementById(strName + "SelUnitValue").value;
    var decCurLst = document.getElementById(strName + "SelCurrValue").value;
    var strUntPrc = "";
    var strAmount = "";
    var strRate;
    var strParentName;
    if (strName.length == 47) {
        strParentName = strName.substr(0, strName.length - 12);
    } else {
        strParentName = strName.substr(0, strName.length - 11);
    }

    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
        strUntPrc = fncRemoveComma(document.getElementById(strName + "txt_UnitPrc").value);
        strAmount = fncRemoveComma(document.getElementById(strName + "txt_Amount").value);
    } else {
        strUntPrc = fncRemoveDot(document.getElementById(strName + "txt_UnitPrc").value);
        strAmount = fncRemoveDot(document.getElementById(strName + "txt_Amount").value);
    }
    if (strUntPrc == "") {
        strUntPrc = 0;
        if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
            document.getElementById(strName + "txt_UnitPrc").value = fncSetComma(f_JudgeRound("fncRoundUp", String(strUntPrc), f_DecLen(0, decCurLst), ".", decCurLst));
        } else {
            document.getElementById(strName + "txt_UnitPrc").value = fncSetComma(f_JudgeRound("fncRoundUp", String(strUntPrc), f_DecLen(0, decCurLst), ",", decCurLst));
        }
    }
    if (isNaN(strUntPrc)) { return false; }

    //単価② = 単価①
    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
        document.getElementById(strName + "txt_DtlPrc").value = fncSetComma(fncRound(String(strUntPrc), f_DecLen(1, decCurLst), "."));
    } else {
        document.getElementById(strName + "txt_DtlPrc").value = fncSetDot(fncRound(String(strUntPrc), f_DecLen(1, decCurLst), ","));
    }
    //掛率 = 単価① / 単価ﾘｽﾄ
    if (strPrcLst == "0" || strPrcLst == "") {
        strRate = "0";
    } else {
        strRate = Number(strUntPrc) / Number(strPrcLst);
    }
    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
        document.getElementById(strName + "txt_Rate").value = fncSetComma(fncRoundUp(String(strRate), "3", "."));
    } else {
        document.getElementById(strName + "txt_Rate").value = fncSetDot(fncRoundUp(String(strRate), "3", ","));
    }
    //数量が入っていたら
    if (!(strAmount == "")) { fncRowCal(strName); }
}

/*******************************************************************************
【関数名】  : fncRowCal
【概要】    : 数量が入力イベント
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncRowCal(strName) {
    var decCurLst = document.getElementById(strName + "SelCurrValue").value;
    var strUntPrc = "";
    var strAmount = "";
    var strParentName;
    if (strName.length == 47) {
        strParentName = strName.substr(0, strName.length - 12);
    } else {
        strParentName = strName.substr(0, strName.length - 11);
    }

    //受注EDI連携は変更不可（非表示のため）
    if (document.getElementById(strParentName + "strHiddenKbn").value == "1") { return true; }
    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
        strUntPrc = fncRemoveComma(document.getElementById(strName + "txt_UnitPrc").value);
        strAmount = fncRemoveComma(document.getElementById(strName + "txt_Amount").value);
    } else {
        strUntPrc = fncRemoveDot(document.getElementById(strName + "txt_UnitPrc").value);
        strAmount = fncRemoveDot(document.getElementById(strName + "txt_Amount").value);
    }
    if (strUntPrc == "") {
        strUntPrc = 0;
        document.getElementById(strName + "txt_UnitPrc").value = strUntPrc;
    }
    if (strAmount == "") {
        strAmount = 0;
        document.getElementById(strName + "txt_Amount").value = strAmount;
    }
    if (isNaN(strAmount)) { return false; }
    //金額 = 単価① * 数量
    var strPrice = Number(strUntPrc) * Number(strAmount);
    if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
        document.getElementById(strName + "txt_Price").value = fncSetComma(fncRound(String(strPrice), f_DecLen(0, decCurLst), "."));
    } else {
        document.getElementById(strName + "txt_Price").value = fncSetDot(fncRound(String(strPrice), f_DecLen(0, decCurLst), ","));
    }
    //海外代理店は消費税を表示しない
    if (document.getElementById(strParentName + "strHiddenKbn").value == "2") {
    } else {
        //消費税 = 金額 * 0.08
        var strTax = fncRoundDown(Number(strPrice) * 0.08, f_DecLen(0, decCurLst), ".");
        if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
            document.getElementById(strName + "txt_Tax").value = fncSetComma(strTax, f_DecLen(0, decCurLst), ".");
        } else {
            document.getElementById(strName + "txt_Tax").value = fncSetDot(strTax, f_DecLen(0, decCurLst), ",");
        }
        //合計 = 単価① * 数量
        var strTotal = fncRound(Number(strPrice) + Number(strTax), f_DecLen(0, decCurLst), ".");
        if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
            document.getElementById(strName + "txt_Total").value = fncSetComma(fncRound(String(strTotal), f_DecLen(0, decCurLst), "."));
        } else {
            document.getElementById(strName + "txt_Total").value = fncSetDot(fncRound(String(strTotal), f_DecLen(0, decCurLst), ","));
        }
    }
    //項目の「合計」に値が入っている場合、それらを合計したものを「縦合計」の「金額」「消費税」「合計」にセットする
    fncVertCal(strName);
}

/*******************************************************************************
【関数名】  : fncVertCal
【概要】    : 合計が入力イベント
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncVertCal(strName) {
    var strChildID;
    var strTotalPrc = "0";
    var strTotalTotal = "0";
    var strPricePrc = "0";
    var strPriceTotal = "0";
    var strTaxPrc = "0";
    var strTaxTotal = "0";
    var strParentName;
    if (strName.length == 47) {
        strParentName = strName.substr(0, strName.length - 12);
        strChildID = strName.substr(0, strName.length - 3);
    } else {
        strParentName = strName.substr(0, strName.length - 11);
        strChildID = strName.substr(0, strName.length - 2);
    }
    var decCurLst = document.getElementById(strChildID + "1" + "_SelCurrValue").value;

    //海外代理店は消費税を表示しない
    if (document.getElementById(strParentName + "strHiddenKbn").value == "2") {
        //縦金額 = 項目ごとの金額の合計
        for (i = 1; i <= parseInt(document.getElementById(strParentName + "intItemRow").value); i++) {
            if ((document.getElementById(strChildID + i + "_txt_Price").value != null) &&
               (document.getElementById(strChildID + i + "_txt_Price").value != "")) {

                //項目ごとの合計のカンマ・ドットを取る
                if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
                    strPricePrc = fncRemoveComma(document.getElementById(strChildID + i + "_txt_Price").value);
                } else {
                    strPricePrc = fncRemoveDot(document.getElementById(strChildID + i + "_txt_Price").value);
                }
                //縦金額 = 縦金額 + 項目ごとの合計
                strPriceTotal = Number(strPriceTotal) + Number(strPricePrc);
                if (decCurLst != document.getElementById(strChildID + i + "_SelCurrValue").value) {
                    decCurLst = "";
                    strTotalTotal = "0";
                }
            }
        }
        if (strPriceTotal == "0") {
            document.getElementById(strParentName + "txt_AmtPrice").value = "";
        } else {
            //計算した縦合計をセットする
            if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
                document.getElementById(strParentName + "txt_AmtPrice").value = fncSetComma(fncRound(strPriceTotal, f_DecLen(0, decCurLst), "."));
            } else {
                //
                document.getElementById(strParentName + "txt_AmtPrice").value = fncSetDot(fncRound(strPriceTotal, f_DecLen(0, decCurLst), ","));
            }
        }
    } else {
        //縦金額 = 項目ごとの金額の合計
        for (i = 1; i <= parseInt(document.getElementById(strParentName + "intItemRow").value); i++) {
            if ((document.getElementById(strChildID + i + "_txt_Total").value != null) &&
               (document.getElementById(strChildID + i + "_txt_Total").value != "")) {

                //項目ごとの合計のカンマ・ドットを取る
                if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
                    strTotalPrc = fncRemoveComma(document.getElementById(strChildID + i + "_txt_Total").value);
                    strPricePrc = fncRemoveComma(document.getElementById(strChildID + i + "_txt_Price").value);
                    strTaxPrc = fncRemoveComma(document.getElementById(strChildID + i + "_txt_Tax").value);
                } else {
                    strTotalPrc = fncRemoveDot(document.getElementById(strChildID + i + "_txt_Total").value);
                    strPricePrc = fncRemoveDot(document.getElementById(strChildID + i + "_txt_Price").value);
                    strTaxPrc = fncRemoveDot(document.getElementById(strChildID + i + "_txt_Tax").value);
                }
                //縦金額 = 縦金額 + 項目ごとの合計
                strTotalTotal = Number(strTotalTotal) + Number(strTotalPrc);
                strPriceTotal = Number(strPriceTotal) + Number(strPricePrc);
                strTaxTotal = Number(strTaxTotal) + Number(strTaxPrc);

                if (decCurLst != document.getElementById(strChildID + i + "_SelCurrValue").value) {
                    decCurLst = "";
                    strTotalTotal = "0";
                    strPriceTotal = "0";
                    strTaxTotal = "0";
                }
            }
        }
        if (strTotalTotal == "0") {
            document.getElementById(strParentName + "txt_AmtPrice").value = "";
            document.getElementById(strParentName + "txt_SumTotal").value = "";
            document.getElementById(strParentName + "txt_AmtTax").value = "";
        } else {
            //計算した縦合計をセットする
            if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
                document.getElementById(strParentName + "txt_AmtPrice").value = fncSetComma(fncRound(strPriceTotal, f_DecLen(0, decCurLst), "."));
                document.getElementById(strParentName + "txt_SumTotal").value = fncSetComma(fncRound(strTotalTotal, f_DecLen(0, decCurLst), "."));
                document.getElementById(strParentName + "txt_AmtTax").value = fncSetComma(fncRound(strTaxTotal, f_DecLen(0, decCurLst), "."));
            } else {
                document.getElementById(strParentName + "txt_AmtPrice").value = fncSetDot(fncRound(strPriceTotal, f_DecLen(0, decCurLst), ","));
                document.getElementById(strParentName + "txt_SumTotal").value = fncSetDot(fncRound(strTotalTotal, f_DecLen(0, decCurLst), ","));
                document.getElementById(strParentName + "txt_AmtTax").value = fncSetDot(fncRound(strTaxTotal, f_DecLen(0, decCurLst), ","));
            }
        }
    }
}
