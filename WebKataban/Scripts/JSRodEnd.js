/*******************************************************************************
【関数名】  : f_OnBlur
【概要】    : テキストボックス・ドロップダウンonblur処理(Mainフレーム)
【引数】    : パターンNo.
: 小数点区分
【戻り値】  : 無し
*******************************************************************************/
function f_RodEnd_OnBlur(strName, EditDiv) {

    var intKKCnt;
    var ArySltKK;
    var AryActSltKK;
    var strRowKK;
    var strRowA;
    var strRowC;
    var strKKVal;
    var strAVal;

    /* KK寸法・A寸法情報取得 */
    ArySltKK = document.getElementById(strName + "_HdnSltKK").value.split(",");
    AryActSltKK = document.getElementById(strName + "_HdnActSltKK").value.split(",");
    strStdKK = document.getElementById(strName + "_HdnStdKK").value;
    strStdA = document.getElementById(strName + "_HdnStdA").value;
    strStdC = document.getElementById(strName + "_HdnStdC").value;
    strRowKK = document.getElementById(strName + "_HdnRowKK").value;
    strRowA = document.getElementById(strName + "_HdnRowA").value;
    strRowC = document.getElementById(strName + "_HdnRowC").value;
    strKKVal = document.getElementById(strName + "_Prod" + strRowKK).value;
    strAVal = document.getElementById(strName + "_Prod" + strRowA).value;

    /* C寸法計算 */
    if (strKKVal != null && strKKVal != "") {
        //KK寸法に値が入っていた場合
        for (i = 1; i < ArySltKK.length; i++) {
            if (ArySltKK[i] == strKKVal) {
                intKKCnt = i;
                break;
            }
        }
        //A寸法に値が入っていた場合
        if (strAVal != null && strAVal != "") {
            document.getElementById(strName + "_Prod" + strRowC).value = parseFloat(strAVal) - parseFloat(AryActSltKK[intKKCnt]);
        } else {

            document.getElementById(strName + "_Prod" + strRowC).value = parseFloat(strStdA) - parseFloat(AryActSltKK[intKKCnt]);
        }
    } else if (strAVal != null && strAVal != "") {
        //A寸法のみに値が入っていた場合
        document.getElementById(strName + "_Prod" + strRowC).value = parseFloat(strAVal) - parseFloat(strStdKK);
    } else {
        //どちらにも値が入っていなかった場合
        document.getElementById(strName + "_Prod" + strRowC).value = parseFloat(strStdC);
    }
}

/*******************************************************************************
【関数名】  : f_OKOnClick
【概要】    : [OK]ボタンonclick処理(Bottomフレーム)
【引数】    : 無し
【戻り値】  : 無し
*******************************************************************************/
function f_RodEnd_OK(strName, strName_child) {
    var RdoId;
    var intListCnt;
    var strProdSize = "";
    var objList;
    var objOtherText;

    /* メインフレームの情報をボトムフレームにセットする */
    for (intLoopCnt = 1; intLoopCnt < parseInt(document.getElementById(strName + "_HdnPtnCnt").value) + 1; intLoopCnt++) {
        // 選択されているパターンの情報をセットする
        if (document.getElementById(strName + "_Rdo" + intLoopCnt).checked) {
            //寸法表情報
            objList = document.getElementById(strName_child + intLoopCnt + "_TblLst");
            if (objList) {
                //寸法表の行数
                intListCnt = document.getElementById(strName_child + intLoopCnt + "_TblLst").rows.length;
                //特注寸法
                for (intLoopCnt2 = 1; intLoopCnt2 < intListCnt; intLoopCnt2++) {
                    strProdSize = strProdSize + '|' + document.getElementById(strName_child + intLoopCnt + "_Prod" + intLoopCnt2).value;
                }
                document.getElementById(strName + "_HdnSelProdSize").value = strProdSize;
            }
            //その他寸法情報
            objOtherText = document.getElementById(strName_child + intLoopCnt + "_OtherProd");
            if (objOtherText != null) {
                //特注寸法
                strProdSize = strProdSize + '|' + document.getElementById(strName_child + intLoopCnt + "_OtherProd").value;
                document.getElementById(strName + "_HdnSelProdSize").value = strProdSize;
            }
        }
    }
}

/*******************************************************************************
【関数名】  : fncDispErrMsg
【概要】    : エラーメッセージ表示
【引数】    : objText   テキストボックス
【戻り値】  : 無し
*******************************************************************************/
function fncDispErrMsg(strName) {
    var strMsg = document.getElementById(strName + "HidMessage").value;
    alert(strMsg);
}