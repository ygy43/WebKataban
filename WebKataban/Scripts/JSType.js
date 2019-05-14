/*******************************************************************************
【関数名】  : TypeDblClick
【概要】    : リストダブルクリックイベント
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function TypeDblClick(strName) {
    document.getElementById(strName + "btnOK").click();
}

/*******************************************************************************
【関数名】  : KatabanKeyDown
【概要】    : Enterキーイベント
【引数】    : strName         <String>    画面名
【引数】    : strID           <String>    画面ID
【戻り値】  : 無し
*******************************************************************************/
function KatabanKeyDown(e, strName, strID) {
    switch (e.keyCode) {
        case 13:
            document.getElementById(strName + "Button4").focus();
            event.Returnvalue = false;
            return false;
        default:
            break;
    }
}