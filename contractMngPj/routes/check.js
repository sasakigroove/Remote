

const mysql2 = require('mysql2/promise');

var HashMap = require('hashmap');


//Mysqlのデータベース設定情報
var mysql_setting = {
    host: 'localhost',
    user: 'root',
    password: 'root',
    database: 'contractPj'

}
require('date-utils');
var hmmsg = new HashMap();
var errorMsgList = new HashMap();

//日付け妥当性チェックと稼働開始日より稼働終了日の方が日付が後かのチェック
exports.check = function ( screenParam) {

    console.log("チェックの方にきたよ");

    var countline = screenParam.get("checklinecount");
    var checkRowArray = countline.split(",");

    var checkcount = screenParam.get("checkcount");
    

    var errCnt = 0;
    var check;
    aftercheck = 4;

    if (checkcount != 0) {

        for (var i = 0; i < checkcount; i++) {

            var lineParam = screenParam.get("lineParam"+i);
            var lineNo = lineParam.get("lineNo"+i);

            var endyear = lineParam.get("endYear" + i);
            var endmonth = lineParam.get("endMonth" + i);
            var endday = lineParam.get("endDay" + i);
            var enddate = endyear + "/" + endmonth + "/" + endday;
            //日付妥当性チェック
            var y = enddate.split("/")[0];
            var m = enddate.split("/")[1] - 1;
            var d = enddate.split("/")[2];
            var date = new Date(y, m, d);

            if (date.getFullYear() != y || date.getMonth() != m || date.getDate() != d) {
                console.log("日付不正と判断");
                errorMsgList.set("msg"+i,(lineNo+1)+"行目で稼働終了日の日付が不正です");
                errCnt++;
                check = 2;
            } else {
                console.log("正常");
                var startdate = lineParam.get("startdate" + i);
                var end = endyear + "年" + endmonth + "月" + endday + "日";
                //稼働開始日より稼働終了日の方があとか
                if (startdate <= end) {
                    msg = "";
                    //稼働終了日が今現在の日付より前かの判定
                    var dt = new Date();
                    var nowdate = dt.toFormat("YYYYMMDD");
                    var enddate = endyear + endmonth + endday;

                    if (nowdate > enddate) {
                        errorMsgList.set("msg"+i,lineNo+"行目の稼働終了日が現在日付より前の日付となっています。");
                        errCnt++; 
                    } else {
                        // 正常
                        check = 4;
                    }

                } else {
                    errorMsgList.set("msg"+i,lineNo+"行目の稼働終了日が稼働開始日より前の日付となっています。");
                    errCnt++;
                    check = 2;
                }

            }



            //エラーか
            if (check == aftercheck) {
                // 同じだったら何もしなくていい
                // 結果が変わらないから
            } else {
                if (aftercheck == 2) {
                    //元々がエラーだったら無視
                } else {
                    // 元々が正常か警告
                    if (aftercheck == 1) {
                        //元々が警告 

                        if (check == 2) {
                            // 読み込んだ分がエラー
                            aftercheck = 2;
                        } else {
                            // 元々が警告で、読み込んだ分が正常 or 警告なのでなにもしない
                        }
                    } else if (aftercheck == 4) {
                        // 元々が正常
                        if (check == 2) {
                            // 読み込んだ分がエラー
                            aftercheck = 2;
                        } else if (check == 1) {
                            // 元々が正常で、読み込んだ分が警告
                            aftercheck = 1;
                        } else {
                            // 元々が正常で、読み込んだ分が正常なのでなにもしない
                        }
                    }
                }
            }
        }
    } else {
        aftercheck = 3;
    }

    hmmsg.set("errorCount",errCnt);
    hmmsg.set("errorMsg",errorMsgList);
    hmmsg.set("checkResultNo", aftercheck);

    return hmmsg;

}






