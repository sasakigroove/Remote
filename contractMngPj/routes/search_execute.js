const mysql2 = require('mysql2/promise');
var mysql_setting = {
  host: 'localhost',
  user: 'root',
  password: 'root',
  database: 'contractPj'

}
require('date-utils');
let now = new Date();
let mycon = null;
var HashMap = require('hashmap');

//初期表示か絞り込みかソートのみか絞り込み＋ソートかを判定して検索するメソッド									
exports.searchcheck = async function (
  keyword,                            //ジャンルとキーワードで検索のキーワード
  period_radio,                       //稼働期間で絞り込みのラジオボタン
  genre_pulldown,                     //ジャンルとキーワードで検索のプルダウン
  sheet_pulldown,                     //シートで検索のプルダウン
  sort_button,                        //検索結果th部分のソートボタン　どのボタンが押されたか判定する用
  start_before_year,                  //稼働開始日で絞り込み年（前）
  start_before_month,                 //稼働開始日で絞り込み月（前）
  start_after_year,                   //稼働開始日で絞り込み年（後）
  start_after_month,                  //稼働終了日で絞り込み月（後）
  end_before_year,                    //稼働終了日で絞り込み年（前）
  end_before_month,                   //稼働終了日で絞り込み月（前）
  end_after_year,                     //稼働終了日で絞り込み年（後）
  end_after_month,                    //稼働終了日で絞り込み月（後）
  sort1, sort2, sort3, sort4, 
  sort5, sort6, sort7, sort8,
  sort9, sort10, sort11, sort12,      //各ソートボタンに画面でASCとDESCの保持をしておく用
  sort13, sort14, sort15,
  sort16, sort17, sort18, sort19,
  errorcount,                           //
  excel_before_year,
  excel_before_month,
  excel_after_year,
  excel_after_month,
  gamenhantei
) {

  var hshMap = new HashMap();
  var [rowswork] = "";                  //検索結果rowsを移し替えるための変数
  var where = "";                       //検索条件が入る
  var summaryWhere = "";                       //検索条件が入る
  var sort = "";                        //ソート条件が入る
  var fromY = 0;                        //稼働開始日・稼働終了日・Excel取り込み年月の絞り込み年（前）が入る
  var fromM = 0;                        //稼働開始日・稼働終了日・Excel取り込み年月の絞り込み月（前）が入る
  var toY = 0;                          //稼働開始日・稼働終了日・Excel取り込み年月の絞り込み年（後）が入る
  var toM = 0;                          //稼働開始日・稼働終了日・Excel取り込み年月の絞り込み年月（後）が入る
  var beforeymd = 0;                    //fromYとfromMを合体させたものが入る
  var afterymd = 0;                     //toYとtoMを合体させたものが入る
  var excel_hantei = 0;                 //YYYY年MM月DD日の形で検索する稼働期間とYYYY/MMの形で検索するExcel取り込み年月を仕分けるための変数                 
  var groupby = "";                     //集計情報を提示する検索条件には入っている
  var summary_yymm = "";                //集計情報を出すための検索条件に入る年月範囲
  var nengatu = "";                     //画面にYYYY年MM月の形で出す文字列を作成する用の変数
  var tougetu_hantei = 0;               //集計情報を出す際、当月で絞り込まれていた場合は当月を取得するため、当月で絞り込まれているかどうかを判定するための変数
  var where_all =0;                     //すべての履歴を表示が選択されたときにwhereが空になってしまいwhereに初期表示用の検索条件が入るのを防止するための変数
  var nameMap = new HashMap();            
  var eigyoudeta = new HashMap();
  var nengetu_hsh = new HashMap();
  var summary_hsh = new HashMap();
  var excel_hantei_summary =0;          //指定なしでシートとExcel取り込み年月を選択したときに当月ではなくExcel取り込み年月で絞り込んだ範囲を


  if (period_radio == "6" || period_radio == "2" || excel_before_year != "0" || excel_after_year != "0" ||
    excel_before_month != "0" || excel_after_month != "0") {

    if (period_radio == "2") {
      if (start_before_year != "0" && start_before_month != "0" && start_after_year != "0" && start_before_month != "0") {
        //指定された稼働開始日～指定された稼働開始日の年月比較チェック
        fromY = start_before_year;
        fromM = start_before_month;
        toY = start_after_year;
        toM = start_after_month;

      } else if (start_before_year != "0" && start_before_month != "0") {
        fromY = start_before_year;
        fromM = start_before_month;

      } else if (start_after_year != "0" && start_after_month != "0") {
        toY = start_after_year;
        toM = start_after_month;

      } else {
        hshMap.set('msg', "稼働開始日の期間のプルダウンは年月の両方を指定してください");
        console.log("年月プルダウンが選択されていません")
      }

    } else if (period_radio == "6") {

      if (end_before_year != "0" && end_before_month != "0" && end_after_year != "0" && end_before_month != "0") {
        //指定された稼働終了日～指定された稼働終了日の年月比較チェック

        fromY = end_before_year;
        fromM = end_before_month;
        toY = end_after_year;
        toM = end_after_month;

      } else if (end_before_year != "0" && end_before_month != "0") {

        fromY = end_before_year;
        fromM = end_before_month;

      } else if (end_after_year != "0" && end_after_month != "0") {

        toY = end_after_year;
        toM = end_after_month;

      } else {
        hshMap.set('msg', "稼働開始日の期間のプルダウンは年月の両方を指定してください");
        console.log("年月プルダウンが選択されていません")
      }

    } else if (excel_before_year != "0" && excel_before_month != "0" && excel_after_year != "0" && excel_after_month != "0") {
      fromY = excel_before_year;
      fromM = excel_before_month;
      toY = excel_after_year;
      toM = excel_after_month;
      excel_hantei = 1;

    } else if (excel_before_year != "0" && excel_before_month != "0") {
      fromY = excel_before_year;
      fromM = excel_before_month;
      excel_hantei = 1;

    } else if (excel_after_year != "0" && excel_after_month != "0") {
      toY = excel_after_year;
      toM = excel_after_month;
      excel_hantei = 1;

    } else {
      hshMap.set('msg', "excel取り込み年月の期間のプルダウンは年月の両方を指定してください");
      console.log("年月プルダウンが選択されていません")
    }

    if (hshMap.get('msg') == null) {
      if (fromY != "0" && fromM != "0" && toY != "0" && toM != "0") {
        //日付比較チェック
        var fromDate = new Date(fromY, fromM);
        var toDate = new Date(toY, toM);

        if (fromDate.getFullYear() == toDate.getFullYear()) {
          // 同一年
          if (toDate.getMonth() < fromDate.getMonth()) {
            // toの方が大きい:NG
            hshMap.set('msg', "指定された年月に誤りがあります");
          } else {
            console.log("年と月が正常:OK");
          }
        } else if (fromDate.getFullYear() > toDate.getFullYear()) {
          // fromの方が上の年数:NG
          hshMap.set('msg', "指定された年月に誤りがあります");
        } else {
          // toの方が上の年数:OK
        }
      }

      //excel取り込み期間で絞り込みが選択されていてエラーも通過している場合は「１」が入る
      if (excel_hantei != 1) {
        if (fromY != "0" && fromM != "0") {
          //fromYとfromMを年月日に直す
          beforeymd = fromY + "年" + fromM + "月" + "1" + "日";
        }
        if (toY != "0" && toM != "0") {
          //月末の取得
         
          var fromDate = new Date(toY, toM, 0);

          afterymd =
            fromDate.getFullYear().toString()
            + "年" + (fromDate.getMonth() + 1).toString()
            + "月" + fromDate.getDate().toString()
            + "日";

        }
      } else if (excel_hantei == 1) {
        if (fromY != "0" && fromM != "0") {

          beforeymd = fromY + "/" + fromM;
        }
        if (toY != "0" && toM != "0") {

          afterymd = toY + "/" + toM;

        }

      }
    }
  }


  //---------------------------------------------------------------------------------------------------------
  //期間で絞り込みの検索条件作成
  if (period_radio != 0) {
    if (hshMap.get('msg') == null) {
      if (period_radio == "1") {
        //稼働開始当月
        if (where == "") {
          where = " WHERE DATE_FORMAT(LAST_DAY(NOW()),'%Y%m') = DATE_FORMAT(STR_TO_DATE(start_date,'%Y年%m月%d日'),'%Y%m')  ";
        } else {
          where = " AND DATE_FORMAT(LAST_DAY(NOW()),'%Y%m') = DATE_FORMAT(STR_TO_DATE(start_date,'%Y年%m月%d日'),'%Y%m')  ";
        }
        tougetu_hantei = 1;
        groupby = " GROUP BY a.year_month_date "
      } else if (period_radio == "2") {
        //指定された稼働開始日～指定された稼働開始日で検索
        if (beforeymd != "0" && afterymd != "0") {
          if (where == "") {
            where = where + " WHERE STR_TO_DATE(start_date,'%Y年%m月%d日') >= STR_TO_DATE('" + beforeymd + "','%Y年%m月%d日') AND STR_TO_DATE(start_date,'%Y年%m月%d日') <= STR_TO_DATE('" + afterymd + "','%Y年%m月%d日') ";
          } else {
            where = where + " STR_TO_DATE(start_date,'%Y年%m月%d日') >= STR_TO_DATE('" + beforeymd + "','%Y年%m月%d日') AND STR_TO_DATE(start_date,'%Y年%m月%d日') <= STR_TO_DATE('" + afterymd + "','%Y年%m月%d日') ";
          }
          
        } else if (beforeymd != "0") {
          //指定された稼働開始日～で検索
          if (where == "") {
            where = where + " WHERE STR_TO_DATE(start_date,'%Y年%m月%d日') >=STR_TO_DATE('" + beforeymd + "','%Y年%m月%d日') ";
          } else {
            where = where + " STR_TO_DATE(start_date,'%Y年%m月%d日') >=STR_TO_DATE('" + beforeymd + "','%Y年%m月%d日') ";
          }

        } else if (afterymd != "0") {
          //～指定された稼働開始日で検索
          if (where == "") {
            where = where + " WHERE STR_TO_DATE(start_date,'%Y年%m月%d日') <= STR_TO_DATE('" + afterymd + "','%Y年%m月%d日') ";
          } else {
            where = where + " STR_TO_DATE(start_date,'%Y年%m月%d日') <= STR_TO_DATE('" + afterymd + "','%Y年%m月%d日') ";
          }
        }

      } else if (period_radio == "3") {
        //稼働終了当月
        if (where == "") {
          where = " WHERE DATE_FORMAT(NOW(),'%Y%m') <= DATE_FORMAT(STR_TO_DATE(end_date,'%Y年%m月%d日'),'%Y%m')	AND DATE_FORMAT(STR_TO_DATE(end_date,'%Y年%m月%d日'),'%Y%m') <= DATE_FORMAT(NOW(),'%Y%m') 	";
        } else {
          where = " AND DATE_FORMAT(NOW(),'%Y%m') >= DATE_FORMAT(STR_TO_DATE(end_date,'%Y年%m月%d日'),'%Y%m')	AND DATE_FORMAT(STR_TO_DATE(end_date,'%Y年%m月%d日'),'%Y%m') <= DATE_FORMAT(NOW(),'%Y%m') ";
        }
        tougetu_hantei = 1;
        groupby = " GROUP BY a.year_month_date "
      } else if (period_radio == "4") {
        //稼働終了まで一か月+一週間未満
        if (where == "") {
          where = where + " WHERE STR_TO_DATE(end_date,'%Y年%m月%d日') <= DATE_ADD(DATE_ADD(STR_TO_DATE(NOW(),'%Y-%m-%d'),INTERVAL 1 MONTH),INTERVAL 7 DAY) ";
        } else {
          where = where + " AND STR_TO_DATE(end_date,'%Y年%m月%d日') <= DATE_ADD(DATE_ADD(STR_TO_DATE(NOW(),'%Y-%m-%d'),INTERVAL 1 MONTH),INTERVAL 7 DAY) ";
        }
      } else if (period_radio == "5") {
        //稼働終了まで一週間+1日未満
        if (where == "") {
          where = where + " WHERE STR_TO_DATE(end_date,'%Y年%m月%d日') <= DATE_ADD(DATE_ADD(STR_TO_DATE(NOW(),'%Y-%m-%d'),INTERVAL 1 WEEK),INTERVAL 1 DAY) ";
        } else {
          where = where + " AND STR_TO_DATE(end_date,'%Y年%m月%d日') <= DATE_ADD(DATE_ADD(STR_TO_DATE(NOW(),'%Y-%m-%d'),INTERVAL 1 WEEK),INTERVAL 1 DAY) ";
        }
      } else if (period_radio == "6") {
        //指定された稼働終了日～指定された稼働終了日で検索
        if (beforeymd != "0" && afterymd != "0") {

          if (where == "") {
            where = where + " WHERE STR_TO_DATE(end_date,'%Y年%m月%d日') >= STR_TO_DATE('" + beforeymd + "','%Y年%m月%d日') AND STR_TO_DATE(end_date,'%Y年%m月%d日') <= STR_TO_DATE('" + afterymd + "','%Y年%m月%d日') ";
          } else {
            where = where + " STR_TO_DATE(end_date,'%Y年%m月%d日') >= STR_TO_DATE('" + beforeymd + "','%Y年%m月%d日') AND STR_TO_DATE(end_date,'%Y年%m月%d日') <= STR_TO_DATE('" + afterymd + "','%Y年%m月%d日') ";
          }
          
        } else if (beforeymd != "0") {
          //指定された稼働終了日～で検索
   
          if (where == "") {
            where = where + " WHERE STR_TO_DATE(end_date,'%Y年%m月%d日') >= STR_TO_DATE('" + beforeymd + "','%Y年%m月%d日')  ";
          } else {
            where = where + " STR_TO_DATE(end_date,'%Y年%m月%d日') >=　 STR_TO_DATE('" + beforeymd + "','%Y年%m月%d日') ";
          }
        } else if (afterymd != "0") {
          //～指定された稼働終了日で検索

          if (where == "") {
            where = where + " WHERE STR_TO_DATE(end_date,'%Y年%m月%d日') <=  STR_TO_DATE('" + afterymd + "','%Y年%m月%d日') ";
          } else {
            where = where + " STR_TO_DATE(end_date,'%Y年%m月%d日') <=　 STR_TO_DATE('" + afterymd + "','%Y年%m月%d日') ";
          }

        }

      } else if (period_radio == "7") {
        //稼働中のみを表示
        if (where == "") {
          where = where + " WHERE LAST_DAY(NOW()) <= STR_TO_DATE(end_date,'%Y年%m月%d日')  AND a.deleted=0 ";
        } else {
          where = where + " AND LAST_DAY(NOW()) <= STR_TO_DATE(end_date,'%Y年%m月%d日')  AND a.deleted=0 ";
        }
      } else if (period_radio == "8") {
        //全ての履歴を表示
        where =  "";
        where_all = 1;

      }
    }
  }


  //----------------------------------------------------------------------------------------------------------
  //ジャンル検索
  if (keyword == null || keyword == "") {
  } else {
   
    if (genre_pulldown == "0") {
      hshMap.set('msg', "ジャンルプルダウンを選択してください");
    }
  }

  if (genre_pulldown != 0) {
    if (keyword == null || keyword == "") {
      //キーワードが入力されていない
      hshMap.set('msg', "キーワードを入力してください");
    } else {
      //キーワードが入力されている

      if (!(genre_pulldown == "0")) {
        //ジャンルで絞り込みが選択されている  
        if (genre_pulldown == "1") {
          //元請
          where = ' WHERE a.general_contractor LIKE \'%' + keyword + '%\' ';
        } else if (genre_pulldown == "2") {
          //案件元営業
          where = ' WHERE a.proposition_manager LIKE \'%' + keyword + '%\' ';
        } else if (genre_pulldown == "3") {
          //所属会社
          where = ' WHERE a.affiliation_company LIKE \'%' + keyword + '%\' ';
        } else if (genre_pulldown == "4") {
          //所属会社営業
          where = ' WHERE a.affiliation_manager LIKE \'%' + keyword + '%\' ';
        } else if (genre_pulldown == "5") {
          //技術者名
          where = ' WHERE a.engineer_name LIKE \'%' + keyword + '%\' ';
        } else if (genre_pulldown == "6") {
          //スキル
          where = ' WHERE a.skill LIKE \'%' + keyword + '%\' ';
        } else if (genre_pulldown == "7") {
          //エンド
          where = ' WHERE a.EU LIKE \'%' + keyword + '%\' ';
        } else {
          hshMap.set('msg', "ジャンルプルダウンを選択してください");
        }
      }

    }
  }

  //----------------------------------------------------------------------------------------------------------
  //シート検索


  if (sheet_pulldown != "0") {
  
    if (!(sheet_pulldown == 0)) {
      //シート別で絞り込みが選択されている
      if (where == "") {
        where = where + " WHERE a.sheet_No = " + sheet_pulldown;
      } else {
        where = where + " AND a.sheet_No = " + sheet_pulldown;
      }
    }
  }
  //----------------------------------------------------------------------------------------------------------
  //Ｅｘｃｅｌ取り込み年月
  if (hshMap.get('msg') == null) {
    if (excel_before_year != 0 && excel_after_month != 0 && excel_after_year != 0 && excel_after_month != 0) {
    
      if (where == "") {
     
        where = where + " WHERE STR_TO_DATE('" + beforeymd + "','%Y/%m') <= STR_TO_DATE(a.year_month_date,'%Y/%m') AND STR_TO_DATE(a.year_month_date,'%Y/%m') <=STR_TO_DATE('" + afterymd + "','%Y/%m')  ";
      } else {
        where = where + " AND  STR_TO_DATE('" + beforeymd + "','%Y/%m') <= STR_TO_DATE(a.year_month_date,'%Y/%m') AND STR_TO_DATE(a.year_month_date,'%Y/%m') <=STR_TO_DATE('" + afterymd + "','%Y/%m')  ";
      }
      
      summaryWhere = summaryWhere + " AND  STR_TO_DATE('" + beforeymd + "','%Y/%m') <= STR_TO_DATE(a.year_month_date,'%Y/%m') AND STR_TO_DATE(a.year_month_date,'%Y/%m') <=STR_TO_DATE('" + afterymd + "','%Y/%m')  ";

      groupby = " GROUP BY a.year_month_date "
      excel_hantei_summary = 1;
    } else if (excel_before_year != 0 && excel_before_month != 0) {
     
      if (where == "") {
        where = where + " WHERE STR_TO_DATE('" + beforeymd + "','%Y/%m') <= STR_TO_DATE(a.year_month_date,'%Y/%m')  ";
      } else {
        where = where + " AND  STR_TO_DATE('" + beforeymd + "','%Y/%m') <= STR_TO_DATE(a.year_month_date,'%Y/%m')  ";
        
      }

      summaryWhere = summaryWhere + " AND  STR_TO_DATE('" + beforeymd + "','%Y/%m') <= STR_TO_DATE(a.year_month_date,'%Y/%m')  ";

      groupby = " GROUP BY a.year_month_date "
      excel_hantei_summary = 1;
    } else if (excel_after_year != 0 && excel_after_month != 0) {

      if (where == "") {
        
        where = where + " WHERE STR_TO_DATE(a.year_month_date,'%Y/%m') <=STR_TO_DATE('" + afterymd + "','%Y/%m')  ";
      } else {
        where = where + " AND STR_TO_DATE(a.year_month_date,'%Y/%m') <=STR_TO_DATE('" + afterymd + "','%Y/%m')  ";
      }
      
      summaryWhere = summaryWhere + " AND STR_TO_DATE(a.year_month_date,'%Y/%m') <=STR_TO_DATE('" + afterymd + "','%Y/%m')  ";

      groupby = " GROUP BY a.year_month_date "
      excel_hantei_summary = 1;
    }
  }

  //----------------------------------------------------------------------------------------------------------
  //ソート
  if (sort_button != 0) {
    console.log("ソート:"+sort_button);
    if (sort_button == 1) {
      //Excel取り込み年月
      if (sort1 == "ASC") {
        sort = "  ORDER BY a.year_month_date ASC ";
        sort1 = "DESC";
      } else if (sort1 == "DESC") {
        sort = "  ORDER BY a.year_month_date DESC ";
        sort1 = "ASC";
      }
    } else if (sort_button == 2) {
      //稼働開始日
      if (sort2 == "ASC") {
        sort = " ORDER BY a.start_date ASC ";
        sort2 = "DESC";
      } else if (sort2 == "DESC") {
        sort = "  ORDER BY a.start_date DESC ";
        sort2 = "ASC";
      }
    } else if (sort_button == 3) {
      //稼働終了日
      if (sort3 == "ASC") {
        sort = " ORDER BY a.end_date ASC ";
        sort3 = "DESC";
      } else if (sort3 == "DESC") {
        sort = " ORDER BY a.end_date DESC ";
        sort3 = "ASC";
      }
    } else if (sort_button == 4) {
      //エンド

      if (sort4 == "ASC") {
        sort = "  ORDER BY a.EU ASC ";
        sort4 = "DESC";
      } else if (sort4 == "DESC") {
        sort = "  ORDER BY a.EU DESC ";
        sort4 = "ASC";
      }
    } else if (sort_button == 5) {
      //元請
      if (sort5 == "ASC") {
        sort = "  ORDER BY a.general_contractor ASC ";
        sort5 = "DESC";
      } else if (sort5 == "DESC") {
        sort = "  ORDER BY a.general_contractor DESC ";
        sort5 = "ASC";
      }
    } else if (sort_button == 6) {
      //案件元
      if (sort6 == "ASC") {
        sort = "  ORDER BY a.proposition_company ASC ";
        sort6 = "DESC";
      } else if (sort6 == "DESC") {
        sort = "  ORDER BY a.proposition_company DESC ";
        sort6 = "ASC";
      }
    } else if (sort_button == 7) {
      //案件元営業  
      if (sort7 == "ASC") {
        sort = "  ORDER BY a.proposition_manager ASC ";
        sort7 = "DESC";
      } else if (sort7 == "DESC") {
        sort = "  ORDER BY a.proposition_manager DESC ";
        sort7 = "ASC";
      }
    } else if (sort_button == 8) {
      //所属会社
      if (sort8 == "ASC") {
        sort = "  ORDER BY a.affiliation_company ASC ";
        sort8 = "DESC";
      } else if (sort8 == "DESC") {
        sort = "  ORDER BY a.affiliation_company DESC ";
        sort8 = "ASC";
      }
    } else if (sort_button == 9) {
      //所属会社営業
      if (sort9 == "ASC") {
        sort = "  ORDER BY a.affiliation_manager ASC ";
        sort9 = "DESC";
      } else if (sort9 == "DESC") {
        sort = "  ORDER BY a.affiliation_manager DESC ";
        sort9 = "ASC";
      }
    } else if (sort_button == 10) {
      //技術者名
      if (sort10 == "ASC") {
        sort = "  ORDER BY a.engineer_name ASC ";
        sort10 = "DESC";
      } else if (sort10 == "DESC") {
        sort = "  ORDER BY a.engineer_name DESC ";
        sort10 = "ASC";
      }
    } else if (sort_button == 11) {
      //スキル
      if (sort11 == "ASC") {
        sort = "  ORDER BY a.skill ASC ";
        sort11 = "DESC";
      } else if (sort11 == "DESC") {
        sort = "  ORDER BY a.skill DESC ";
        sort11 = "ASC";
      }
    } else if (sort_button == 12) {
      //売上
      if (sort12 == "ASC") {
        sort = "  ORDER BY a.earnings ASC ";
        sort12 = "DESC";
      } else if (sort12 == "DESC") {
        sort = "  ORDER BY a.earnings DESC ";
        sort12 = "ASC";
      }
    } else if (sort_button == 13) {
      //仕入れ 
      if (sort13 == "ASC") {
        sort = "  ORDER BY a.purchase_price ASC ";
        sort13 = "DESC";
      } else if (sort13 == "DESC") {
        sort = "  ORDER BY a.purchase_price DESC ";
        sort13 = "ASC";
      }
    } else if (sort_button == 14) {
      //法定福利
      if (sort14 == "ASC") {
        sort = "  ORDER BY a.legal_welfare_expense ASC ";
        sort14 = "DESC";
      } else if (sort14 == "DESC") {
        sort = "  ORDER BY a.legal_welfare_expense DESC ";
        sort14 = "ASC";
      }
    } else if (sort_button == 15) {
      //粗利益
      if (sort15 == "ASC") {
        sort = "  ORDER BY a.gross_profit ASC ";
        sort15 = "DESC";
      } else if (sort15 == "DESC") {
        sort = "  ORDER BY a.gross_profit DESC ";
        sort15 = "ASC";
      }
    } else if (sort_button == 16) {
      //粗利率
      if (sort16 == "ASC") {
        sort = "  ORDER BY a.gross_margin ASC ";
        sort16 = "DESC";
      } else if (sort16 == "DESC") {
        sort = "  ORDER BY a.gross_margin DESC ";
        sort16 = "ASC";
      }
    } else if (sort_button == 17) {
      //粗利率
      if (sort17 == "ASC") {
        sort = "  ORDER BY a.proposition_number ASC ";
        sort17 = "DESC";
      } else if (sort17 == "DESC") {
        sort = "  ORDER BY a.proposition_number DESC ";
        sort17 = "ASC";
      }
    } else if (sort_button == 18) {
      //粗利率
      if (sort18 == "ASC") {
        sort = "  ORDER BY a.real_earnings ASC ";
        sort18 = "DESC";
      } else if (sort18 == "DESC") {
        sort = "  ORDER BY a.real_earnings DESC ";
        sort18 = "ASC";
      }
    } else if (sort_button == 19) {
      //粗利率
      if (sort19 == "ASC") {
        sort = "  ORDER BY a.real_gross_profit ASC ";
        sort19 = "DESC";
      } else if (sort19 == "DESC") {
        sort = "  ORDER BY a.real_gross_profit DESC ";
        sort19 = "ASC";
      }
    }
  }
  hshMap.set('sort1', sort1);
  hshMap.set('sort2', sort2);
  hshMap.set('sort3', sort3);
  hshMap.set('sort4', sort4);
  hshMap.set('sort5', sort5);
  hshMap.set('sort6', sort6);
  hshMap.set('sort7', sort7);
  hshMap.set('sort8', sort8);
  hshMap.set('sort9', sort9);
  hshMap.set('sort10', sort10);
  hshMap.set('sort11', sort11);
  hshMap.set('sort12', sort12);
  hshMap.set('sort13', sort13);
  hshMap.set('sort14', sort14);
  hshMap.set('sort15', sort15);
  hshMap.set('sort16', sort16);
  hshMap.set('sort17', sort17);
  hshMap.set('sort18', sort18);
  hshMap.set('sort19', sort19);



  //----------------------------------------------------------------------------------------------------------
  try {
    var summary_judgement = "";

    if (where_all == 1) {


    }else if(where == "" ) {
      //初期表示時は稼働終了日が当月かつシートNoが１のもの　ソート順は稼働終了日昇順

      where = where + " WHERE DATE_FORMAT(NOW(),'%Y%m') >= DATE_FORMAT(STR_TO_DATE(end_date,'%Y年%m月%d日'),'%Y%m') AND DATE_FORMAT(STR_TO_DATE(end_date,'%Y年%m月%d日'),'%Y%m') >= DATE_FORMAT(NOW(),'%Y%m') ";

      tougetu_hantei =1;
      groupby = " GROUP BY a.year_month_date "
    } 

    if (gamenhantei == "0") {　//gamenhantei==0（画面出力用・稼働一覧のシートの情報のみを出力）
      if (where == "") {
        where = where + " WHERE a.sheet_No =1";
      } else {
        if (sheet_pulldown == "0") {
          where = where + " AND a.sheet_No =1";
        }
      }

    } else if (gamenhantei == "1") {　//gamenhantei==1（Excel出力用・全シートの情報を出力）

    }
    if (gamenhantei == 1) {
      sort = " ORDER BY a.sheet_No ASC ";
    } else if (gamenhantei == 0 && sort == "") {
      sort = " ORDER BY a.end_date ASC ";
    }


    mycon = await mysql2.createConnection(mysql_setting);



    const [rows] = await mycon.query(
      `SELECT
      a.year_month_date,
      a.start_date,
      a.end_date,
      a.engineer_name, 
      a.EU,
      a.general_contractor,
      a.proposition_company,
      a.earnings,
      a.proposition_manager,
      a.affiliation_company,
      a.affiliation_manager,
      a.skill,
      a.legal_welfare_expense,
      a.purchase_price,
      a.gross_margin,
      a.gross_profit,
      a.deleted,
      a.proposition_number,
      a.real_earnings,
      a.real_gross_profit,
      a.sheet_No,
      (CASE WHEN STR_TO_DATE(end_date,'%Y年%m月%d日')<= DATE_ADD(DATE_ADD(STR_TO_DATE(NOW(),'%Y-%m-%d'),
      INTERVAL 1 WEEK),INTERVAL 1 DAY)
      THEN "warning"
      ELSE "no" END) AS warning_red,
      (CASE WHEN STR_TO_DATE(end_date,'%Y年%m月%d日') <= DATE_ADD(DATE_ADD(STR_TO_DATE(NOW(),'%Y-%m-%d'),
      INTERVAL 1 MONTH),INTERVAL 7 DAY)
      THEN "warning"
      ELSE "no" END) AS warning_yellow
FROM
management_history a LEFT JOIN Summary_table b ON a.sheet_No=b.sheet_No  AND a.year_month_date=b.year_month_date`
      + where
      + " GROUP BY  a.year_month_date,a.engineer_name,a.proposition_company,a.affiliation_manager,a.affiliation_company,a.affiliation_manager"
      + sort
      );

    rowswork = rows;
    hshMap.set('rowsKey', rows);
    hshMap.set('size', rows.length);
    console.log("検索結果件数"+rows.length);
    if (rows == 0 && gamenhantei != "1") {
      hshMap.set('msg', "該当する検索結果が存在しません");


      //初期表示時は稼働終了日が当月かつシートNoが１のもの　ソート順は稼働終了日昇順
      where = " WHERE DATE_FORMAT(NOW(),'%Y%m') <= DATE_FORMAT(STR_TO_DATE(end_date,'%Y年%m月%d日'),'%Y%m') AND DATE_FORMAT(STR_TO_DATE(end_date,'%Y年%m月%d日'),'%Y%m') <= DATE_FORMAT(NOW(),'%Y%m') ";
      sort = " ORDER BY a.end_date ASC ";

      if (gamenhantei == "0") {　//画面判定用

        where = where + " AND a.sheet_No =1 "
      } else if (gamenhantei == "1") {　//Excel出力用
        where = "";

      }

      mycon = await mysql2.createConnection(mysql_setting);


      const [rows] = await mycon.query(
        `SELECT
              a.year_month_date,
              a.start_date,
              a.end_date,
              a.engineer_name, 
              a.EU,
              a.general_contractor,
              a.proposition_company,
              a.earnings,
              a.proposition_manager,
              a.affiliation_company,
              a.affiliation_manager,
              a.skill,
              a.legal_welfare_expense,
              a.purchase_price,
              a.gross_margin,
              a.gross_profit,
              a.deleted,
              a.proposition_number,
              a.real_earnings,
              a.real_gross_profit,
              a.sheet_No,
              (CASE WHEN STR_TO_DATE(end_date,'%Y年%m月%d日')<= DATE_ADD(DATE_ADD(STR_TO_DATE(NOW(),'%Y-%m-%d'),
              INTERVAL 1 WEEK),INTERVAL 1 DAY)
              THEN "warning"
              ELSE "no" END) AS warning_red,
              (CASE WHEN STR_TO_DATE(end_date,'%Y年%m月%d日') <= DATE_ADD(DATE_ADD(STR_TO_DATE(NOW(),'%Y-%m-%d'),
              INTERVAL 1 MONTH),INTERVAL 7 DAY)
              THEN "warning"
              ELSE "no" END) AS warning_yellow
        FROM
            management_history a LEFT JOIN Summary_table b ON a.sheet_No=b.sheet_No  AND a.year_month_date=b.year_month_date `
        + where
        + sort);

      hshMap.set('rowsKey', rows);
      hshMap.set('size', rows.length);
      console.log("検索結果件数"+rows.length);
      [rowswork] = [rows];



    }
    //-------------------------------------------------------------------------------------------------------------
   //合計を出す
    //ジャンル検索が選択されていた場合は合計の欄を出さない
    if (genre_pulldown == "0") {

      const [sum] = await mycon.query(
        `SELECT
            sum(earnings) as sum_earnings ,
            sum(legal_welfare_expense) as sum_legal_welfare_expense,
            sum(purchase_price) as sum_purchase_price,
            sum(gross_profit) as sum_gross_profit,
            format(sum(gross_profit)/sum(earnings)*100,1) as sum_gross_margin,
            sum(proposition_number)as sum_proposition_number,
            sum(real_earnings)as sum_real_earnings,
            sum(real_gross_profit)as sum_real_gross_profit
            
      FROM
            management_history a INNER JOIN Summary_table b ON a.sheet_No=b.sheet_No AND a.year_month_date=b.year_month_date
            INNER JOIN Sheet_master c ON a.sheet_No=c.sheet_No `+ where);

      hshMap.set('sum', sum);
     


    }


    //-------------------------------------------------------------------------------------------------------------
    //集計情報を出す

    //検索該当なしだった場合は集計情報で使用する変数の値を初期表示用のものに戻す
    if (rows == 0) {

      sheet_pulldown = "0";
      tougetu_hantei =1;
      where ="";
      groupby = " GROUP BY a.year_month_date ";
    }

  if(sheet_pulldown !="0" || sheet_pulldown !="1"){
    if(period_radio == "0"){
      if(excel_hantei_summary != 1){
        tougetu_hantei =1;
        groupby = " GROUP BY a.year_month_date "
      }
    }
  }

      //ここから営業別受注件数の判定


      //範囲検索で一か月で絞り込めることが可能な検索条件にはgroupbyの変数が入っている
      //Excel取り込み年月、稼働開始当月、稼働終了日当月のみ
      if (groupby != "") {
   
        if (tougetu_hantei == 1) {
          var dt = new Date();
          summary_yymm = dt.toFormat("YYYY/MM")
          where = " WHERE a.year_month_date = '" + summary_yymm + "' "
          nengatu = dt.toFormat("YYYY年MM月")
          console.log("集計情報の表示--当月判定")
          

        }

        if (sheet_pulldown != "1" && sheet_pulldown != "0") {
          //検索でシートが稼働一覧以外に選択されている
          //どの集計情報を出力すればよいかsummary_itemを検索してチェックする
          console.log("集計情報の表示--シートが指定されている")
          const [check] = await mycon.query(
            `SELECT
                    summary_item
              FROM
                    sheet_master
              WHERE sheet_No = `+ sheet_pulldown);

        
          if (check[0].summary_item == "2") {
            //summary_item ==2(営業別)
            hshMap.set('summary_judgement', "2");
            if(where==""){
              where = "WHERE a.sheet_No = "+ sheet_pulldown

            }else{
              where =  where + " AND a.sheet_No = "+ sheet_pulldown

            }


            const [summary_check] = await mycon.query(
              ` SELECT 
                        a.year_month_date
                  FROM  
                        management_history a INNER JOIN Summary_table b ON a.sheet_No=b.sheet_No AND a.year_month_date=b.year_month_date
                        INNER  JOIN Sheet_master c ON a.sheet_No=c.sheet_No `
              + where
              + groupby);

    

              var ymdArray = new HashMap();
              var b = "";
  
              for (var i = 0; i < summary_check.length; i++) {
  
                ymdArray.set(i, summary_check[i].year_month_date.toString())
                summary_yymm = summary_check[i].year_month_date.toString();
                b = summary_yymm.replace('/', '年');
                nengatu = b + "月"



                
            const [summary] = await mycon.query(

              `SELECT
                      b.average_sales,
                      b.average_profit,
                      b.operating_count,
                      b.a_star,
                      c.summary_item,
                      c.sales_name
                FROM
                      summary_table b INNER JOIN sheet_master c ON b.sheet_No=c.sheet_No
                WHERE c.sheet_No = `+ sheet_pulldown + " ORDER BY year_month_date DESC");

                
              nengetu_hsh.set('average_sales' + i , nengatu + "　" + summary[0].sales_name+ "　" + "平均売上");
              nengetu_hsh.set('average_profit' + i , nengatu + "　" + summary[0].sales_name+ "　" +  "平均利益");
              nengetu_hsh.set('operating_count' + i , nengatu + "　" + summary[0].sales_name+ "　" +  "稼働件数");
              nengetu_hsh.set('a_star' + i , nengatu + "　" + summary[0].sales_name+ "　" +  "A-STAR");
              summary_hsh.set('summary_average_sales' + i , summary[0].average_sales);
              summary_hsh.set('summary_average_profit' + i , summary[0].average_profit);
              summary_hsh.set('summary_operating_count' + i , summary[0].operating_count);
              summary_hsh.set('summary_a_star' + i , summary[0].a_star);

              }
              hshMap.set('summary_check',summary_check.length)
              hshMap.set('now_mm_hsh',nengetu_hsh)
              hshMap.set('summary_hsh',summary_hsh)
            }
              
//==================ここまでシートNo+シートプルダウンが押された場合の集計情報出力=================
          }else if(sheet_pulldown=="1" || sheet_pulldown=="0"){


          const [summary_check] = await mycon.query(
            ` SELECT 
                      a.year_month_date
              FROM    management_history a INNER JOIN Summary_table b ON a.sheet_No=b.sheet_No AND a.year_month_date=b.year_month_date
                      INNER  JOIN Sheet_master c ON a.sheet_No=c.sheet_No `
                      + where
                      + groupby);
                      
            var ymdArray = new HashMap();
            var b = "";

            for (var i = 0; i < summary_check.length; i++) {

              ymdArray.set(i, summary_check[i].year_month_date.toString())
              summary_yymm = summary_check[i].year_month_date.toString();
              b = summary_yymm.replace('/', '年');
              nengatu = b + "月"
           

              const [eigyoubetu] = await mycon.query(
                `SELECT
                          b.year_month_date,
                          b.sheet_No,
                          b.ordercount_sales,
                          c.sales_name
                  FROM
                          summary_table b INNER JOIN Sheet_master c  ON b.sheet_No=c.sheet_No
                  WHERE   c.summary_item=2 
                          AND     NOT c.sheet_No=1 
                          AND     ( b.year_month_date = '`+ summary_yymm + "') ORDER BY b.year_month_date ASC");


              const [summary] = await mycon.query(
                `SELECT
                        b.year_month_date,
                        b.average_sales,
                        b.average_profit,
                        b.operating_count,
                        b.month_end_count,
                        b.ordercount_new,
                        b.ordercount_new_target,
                        b.ordercount_GAP,
                        c.summary_item
                  FROM 
                        summary_table b LEFT JOIN sheet_master c ON b.sheet_No=c.sheet_No
                WHERE   ( b.year_month_date = '`+ summary_yymm + `') AND c.sheet_No=1 `);



 


              nengetu_hsh.set('average_sales' + i , nengatu + "　" + "平均売上");
              nengetu_hsh.set('average_profit' + i , nengatu + "　" + "平均利益");
              nengetu_hsh.set('operating_count' + i , nengatu + "　" + "稼働件数");
              nengetu_hsh.set('month_end_count' + i , nengatu + "　" + "末落ち件数");
              nengetu_hsh.set('ordercount_new' + i , nengatu + "　" + "新規受注件数");
              nengetu_hsh.set('ordercount_new_target' + i , nengatu + "　" + "新規受注件数目標");
              nengetu_hsh.set('ordercount_GAP' + i , nengatu + "　" + "受注件数GAP");
              summary_hsh.set('summary_average_sales' + i , summary[0].average_sales);
              summary_hsh.set('summary_average_profit' + i , summary[0].average_profit);
              summary_hsh.set('summary_operating_count' + i , summary[0].operating_count);
              summary_hsh.set('summary_month_end_count' + i , summary[0].month_end_count);
              summary_hsh.set('summary_ordercount_new' + i , summary[0].ordercount_new);
              summary_hsh.set('summary_ordercount_new_target' + i , summary[0].ordercount_new_target);
              summary_hsh.set('summary_ordercount_GAP' + i , summary[0].ordercount_GAP);

              



              for (var y = 0; y < eigyoubetu.length; y++) {

                nameMap.set('name' + i + y, nengatu + "　" + eigyoubetu[y].sales_name + "受注件数")
                eigyoudeta.set('deta' + i + y, eigyoubetu[y].ordercount_sales)
              }

            }
          hshMap.set('eigyoubetu_check', y)
          hshMap.set('summary_check', summary_check.length)
          hshMap.set('now_mm_hsh', nengetu_hsh)
          hshMap.set('summary_hsh', summary_hsh)
          hshMap.set('summary_judgement', "1");
          //eigyoubetu_hanteiは画面に営業別受注件数を表示するかしないかの判定をするためのキー
          hshMap.set('eigyoubetu_hantei', "0")
          hshMap.set('eigyoubetu_name', nameMap)
          hshMap.set('eigyoubetu_deta', eigyoudeta)

        } else if (check[0].summary_item == "0" || check[0].summary_item == null) {
          //summary_item ==0(集計情報を出力しない)
      
          hshMap.set('summary_judgement', "3");
      
        }
      


    }else{
      hshMap.set('eigyoubetu_hantei',"1")
      hshMap.set('summary_judgement', "3");

    }


    //-------------------------------------------------------------------------------------------------------------
    //エラー行と警告行と削除行チェック
    var madika = new HashMap();


    // 更新のエラー行チェック



    if (errorcount == null) {

      for (var i = 0; i < rowswork.length; i++) {

        if (rowswork[i].warning_red == "warning") {
          // 赤で表示するやつ(契約終了間近)
          // 2:1週間+1日、1:一カ月+一週間、0:まだ
          madika.set("madi" + i, "2");



        } else if (rowswork[i].warning_yellow == "warning") {
          madika.set("madi" + i, "1");
        } else {
          madika.set("madi" + i, "0");
        }

        if (rowswork[i].deleted == "1") {
          // 削除済み
          madika.set("madi" + i, "3");
        }

      }
    } else {

      var err_split = errorcount.split(",");
      for (var i = 0; i < err_split.length; i++) {

        madika.set("madi" + err_split[i], "4");
      }
      for (var i = 0; i < rowswork.length; i++) {
     

        if (rowswork[i].warning_red == "warning") {
          // 赤で表示するやつ(契約終了間近)
          // 2:1週間+1日、1:一カ月+一週間、0:まだ
          madika.set("madi" + i, "2");

        } else if (rowswork[i].warning_yellow == "warning") {
          madika.set("madi" + i, "1");
        } else {
          madika.set("madi" + i, "0");
        }

        if (rowswork[i].deleted == "1") {
          // 削除済み
          madika.set("madi" + i, "3");
        }

      }

    }

    hshMap.set('madika', madika);
    //-----------------------------------------------------------------------------------------------------
    //excelに出力する用の合計
    const [excel_sum] = await mycon.query(
      ` SELECT
      COALESCE(sum(earnings),0) as sum_earnings ,
      COALESCE(sum(legal_welfare_expense),0) as sum_legal_welfare_expense,
      COALESCE(sum(purchase_price),0) as sum_purchase_price,
      COALESCE(sum(gross_margin),0) as sum_gross_margin,
      COALESCE(sum(gross_profit),0) as sum_gross_profit,
      COALESCE(sum(proposition_number),0) as sum_proposition_number,
      COALESCE(sum(real_earnings),0) as sum_real_earnings,
      COALESCE(sum(real_gross_profit),0) as sum_real_gross_profit,
      c.sheet_No,
      a.year_month_date
      FROM
      Sheet_master c LEFT JOIN (SELECT * FROM management_history a `+where+`  ) a
      ON a.sheet_No=c.sheet_No
      LEFT JOIN Summary_table b
      ON a.sheet_No=b.sheet_No AND a.year_month_date=b.year_month_date 
      GROUP BY c.sheet_No ORDER BY c.sheet_No ASC `);


    hshMap.set('excel_sum', excel_sum);
    //-------------------------------------------------------------------------------------------------------------
    // DB接続解除										
    mycon.end();

  } catch (err) {

    // 例外発生時の処理										
    console.log(err);
    hshMap.set('msg', '検索に失敗しました');
  }



  return hshMap;
}


