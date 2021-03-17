//cd C:\Users\7F-002\Desktop\search_pro\Search_List
const mysql2 = require('mysql2/promise');

var mysql_setting = {
  host: 'localhost',
  user: 'root',
  password: 'passpass',
  database: 'contractPj'
}
let mycon = null;
var HashMap = require('hashmap');

//最初の画面を表示する-------------------------------------------
var express = require('express');
var router = express.Router();
//スラッシュのみはapp.jsで定義したur「search_list」が省略されている
router.get('/', (req, res, next) => {
  //初期表示時はgetに来る
  //初期表示時にgetが走ってしまうため、メッセージの定義をgetでもしておく

  Initial_Display(res);

});


//値を受け取る--------------------------------------------------
// //formタグの情報はreqの中にある 

router.post('/kensaku', (req, res, next) => {
  //ボタンが押されたらpostに来る　postメソッドはボタン分必要
  console.log('検索のpostに来た');
  var radioCond = req.body.radioCond;
  var start_before_year = req.body.start_before_pulldown1;
  var start_before_month = req.body.start_before_pulldown2;
  var start_after_year = req.body.start_after_pulldown1;
  var start_after_month = req.body.start_after_pulldown2;
  var end_before_year = req.body.end_before_pulldown1;
  var end_before_month = req.body.end_before_pulldown2;
  var end_after_year = req.body.end_after_pulldown1;
  var end_after_month = req.body.end_after_pulldown2;
  var genre = req.body.genre_pulldown;
  var sheetsearch = req.body.sheet_pulldown;
  var keyword = req.body.keyword;
  var sort_button = req.body.sort_button;
  var sort1 = req.body.sort1;
  var sort2 = req.body.sort2;
  var sort3 = req.body.sort3;
  var sort4 = req.body.sort4;
  var sort5 = req.body.sort5;
  var sort6 = req.body.sort6;
  var sort7 = req.body.sort7;
  var sort8 = req.body.sort8;
  var sort9 = req.body.sort9;
  var sort10 = req.body.sort10;
  var sort11 = req.body.sort11;
  var sort12 = req.body.sort12;
  var sort13 = req.body.sort13;
  var sort14 = req.body.sort14;
  var sort15 = req.body.sort15;
  var sort16 = req.body.sort16;
  var sort17 = req.body.sort17;
  var sort18 = req.body.sort18;
  var sort19 = req.body.sort19;
  var checklinecount = req.body.checklinecount;
  var errorcount = req.body.errorcount;
  var excel_before_year = req.body.excel_before_year;
  var excel_before_month = req.body.excel_before_month;
  var excel_after_year = req.body.excel_after_year;
  var excel_after_month = req.body.excel_after_month;



  search_Display(res, keyword, radioCond, genre, sheetsearch, sort_button,
    start_before_year, start_before_month, start_after_year, start_after_month,
    end_before_year, end_before_month, end_after_year, end_after_month,
    sort1, sort2, sort3, sort4, sort5, sort6, sort7, sort8, sort9, sort10, sort11,
    sort12, sort13, sort14, sort15, sort16, sort17, sort18, sort19
    , checklinecount, errorcount, excel_before_year, excel_before_month, excel_after_year, excel_after_month);


});

//絞り込み検索とソートのためのメソッド------------------------------------------------
async function search_Display(res, keyword, radioCond, genre, sheetsearch, sort_button,
  start_before_year, start_before_month, start_after_year, start_after_month,
  end_before_year, end_before_month, end_after_year, end_after_month,
  sort1, sort2, sort3, sort4, sort5, sort6, sort7, sort8, sort9, sort10, sort11, sort12, sort13, sort14, sort15, sort16, sort17, sort18, sort19
  , checklinecount, errorcount, excel_before_year, excel_before_month, excel_after_year, excel_after_month) {
  var sheet_data_method = new HashMap();
  var recordCheckList = new HashMap();
  console.log(' 絞り込み検索とソートのためのメソッドにきた');

  var sheet_search = require('./sheet_pulldown.js');
  var search_and_sort = require('./search_execute.js');

  //シートとシート別で絞り込みプルダウンを連動させるためのメソッド
  sheet_data_method = await sheet_search.sheet_pulldown_search();
  var sheet_data = sheet_data_method.get('rowsKey');


  //検索結果を出すためのプルダウン
  var result_method = await search_and_sort.searchcheck(
    keyword, radioCond, genre, sheetsearch, sort_button,
    start_before_year, start_before_month, start_after_year, start_after_month,
    end_before_year, end_before_month, end_after_year, end_after_month,
    sort1, sort2, sort3, sort4, sort5, sort6, sort7, sort8, sort9, sort10, sort11,
    sort12, sort13, sort14, sort15, sort16, sort17, sort18, sort19
    , errorcount, excel_before_year, excel_before_month, excel_after_year, excel_after_month, 0
  );
  var result = result_method.get('rowsKey');
  var sum_data = result_method.get('sum');

  for (var i = 0; i < result.length; i++) {
    recordCheckList.set("record" + i, 0);
  }

  var data = {
    'ROWS': result, 'sheet_data': sheet_data, 'sum_data': sum_data, updatemsg: '', deletemsg: ''
    , 'pushButton': 0
    , 'sort_button': sort_button
    , 'summary_data': result_method.get('summary_hsh')
    , 'summary_judgement': result_method.get('summary_judgement')
    , 'madika': result_method.get('madika')
    , 'msg': result_method.get('msg')
    , 'keyword': keyword
    , 'errorcount': errorcount
    , 'warningcount': 0
    , 'checklinecount': checklinecount
    , 'start_before_pulldown1': start_before_year
    , 'start_before_pulldown2': start_before_month
    , 'start_after_pulldown1': start_after_year
    , 'start_after_pulldown2': start_after_month
    , 'end_before_pulldown1': end_before_year
    , 'end_before_pulldown2': end_before_month
    , 'end_after_pulldown1': end_after_year
    , 'end_after_pulldown2': end_after_month
    , 'radioCond': radioCond
    , 'genre_pulldown': genre
    , 'sheet_pulldown': sheetsearch
    , 'sort1': result_method.get('sort1')
    , 'sort2': result_method.get('sort2')
    , 'sort3': result_method.get('sort3')
    , 'sort4': result_method.get('sort4')
    , 'sort5': result_method.get('sort5')
    , 'sort6': result_method.get('sort6')
    , 'sort7': result_method.get('sort7')
    , 'sort8': result_method.get('sort8')
    , 'sort9': result_method.get('sort9')
    , 'sort10': result_method.get('sort10')
    , 'sort11': result_method.get('sort11')
    , 'sort12': result_method.get('sort12')
    , 'sort13': result_method.get('sort13')
    , 'sort14': result_method.get('sort14')
    , 'sort15': result_method.get('sort15')
    , 'sort16': result_method.get('sort16')
    , 'sort17': result_method.get('sort17')
    , 'sort18': result_method.get('sort18')
    , 'sort19': result_method.get('sort19')
    , 'excel_before_year': excel_before_year
    , 'excel_before_month': excel_before_month
    , 'excel_after_year': excel_after_year
    , 'excel_after_month': excel_after_month
    , 'eigyoubetu_name': result_method.get('eigyoubetu_name')
    , 'eigyoubetu_deta': result_method.get('eigyoubetu_deta')
    , 'size': result_method.get('size')
    , 'now_mm': result_method.get('now_mm_hsh')
    , 'eigyoubetu_hantei': result_method.get('eigyoubetu_hantei')
    , 'summary_check': result_method.get('summary_check')
    , 'eigyoubetu_check': result_method.get('eigyoubetu_check')
    , 'recordCheckList': recordCheckList

  };
  res.render('Search_List_view', data);

}

//初期表示のためのメソッド------------------------------------------------
async function Initial_Display(res) {

  var sheet_data_method = new HashMap();
  console.log(' 初期表示のためのメソッドにきた');

  var sheet_search = require('./sheet_pulldown.js');
  var initial_display_search = require('./search_execute.js');

  //シートとシート別で絞り込みプルダウンを連動させるためのメソッド
  sheet_data_method = await sheet_search.sheet_pulldown_search();
  var sheet_data = sheet_data_method.get('rowsKey');

  var recordCheckList = new HashMap();

  //初期表示時の検索結果を出すためのプルダウン

  var result_method = await initial_display_search.searchcheck(

    null//キーワード
    , 0　　//ラジオボタン
    , 0   //ジャンルで絞り込み
    , 0   //シートで絞り込み
    , 0   //ソートボタン
    , 0   //稼働開始年（前）
    , 0   //稼働開始月（前）
    , 0   //稼働開始年（後）
    , 0　　//稼働開始月（後）
    , 0   //稼働終了年（前）
    , 0   //稼働終了月（前）
    , 0   //稼働終了年（後）
    , 0　　//稼働終了月（後）
    , 0   //ソートボタン１
    , 0   //ソートボタン２
    , 0   //ソートボタン３
    , 0   //ソートボタン４
    , 0   //ソートボタン５
    , 0   //ソートボタン６
    , 0   //ソートボタン７
    , 0   //ソートボタン８
    , 0   //ソートボタン９
    , 0   //ソートボタン１０
    , 0   //ソートボタン１１
    , 0   //ソートボタン１２
    , 0   //ソートボタン１３
    , 0   //ソートボタン１４
    , 0    //ソートボタン１５
    , 0   //ソートボタン１６
    , 0   //ソートボタン１７
    , 0    //ソートボタン１８
    , 0   //ソートボタン１９
    , null//エラーカウント
    , 0   //excel取り込み年（前）
    , 0   //excel取り込み月（前）
    , 0   //excel取り込み年（後）
    , 0   //excel取り込み月（後）
    , 0   //画面判定用変数（画面「０」Excel出力「１」）
  );

  var result = result_method.get('rowsKey');
  var sum_data = result_method.get('sum');

  for (var i = 0; i < result.length; i++) {
    recordCheckList.set("record" + i, 0);
  }



  var data = {
    updatemsg: '', deletemsg: '',
    'ROWS': result,
    'sheet_data': sheet_data
    , 'sum_data': sum_data
    , 'summary_data': result_method.get('summary_hsh')
    , 'madika': result_method.get('madika')
    , 'msg': ""
    , 'summary_judgement': "1"
    , 'keyword': ""
    , 'errorcount': "0"
    , 'warningcount': "0"
    , 'checklinecount': ""
    , 'start_before_pulldown1': "0"
    , 'start_before_pulldown2': "0"
    , 'end_before_pulldown1': "0"
    , 'end_before_pulldown2': "0"
    , 'start_after_pulldown1': "0"
    , 'start_after_pulldown2': "0"
    , 'end_after_pulldown1': "0"
    , 'end_after_pulldown2': "0"
    , 'radioCond': "0"
    , 'genre_pulldown': "0"
    , 'sheet_pulldown': "0"
    , 'sort_button': "0"
    , 'sort1': "ASC"
    , 'sort2': "ASC"
    , 'sort3': "ASC"
    , 'sort4': "ASC"
    , 'sort5': "ASC"
    , 'sort6': "ASC"
    , 'sort7': "ASC"
    , 'sort8': "ASC"
    , 'sort9': "ASC"
    , 'sort10': "ASC"
    , 'sort11': "ASC"
    , 'sort12': "ASC"
    , 'sort13': "ASC"
    , 'sort14': "ASC"
    , 'sort15': "ASC"
    , 'sort16': "ASC"
    , 'sort17': "ASC"
    , 'sort18': "ASC"
    , 'sort19': "ASC"
    , 'excel_before_year': "0"
    , 'excel_before_month': "0"
    , 'excel_after_year': "0"
    , 'excel_after_month': "0"
    , 'eigyoubetu_name': result_method.get('eigyoubetu_name')
    , 'eigyoubetu_deta': result_method.get('eigyoubetu_deta')
    , 'size': result_method.get('size')
    , 'now_mm': result_method.get('now_mm_hsh')
    , 'eigyoubetu_hantei': result_method.get('eigyoubetu_hantei')
    , 'summary_check': result_method.get('summary_check')
    , 'eigyoubetu_check': result_method.get('eigyoubetu_check')
    , 'pushButton': 0
    , 'recordCheckList': recordCheckList

  };

  res.render('Search_List_view', data);

}




module.exports = router;
