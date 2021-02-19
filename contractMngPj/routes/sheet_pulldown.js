const mysql2 = require('mysql2/promise');	                    			
var mysql_setting = {                       			
    host:'localhost',                       			
    user:'root',                        			
    password:'root',                        			
    database:'contractPj'                      			
}       

let mycon = null;  
var HashMap = require('hashmap'); 

//onload時にシート別のプルダウンに値を埋め込むための検索
exports.sheet_pulldown_search=async function() {
  console.log("プルダウン表示のメソッド");		
try{				
  var hshMap = new HashMap();							
   var sheet_data= '';	
 																
              mycon = await mysql2.createConnection(mysql_setting);										 
									
              const [rows] = await mycon.query(
            " SELECT sheet_No,sheet_name FROM sheet_master ");										
            hshMap.set('rowsKey',rows);
            
              // DB接続解除										
              mycon.end();										
                      
      }catch(err){										
                      
          // 例外発生時の処理										
          console.log(err);										
          msg = '検索に失敗しました';										
      }		
                                        


      return hshMap;
}