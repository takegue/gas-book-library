spreadsheet_id = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
slack_params = {
   "token" : "xxxxxxxxxxxxxxxxxxxxxxxx",
   "channel" : {
      "book-rent" : "channel_id",
      "book-order" : "channel_id",
      "test_bot" : "channel_id",
  },
  "post_url" : "https://hooks.slack.com/services/xxxxxxxxx/xxxxxxxxx/xxxxxxxxxxxxxxxxxxxxxxxx"
}

n_columns = 9;
DEBUG = 0;

function getSpreadSheet(){
  if(getSpreadSheet.sheet){
    return getSpreadSheet.sheet;
  }
  getSpreadSheet.sheet = SpreadsheetApp.openById(spreadsheet_id);
  return getSpreadSheet.sheet
}

function getListSheet(){
  if(getListSheet.sheet){
    return getListSheet.sheet;
  }
  getListSheet.sheet = getSpreadSheet().getSheetByName(DEBUG ? 'DEBUG' : '蔵書一覧');
  return getListSheet()
}

function getLogSheet(){
  if(getLogSheet.sheet){
    return getLogSheet.sheet;
  }
  getLogSheet.sheet = getSpreadSheet().getSheetByName('log');
  return getLogSheet.sheet
}

function cacheWrapper(key, f){
   var cache = CacheService.getScriptCache();
   var cached = cache.get(key)
   if ( cached != null) {
     Logger.log("CACHE HIT:" + cached)
     return JSON.parse(cached)
   }
   
   content = f()
   Logger.log(content)
   cache.put(key, JSON.stringify(content), 21600)
   return content
}

function getDataValues(){
  var cache = CacheService.getScriptCache();
  var keys =  getDataValues.keys()
  
  var cached = Array(cache.getAll(keys))[0]
  if ( Object.keys(cached).length > 0) {
    contents = new Array(keys.length)
    for(i=0; i < keys.length; i++){
       contents[i] = JSON.parse(cached[keys[i]])
    }
    return transpose(contents)
  } 
  
  Logger.log('SETUP')
  n_rows = getListSheet().getLastRow();
  var contents = getListSheet().getRange(2, 1, n_rows, n_columns).getValues();
  
  values = {} 
  contents_t = transpose(contents)
  for(i=0;i<n_columns;i++){
      values[keys[i]] = JSON.stringify(contents_t[i])
  }
  cache.putAll(values, 21600);
  Logger.log("e")
  return contents;
}

getDataValues.clear = function(){
  var cache = CacheService.getScriptCache(); 
  keys = getDataValues.keys();
  cache.removeAll(keys);
}

getDataValues.keys = function(){
  keys = new Array(n_columns)
  for(i=0;i<keys.length;i++){
    keys[i] = "books-contents-" + i
  }
  return keys
}

function transpose(arrays){
  new_arrays = []
  for(i=0;i<arrays[0].length;i++){
    record = new Array(arrays.length)
    for(j=0; j<arrays.length;j++){
      record[j] = arrays[j][i]
    }
    new_arrays.push(record)
  }
  return new_arrays
}

function log2sheet(msg, level) {
    msg = msg === undefined ? '' : msg;
    level = level === undefined ? '[DEBUG]' : level;

    var logsheet = getLogSheet('log');  
    // 最終行に追加
    logsheet.appendRow([new Date(), level, msg]);
}

function sendMsg(url, msg){
  var urlFetchOption = {
    'method' : 'POST',    
    'muteHttpExceptions' : true,
    'payload' : JSON.stringify(msg),
  };
  
  var response = UrlFetchApp.fetch(url, urlFetchOption); 
  try {  
    return {
      responseCode : response.getResponseCode(),
      rateLimit : {
        limit : response.getHeaders()['X-RateLimit-Limit'],
        remaining : response.getHeaders()['X-RateLimit-Remaining'],
      },
      parseError : false,
      // body : JSON.parse(response.getContentText()),
    };
  } catch(e) {
    log2sheet(e, '[ERROR]')
  }
}

function findRows(sheet, func){
  Logger.log('findRows: Start')
  var range = getDataValues();
  var res = []
  Logger.log('findRows: Get values')
  Logger.log(range.length)
  for(i=0;i<range.length;i++){
    if(range[i][0] && func(range[i])){
      res.push([i+2, range[i]]);
    } 
  }
  Logger.log('findRows: End')
  return res;
}

function registNewBook(params){
  var sheet = getListSheet(),
        c2n = getColumn2Num(sheet); 

  for(var n in links){ 
    link = links[n].url      
    log2sheet('NEW:' + link, '[INFO]')
    
    rnum = sheet.getLastRow()+1
    r = sheet.getRange(
      rnum, 1, 1, sheet.getLastColumn()
    )
    
    new_record = new Array(sheet.getLastColumn())
    for(i=0;i<new_record.length;i++){new_record[i] = ''}
    
    new_record[c2n['ID']] = rnum
    new_record[c2n['URL']] = link
    new_record[c2n['URL'] + 1] = '=REGEXREPLACE(' + 'J'+ rnum + ', "amazon.co.jp/[^/]*/dp", "amazon.co.jp/dp")'
    new_record[c2n['URL'] + 2] = '=IMPORTXML(' + 'K' + rnum + ', "//title")' 
    new_record[c2n['URL'] + 3] = '=SPLIT(SUBSTITUTE('  + 'L' + rnum + ', " :", "| ") , "|")'
    new_record[c2n['購入日']] = new Date()
    new_record[c2n['状態']] = '購入中'
    new_record[c2n['書籍名']] = '=' + 'L' + rnum
 
    r.setValues([new_record])
  }
  return {}
}

function doPost(e) {
  
  // トークンが不一致なら処理終了
  var params = 'token' in e.parameter ? e.parameter : JSON.parse(e.postData.getDataAsString());
  var token = params.token;

  if (token != slack_params['token']) {
    return ContentService.createTextOutput("Invalid Token");
  }
  
  if(params.type == 'url_verification'){
    return ContentService.createTextOutput(
      JSON.stringify({"challenge" : params.challenge})
    ).setMimeType(ContentService.MimeType.JSON);
  }else if(params.type == 'event_callback'){
    channel = params.event.channel  
    if(
      channel == slack_params['channel']['book-order']
       || channel == slack_params['channel']['test_bot']
      ){
      registNewBook(params);
    }
    return {}
  }

  // Parameter example:
  //{
  //  "channel_name": "xxxxxxxxxxxx",
  //  "user_id": "xxxxxxxxx",
  //  "user_name": "xxxxxxxxxxxxxxx",
  //  "team_domain": "xxxxx",
  //  "team_id": "xxxxxxxxx",
  //  "text": "help",
  //  "channel_id": "xxxxxxxxx",
  //  "command": "/book",
  //  "token": "xxxxxxxxxxxxxxxxxxxxxxxx",
  //  "response_url": "https://hooks.slack.com/services/xxxxxxxxx/xxxxxxxxx/xxxxxxxxxxxxxxxxxxxxxxxx"
  //}  
  var args = e.parameter.text.split(' ');  
  var subcmd = args[0];
  if(args.length > 1){
    var args = args.slice(1);
  }else{
    var args = [""];
  }  

  Logger.log('Receive from ' +  e.parameter.user_id + '@' + e.parameter.channel_name);
  subCmds = {
    "checkout" : bookCmdCheckout,
    "search"   : bookCmdSearch,
    "reserve"  : bookCmdReserve,
    "return"   : bookCmdReturn,
    "help"     : bookCmdHelp,
  }
  if(!subcmd || !(subcmd in subCmds)){
    subcmd = "help"
  }
  
  sendMsg(e.parameter.response_url,     
      JSON.stringify({
        "response_type": "ephemeral", 
        "text": "Please wait a minutes"
  }));
  
  time_s = new Date()
  resp = subCmds[subcmd](args, e.parameter)
  sendMsg(e.parameter.response_url, resp)
  
  time_e = new Date()
  
  return ContentService.createTextOutput(
    JSON.stringify({
      "response_type": "ephemeral", 
      "text": "Complete"
    })).setMimeType(ContentService.MimeType.JSON);
}


function bookCmdSearch(args, kwargs){
  var query     = args[0],
      sheet     = getListSheet(),
      c2n       = getColumn2Num(sheet),
      resp = {
        "response_type": "ephemeral", 
        "text": "Sorry, that didn't work. Please try again."
      }
  
  res = searchBooks(sheet, query)
  
  r = ['Search result For *' + query + '* (' + res.length +' Hits)'+':']
  for(i=0;i<res.length;i++){
    record = res[i][1]
    status = record[c2n['状態']] == '' ? '' : '[' + record[c2n['状態']] + ']'
    r.push(status + ' `' + record[c2n['書籍名']] + '`(' + record[c2n['ID']] + ')')
  }

  // log2sheet('Search: `' + query + '`', '[INFO]')
  return {"text" : r.join('\n')};
}

function bookCmdCheckout(args, kwargs){

  var user_id   = kwargs.user_id,
      user_name = kwargs.user_name,
      query     = args[0],
      sheet     = getListSheet(),
      c2n       = getColumn2Num(sheet),
      resp = {
        "response_type": "ephemeral", 
        "text": "Sorry, that didn't work. Please try again."
      }
  
  ret = checkoutBook(sheet, user_name, query);
  if(ret && ret.length == 1){
    // resp['response_type'] = 'in_channel';
    msg = 'User <@' + user_id + '> checks out a book '; 
    resp['text'] = msg
    resp['attachments'] =[{
      'text' : ret[0][c2n['書籍名']]
    }]
    sendMsg(slack_params['post_url'], resp)
    log2sheet('CHECKOUT:' + ret[0][c2n['書籍名']], '[INFO]')
    resp['text'] = 'You have checked out following books'
  }else if(ret.length > 1){
    resp['text'] = 'Query *' + query + '*  is ambigious. Plase specify the BOOK_ID from following books'
    resp['attachments'] = []
    for(i=0; i<ret.length; i++){
      record = ret[i][1]
      msg = 'BOOK_ID='+ record[c2n['ID']] + ' : `' +record[c2n['書籍名']] + '`';
      resp['attachments'].push({'text' : msg})
    }     
  }
  else{
    resp['text'] = 'Failed to checkout'
    log2sheet('Failed', '[INFO]')
  }
  // TODO: Implements more efficient cache update 
  getDataValues.clear()
  return resp
}


function bookCmdReserve(args, kwargs){

  var user_id   = kwargs.user_id,
      user_name = kwargs.user_name,
      query     = args[0],
      sheet     = getListSheet(),
      c2n       = getColumn2Num(sheet),
      resp = {
        "response_type": "ephemeral", 
        "text": "Sorry, that didn't work. Please try again."
      }
  Logger.log(user_id);
  Logger.log(user_name)
  ret = reserveBook(sheet, user_name, query);
  if(ret && ret.length == 1){
    msg = 'User <@' + user_id + '> reserve a book '; 
    resp['text'] = msg
    resp['attachments'] =[{
      'text' : ret[0][c2n['書籍名']]
    }]
    sendMsg(slack_params['post_url'], resp)
    log2sheet('RESERVED:' + ret[0][c2n['書籍名']], '[INFO]')
    resp['text'] = 'You have reserved out following books'
  }else if(ret.length > 1){
    resp['text'] = 'Query *' + query + '*  is ambigious. Plase specify the BOOK_ID from following books'
    resp['attachments'] = []
    for(i=0; i<ret.length; i++){
      record = ret[i][1]
      msg = 'BOOK_ID='+ record[c2n['ID']] + ' : `' +record[c2n['書籍名']] + '`';
      resp['attachments'].push({'text' : msg})
    }     
  }
  else{
    resp['text'] = 'Failed to reserve'
    log2sheet('Failed', '[INFO]')
  }
  // TODO: Implements more efficient cache update 
  getDataValues.clear()
  return resp
}



function bookCmdReturn(args, kwargs){

  var user_id   = kwargs.user_id,
      user_name = kwargs.user_name,
      query     = args[0] === undefined ? false : args[0],
      sheet     = getListSheet(),
      c2n       = getColumn2Num(sheet),
      resp      = {
        "response_type": "ephemeral", 
        "text": "Sorry, that didn't work. Please try again."
      },
      name2id   = getMembers()
  
  ret = returnBook(sheet, user_name, query);
  
  if(ret && ret.length > 0){
    // In-channel text message for returning
    resp['text'] = 'User <@' + user_id  + '> returns following books';
    resp['attachments'] = []
    Logger.log(ret)
    for(i=0; i<ret.length; i++){
      msg = ret[i][c2n['書籍名']];
      
      subscriber = ret[i][c2n['予約者']]
      if(subscriber != ''){
        msg = '[<@' + name2id[subscriber] + '>]  ' + msg 
      }
      
      resp['attachments'].push({'text' : msg})

    }
    sendMsg(slack_params['post_url'], resp)
    
    // Ephemeral text message for returning
    resp['text'] = 'You have returned following books'
    resp['attachments'] = []
    Logger.log(ret)
    for(i=0; i<ret.length; i++){
      msg = "<" + makeReportLink(ret[i][c2n['書籍名']], user_name) + "|[Qiitaで書評を書く]> " + ret[i][c2n['書籍名']];
      resp['attachments'].push({'text' : msg})
    }
    
  }else{
    log2sheet('Failed', '[INFO]')
  }
  getDataValues.clear()
  return resp
}

function bookCmdHelp(args, kwargs){
  // Response example
  // {
  //  "response_type": "ephemeral", 
  //  "text": "Sorry, that didn't work. Please try again."
  // } 
  return { 
    "text" : [
      "蔵書管理要のAPIです",
      "`search hogehoge`\t: *hogehoge* をキーワードに検索します",
      "`checkout hogehoge`\t: *hogehoge* を借ります (hogehoge は キーワード or BOOK_ID)",
      "`reserve hogehoge`\t: *hogehoge* を予約します (hogehoge は キーワード or BOOK_ID)",
      "`return [hogehoge]`\t: [hogehoge]もしくは借りている全ての本を返します  (hogehoge は キーワード or BOOK_ID)*",
  ].join("\n")};
}
      
function getColumn2Num(sheet){
  return cacheWrapper("c2n", function(){
    range = sheet.getRange(sheet.getFrozenRows(), 1, 1, sheet.getLastColumn());
    columns = range.getValues()[0];
  
    var c2n = {}
    for(i=0; i<columns.length;i++){
      c2n[columns[i]] = i;
    }
    return c2n;
  })
}

function searchBooks(sheet, query, ignore_rentaled){
  // log2sheet(arguments.callee.name);
  Logger.log("Search Start")
  ignore_rentaled = ignore_rentaled === undefined ? 1 : ignore_rentaled;
  c2n = getColumn2Num(sheet)
  Logger.log("Setup")
  
  if( isFinite(query) ){
    num = parseInt(query);
    if(num == 0){
      return []
    }    
    res = findRows(
      sheet, function(e){
        return e[c2n['ID']] == num && (ignore_rentaled || e[c2n['状態']] != '貸出中' )
      });
  }else{
    word = query;
    res = findRows(
      sheet,
      function(e){
        return e[c2n['書籍名']].toLowerCase().search(word.toLowerCase()) > -1 && (ignore_rentaled || e[c2n['状態']] != '貸出中' )
      });
  }
  Logger.log("End")
  return res;
}

function changeBookState(sheet, user, query, updater, ignore_rentaled){ 
  c2n = getColumn2Num(sheet);
  
  var res = searchBooks(sheet, query, ignore_rentaled);
  Logger.log(res);
  if(!res || !res.length){
    return []
  }else if(res.length > 1){
    return res
  }
  
  row = res[0][0], record = res[0][1];
  new_record = updater(record, c2n)
  Logger.log(new_record)
  sheet.getRange(row, 1, 1, new_record.length).setValues([new_record]);
  return [new_record];
}

function checkoutBook(sheet, user, query){ 
  return changeBookState(sheet, user, query, function(r, c2n){
    r[c2n['貸出先']] = user;
    r[c2n['貸出日']] = new Date();
    r[c2n['返却日']] = "";
    r[c2n['状態']] = '貸出中';
    return r
  })
}

function reserveBook(sheet, user, query){
  return changeBookState(sheet, user, query, function(r, c2n){
    r[c2n['予約者']] = user;
    return r
  }, true)
}

function returnBook(sheet, user, query){ 
  query = query === undefined ? false : query;
  
  c2n = getColumn2Num(sheet);

  if( isFinite(query) ){
    num = parseInt(query);
    if(num == 0){
      return []
    }    
    res = findRows(
      sheet, function(e){
        return e[c2n['貸出先']] == user && (!query || e[c2n['ID']] == num) && e[c2n['状態']] == '貸出中'
      });
  }else{
    word = query;
    res = findRows(
      sheet,
      function(e){
        return e[c2n['貸出先']] == user && ( !query || e[c2n['書籍名']].toLowerCase().search(word.toLowerCase()) > -1) &&  e[c2n['状態']] == '貸出中'
      });
  }

  var ret = []
  for(i = 0; i<res.length; i++){
    row = res[i][0], record = res[i][1];
    record[c2n['返却日']] = new Date();
    record[c2n['状態']] = '返却済';
    record[c2n['読んだ']] =  record[c2n['読んだ']] + 1;
    sheet.getRange(row, 1, 1, record.length).setValues([record]);  
    ret.push(record)
  }
  return ret;
}

function updateCache(){
  getDataValues.clear()
  getDataValues()
}

function doGet(e) {
  return ContentService.createTextOutput('Got it!');
}

function makeReportLink(book_title, user_name){
  endpoint = "https://teamretty.qiita.com/drafts/new?"
  title = "[読了] " + book_title + "@" + user_name
  tags = "book"
  
  body = "#書籍名:" + book_title + "\n\n" + 
　　  "# あらすじ\n\n" +
   "# なぜ読もうと思ったか?\n\n" +
   "# 誰にどこを読んで欲しいか？\n\n" + 
   "# 所感 \n\n" + 
   "# 次に読みたい本\n\n"
   params = "title=" + encodeURIComponent(title) + "&body=" + encodeURIComponent(body) + "&tags=" + encodeURIComponent(tags) 
   return endpoint + params
}



function getMembers(){
  return cacheWrapper('members', function(){
    var urlFetchOption = {
      'method' : 'POST',    
      'muteHttpExceptions' : true,
    };
    url = "https://slack.com/api/users.list?token=slack_token_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx&pretty=1"
    var response = UrlFetchApp.fetch(url, urlFetchOption);
    members = JSON.parse(response.getContentText())['members']
    var name2id = {}
    for(n in members){
      member = members[n]
      name2id[member['name']] = member['id']
    }
    return name2id
  })
}


function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "検索する蔵書の更新",
    functionName : "updateCache"
  }];
  spreadsheet.addMenu("検索ツール", entries);
};


function testFunc(){
   log2sheet('test', 'DEBUG');
   var sheet = getListSheet()

  Logger.log("start")
  Logger.log(returnBook(getListSheet(), "shunsuke_takeno"))
  Logger.log("end")

}
