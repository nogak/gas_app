/* 
勤怠予定表からメンバーの今日の勤務地を取得し、Slackに投げる
※インストーラブルトリガーでsetTrigger()関数を呼び、スクリプト(main_kintai)の回る時刻を指定
(インストーラブルトリガーのみだと厳密な時刻指定ができないため)
【機能フロー】
    ①シート集合？を開く
    ②シートを選択
    ③対象日付のセルを取得
    ④各メンバーの名前と予定のリストを作成
    ⑤メッセージの生成
    ⑥スラックに送る
*/
// ss : spread_sheet
// sh : sheet

// main関数(main_kintai)の実行される時刻を指定
function setTrigger(){
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth();
    const day = today.getDate();
    const youbi = today.getDay();
    
    // 土日の判定
    if(youbi === 0 || youbi === 6) return;
    // 祝日の判定
    const id = 'ja.japanese#holiday@group.v.calendar.google.com'
    const cal = CalendarApp.getCalendarById(id);
    const events = cal.getEventsForDay(today);
    if(events.length) return;

    // console.log(year, month, day)
    const hour = '9';
    const minute = '00'
    // Dateオブジェクトに変換
    const setTime = new Date(year, month, day, hour, minute);
    console.log(setTime);
    // 特定時刻にスクリプトが起動するように設定
    // ScriptApp.newTrigger('main_kintai').timeBased().at(setTime).create();

    // 上の特定時刻での起動がうまく動かなかったためとりあえず。
    main_kintai();
}

// 実行済みのトリガーを削除
function deleteTrigger(){
    const triggers = ScriptApp.getProjectTriggers();
    for(const trigger of triggers){
        if(trigger.getHandlerFunction() === 'main_kintai') ScriptApp.deleteTrigger(trigger);
    }
}

function main_kintai(){
    // 対象シートの取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh;
    try{
        sh = ss.getSheetByName('勤怠予定表');
    }catch(e){
        console.log('not found sheet');
        sendSlack('not found sheet');
        return;
    }

    // 今日の日付を取得
    let today = new Date();
    // console.log(today)
    today = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd");
    // console.log(today);

    // スプレッドシートから今日の日付のセルの位置を探す
    const last_row = sh.getLastRow(); // 最終行のセルの位置を取得
    // console.log(typeof last_row)
    const date_list_tmp = sh.getRange("C1:C" + String(last_row)).getValues(); // 生の日付情報をリストで取得（変換が必要）
    // console.log(Utilities.formatDate(date_list_tmp[3][0], "Asia/Tokyo", "yyyy/MM/dd"))
    const date_list = new Array(last_row) // 変換後の日付を格納するリスト
    for(let i=3; i < last_row; i++){
        date_list[i] = Utilities.formatDate(date_list_tmp[i][0], "Asia/Tokyo", "yyyy/MM/dd");
    }
    // console.log(date_list)
    // console.log(today);
    let today_cell_row; // 今日の日付のセルの行番号
    for(let i=0; i < last_row; i++){
        if(date_list[i] == today) today_cell_row = i+1; // スプレッドシートだと1から始まるためズレを修正
    }
    // console.log(today_cell_row);
    
    // 各メンバーの名前と予定のリストを生成
    const last_col = sh.getLastColumn();
    const member_num = last_col - 3;
    // console.log(member_num);

    let last_member_cell = sh.getRange(3, last_col).getA1Notation();
    let first_place_cell = sh.getRange(today_cell_row, 4).getA1Notation();
    let last_place_cell = sh.getRange(today_cell_row, last_col).getA1Notation();
    // console.log(last_member_cell);
    // console.log(first_place_cell);
    // console.log(last_place_cell);

    let member_list = sh.getRange("D3:" + String(last_member_cell)).getValues()[0];
    let place_list = sh.getRange(String(first_place_cell) + ":" + String(last_place_cell)).getValues()[0];
    let place_list_color = sh.getRange(String(first_place_cell) + ":" + String(last_place_cell)).getFontColors()[0];
    // console.log(member_list)
    // console.log(place_list)
    // console.log(place_list_color)

    for(let i=0; i < member_num; i++){
        place_list[i] = split_string(place_list[i]); // カンマ区切りで午前・午後判定をして文字列変換
        place_list[i] = check_color(place_list[i], place_list_color[i]) // カンマ区切りがない場合に先頭の文字色から午前・午後・終日判定をして文字列変換
    }
    
    // Slackメッセージの生成
    let slack_msg = "本日の勤務地\n";
    for(let i=0; i < member_num; i++){
        if(place_list[i] != '') slack_msg += String(member_list[i]) + " -> " +String(place_list[i]) + "\n";
    }
    if(slack_msg === "本日の勤務地\n") return;
    console.log(slack_msg);
    sendSlack(slack_msg);
    // deleteTrigger();
}

function sendSlack(slack_msg){
    // 周りから見えないように環境変数で設定したほうが良い
    let webHookUrl = "各自で取得";

    let jsonData = {
        "channel" : "#times-nogami", 
        // "icon_emoji" : "", 
        "text" : slack_msg, 
        "username" : "nogami_test"
    };

    let payload = JSON.stringify(jsonData);

    let options = {
        "method" : "post", 
        "contentType" : "application/json", 
        "payload" : payload, 
    };

    UrlFetchApp.fetch(webHookUrl, options); // リクエスト
}

// カンマ区切りで午前・午後を判定して文字列を変換する
function split_string(input_text){
    char_num = input_text.length;
    let split_index = char_num;
    let split_flag = false;

    for(let i=0; i < char_num; i++){
        if(input_text[i] == ',' || input_text[i] == '、' || input_text[i] == '/'){
            split_index = i;
            split_flag = true;
        }
    }
    
    // カンマ区切りでの判定
    let place_string = new String();
    if(split_flag){
        place_string = "午前 : ";
        for(let i=0; i < char_num; i++){
            if(i == split_index){
                place_string += ", 午後 : ";
            }else{
                place_string += input_text[i];
            }  
        }
    }else{
        place_string = input_text
    }
    return place_string
}

// 先頭文字の色で午前・午後を判定して文字列を変換
function check_color(input_text, color){
    if(input_text.substr(0, 4) != "午前 :"){
        if(color == '#0000ff' && input_text != ''){
            input_text = "午前 : " + input_text;
        }else if(color == '#ff0000' && input_text != ''){
            input_text = "午後 : " + input_text;
        }else if(color == '#000000' && input_text != ''){
            input_text = "終日 : " + input_text;
        }else if(color == '#9900ff' && input_text != ''){
            input_text = "夜勤 : " + input_text;
        }
    }
    return input_text
}