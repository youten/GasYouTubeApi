// YouTube Data API v3とGASでSpreadSheetをいじる練習をしながら
// REALITYの歌ってみたリストをつくりたい話
// https://docs.google.com/spreadsheets/d/1wKa3VCHcaDbEfxKouLfGT7m42VRUS7cGB1JfhBHfPY4

// スプレッドシートにメニュー追加
function onOpen() {
  var entries = [
    {name: "チャンネルの動画をリストアップ", functionName: "listByChannelId"},
    {name: "プレイリストの動画をリストアップ", functionName: "listByPlaylistId"},
    {name: "チャンネルとプレイリストの最近10日以内の動画をリストアップ", functionName: "listByChannelorPlaylistId"}
    ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("YouTube", entries);
}

// channelIdからvideoリストをつくる
// アクティブなセルに書かれた文字列をchannelIdとして読み取り、2行目からリストを生成する。
function listByChannelId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var value = sheet.getActiveCell().getValue();
  mylog("activeCellValue=" + value);

  appendLines(sheet, 2, getVideosFromChannel(value));
}

// playlistIdからvideoリストをつくる
// アクティブなセルに書かれた文字列をchannelIdとして読み取り、2行目からリストを生成する。
function listByPlaylistId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var value = sheet.getActiveCell().getValue();
  mylog("activeCellValue=" + value);

  appendLines(sheet, 2, getVideosFromPlaylistId(value));
}

// channdlIdまたはPlaylistIdのリストからvideoリストをつくる
// アクティブなセルの文字列リストを読み取り、2行目からリストを生成する
// "PL"から始まる文字列の際にPlaylistIdと判断し、それ以外をchannelIdとする
function listByChannelorPlaylistId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getActiveRange().getValues();
  mylog("activeCellValues.Count=" + values.length);

  var line = 2;
  for (var i = 0; i < values.length; i++) {
    var id = values[i] + "";
    mylog("id=" + id);

    // 最近10daysに絞る
    var date10DaysAgo = new Date();
    date10DaysAgo.setDate(date10DaysAgo.getDate() - 10);

    if (id.indexOf("PL") === 0) { // Playlist id
      line = appendLines(sheet, line, getVideosFromPlaylistId(id, date10DaysAgo));
    } else { // Channel id
      line = appendLines(sheet, line, getVideosFromChannel(id, date10DaysAgo));
    }
  }
}

function appendLines(sheet, startLine, videoItems) {
  var line = startLine;
  for (var j = 0; j < videoItems.length; j++) {
    var video = videoItems[j];
    if (video != null) {
      var public = true;
      if (video.status != null && video.status.privacyStatus == "private") {
        public = false;
      }
      if (public) {
        appendLine(sheet, line, video);
        line++;
      }
    }
  }
  return line;
}

function appendLine(sheet, line, video) {
  var snippet = video.snippet;
  mylog(snippet.title);
  var videoId = "";
  // 動画公開日、Search APIではsnippet.publishedAtで取得できる
  var date = Utilities.formatDate(new Date(snippet.publishedAt), "JST", "yyyy/MM/dd");
  if (video.id != null) {
    videoId = video.id.videoId;
  }
  if (video.contentDetails != null) {
    // 動画公開日、PlaylistItems APIではcontentDetails.videoPublishedAtで取得できる。
    // snippet.publishedAtはプレイリストへの追加日で違うので上書きする。
    date = Utilities.formatDate(new Date(video.contentDetails.videoPublishedAt), "JST", "yyyy/MM/dd");
    videoId = video.contentDetails.videoId;
  }
  var col = 4;
  sheet.getRange(line, col).setValue(date);
  var imageUrl = snippet.thumbnails["default"].url;
  sheet.getRange(line, col + 1).setValue("=IMAGE(\"" + imageUrl + "\")"); // セル内に画像を貼る
  sheet.getRange(line, col + 2).setValue(snippet.title);
  sheet.getRange(line, col + 3).setValue("https://www.youtube.com/watch?v=" + videoId);
  // note
  var nowDate = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  sheet.getRange(line, col + 5).setValue(nowDate); // register date
  sheet.getRange(line, col + 6).setValue("\n\n\n"); // dummy \n x3
}
    
function getVideosFromChannel(channelId, publishedAfterDate) {  
  var nextToken = "";
  var items = []

  var publishedAfter = "";
  if (publishedAfterDate != null){
    publishedAfter = publishedAfterDate.toISOString();
    mylog("publishedAfter=" + publishedAfter);
  }

  while (nextToken != null) { 
    try {
      var res = YouTube.Search.list('id,snippet',{
        maxResults: 50,
        publishedAfter: publishedAfter,
        order: `date`,
        type: `video`, // 指定しないとchannelやplaylistもとれてしまう
        channelId: channelId,
        pageToken: nextToken
      });
      for (var i = 0; i < res.items.length; i++) {
        items.push(res.items[i]);
      }
      nextToken = res.nextPageToken;      
    } catch(e) {
      mylog(e);
      break;
    }
  }
  mylog("channelId=" + channelId + " count=" + items.length);
  return items;
}

function getVideosFromPlaylistId(playlistId, publishedAfterDate) {
  var nextToken = "";
  var items = [];
  while (nextToken != null) {
   try {
    var res = YouTube.PlaylistItems.list('snippet,status,contentDetails',{
        maxResults: 50,
        playlistId: playlistId,
        pageToken: nextToken
      });
      for (var i = 0; i < res.items.length; i++) {
        // publishedAfterDateが指定されていたら、それより前の動画は追加しない
        if (publishedAfterDate != null) {
          var videoPublishedDate = new Date(res.items[i].contentDetails.videoPublishedAt);
          if (publishedAfterDate > videoPublishedDate){
            mylog ("skip, old videoPublishedAt=" + videoPublishedDate);
            continue;
          }
        }        
        items.push(res.items[i]);
      }
      nextToken = res.nextPageToken;
   } catch(e) {
     mylog(e);
     break;
   }
  }
  mylog("playlistId=" + playlistId + " count=" + items.length);
  return items;
}

// debug用。openByIdの引数は書き込み権限のある（本スクリプトを動作させる）スプレッドシートのIDに書き換えて使う。
function mylog(value) {
  var ss = SpreadsheetApp.openById("1wKa3VCHcaDbEfxKouLfGT7m42VRUS7cGB1JfhBHfPY4");
  var sheet = ss.getSheetByName("log");
  sheet.appendRow([new Date(), value]);
}

