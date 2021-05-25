// YouTube Data API v3とGASでSpreadSheetをいじる練習をしながら
// REALITYの歌ってみたリストをつくりたい話
// https://docs.google.com/spreadsheets/d/1wKa3VCHcaDbEfxKouLfGT7m42VRUS7cGB1JfhBHfPY4

// スプレッドシートにメニュー追加
function onOpen() {
  var entries = [
    {name: "listByChannelId", functionName: "listByChannelId"},
    {name: "listByPlaylistId", functionName: "listByPlaylistId"}
    ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("YouTube", entries);
}

// channelIdからvideoリストをつくる
// アクティブなセルに書かれた文字列をchannelIdとして読み取り、2行目からリストを生成する。
function listByChannelId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var value = sheet.getActiveCell().getValue();
  mylog("activeCellValue=" + value);

  var videoItems = getVideosFromChannel(value);
  appendLines(sheet, videoItems);  
}

// playlistIdからvideoリストをつくる
// アクティブなセルに書かれた文字列をchannelIdとして読み取り、2行目からリストを生成する。
function listByPlaylistId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var value = sheet.getActiveCell().getValue();
  mylog("activeCellValue=" + value);

  var videoItems = getVideos(value);
  appendLines(sheet, videoItems);  
}

function appendLines(sheet, videoItems) {
  var line = 1;
  for (var j = 0; j < videoItems.length; j++) {
    var video = videoItems[j];
    if (video != null) {
      var public = true;
      if (video.status != null && video.status.privacyStatus == "private") {
        public = false;
      }
      if (public) {
        line++;
        appendLine(sheet, line, video);
      }
    }
  }
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
    // snippet.publishedAtはプレイリストへの追加日なことに注意。
    date = Utilities.formatDate(new Date(video.contentDetails.videoPublishedAt), "JST", "yyyy/MM/dd");
    videoId = video.contentDetails.videoId;
  }
  sheet.getRange(line, 1).setValue(date);
  var imageUrl = snippet.thumbnails["default"].url;
  sheet.getRange(line, 2).setValue("=IMAGE(\"" + imageUrl + "\")"); // セル内に画像を貼る
  sheet.getRange(line, 3).setValue(snippet.title);
  sheet.getRange(line, 4).setValue("https://www.youtube.com/watch?v=" + videoId);
}
    
function getVideosFromChannel(channelId) {
  var nextToken = "";
  var items = [];
  while (nextToken != null) { 
   var res = YouTube.Search.list('id,snippet',{
      maxResults: 50,
      order: `date`,
      type: `video`, // 指定しないとchannelやplaylistもとれてしまう
      channelId: channelId,
      pageToken: nextToken
    });
    for (var i = 0; i < res.items.length; i++) {
      items.push(res.items[i]);
    }
    nextToken = res.nextPageToken;
  }
  mylog("channelId=" + channelId + " count=" + items.length);
  return items;
}

function getVideos(playlistId) {
  var nextToken = "";
  var items = [];
  while (nextToken != null) { 
   var res = YouTube.PlaylistItems.list('snippet,status,contentDetails',{
      maxResults: 50,
      playlistId: playlistId,
      pageToken: nextToken
    });
    for (var i = 0; i < res.items.length; i++) {
      items.push(res.items[i]);
    }
    nextToken = res.nextPageToken;
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
