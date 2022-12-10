// YouTube Data API v3とGASでSpreadSheetをいじる練習をしながら
// REALITYの歌ってみたリストをつくりたい話
// https://docs.google.com/spreadsheets/d/1wKa3VCHcaDbEfxKouLfGT7m42VRUS7cGB1JfhBHfPY4

// スプレッドシートにメニュー追加
function onOpen() {
  var entries = [
    {name: "チャンネルの動画をリストアップ", functionName: "listByChannelId"},
    {name: "プレイリストの動画をリストアップ", functionName: "listByPlaylistId"},
    {name: "チャンネルとプレイリストの最近10日以内の動画をリストアップ", functionName: "listByChannelorPlaylistId"},
    {name: "HTML1セル生成", functionName: "generateMusicHtml"},
    {name: "選択範囲を結合コピー", functionName: "concatRangeTextTo2N"},
    ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("YouTube", entries);
}

// ハンドル（@なんたら）からキーワードでチャンネル検索して先頭のものを返す
function getChannelIdFromKeyword(value){
  try {
    var res = YouTube.Search.list('snippet',{
      maxResults: 10,
      type: `channel`,
      q: value
    });
    
    mylog("q=" + value + " length=" + res.items.length);
    for (var i = 0; i < res.items.length; i++) {
      mylog("[" + i + "] channelId=" + res.items[i].id.channelId
       + " title=" + res.items[i].snippet.channelTitle
       + " description=" + res.items[i].snippet.description);
    }

    if (res.items.length > 0) {
      return res.items[0].id.channelId;
    }
  } catch(e) {
    // do nothing
  }
  return null;
}

// 選択された行からHTMLを1セル生成する
// 8:urlYouTubeMusic
// 2:name
// 3:urlReality
// 4:urlTwitter
// 5:urlYouTubeChannel
// 6:publishDate
// 9:title
// 13:html
function generateMusicHtml() {
  // 1曲分のcolな部分HTML文字列の構成要素
  const HTML_PREFIX_URL_YOUTUBE_MUSIC = "<div class=\"col\"> <div class=\"card shadow-sm\"> <div class=\"ratio ratio-16x9\"> <iframe loading=\"lazy\" src=\"https://www.youtube.com/embed/";
  const HTML_PREFIX_TITLE = "?controls=0\" frameborder=\"0\" allow=\"accelerometer; clipboard-write; encrypted-media; gyroscope; picture-in-picture\" allowfullscreen></iframe> </div> <div class=\"card-body\"><p class=\"card-text\">";
  const HTML_PREFIX_NAME = "<br><small class=\"text-muted\">";
  const HTML_PREFIX_URL_REALITY = "</small></p> <div class=\"d-flex justify-content-between align-items-center\"> <div class=\"btn-group\">\n<a href=\"";
  const HTML_PREFIX_URL_TWITTER = "\"><img src=\"./img/icon-reality.png\"></a> <a href=\"";
  const HTML_PREFIX_URL_YOUTUBE_CHANNEL = "\"><img src=\"./img/icon-twitter.png\"></a> <a href=\"";
  const HTML_PREFIX_PUBLISH_DATE = "\"><img src=\"./img/icon-youtube.png\"></a> </div><small class=\"text-muted\">";
  const HTML_SUFFIX = "</small></div></div></div></div>\n\n";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var firstLine = range.getRow(); // 先頭行番号
  var rows = range.getNumRows(); // 行数
  mylog("generateMusicHtml() start line=" + line + " rows=" + rows);

  for (var line = firstLine; line < firstLine + rows; line++) {
    var publishDate = sheet.getRange(line, 6).getDisplayValue();
    var title = sheet.getRange(line, 8).getValue();
    var name = sheet.getRange(line, 2).getValue();
    var html = "<!-- " + publishDate + " " + title + " " + name + " -->\n"; // ヘッダHTMLコメント行

    var musicId = sheet.getRange(line, 9).getValue().replace("https://www.youtube.com/watch?v=", "");
    html += HTML_PREFIX_URL_YOUTUBE_MUSIC + musicId;
    html += HTML_PREFIX_TITLE + title;
    html += HTML_PREFIX_NAME + name;

    var urlReality = sheet.getRange(line, 3).getValue();
    if (urlReality.indexOf("http") < 0) {
      urlReality = ""; // イレギュラー表示内容の際にはURLを空にする
    }
    html += HTML_PREFIX_URL_REALITY + urlReality;

    var urlTwitter = sheet.getRange(line, 4).getValue();
    if (urlTwitter.indexOf("http") < 0) {
      urlTwitter = ""; // イレギュラー表示内容の際にはURLを空にする
    }
    html += HTML_PREFIX_URL_TWITTER + urlTwitter;

    html += HTML_PREFIX_URL_YOUTUBE_CHANNEL + sheet.getRange(line, 5).getValue();
    html += HTML_PREFIX_PUBLISH_DATE + publishDate;
    html += HTML_SUFFIX;
    sheet.getRange(line, 13).setValue(html);
  }
}

// 選択セルの表示値を文字列連結して特定セルに書き込む
function concatRangeTextTo2N() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getActiveRange().getDisplayValues();
  var text = "";
  for (var i = 0; i < values.length; i++) {
    text += values[i];
  }
  sheet.getRange("N2").setValue(text);
}

// channelIdまたはハンドルから検索結果でvideoリストをつくる
// アクティブなセルに書かれた文字列をchannelIdまたはハンドルとして読み取り、2行目からリストを生成する。
function listByChannelId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var value = sheet.getActiveCell().getValue();
  mylog("activeCellValue=" + value);

  // @から始まる際にはChannelIdじゃなくてハンドルとして検索からChannelId化を試みる
  if (value.startsWith("@")) {
    value = getChannelIdFromKeyword(value);
  }

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
    // 4列目のセルが空じゃない際には上書きしないで次の行にする
    while (!sheet.getRange(line, 4).isBlank()) {
      line++;
    }

    var video = videoItems[j];
    if (video != null) {
      if (video.status == null || video.status.privacyStatus != "private") {
        // private属性が確認できない、publicの際のみappendする
        var result = appendLine(sheet, line, video);
        if (result) {
          line++;
        }
      }
    }
  }
  return line;
}

function appendLine(sheet, line, video) {
  var snippet = video.snippet;
  if (!snippet.thumbnails["default"] || !snippet.thumbnails["default"].url || snippet.title == "Deleted video") {
    return false;
  }
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
  sheet.getRange(line, col + 6).setValue("\n\n\n\n"); // dummy \n x
  return true;
}
    
function getVideosFromChannel(channelId, publishedAfterDate) {
  var nextToken = "";
  var items = [];

  var publishedAfter = "2000-01-01T00:00:00Z";
  if (publishedAfterDate != null) {
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
        // liveBroadcastContent:"upcoming" はプレミア公開前なので除外
        if (res.items[i].snippet.liveBroadcastContent === "upcoming") {
          mylog ("skip, upcoming title=" + res.items[i].snippet.title);
          continue;
        }
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
            mylog ("skip, old videoPublishedAt=" + videoPublishedDate + " title=" + res.items[i].snippet.title);
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
  console.log(value);
  var ss = SpreadsheetApp.openById("1wKa3VCHcaDbEfxKouLfGT7m42VRUS7cGB1JfhBHfPY4");
  var sheet = ss.getSheetByName("log");
  sheet.appendRow([new Date(), value]);
}

