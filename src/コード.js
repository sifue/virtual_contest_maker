// 制約:
// 時間のパースの設定上スクリプトの実行環境がJSTでなくてはならない
// Logger.log('hoge'); 
// の時のログの表示がJSTとなっていることが実行条件となる。
// バージョン:
// v1.0.0

/**
 * 作成シートでのコンテスト作成を実行する
 */
function createNewVirtualCotest() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var createSheet = spreadsheet.getActiveSheet();

  var title = createSheet.getRange('B2').getValue();
  var contestId = createSheet.getRange('B3').getValue();
  var beginTime = createSheet.getRange('B4').getValue();
  var endTime = createSheet.getRange('B5').getValue();

  // シートをコピーして作成
  var copyOriginSheet = spreadsheet.getSheetByName('コピー元');
  copyOriginSheet.activate();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();

  sheet.setName(title);
  sheet.getRange('B1').setValue(title);
  sheet.getRange('B2').setValue(contestId);
  sheet.getRange('B3').setValue(beginTime);
  sheet.getRange('B4').setValue(endTime);

  // コンテストの情報の取得して設定する
  // 例: https://beta.atcoder.jp/contests/abc052/tasks
  var urlTasks = 'https://beta.atcoder.jp/contests/' + contestId + '/tasks';
  var requestTasks = UrlFetchApp.fetch(urlTasks);
  var contentTasks = requestTasks.getContentText();

  // 問題の全レベルを取得
  var taskLevels = [];
  var taskLevelRegex = new RegExp('<td class="text-center no-break"><a href=\'/contests/[a-z0-9_]+/tasks/[a-z0-9_]+\'>([A-Z]+)</a></td>', 'g');
  var isEndTask = true;
  while (isEndTask) {
    var selectedLevel = taskLevelRegex.exec(contentTasks);
    if (selectedLevel) {
      taskLevels.push(selectedLevel[1]);
    } else {
      isEndTask = false;
    }
  }

  // 問題の全タイトルとリンクを取得
  var tasks = [];
  var taskLinks = [];
  var taskRegex = new RegExp('<td><a href=\'(/contests/[a-z0-9_]+/tasks/[a-z0-9_]+)\'>(.+)</a></td>', 'g');
  isEndTask = true;
  while (isEndTask) {
    var selectedTask = taskRegex.exec(contentTasks);
    if (selectedTask) {
      tasks.push(selectedTask[2]);
      taskLinks.push(selectedTask[1]);
    } else {
      isEndTask = false;
    }
  }

  // 問題とレベルを結合しながら、表と問題へのリンクを更新
  for (var index = 0; index < tasks.length; index++) {
    tasks[index] = taskLevels[index] + ' - ' + tasks[index];
    sheet.getRange(5, 4 + index).setFormula(
      '=HYPERLINK("https://beta.atcoder.jp' + taskLinks[index] + '","' + tasks[index] + '")');
  }

  Logger.log('Virtual Contest が作成されました:' + title);
}

/**
 * コンテストシートでの参加を実行する
 */
function join() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var sheetName = sheet.getName();

  var currentParticipantCount = sheet.getRange('H1').getValue();
  var writeRange = sheet.getRange(6 + currentParticipantCount, 1);
  writeRange.setValue(Session.getActiveUser());
  sheet.getRange('H1').setValue(currentParticipantCount + 1);

  // 記入セルを選択
  sheet.getRange(6 + currentParticipantCount, 2).activate();

  Browser.msgBox(Session.getActiveUser() +
    'を追加しました。AtCoder IDをシートに記入してください。', Browser.Buttons.OK);
}

/**
 * AtCoder形式のString文字列をパースしてDateオブジェクトにする
 * @param {*} dataString 例: 2018-04-25 18:09:54+0900
 */
function parseDate(dataString) {
  var r = /(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})\+0900/.exec(dataString);
  return new Date(
    parseInt(r[1], 10),
    parseInt(r[2], 10) - 1,
    parseInt(r[3], 10),
    parseInt(r[4], 10),
    parseInt(r[5], 10),
    parseInt(r[6], 10)
  );
}

/**
 * 数値を2桁0埋めした文字列に変換する
 * @param {*} number 
 */
function formatDigit(number) {
  return ('0' + number).slice(-2);
}

/**
 * 直近の点数獲得時間をフォーマットする
 * @param {*} latestScoreTime 
 * @param {*} beginTIme 
 */
function formatLatestScoreTime(latestScoreTime, beginTime) {
  var secs = (latestScoreTime.getTime() - beginTime.getTime()) / 1000;
  var min = Math.floor(secs / 60);
  var sec = secs % 60;
  return ' [' + min + ':' + formatDigit(sec) + ']';
}

/**
 * コンテストの参加者の点数の状況を更新する
 */
function updateStatus() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // 最終更新を設定
  sheet.getRange('H2').setValue(new Date());

  // AtCoderユーザー一覧を作成する
  var valuesParticipant = sheet.getRange('B6:B999').getValues();
  var participants = [];
  for (var i = 0; i < valuesParticipant.length; i++) {
    if (valuesParticipant[i][0]) {
      participants.push(valuesParticipant[i][0]);
    }
  }

  // 問題一覧を作る
  var valuesTask = sheet.getRange('D5:Z5').getValues();
  var tasks = [];
  for (var i = 0; i < valuesTask[0].length; i++) {
    if (valuesTask[0][i]) {
      tasks.push(valuesTask[0][i]);
    }
  }

  // コンテストID, 開始, 終了を取得する
  var contestId = sheet.getRange('B2').getValue();
  var beginTime = sheet.getRange('B3').getValue(); // Date
  var endTime = sheet.getRange('B4').getValue(); // Date

  // 全ユーザーの提出情報をオブジェクトにまとめる
  // Key: ユーザー名 Value: Submission[]
  // Submission: {time, task, id, score, status}
  var allSubmissions = {};
  var allSelectedSubmissions = {};
  // Key: ユーザー名 Value: {taskScoreMaxs: {}, notAcCounts: {}, totalScore}
  var allResult = {};

  for (var index = 0; index < participants.length; index++) {
    var user = participants[index];
    // 提出を取得する
    // 例: https://beta.atcoder.jp/contests/abc054/submissions?f.Task=&f.Language=&f.Status=&f.User=sifue
    var urlSubmissions = 'https://beta.atcoder.jp/contests/' + contestId + '/submissions?f.Task=&f.Language=&f.Status=&f.User=' + user;
    var requestSubmissions = UrlFetchApp.fetch(urlSubmissions);
    var contentTasks = requestSubmissions.getContentText();

    // 時間を取得
    var times = [];
    var timesRegex = new RegExp('<td class="no-break"><time class=\'fixtime fixtime-second\'>(.+)</time></td>', 'g');
    var isEnd = true;
    while (isEnd) {
      var selectedTime = timesRegex.exec(contentTasks);
      if (selectedTime) {
        times.push(parseDate(selectedTime[1]));
      } else {
        isEnd = false;
      }
    }

    // 問題を取得
    var submittedTasks = [];
    var submittedTaskssRegex = new RegExp('<td><a href=\'/contests/[a-z0-9_]+/tasks/[a-z0-9_]+\'>(.+)</a></td>', 'g');
    isEnd = true;
    while (isEnd) {
      var selectedTask = submittedTaskssRegex.exec(contentTasks);
      if (selectedTask) {
        submittedTasks.push(selectedTask[1]);
      } else {
        isEnd = false;
      }
    }

    // スコアとIDを取得
    var scores = [];
    var ids = [];
    var scoresRegex = new RegExp('<td class="text-right submission-score" data-id="([0-9]+)">([0-9]+)</td>', 'g');
    isEnd = true;
    while (isEnd) {
      var selectedScore = scoresRegex.exec(contentTasks);
      if (selectedScore) {
        scores.push(selectedScore[2]);
        ids.push(selectedScore[1]);
      } else {
        isEnd = false;
      }
    }

    // ステータス
    var statuses = [];
    var statusesRegex = new RegExp('<span class=\'.+\' aria-hidden=\'true\' data-toggle=\'tooltip\' data-placement=\'top\' title=".+">(.+)</span></td>', 'g');
    isEnd = true;
    while (isEnd) {
      var selectedStatus = statusesRegex.exec(contentTasks);
      if (selectedStatus) {
        statuses.push(selectedStatus[1]);
      } else {
        isEnd = false;
      }
    }

    // 全部の提出と時間帯にマッチする提出を選択
    var userSubmissions = [];
    var selectedUserSubmissions = [];
    for (var j = 0; j < times.length; j++) {
      var submission = {
        time: times[j],
        task: submittedTasks[j],
        id: ids[j],
        score: scores[j],
        status: statuses[j]
      };
      userSubmissions.push(submission);

      if (beginTime <= submission.time && submission.time <= endTime) {
        selectedUserSubmissions.push(submission);
      }
    }
    allSubmissions[user] = userSubmissions;
    allSelectedSubmissions[user] = selectedUserSubmissions;

    // 集計結果
    var result = {
      taskScoreMaxs: {},
      taskLatestScoreTime: {},
      notAcCounts: {},
      totalScore: 0,
      latestScoreTime: new Date(0)
    };

    // それぞれのスコア最大値やAC以外の数をを計算
    for (var k = 0; k < tasks.length; k++) {
      var task = tasks[k];

      for (var l = 0; l < selectedUserSubmissions.length; l++) {
        var submission = selectedUserSubmissions[l];
        if (task === submission.task) {

          // 無ければ初期化
          if (!result.taskScoreMaxs[task]) {
            result.taskScoreMaxs[task] = 0;
          }
          if (!result.notAcCounts[task]) {
            result.notAcCounts[task] = 0;
          }

          var score = parseInt(submission.score);
          if (score > result.taskScoreMaxs[task]) {
            result.taskScoreMaxs[task] = score;
            result.taskLatestScoreTime[task] = submission.time;
          }

          if (submission.status !== 'AC') {
            result.notAcCounts[task] += 1;
          }
        }
      }
    } // task

    // 集計スコアと最終AC時間を取得する
    var totalScore = 0;
    for (var m = 0; m < tasks.length; m++) {
      if (isFinite(result.taskScoreMaxs[tasks[m]])) {
        totalScore += result.taskScoreMaxs[tasks[m]];
        if (result.taskLatestScoreTime[tasks[m]] > result.latestScoreTime) {
          result.latestScoreTime = result.taskLatestScoreTime[tasks[m]];
        }
      }
    }
    result.totalScore = totalScore;
    allResult[user] = result;

    // スプレッドシートに更新を行う
    var values;
    if (result.totalScore) {
      values =  [[ '' + result.totalScore + formatLatestScoreTime(result.latestScoreTime, beginTime)]];
    } else {
      values =  [[ 0 ]];
    }
    
    for (var n = 0; n < tasks.length; n++) {
      var value = '';
      if (isFinite(result.taskScoreMaxs[tasks[n]])) {
        value += result.taskScoreMaxs[tasks[n]];
      }

      if (result.notAcCounts[tasks[n]] > 0) {
        value += ' (' + result.notAcCounts[tasks[n]] + ')';
      }

      if (isFinite(result.taskScoreMaxs[tasks[n]])
          && result.taskLatestScoreTime[tasks[n]]
    ) {
        value +=  formatLatestScoreTime(result.taskLatestScoreTime[tasks[n]], beginTime);
      }
      values[0].push(value);
    }
    sheet.getRange(6 + index, 3, 1, 1 + tasks.length).setValues(values);

    // TODO WANT ユーザーごとの更新時間をPropertiesServiceに保存、1分以内のものを無視する
    // 今のところ1名5秒程度なので360秒制限は、参加者70名ぐらいにならないと顕在化しなさそう
  } // user

  // Logger.log('allSubmissions:');
  // Logger.log(allSubmissions);
  Logger.log('allSelectedSubmissions:');
  Logger.log(allSelectedSubmissions);
  Logger.log('allResult:');
  Logger.log(allResult);
  Logger.log(participants.length + '名の参加者の点数の更新を行いました: ' + participants);
}
