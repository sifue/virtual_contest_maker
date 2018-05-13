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
  var urlTasks = 'https://beta.atcoder.jp/contests/' + contestId + '/tasks';
  var requestTasks = UrlFetchApp.fetch(urlTasks);
  var contentTasks = requestTasks.getContentText();

  // 問題の全レベルを取得
  var taskLevels = [];
  var taskLevelRegex = new RegExp('<td class="text-center no-break"><a href=\'/contests/[a-z0-9_]+/tasks/[a-z0-9_]+\'>([A-Z]+)</a></td>', 'g'); 
  var isEndTask = true;
  while(isEndTask) {
    var selectedLevel = taskLevelRegex.exec(contentTasks);
    if(selectedLevel) {
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
  while(isEndTask) {
    var selectedTask = taskRegex.exec(contentTasks);
    if(selectedTask) {
      tasks.push(selectedTask[2]);
      taskLinks.push(selectedTask[1]);
    } else {
      isEndTask = false;
    }
  }

  // 問題とレベルを結合しながら、表と問題へのリンクを更新
  for (var index = 0; index < tasks.length; index++) {
    tasks[index] = taskLevels[index] + ' - ' +  tasks[index];
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

  Browser.msgBox(Session.getActiveUser() +
    'を追加しました。AtCoder IDをシートに記入してください。', Browser.Buttons.OK);
}

/**
 * コンテストの参加者の点数の状況を更新する
 */
function updateStatus() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

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
  var beginTime = sheet.getRange('B3').getValue();
  var endTime = sheet.getRange('B4').getValue();

  // 全ユーザーの提出情報をオブジェクトにまとめる
  // Key: ユーザー名 Value: Submission[]
  // Submission: {time, task, id, score, status}
  var allSubmissions = {};
  // Key: ユーザー名 Value: {taskScoreMaxs: {}, notAcCounts: {}, totalScore}
  var allResult = {};

  for (var index = 0; index < participants.length; index++) {
    var user = participants[index];
    // 提出を取得する
    var urlSubmissions = 'https://beta.atcoder.jp/contests/' + contestId + '/submissions?f.Task=&f.Language=&f.Status=&f.User=' + user;
    var requestSubmissions = UrlFetchApp.fetch(urlSubmissions);
    var contentTasks = requestSubmissions.getContentText();

    // 時間を取得
    var times = [];
    var timesRegex = new RegExp('<td class="no-break"><time class=\'fixtime fixtime-second\'>(.+)</time></td>', 'g'); 
    var isEnd = true;
    while(isEnd) {
      var selectedTime = timesRegex.exec(contentTasks);
      if(selectedTime) {
        times.push(selectedTime[1]);
      } else {
        isEnd = false;
      }
    }

    // 問題を取得
    var submittedTasks = [];
    var submittedTaskssRegex = new RegExp('<td><a href=\'/contests/[a-z0-9_]+/tasks/[a-z0-9_]+\'>(.+)</a></td>', 'g'); 
    isEnd = true;
    while(isEnd) {
      var selectedTask = submittedTaskssRegex.exec(contentTasks);
      if(selectedTask) {
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
    while(isEnd) {
      var selectedScore = scoresRegex.exec(contentTasks);
      if(selectedScore) {
        scores.push(selectedScore[2]);
        ids.push(selectedScore[1]);
      } else {
        isEnd = false;
      }
    }

    // ステータス
    var statuses = [];
    var statusesRegex = new RegExp('<td class=\'text-center\'><span class=\'.+\' aria-hidden=\'true\' data-toggle=\'tooltip\' data-placement=\'top\' title=".+">(.+)</span></td>', 'g'); 
    isEnd = true;
    while(isEnd) {
      var selectedStatus = statusesRegex.exec(contentTasks);
      if(selectedStatus) {
        statuses.push(selectedStatus[1]);
      } else {
        isEnd = false;
      }
    }

    var userSubmissions = [];
    for (var j = 0; j < times.length; j++) {
      userSubmissions.push({
        time : times[j],
        task : submittedTasks[j],
        id : ids[j],
        score : scores[j],
        status : statuses[j]
      });
    }
    allSubmissions[user] = userSubmissions;

    // 集計結果
    var result = {
      taskScoreMaxs : {},
      notAcCounts : {},
      totalScore : 0
    };

    // それぞれのスコア最大値やAC以外の数をを計算
    for (var k = 0; k < submittedTasks.length; k++) {
      var task = submittedTasks[k];
      if(!result.taskScoreMaxs[task]) {
        result.taskScoreMaxs[task] = 0;
      }
      if(!result.notAcCounts[task]) {
        result.notAcCounts[task] = 0;
      }

      for (var l = 0; l < userSubmissions.length; l++) {
        var submission = userSubmissions[l];
        if(task === submission.task) {
          var score = parseInt(submission.score);
          if(score > result.taskScoreMaxs[task]) {
            result.taskScoreMaxs[task] = score;
          }

          if(submission.status !== 'AC') {
            result.notAcCounts[task] += 1;
          }
        }
      }
    } // task

    var totalScore = 0;
    for (var m = 0; m < tasks.length; m++) {
      if(isFinite(result.taskScoreMaxs[tasks[m]])) {
        totalScore += result.taskScoreMaxs[tasks[m]];
      }
    }
    result.totalScore = totalScore;
    allResult[user] = result;

    // スプレッドシートに更新を行う
    var values = [[result.totalScore]];
    for (var n = 0; n < tasks.length; n++) {
      var value = '';
      if (isFinite(result.taskScoreMaxs[tasks[n]])) {
        value += result.taskScoreMaxs[tasks[n]];
      }

      if (result.notAcCounts[tasks[n]] > 0) {
        value += '(' + result.notAcCounts[tasks[n]] + ')';
      }
      values[0].push(value);
    }
    sheet.getRange(6 + index, 3, 1, 1 + tasks.length).setValues(values);

    // TODO 更新時間を保存、1分以内のものを無視する
  } // user

  // TODO 時間で絞る

  // Logger.log(allSubmissions);
  Logger.log(allResult);
  Logger.log(participants.length + '名の参加者の点数の更新を行いました: ' + participants);
}
