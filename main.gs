// スプレッドシートの取得
var spreadSheet = SpreadsheetApp.openByUrl('ここにスプレッドシートのURLを挿入');

// シートの取得
var participantSheet = spreadSheet.getSheetByName("参加者");
var historySheet = spreadSheet.getSheetByName("履歴");

// チーム分けの実行
function createGroup() {
  var participantIds = setParticipantIds();
  var bestScore = 0;
  var bestGroup = [];

  for(var i = 0; i < 200; i++) {
    var groups = slicesArray(participantIds, participantSheet.getRange(2, 3).getValue());
    var patternScore = 0;

    for(var j = 0; j < groups.length; j++) {
      var groupMembers = groups[j];
      var groupScore = 0;

      for(var k = 0; k < groupMembers.length - 1; k++) {
        var score = 0;
        var xHistory = historySheet.getRange(2, groupMembers[k], historySheet.getLastRow()).getValues();

        for(var l = k + 1; l < groupMembers.length; l++){
          var yHistory = historySheet.getRange(2, groupMembers[l], historySheet.getLastRow()).getValues();

          for(var m = 0; m < xHistory.length; m++) {
            if(xHistory[m][0] == "" || yHistory[m][0] == ""){
              score *= 0.5;
            } else if(xHistory[m][0] == yHistory[m][0]){
              score += 1;
            } else {
              score *= 0.5;
            }
          }
          groupScore += score;
        }
      }
      patternScore += groupScore;
    }
    if(bestScore == 0) {
      bestScore = patternScore;
      bestGroup = groups;
    } else if(bestScore > patternScore) {
      bestScore = patternScore;
      bestGroup = groups;
    }
    shuffle(participantIds);
  }
  Logger.log(bestGroup);
}

// ランチ参加者のIDを一次配列で取得する
function setParticipantIds() {
  const lastRow = participantSheet.getLastRow();
  const idsArray = participantSheet.getRange(2, 2, lastRow - 1).getValues();
  const participantIds = ArrayLib.transpose(idsArray)[0];
  const participantIdsWithoutAbsence = participantIds.filter(id => id !== '' && id !== 0 && id !== 'x');

  return participantIdsWithoutAbsence;
}

// 配列 array を、指定した最大個数 limitLength ずつに分割する
// 二次元配列になって返る
function slicesArray(array, limitLength) {
  var index = 0;
  var results = [];

  while (array.length > index + limitLength) {
    const result = array.slice(index, index + limitLength);
    results.push(result);
    index += limitLength;
  }

  const rest = array.slice(index, array.length + 1);
  results.push(rest);
  return results;
}

// 配列をシャッフルする
function shuffle(array) {
  for (var i = array.length - 1; i > 0; --i) {
    var r = Math.floor(Math.random() * (i + 1));
    var tmp = array[i];
    array[i] = array[r];
    array[r] = tmp;
  }
}
