// スプレッドシートの取得
var spreadSheet = SpreadsheetApp.openByUrl('ここにスプレッドシートのURLを挿入');

// シートの取得
var participantSheet = spreadSheet.getSheetByName("参加者");
var historySheet = spreadSheet.getSheetByName("履歴");

// チーム分けの実行：最良のグループの組み合わせを求めてログに吐き出す
function createGroup() {
  const peopleLimitOfGroup = participantSheet.getRange(2, 3).getValue(); // 1グループの最大人数
  const historyLastRowNumber = historySheet.getLastRow();
  const historyLastColumnNumber = historySheet.getLastColumn();
  const historySheetValues = historySheet.getRange(2, 1, historyLastRowNumber - 1, historyLastColumnNumber).getValues();

  let participantIds = setParticipantIds();
  let bestGroups = [];
  let bestGroupsPatternScore = 0;

  const shuffleLimitCount = 1000;

  // グループのシャッフルを shuffleLimitCount 回繰り返すことで、最良のグループを見つける仕組みになっている
  for (let shuffleTryCount = 1; shuffleTryCount <= shuffleLimitCount; ++shuffleTryCount) {
    const groups = slicesArray(participantIds, peopleLimitOfGroup);
    let currentPatternScore = 0; // この値が低いほど、よいパターンと言える

    for (let groupId = 0; groupId < groups.length; ++groupId) {
      const groupMembers = groups[groupId];
      let groupScore = 0; // この値が低いほど理想的なグループ

      for (let groupMemberX = 0; groupMemberX < groupMembers.length - 1; ++groupMemberX) {
        let score = 0;
        const columnX = groupMembers[groupMemberX];
        const xHistory = transposeDoubleArray(historySheetValues)[columnX - 1]; // グループ groupId 内の X さんの参加履歴

        for (let groupMemberY = groupMemberX + 1; groupMemberY < groupMembers.length; ++groupMemberY) {
          const columnY = groupMembers[groupMemberY];
          const yHistory = transposeDoubleArray(historySheetValues)[columnY - 1]; // グループ groupId 内の Y さんの参加履歴

          for (let historyRow = 0; historyRow < xHistory.length; ++historyRow) {
            if (xHistory[historyRow] === '' || yHistory[historyRow] === '') {
              // 同じ値であっても、どっちも不参加の場合はスコア 0.5 倍の処理
              score *= 0.5;
            } else if (xHistory[historyRow] === yHistory[historyRow]) {
              // 同じ値である（同じグループになっている）ならスコア +1 の処理
              score += 1;
            } else {
              // それ以外ならスコア 0.5 倍の処理
              score *= 0.5;
            }
          }
          groupScore += score;
        }
      }
      currentPatternScore += groupScore;
    }

    if (bestGroupsPatternScore === 0) {
      bestGroupsPatternScore = currentPatternScore;
      bestGroups = groups;
    } else if (bestGroupsPatternScore > currentPatternScore) {
      bestGroupsPatternScore = currentPatternScore;
      bestGroups = groups;
    }

    // [1, 2, 3] ... のような順番のまま出力されたらバグがあると気付けるよう、
    // shuffle はループの最後におこなう
    shuffle(participantIds);
  }
  Logger.log(bestGroups);

  let groupCharactersArray = [];
  bestGroups.forEach((memberIds, groupIdx) => {
    memberIds.forEach(memberId => {
      groupCharactersArray[memberId - 1] = String.fromCharCode(65 + groupIdx);
    })
  });
  Logger.log(groupCharactersArray);
  Logger.log(groupCharactersArray.length);
  historySheet.appendRow(groupCharactersArray);
}

// ランチ参加者のIDを一次配列で取得する
function setParticipantIds() {
  const lastRow = participantSheet.getLastRow();
  const idsArray = participantSheet.getRange(2, 2, lastRow - 1).getValues();
  const participantIds = transposeDoubleArray(idsArray)[0];
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

// 二次元配列の行と列を入れ替える
// オリジナル: https://script.google.com/home/projects/1r9wNWbta3ebuYL4ENAdIp4UYKmyNiWf1AqsXYzfXduRHhTZEeTxS9MhZ/edit
function transposeDoubleArray(data) {
  if (data.length > 0) {
    var r = [];
    for (var i = 0; i < data[0].length; i++) {
      var newRow = [];
      for (var j = 0; j < data.length; j++) {
        newRow[j] = data[j][i];
      }
      r[i] = newRow;
    }
    return r;
  }
  else {
    return data;
  }
}
