function distributePeopleAndScores() {
  //アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // '<レベルと名前が書いてあるシート>'という名前のシートを取得
  var sheet = ss.getSheetByName('<レベルと名前が書いてあるシート>');
  // '<結果を表示するシート>'という名前のシートを取得
  var resultSheet = ss.getSheetByName('<結果を表示するシート>');
  // セルA2からB14までのデータを取得
  var data = sheet.getRange("A2:B14").getValues();

  // 空の行または'欠席'と記載されている行を除外するフィルタリング処理
  data = data.filter(function (row) {
    return row[0] !== "" && row[1] !== "" && row[0] !== "欠席" && row[1] !== "欠席";
  });

  // <AAA>と<BBB>はそれぞれ第一、第二のグループに固定で入るメンバーを示す
  var fixedMembers = { "<AAA>": 0, "<BBB>": 1 };
  // グループを作成（2つの空配列）
  var groups = [[], []];

  // 特定のメンバーを固定したグループに分ける。それ以外のメンバーを新たなdata配列に戻す。
  data = data.filter(function (person) {
    var name = person[0];
    // もし名前が固定メンバーの中にあれば
    if (name in fixedMembers) {
      // その人を指定したグループに追加
      groups[fixedMembers[name]].push(person);
      // フィルターから除外（新たなdata配列には追加しない）
      return false;
    }
    // それ以外の人は新たなdata配列に追加する（フィルターから除外しない）
    return true;
  });

  // 固定メンバーを保存し、それらをグループから一時的に除外する
  var fixedGroupMembers = [];
  for (var name in fixedMembers) {
    var groupIndex = fixedMembers[name];
    // 名前が一致するメンバーのインデックスを探す
    var memberIndex = groups[groupIndex].findIndex((person) => person[0] === name);
    // そのメンバーをグループから除外し、固定メンバー配列に保存
    fixedGroupMembers.push(groups[groupIndex].splice(memberIndex, 1)[0]);
  }

  // データをランダムに並べ替える
  data.sort(function () {
    return 0.5 - Math.random();
  });

  // 点数の総数を計算
  var totalScore = data.reduce(function (sum, person) {
    return sum + person[1];
  }, 0);

  // 総点数の半分を計算
  var halfScore = totalScore / 2;

  // 現在のグループと現在のスコアを追跡するための変数を初期化
  var currentGroup = 0;
  var currentScore = 0;

  //  データを巡回し、現在のグループの点数が半分を超えたらグループを切り替え
  for (var i = 0; i < data.length; i++) {
    if (currentScore + data[i][1] > halfScore && currentGroup == 0) {
      currentGroup = 1;
    }
    groups[currentGroup].push(data[i]);
    currentScore += data[i][1];
  }

  // グループ間の人数の差が2人以上であれば調整
  while (Math.abs(groups[0].length - groups[1].length) >= 2) {
    if (groups[0].length > groups[1].length) {
      // グループをスコアの昇順にソート
      groups[0].sort((a, b) => a[1] - b[1]);

      // 中間のスコアの人物のインデックスを取得
      var middleScoreIndex = Math.floor(groups[0].length / 2);
      // 中間のスコアの人物を取り出す
      var middleScorePerson = groups[0].splice(middleScoreIndex, 1)[0];

      // 中間のスコアの人物をグループ1に移す
      groups[1].push(middleScorePerson);
    } else {
      // グループをスコアの昇順にソート
      groups[1].sort((a, b) => a[1] - b[1]);

      // 中間のスコアの人物のインデックスを取得
      var middleScoreIndex = Math.floor(groups[1].length / 2);

      // 中間のスコアの人物を取り出す
      var middleScorePerson = groups[1].splice(middleScoreIndex, 1)[0];

      // 中間のスコアの人物をグループ0に移す
      groups[0].push(middleScorePerson);
    }
  }

  // 固定メンバーをグループに戻す
  for (var member of fixedGroupMembers) {
    var name = member[0];
    var groupIndex = fixedMembers[name];
    groups[groupIndex].push(member);
  }

  // 結果をシートに書き出す
  for (var i = 0; i < groups.length; i++) {
    // グループ0の場合は列1から、グループ1の場合は列4から書き出し開始
    var startColumn = (i == 0) ? 1 : 4;
    var startRow = 2;
    for (var j = 0; j < groups[i].length; j++) {
      // 名前をセット
      resultSheet.getRange(startRow + j, startColumn).setValue(groups[i][j][0]);
      //以下のコメントアウトを外すとレベルが表示される(デバック用)
      //resultSheet.getRange(startRow + j, startColumn + 1).setValue(groups[i][j][1]);
    }

    // 以降のセルをクリアする
    var lastRow = startRow + groups[i].length;
    var numRows = resultSheet.getLastRow() - lastRow + 1;

    // クリアする行が存在する場合のみクリアを行う
    if (numRows > 0) {
      // 以降のセルの内容をクリア
      resultSheet.getRange(lastRow, startColumn, numRows, 2).clearContent();
    }
  }
}
