function distributePeopleAndScores() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sheet = ss.getSheetByName('levelSheet');
	var resultSheet = ss.getSheetByName('resultSheet');
	var data = sheet.getRange("A2:B80").getValues();

	// 空の行または'欠席'と記載されている行を除外
	data = data.filter(function (row) {
		return row[0] !== "" && row[1] !== "" && row[0] !== "欠席" && row[1] !== "欠席";
	});

	// aaaさんは第一インスタンス固定、bbbは第二固定
	var fixedMembers = { "aaa": 0, "bbb": 1 };
	var groups = [[], []];

	// 特定のメンバーを固定したグループに分ける。それ以外のメンバーを新たなdata配列に戻す。
	data = data.filter(function (person) {
		var name = person[0];
		if (name in fixedMembers) {
			groups[fixedMembers[name]].push(person);
			return false;
		}
		return true;
	});

	// 固定メンバーをグループから削除
	var fixedGroupMembers = [];
	for (var name in fixedMembers) {
		var groupIndex = fixedMembers[name];
		var memberIndex = groups[groupIndex].findIndex((person) => person[0] === name);
		fixedGroupMembers.push(groups[groupIndex].splice(memberIndex, 1)[0]);
	}

	// ランダムに並べる
	data.sort(function () {
		return 0.5 - Math.random();
	});

	// 点数の総数を計算
	var totalScore = data.reduce(function (sum, person) {
		return sum + person[1];
	}, 0);

	var halfScore = totalScore / 2;

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
			groups[0].sort((a, b) => a[1] - b[1]);

			var middleScoreIndex = Math.floor(groups[0].length / 2);

			var middleScorePerson = groups[0].splice(middleScoreIndex, 1)[0];

			groups[1].push(middleScorePerson);
		} else {
			groups[1].sort((a, b) => a[1] - b[1]);

			var middleScoreIndex = Math.floor(groups[1].length / 2);

			var middleScorePerson = groups[1].splice(middleScoreIndex, 1)[0];

			groups[0].push(middleScorePerson);
		}
	}

	// 固定メンバーをグループに戻す
	for (var member of fixedGroupMembers) {
		var name = member[0];
		var groupIndex = fixedMembers[name];
		groups[groupIndex].push(member);
	}

	// Doneだyo!
	for (var i = 0; i < groups.length; i++) {
		var startColumn = (i == 0) ? 1 : 4;
		var startRow = 2;
		for (var j = 0; j < groups[i].length; j++) {
			resultSheet.getRange(startRow + j, startColumn).setValue(groups[i][j][0]);
			//以下のコメントアウトを外すとレベルが表示される(デバック用)
			resultSheet.getRange(startRow + j, startColumn + 1).setValue(groups[i][j][1]);
		}

		// 以降のセルをクリアする
		var lastRow = startRow + groups[i].length;
		var numRows = resultSheet.getLastRow() - lastRow + 1;

		// クリアする行が存在する場合のみクリアを行う
		if (numRows > 0) {
			resultSheet.getRange(lastRow, startColumn, numRows, 2).clearContent();
		}
	}
}

