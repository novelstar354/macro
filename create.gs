
function createQuestionOutput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 元データ
  const unit4 = sheet.getRange("A1:H25").getValues();
  const unit5 = sheet.getRange("I1:P25").getValues();

  // 通常版と逆版を作成
  createOutput(sheet, unit4, unit5, false, 1);
  createOutput(sheet, unit4, unit5, true, 9);

  SpreadsheetApp.flush();
}


// 出力処理
function createOutput(sheet, unit4, unit5, reverse, destinationCol) {
  const result = [];

  for (let r = 0; r < 25; r++) {
    const row = [];

    for (let b = 0; b < 4; b++) {
      const number = b * 25 + r + 1;

      // 通常版：奇数unit4、偶数unit5
      // 逆版：偶数unit4、奇数unit5
      const useUnit4 = reverse
        ? number % 2 === 0
        : number % 2 === 1;

      const source = useUnit4 ? unit4 : unit5;

      const block = Math.floor((number - 1) / 25);
      const sourceRow = (number - 1) % 25;
      const col = block * 2;

      row.push(
        source[sourceRow][col],
        source[sourceRow][col + 1]
      );
    }

    result.push(row);
  }

  // 内容を書き込み
  sheet.getRange(29, destinationCol, 25, 8)
    .setValues(result);

  // 書式をコピー
  for (let r = 0; r < 25; r++) {
    for (let b = 0; b < 4; b++) {
      const number = b * 25 + r + 1;

      const useUnit4 = reverse
        ? number % 2 === 0
        : number % 2 === 1;

      const sourceCol =
        (useUnit4 ? 1 : 9) +
        Math.floor((number - 1) / 25) * 2;

      const sourceRow = (number - 1) % 25 + 1;
      const destCol = destinationCol + b * 2;

      sheet.getRange(sourceRow, sourceCol, 1, 2)
        .copyTo(
          sheet.getRange(29 + r, destCol, 1, 2),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
    }
  }

  // 行の高さ
  for (let r = 0; r < 25; r++) {
    sheet.setRowHeight(
      29 + r,
      sheet.getRowHeight(r + 1)
    );
  }
}


// メニュー
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("作成")
    .addItem("作成", "createQuestionOutput")
    .addToUi();
}

