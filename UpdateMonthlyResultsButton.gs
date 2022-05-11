//スマホアプリ「時刻カウンタ」で記録しGoogleドライブのフォルダに保存したタイムスタンプファイルから、
//実行したシート当月分の毎時ごとの件数を集計し、シートに記入する
function updateMonthlyResults() {
  const _ = Underscore.load();
  const Month = SpreadsheetApp.getActiveSheet().getSheetName();
  const StampFiles = getTimeStampFolder(Month).getFilesByType("text/plain");
  const SheetValuesArray = convertDateToString(SpreadsheetApp.getActiveSheet().getDataRange().getValues());
  const TransposedSheetValuesArray = _.zip.apply(_, SheetValuesArray);
  const IncentiveStartCol = SheetValuesArray[TitleRow].indexOf(IncentiveStartValue);    //時間帯インセンティブ開始列の数値表現(0-index)

  generateResultsByIncentive(StampFiles, SheetValuesArray, TransposedSheetValuesArray, IncentiveStartCol)
    .updateMonthlyResultsOnSheet(SheetValuesArray, TransposedSheetValuesArray);

  SpreadsheetApp.getUi().alert("今月の稼働実績を更新しました")
}


function getTimeStampFolder(Month) {
  const Folders = DriveApp.getFolderById(TimeStampRootID).getFolders();

  while (Folders.hasNext()) {
    let childFolder = Folders.next();
    if (childFolder.getName() === Month) return childFolder
  }

  SpreadsheetApp.getUi().alert("対象月のタイムスタンプ格納フォルダが見つかりません")
}


//indexOfメソッドが文字列型しか扱えないためDate型を予め変換しておく
function convertDateToString(array) {
  return array.map( arr => { 
    const newArr = arr.map( v => {
      let type = Object.prototype.toString.call(v);
      return type === "[object Date]" ? Utilities.formatDate(v, 'JST', 'yyyy/MM/dd') : v;
    })
    return newArr
  })
}


function readStampFileToStringArray(stampFile) {
  return stampFile.getBlob().getDataAsString("utf-8").split(/\n/)
}


function getIncentiveByRow(SheetValuesArray, TitleRow, IncentiveStartCol, IncentiveSectionHours, row) {
  const IncentiveSectionArray = SheetValuesArray[TitleRow].slice(IncentiveStartCol, IncentiveStartCol + IncentiveSectionHours);
  const IncentiveArray = SheetValuesArray[row].slice(IncentiveStartCol, IncentiveStartCol + IncentiveSectionHours);
  const incentive = new Map();
  IncentiveSectionArray.map((key, i) => {incentive.set(key, IncentiveArray[i])});
  return incentive
}


function aggregateDailyResultByIncentive(dailyStamps, incentive) {
  const result = new Map();
  const dailyStampsIterator = dailyStamps[Symbol.iterator]();
  let prevousDay = "";

  while (1) {
    let iteratedStamp = dailyStampsIterator.next();
    let stamp = iteratedStamp.value;
    let isDone = (iteratedStamp.done) || (stamp === "");

    if (isDone) break
    if (stamp.split(",")[10][0] === '1') continue   //タイムスタンプの無効フラグが立っている

    let day = stamp.split(" ")[0].slice(-2);
    if (prevousDay === "") prevousDay = day
    let isSameDay = (prevousDay === day);
    isSameDay ? prevousDay = day : () => {throw new Error("タイムスタンプファイルに別日が混在しています")};
    let hour = Number(stamp.split(" ")[1].slice(0, 2));
    let inc = incentive.get(hour + "時");
    if (!inc) throw new Error(day + "日のインセンティブが入力されていません");

    result.has(inc) ? result.set(inc, result.get(inc) + 1) : result.set(inc, 1)
  }

  return result
}


function generateResultsByIncentive(StampFiles, SheetValuesArray, TransposedSheetValuesArray, IncentiveStartCol) {
  this.results = []

  while (StampFiles.hasNext()) {
      let stampFile = StampFiles.next();
      let dailyStamps = readStampFileToStringArray(stampFile);
      let day = dailyStamps[0].split(",")[3].split(" ")[0].replace(/-/g, "/");
      let row = TransposedSheetValuesArray[0].indexOf(day);
      let incentive = getIncentiveByRow(SheetValuesArray, TitleRow, IncentiveStartCol, IncentiveSectionHours, row);
      let dailyResultByIncentive = aggregateDailyResultByIncentive(dailyStamps, incentive);
      results.push([day, dailyResultByIncentive])
    }

  return this
}


function updateMonthlyResultsOnSheet(SheetValuesArray, TransposedSheetValuesArray) {
  const Sheet = SpreadsheetApp.getActiveSheet();
  const ResultsIterator = this.results[Symbol.iterator]();

  while (1) {
    let iteratedResults = ResultsIterator.next();
    if (iteratedResults.done) break
    let [day, result] = iteratedResults.value;
    let row = TransposedSheetValuesArray[0].indexOf(day);
    let resultIterator = result[Symbol.iterator]();

    while (1) {
      let iteratedResult = resultIterator.next();
      if (iteratedResult.done) break
      let [incentive, value] = iteratedResult.value;
      let col = SheetValuesArray[TitleRow].indexOf(incentive);
      Sheet.getRange(row + 1, col + 1).setValue(value)
    }
  }
}
