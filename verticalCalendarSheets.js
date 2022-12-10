const monthText = new Map([
  [0, "Enero"],
  [1, "Febrero"],
  [2, "Marzo"],
  [3, "Abril"],
  [4, "Mayo"],
  [5, "Junio"],
  [6, "Julio"],
  [7, "Agosto"],
  [8, "Septiembre"],
  [9, "Octubre"],
  [10, "Noviembre"],
  [11, "Diciembre"],
])

const dayText = new Map([
  [0, "Lunes"],
  [1, "Martes"],
  [2, "Miércoles"],
  [3, "Jueves"],
  [4, "Viernes"],
  [5, "Sábado"],
  [6, "Domingo"],
])


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
    .addItem('iniciar', 'init')
    

    .addToUi();
}



function init() {

  let yearInit = 2019
  let yearFinally = 2026
  let deltaYear = (yearFinally - yearInit)

  const iniRow = 7;
  let row = iniRow;
  let initYearRow = row

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  //sheet.autoResizeColumn(4);



  const date = new Date(yearInit, 0, 1);

  for (let k = 1; k <= deltaYear; k++) {

    row = showForYear(sheet, date, row);
    //Logger.log(initYearRow + " : " + (row-initYearRow) )
    mergeCell(sheet, [initYearRow, 1, (row - initYearRow), 1]);
    createGroups(sheet, [initYearRow, 1, (row - initYearRow)-1, 1], 1);
    initYearRow = row

  }

  sheet.getRange(1, 1, 2, 3).setValues([["Rows", "Cols", "init"], [row, 6, iniRow]]);

  sheet.getRange(iniRow - 1, 1, 1, 6).setValues([["Year", "nYear", "Month", "nMonth", "Day", "Date"]]);



}

function showForYear(sheet, date, row) {
  let initRow = row;
  for (let i = 1; i <= 12; i++) {
    row = showForMonth(sheet, date, row);
    //Logger.log(initRow + " : " + (row-initRow) )

    mergeCell(sheet, [initRow, 3, (row - initRow), 1]);
    mergeCell(sheet, [initRow, 2, (row - initRow), 1]);
    createGroups(sheet, [initRow, 2, (row - initRow)-1, 1], 1);
    initRow = row;

  }
  return row;
}


function showForMonth(sheet, date, row) {

  const maxDays = 32
  let initMonth = date.getMonth();
  let nWeekMonth = 1;

  let countWeek = 0;

  for (let j = 1; j <= maxDays; j++) {

    cells = sheet.getRange(row, 1, 1, 6);
    dateCurrent = date.getDate();
    dayCurrent = date.getDay();

    if (!isNextMonth(initMonth, date)) {
      countWeek++;
      showValues(
        cells,
        date.getFullYear(),
        date.getMonth(),
        nWeekMonth,
        date.getDay(),
        date.getDate());

      if (isNextWeek(dayCurrent + 1)) {
        row++;
        showIndexGrps(
          sheet.getRange(row, 1, 1, 6),
          date.getFullYear(),
          date.getMonth(),
          nWeekMonth)


        mergeCell(sheet, [(row - countWeek), 4, countWeek + 1, 1]);
        createGroups(sheet, [(row - countWeek), 4, countWeek , 1],1);

        nWeekMonth++;
        countWeek = 0;
      }

    } else {
      //Logger.log((row-countWeek)+":"+countWeek + "->"+ nWeekMonth)
      if (countWeek != 0) {

        mergeCell(sheet, [(row - countWeek), 4, countWeek + 1, 1]);
        createGroups(sheet, [(row - countWeek), 4, countWeek, 1], 1)
        row++;
        showIndexGrps(
          sheet.getRange(row, 1, 1, 6),
          date.getFullYear(),
          date.getMonth(),
          nWeekMonth)
      } else {
        showIndexGrps(
          sheet.getRange(row, 1, 1, 6),
          date.getFullYear(),
          date.getMonth(),
          nWeekMonth)
      }

      nWeekMonth = 1;
      break;

    }


    // New day----
    date.setDate(date.getDate() + 1);
    row++;
    //--------

  }

  return row;
}


function isNextMonth(currentMonth, date) {
  return currentMonth != date.getMonth()
}

function isNextWeek(dayCurrent) {
  return dayCurrent % 7 == 0 && dayCurrent != 0;
}

function createGroups(sheet, arr, index=-1){
  sheet.getRange(arr[0], arr[1], arr[2], arr[3]).shiftRowGroupDepth(index);

}


function showIndexGrps(cells, yearCurrent, monthCurrent, countWeekForMonth) {
  const currentValues = [[yearCurrent, (monthCurrent + 1), monthNumberToText(monthCurrent), countWeekForMonth, undefined, undefined]];
  cells.setValues(currentValues);
}

function mergeCell(sheet, arr) {
  sheet.getRange(arr[0], arr[1], arr[2], arr[3]).mergeVertically();
}

function showValues(cells, yearCurrent, monthCurrent, countWeekForMonth, dayCurrent, dateCurrent) {
  const currentValues = [[yearCurrent, (monthCurrent + 1), monthNumberToText(monthCurrent), countWeekForMonth, dayToText(dayCurrent), dateCurrent]];
  cells.setValues(currentValues);
}


function monthNumberToText(n) {
  return monthText.get(n);
}
function dayToText(n) {
  return dayText.get(n);
}




















