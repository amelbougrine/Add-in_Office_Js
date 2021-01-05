// =RecSheets()  --> Feuilles
// =RecRows()    --> Lignes
// =RecColumns() --> Colonnes
// =RecSum(Table, Champ, Filtre) --> Somme
// =RecCount(Table, Filtre1, Filtre2) --> Compter le nombre d'enregistrement
//=RecDistinctRows()
var data1;
async function call() {
  await ShowData();
}
call();
/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [[data1.Date], [data1.Montant], ['third']];
}
/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [[data1.Date, data1.Montant, 'third']];
}
// /**
//   * Gets the star count for a given Github repository.
//   * @customfunction 
//   * @return {string} Data
//   */
 function ShowData() {
  var ourRequest = new XMLHttpRequest();
  ourRequest.open('GET', 'https://raw.githubusercontent.com/amelbougrine/Office-test/main/bank.json');
  ourRequest.onload = function() {
    if (ourRequest.status >= 200 && ourRequest.status < 400) {
      var ourData = JSON.parse(ourRequest.responseText);
      data1 = ourData[0];
      return;
    } else {
      let data = "We connected to the server, but it returned an error: " + ourRequest.status;
      data1 = data;
      return;
    }
  };
  ourRequest.onerror = function() {
    let data = "Connection error";
    data1 = data;
    return;
  };
  ourRequest.send();
}

/**
 * Select Worksheets
 * @customfunction
 * @param {string} Feuilles  Worksheet
 */
function Sheets(Feuilles) {
}
/**
 * Select Rows
 * @customfunction
 * @param {number[][]} Lignes Rows
 */
function Rows(Lignes) {
}
/**
 * Select Columns
 * @customfunction
 * @param {number[][]} Colonnes Columns
 */
function Columns(Colonnes) {
}
/**
 * Sum of a selected Range 
 * @customfunction
 * @param {string} Table  Worksheet
 * @param {number[][]} Champ  Range
 * @param {string} Filter  Cell
 */
function Sum(Table, Champ, Filtre) {
  // selectedSheet.getUsedRange().getFormat().autofitColumns();
  let selectedSheet = context.workbook.worksheets.getItem("Sheet1");
  let worksheet = worksheet
  let singleRange = header;
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
/**
 * Count the number of records 
 * @customfunction
 * @param {string} Table  Worksheet
 * @param {string} Filter1  Cell
 * @param {string} Filtre2  Cell
 */
function Count(Table, Filtre1, Filtre2) {
}