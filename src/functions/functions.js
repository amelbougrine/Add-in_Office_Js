// =RecSheets()  --> Feuilles
// =RecRows()    --> Lignes
// =RecColumns() --> Colonnes
// =RecSum(Table, Champ, Filtre) --> Somme
// =RecCount(Table, Filtre1, Filtre2) --> Compter le nombre d'enregistrement
/**
 * Sum of a single range
 * @customfunction
 * @param {number[][]} singleRange  a single range
 */
function Sheets() {
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
}/**
* The sum of all of the numbers.
* @customfunction
* @param {number[][][]} operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
* @returns {number} The sum of all of the numbers.
*/
function Sum(operands) {
  let total = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
/**
 * Sum of a single range
 * @customfunction
 * @param {number[][]} singleRange  a single range
 */
function Sum(header) {
  // selectedSheet.getUsedRange().getFormat().autofitColumns();
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
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function SecondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 */
async function GetChartTitlee () {
  var context = new Excel.RequestContext();
  var sheets = context.workbook.worksheets;
  sheets.load("name");
  var array = [];
  return context.sync().then(function() {
    for (var i = 0; i < sheets.items.length; i++) {
      array.push(sheets.items[i].name);
    }
    return array;
  });
}