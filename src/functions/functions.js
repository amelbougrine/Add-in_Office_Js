// =RecSheets()  --> Feuilles
// =RecRows()    --> Lignes
// =RecColumns() --> Colonnes
// =RecSum(Table, Champ, Filtre) --> Somme
// =RecCount(Table, Filtre1, Filtre2) --> Compter le nombre d'enregistrement
//=RecDistinctRows()
/**
  * Gets the star count for a given Github repository.
  * @customfunction 
  */
 function ShowData() {
  getData().then((a) => {
    console.log(a);
  });
};
  // var array= []; 
  async function getData() {
    try {
      const url = "https://api.github.com/amelbougrine/Office-test";
      const response =  await fetch(url).then(
        res => res.ok ? 
          res.json().then(res => Promise.resolve(res)) 
          : res.json().then(err => Promise.reject(err)));
        const data = response; 
        return data;
    } catch (error) {
      return error;
    }
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