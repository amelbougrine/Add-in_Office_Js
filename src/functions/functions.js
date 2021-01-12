// =RecSheets()  --> Feuilles
// =RecRows()    --> Lignes
// =RecColumns() --> Colonnes
// =RecSum(Table, Champ, Filtre) --> Somme
// =RecCount(Table, Filtre1, Filtre2) --> Compter le nombre d'enregistrement
//=RecDistinctRows()
var bank;
async function call() {
  await ShowData();
}
call();

 function ShowData() {
  var ourRequest = new XMLHttpRequest();
  ourRequest.open('GET', 'https://raw.githubusercontent.com/amelbougrine/Office-test/main/bank.json');
  ourRequest.onload = function() {
    if (ourRequest.status >= 200 && ourRequest.status < 400) {
      bank = JSON.parse(ourRequest.responseText);
      return;
    } else {
      let data = "We connected to the server, but it returned an error: " + ourRequest.status;
      bank = data;
      return;
    }
  };
  ourRequest.onerror = function() {
    let data = "Connection error";
    bank = data;
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
 * Show data by Columns
 * @customfunction
 * @param {string} table  Table 
 * @param {string} champ  Range 
 * @returns {string[][]} A dynamic array with multiple results.
 */
function Columns(table, champ) {
  var list = [];
  if (table == "Bank" || table == "bank") {
    for (let i=0; i<bank.length; i++) {
      switch (champ) {
        case "Date":
        case "date":
          var  element = Array.of(JSON.stringify(bank[i].Date));
          break;
        case "Montant":
        case "montant":
          var  element = Array.of(JSON.stringify(bank[i].Montant));
          break;
        case "N_Compte":
        case "n_compte":
          var  element = Array.of(JSON.stringify(bank[i].N_Compte));
          break;
        default:
          return "There is no " + champ + " range, Try 'Date', 'Montant' or 'N_Compte'.";
      }
      list.push(element);
    };
    return list;
  } else {
    return "There is no " + table + "table, Try 'Bank'"; 
  }
}

/**
 * Show data by Rows
 * @customfunction
 * @param {string} table  Table 
 * @param {string} champ  Range 
 * @returns {string[][]} A dynamic array with multiple results.
 */
function Rows (table, champ) {
  var list = [];
  if (table == "Bank" || table == "bank") {
    for (let i=0; i<bank.length; i++) {
      switch (champ) {
        case "Date":
        case "date":
          var  element = JSON.stringify(bank[i].Date);
          break;
        case "Montant":
        case "montant":
          var  element = JSON.stringify(bank[i].Montant);
          break;
        case "N_Compte":
        case "n_compte":
          var  element = JSON.stringify(bank[i].N_Compte);
          break;
        default:
          return "There is no " + champ + " range, Try 'Date', 'Montant' or 'N_Compte'.";
      }
      list.push(element);
    };
    return Array.of(list);
  } else {
    return "There is no " + table + "table, Try 'Bank'"; 
  }
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