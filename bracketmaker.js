// This script works with the Brackets Test spreadsheet to create a tournament bracket
// given a list of players or teams.

var RANGE_PLAYER1 = 'FirstPlayer';
var SHEET_PLAYERS = 'Players';
var SHEET_BRACKET = 'Bracket';
var CONNECTOR_WIDTH = 15;
var BRACKET_HEIGHT = 44;
var BRACKET_WIDTH = 222;

/**
 * This method creates the brackets based on the data provided on the players
 */
function createBracket() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rangePlayers = ss.getRangeByName(RANGE_PLAYER1);
  var sheetControl = ss.getSheetByName(SHEET_PLAYERS);
  var sheetResults = ss.getSheetByName(SHEET_BRACKET);

  // Get the players from column A.  We assume the entire column is filled here.
  rangePlayers = rangePlayers.offset(0, 0, sheetControl.getMaxRows() -
    rangePlayers.getRowIndex() + 1, 1);
  var players = rangePlayers.getValues();

  // Now figure out how many players there are(ie don't count the empty cells)
  var numPlayers = 0;
  for (var i = 0; i < players.length; i++) {
    if (!players[i][0] || players[i][0].length == 0) {
      break;
    }
    numPlayers++;
  }
  players = players.slice(0, numPlayers);

  // Provide some error checking in case there are too many or too few players/teams.
  if (numPlayers > 64) {
    Browser.msgBox('Sorry, this script can only create brackets for 64 or fewer players.');
    return; // Early exit
  }

  if (numPlayers < 3) {
    Browser.msgBox('Sorry, you must have at least 3 players.');
    return; // Early exit
  }

  // First clear the results sheet and all formatting
  sheetResults.clear();
  sheetResults.setRowHeights(1, 50, 21);

  var upperPower = Math.ceil(Math.log(numPlayers) / Math.log(2));

  // Find out what is the number that is a power of 2 and lower than numPlayers.
  var countNodesUpperBound = Math.pow(2, upperPower);

  // Find out what is the number that is a power of 2 and higher than numPlayers.
  var countNodesLowerBound = countNodesUpperBound / 2;

  // This is the number of nodes that will not show in the 1st level.
  var countNodesHidden = numPlayers - countNodesLowerBound;

  var bracketInfo = new Bracket();
  for (var i = 0; i < countNodesLowerBound; i++) {
    if (i < countNodesHidden) {
      // Must be on the first level
      var rng = sheetResults.getRange(i * 6 + 1, 1);
      setBracketItem_(sheetResults, rng, players);
      setBracketItem_(sheetResults, rng.offset(4, 0, 1, 1), players);
      setConnector_(sheetResults, rng.offset(1, 0, 4, 1));
      var nextBracketAddress = rng.offset(2, 1, 1, 1);
      bracketInfo.AddRowIndex(nextBracketAddress.getRowIndex());
      setBracketItem_(sheetResults, nextBracketAddress);
    } else {
      // This player gets a bye
      setBracketItem_(sheetResults.getRange(i * 6 + 2, 3), players);
    }
  }

  // Now fill in the rest of the bracket
  upperPower--;
  for (var columnIndex = 0; columnIndex < upperPower; columnIndex++) {
    var numOfBrackets = bracketInfo.Count(); // keep it this way!
    // processing both rows at this bracket therefore 'rowIndex +=2'
    for (var rowIndex = 0; rowIndex < numOfBrackets; rowIndex += 2) {
      // calculating middle cell of the two brackets
      var middleCell = Math.ceil(bracketInfo.GetRowsDifference() / 2);
      var nextRowIndex = bracketInfo.TopIndex() + middleCell - 1;
      // the bracket in the first two columns are already created therefore 'columnIndex + 3'
      setBracketItem_(sheetResults, sheetResults.getRange(nextRowIndex, columnIndex + 3));

      // the connector must start one row below top bracket and end one row above the bottom bracket
      // also we use right aligned border therefore 'columnIndex+2' if we used left aligned then it would be 'columnIndex+3'
      var rngConnector = sheetResults.getRange(bracketInfo.TopIndex(), columnIndex + 2, bracketInfo.GetRowsDifference());
      setConnector_(sheetResults, rngConnector.offset(1, 0, bracketInfo.GetRowsDifference() - 1, 1));

      // Processed the top two indices so it can be removed. Also adding row index of the next column for next loop.
      bracketInfo.RemoveIndex(2);
      bracketInfo.AddRowIndex(nextRowIndex);
    }
  }
}

/**
 * Sets the value of an item in the bracket and the formatting.
 * @param {Sheet} sheet The spreadsheet to setup.
 * @param {Range} rng The Spreadsheet Range.
 * @param {string[]} content A string or formula to add.
 */
function setBracketItem_(sheet, rng, content) {
  if (content) {
    if (content[0] == '=') {
      rng.setFormula(content);
    }
    else {
      rng.setValue(content);
    }
  }
  rng.setBorder(null, null, true, null, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(rng.getRowIndex(), BRACKET_HEIGHT)
  sheet.setColumnWidth(rng.getColumnIndex(), BRACKET_WIDTH);
}

/**
 * Sets the color and width for connector cells.
 * @param {Sheet} sheet The spreadsheet to setup.
 * @param {Range} rng The spreadsheet range.
 */
function setConnector_(sheet, rng) {
  rng.setBorder(null, null, null, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * A helper class that encapsules some logic operations for calculating position of a bracket.
 */
function Bracket() {
  // stores row indices where a name of a team/player should be places
  let rowIndices = [];

  /**
   * Adds given row index to internal array of row indices
   * @param {*} rowIndex a row index to be added
   */
  this.AddRowIndex = function (rowIndex) {
    rowIndices.push(rowIndex)
  }

  /**
   * Returns difference between the second index and first index + 1.
   * Example: 9 - 3 + 1 = 7 cells between these indices.
   */
  this.GetRowsDifference = function () {
    if (addresses.length <= 1) return 1;
    return this.BottomIndex() - this.TopIndex() + 1;
  }

  /**
   * Returns count of the indices stored in array.
   */
  this.Count = function () {
    return addresses.length;
  }

  /**
   * Returns the first index from the array that represents the upper
   * position of a bracket in a sheet.
   *   1   2   3   4
   * A ___ - this is the top index[0]
   * B
   * C
   * D ___ - this is the bottom index[1]
   */
  this.TopIndex = function () {
    return addresses[0];
  }

  /**
   * Returns the second index from the array that represents the bottom
   * position of a bracket in a sheet.
   *   1   2   3   4
   * A ___ - this is the top index[0]
   * B
   * C
   * D ___ - this is the bottom index[1]
   */
  this.BottomIndex = function () {
    return addresses[1];
  }

  /**
   * Removes given number of indices from the array.
   * @param {*} numOfIndices Num of indices to be removed
   */
  this.RemoveIndex = function (numOfIndices) {
    if (!numOfIndices) return;
    if (addresses.length < numOfIndices) return;
    for (var i = 0; i < numOfIndices; i++) {
      addresses.shift();
    }
  }
}