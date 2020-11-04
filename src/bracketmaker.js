var SHEET_BRACKET = 'Bracket';
var CELLS_INSIDE_BRACKETS = 5; // total number of cells including top & bottom bracket cells
var CELLS_BETWEEN_BRACKETS = 6; // this + last bracket + 1 = next bracket
var BRACKET_HEIGHT = 44;
var BRACKET_WIDTH = 222;

function testBracket() {

  var players = Browser.inputBox('Type number of players/teams');
  try {
    players = parseInt(players);
    createStandardBracket(players);
  } catch (error) {
    Browser.msgBox(`Try again.${error}`);
  }
}

/**
 * This method creates standard bracket.
 * If number of players is not a power 2 some players will get 'bye'.
 */
function createStandardBracket(numPlayers) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetResults = ss.getSheetByName(SHEET_BRACKET);
  sheetResults.setHiddenGridlines(true);

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
  sheetResults.setRowHeights(1, 50, BRACKET_HEIGHT);

  var upperPower = Math.ceil(Math.log(numPlayers) / Math.log(2));

  // Find out what is the number that is a power of 2 and lower than numPlayers.
  var countNodesUpperBound = Math.pow(2, upperPower);

  // Find out what is the number that is a power of 2 and higher than numPlayers.
  var countNodesLowerBound = countNodesUpperBound / 2;

  // This is the number of nodes that will not show in the 1st level.
  var countNodesHidden = numPlayers - countNodesLowerBound;

  var player = 1;
  var matchIndex = 1;
  var bracketInfo = new Bracket();
  for (var i = 0; i < countNodesLowerBound; i++) {
    if (i < countNodesHidden) {
      // Must be on the first level
      var topBracketCell = sheetResults.getRange(i * CELLS_BETWEEN_BRACKETS + 1, 1);
      // top node
      setBracketItem_(sheetResults, topBracketCell, `Team #${player++}`);
      // bottom node
      var bottomBracketCell = topBracketCell.offset(CELLS_INSIDE_BRACKETS - 1, 0, 1, 1);
      setBracketItem_(sheetResults, bottomBracketCell, `Team #${player++}`);

      var middleBracketCell = topBracketCell.offset(Math.ceil(CELLS_INSIDE_BRACKETS / 2) - 1, 0, 1, 1);
      setMiddleText_(middleBracketCell, matchIndex++);
      setConnector_(topBracketCell.offset(1, 0, CELLS_INSIDE_BRACKETS - 1, 1));

      var nextBracketAddress = middleBracketCell.offset(0, 1, 1, 1);
      bracketInfo.AddRowIndex(nextBracketAddress.getRowIndex());
      setBracketItem_(sheetResults, nextBracketAddress);
    } else {
      // This player gets a bye
      setBracketItem_(sheetResults.getRange(i * CELLS_BETWEEN_BRACKETS + 1, 1), players);
    }
  }

  // Now fill in the rest of the bracket
  var bronzeMedalTopCell;
  var middleBracketCell;
  upperPower--;
  for (var columnIndex = 0; columnIndex < upperPower; columnIndex++) {
    var numOfBrackets = bracketInfo.Count(); // keep it this way!
    // processing both rows at this bracket therefore 'rowIndex +=2'
    for (var rowIndex = 0; rowIndex < numOfBrackets; rowIndex += 2) {
      // calculating middle cell of the two brackets
      var middleCellIndex = Math.ceil(bracketInfo.GetRowsDifference() / 2);
      var nextRowIndex = bracketInfo.TopIndex() + middleCellIndex - 1;
      // the bracket in the first two columns are already created therefore 'columnIndex + 3'
      setBracketItem_(sheetResults, sheetResults.getRange(nextRowIndex, columnIndex + 3));

      middleBracketCell = sheetResults.getRange(nextRowIndex, columnIndex + 2);
      setMiddleText_(middleBracketCell, matchIndex++);

      // the connector must start one row below top bracket and end one row above the bottom bracket
      // also we use right aligned border therefore 'columnIndex+2' if we used left aligned then it would be 'columnIndex+3'
      var rngConnector = sheetResults.getRange(bracketInfo.TopIndex(), columnIndex + 2, bracketInfo.GetRowsDifference());
      setConnector_(rngConnector.offset(1, 0, bracketInfo.GetRowsDifference() - 1, 1));

      if (columnIndex == (upperPower - 1)) {
        bronzeMedalTopCell = middleBracketCell.offset((bracketInfo.GetRowsDifference() / 2) + 5, 0, 1, 1);
      }

      // Processed the top two indices so it can be removed. Also adding row index of the next column for next loop.
      bracketInfo.RemoveIndex(2);
      bracketInfo.AddRowIndex(nextRowIndex);

    } // rowIndex
  } // columnIndex

  // finish gold medal match
  middleBracketCell.setValue(matchIndex);
  var winnerCell = middleBracketCell.offset(1, 1, 1, 1);
  setMiddleText_(winnerCell, 'WINNER');
  winnerCell.setBackground('yellow');

  // add bronze medal match
  setBracketItem_(sheetResults, bronzeMedalTopCell);
  var bottomBracketCell = bronzeMedalTopCell.offset(CELLS_INSIDE_BRACKETS - 1, 0, 1, 1);
  setBracketItem_(sheetResults, bottomBracketCell);
  var middleBracketCell = bronzeMedalTopCell.offset(Math.ceil(CELLS_INSIDE_BRACKETS / 2) - 1, 0, 1, 1);
  setMiddleText_(middleBracketCell, matchIndex - 1);
  setConnector_(bronzeMedalTopCell.offset(1, 0, CELLS_INSIDE_BRACKETS - 1, 1));
  setBracketItem_(sheetResults, middleBracketCell.offset(0, 1, 1, 1));
  setMiddleText_(middleBracketCell.offset(1, 1, 1, 1), 'BRONZE');
  middleBracketCell.offset(1, 1, 1, 1).setBackground('orange');
}

/**
 * Sets the value of an item in the bracket and the formatting.
 * @param {Sheet} sheet The spreadsheet to setup.
 * @param {Range} rng The Spreadsheet Range.
 * @param {string} content A string or formula to add.
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
  rng.setBorder(null, null, true, null, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(rng.getRowIndex(), BRACKET_HEIGHT)
  sheet.setColumnWidth(rng.getColumnIndex(), BRACKET_WIDTH);
}

/**
 * Sets the text in the middle of the two brackets
 * @param {Range} rng The spreadsheet range.
 * @param {string} content Te text to be inserted.
 */
function setMiddleText_(rng, content) {
  rng.setValue(content);
  rng.setFontWeight("bold");
  rng.setFontSize(10);
  rng.setHorizontalAlignment("center");
  rng.setVerticalAlignment("bottom");
}

/**
 * Sets the color and width for connector cells.
 * @param {Range} rng The spreadsheet range.
 */
function setConnector_(rng) {
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
    if (rowIndices.length <= 1) return 1;
    return this.BottomIndex() - this.TopIndex() + 1;
  }

  /**
   * Returns count of the indices stored in array.
   */
  this.Count = function () {
    return rowIndices.length;
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
    return rowIndices[0];
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
    return rowIndices[1];
  }

  /**
   * Removes given number of indices from the array.
   * @param {*} numOfIndices Num of indices to be removed
   */
  this.RemoveIndex = function (numOfIndices) {
    if (!numOfIndices) return;
    if (rowIndices.length < numOfIndices) return;
    for (var i = 0; i < numOfIndices; i++) {
      rowIndices.shift();
    }
  }
}