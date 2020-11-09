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

  // if the number of players is not power of 2, this calculates how many matches needs to played
  // before it's power of 2
  var numOfPreBracketMatches = numPlayers - Math.pow(2, Math.trunc(Math.log(numPlayers) / Math.log(2)));

  var upperPower = Math.ceil(Math.log(numPlayers) / Math.log(2));

  // Find out what is the number that is a power of 2 and higher than numPlayers.
  var countNodesUpperBound = Math.pow(2, upperPower);

  // Find out what is the number that is a power of 2 and lower than numPlayers.
  var countNodesLowerBound = countNodesUpperBound / 2;

  // This is the number of nodes that will not show in the 1st level.
  var countNodesHidden = numPlayers - countNodesLowerBound;

  var startRow = 1;
  var startColumn = 1;
  if (numOfPreBracketMatches > 0) {
    var player = numPlayers;
    startRow = 4;
    for (var i = 0; i < numOfPreBracketMatches; i++) {

      // Must be on the first level
      var topBracketCell = sheetResults.getRange(i * CELLS_BETWEEN_BRACKETS + 1 + startRow, 1);
      // top node
      setBracketItem_(sheetResults, topBracketCell, `Team #${player--}`);
      // bottom node
      var bottomBracketCell = topBracketCell.offset(CELLS_INSIDE_BRACKETS - 1, 0, 1, 1);
      setBracketItem_(sheetResults, bottomBracketCell, `Team #${player--}`);

      var middleBracketCell = topBracketCell.offset(Math.ceil(CELLS_INSIDE_BRACKETS / 2) - 1, 0, 1, 1);
      setMiddleText_(middleBracketCell, matchIndex++);
      setConnector_(topBracketCell.offset(1, 0, CELLS_INSIDE_BRACKETS - 1, 1));
    }
    startRow = 1;
    startColumn = 2;
  }

  player = 1;
  var matchIndex = 1;
  var bracketInfo = new Bracket();
  for (var i = 0; i < countNodesLowerBound; i++) {
    var cellsBetweenBrackets = CELLS_BETWEEN_BRACKETS;
    var cellsInsideBracket = CELLS_INSIDE_BRACKETS;
    if (numOfPreBracketMatches > 0) {
      cellsBetweenBrackets *= 2;
      cellsInsideBracket += 2;
    }
    // Must be on the first level
    var topBracketCell = sheetResults.getRange(i * cellsBetweenBrackets + 1, startColumn);
    // top node
    setBracketItem_(sheetResults, topBracketCell, `Team #${player++}`);
    // bottom node
    var bottomBracketCell = topBracketCell.offset(cellsInsideBracket - 1, 0, 1, 1);
    setBracketItem_(sheetResults, bottomBracketCell, `Team #${player++}`);

    var middleBracketCell = topBracketCell.offset(Math.ceil(cellsInsideBracket / 2) - 1, 0, 1, 1);
    setMiddleText_(middleBracketCell, matchIndex++);
    setConnector_(topBracketCell.offset(1, 0, cellsInsideBracket - 1, 1));

    var nextBracketAddress = middleBracketCell.offset(0, 1, 1, 1);
    bracketInfo.AddRowIndex(nextBracketAddress.getRowIndex());
    setBracketItem_(sheetResults, nextBracketAddress);

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
