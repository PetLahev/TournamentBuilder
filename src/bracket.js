const  NAME_PREFIX = 'BracketRange';
/**
 * Provides information about a bracket (match) in a sheet.
 */
function Bracket() {

    let namedRangeString;
    let namedRange;
    let bracketInColumn;

    /**
     * Gets name of the bracket.
     */
    this.name = function () {
        return namedRangeString;
    }

    /**
     * Gets index of the column where the bracket is placed.
     */
    this.columnIndex = function () {
        if (namedRange == null) return 0;
        return namedRange.getColumn();
    }

    /**
     * Returns first cell of the bracket as a range.
     */
    this.topBracket = function () {
        if (namedRange == null) return null;
        return namedRange.getCell(1, 1);
    }

    /**
     * Returns the last cell of the bracket as a range.
     */
    this.bottomBracket = function () {
        if (namedRange == null) return null;
        return namedRange.getCell(namedRange.getLastRow(), 1);
    }

    /**
     * Returns the middle cell of the bracket as a range.
     */
    this.middleCell = function () {
        if (namedRange == null) return null;
        var middle = Math.ceil((namedRange.getLastRow() - namedRange.getCell(1, 1).getRowIndex() + 1) / 2);
        return namedRange.getCell(middle, 1);
    }

    /**
     * Returns difference between the second index and first index + 1.
     * Example: 9 - 3 + 1 = 7 cells between these indices.
     */
    this.getRowsDifference = function () {
        if (namedRange == null) return 1;
        return namedRange.getLastRow() - namedRange.getCell(1, 1).getRowIndex() + 1;
    }

    /**
     *  Returns index of the bracket in the column.
     *  If it's first bracket in the column the value will be equal 1
     */
    this.bracketIndex = function () {
        return bracketInColumn;
    }

    /**
     * Adds a named range to given sheet and sets the internal values of the whole bracket.
     * @param {spreadsheet} spreadsheet a reference to active spreadsheet
     * @param {worksheet} sheet a reference the the active sheet where to insert named range
     * @param {number} matchIndex an index of the match in the tournament (named range will be created with this index)
     * @param {number} bracketIndex order of the bracket in a column (each column starts from 1)
     * @param {Range} cell1 a reference to the top cell of the bracket
     * @param {Range} cell2 a reference to the bottom cell of the bracket
     */
    this.addNamedRange = function (spreadsheet, sheet, matchIndex, bracketIndex, cell1, cell2) {
        namedRangeString = `${NAME_PREFIX}${matchIndex}`;
        bracketInColumn = bracketIndex;
        var range = cell1;
        if (cell2) {
            range = sheet.getRange(cell1.getRowIndex(), cell1.getColumn(), cell2.getRowIndex() - cell1.getRowIndex() + 1);
        }
        spreadsheet.setNamedRange(namedRangeString, range);
        namedRange = spreadsheet.getRangeByName(namedRangeString);
    }

    /**
     * Adds existing named range to the class
     * @param {Range} range a reference to a range of named range
     */
    this.addNameToClass = function (range) {
        namedRangeString = range.getName();
        namedRange = range;
    }

    /**
     * Sets the border of given range or saved named range.
     * @param {Range} rng a reference to range of cells where border should be applied
     */
    this.setConnector = function (rng) {
        if (!rng) {
            rng = namedRange.offset(1, 0, this.getRowsDifference() - 1, 1);
        }
        rng.setBorder(null, null, null, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
    }

    /**
    * Sets the text in the middle of the two brackets
    * @param {string} content Te text to be inserted.
    */
    this.setMiddleText = function(content) {
        if (namedRange == null) return;
        var rng = this.middleCell();
        rng.setValue(content);
        rng.setFontWeight("bold");
        rng.setFontSize(10);
        rng.setHorizontalAlignment("center");
        rng.setVerticalAlignment("bottom");
    }
}