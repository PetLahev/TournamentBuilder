
/**
 *
 */
const BRACKET_NAME_PREFIX: string = 'BracketRange';

/**
 *
 */
class Bracket {

    namedRangeString: string;
    namedRange: any;
    bracketInColumn: number;

    /**
     * Gets name of the bracket.
     */
    name() {
        return this.namedRangeString;
    }

    /**
     * Gets the range of the named range.
     */
    range() {
        return this.namedRange;
    }

    /**
     * Gets index of the column where the bracket is placed.
     */
    columnIndex() {
        if (this.namedRange == null) return 0;
        return this.namedRange.getColumn();
    }

    /**
     * Returns index of the first row of the stored named range.
     */
    rowIndex() {
        if (this.namedRange == null) return 0;
        this.namedRange.getCell(1, 1).getRowIndex();
    }

    /**
     * Returns first cell of the bracket as a range.
     */
    topBracket() {
        if (this.namedRange == null) return null;
        return this.namedRange.getCell(1, 1);
    }

    /**
     * Returns the last cell of the bracket as a range.
     */
    bottomBracket() {
        if (this.namedRange == null) return null;
        return this.namedRange.getCell(this.getRowsDifference(), 1);
    }

    /**
     * Returns the middle cell of the bracket as a range.
     */
    middleCell() {
        if (this.namedRange == null) return null;
        let middle = Math.ceil((this.namedRange.getLastRow() - this.namedRange.getCell(1, 1).getRowIndex() + 1) / 2);
        return this.namedRange.getCell(middle, 1);
    }

    /**
     * Returns difference between the second index and first index + 1.
     * Example: 9 - 3 + 1 = 7 cells between these indices.
     */
    getRowsDifference() {
        if (this.namedRange == null) return 1;
        return this.namedRange.getLastRow() - this.namedRange.getCell(1, 1).getRowIndex() + 1;
    }

    /**
     *  Returns index of the bracket in the column.
     *  If it's first bracket in the column the value will be equal 1
     */
    bracketIndex() {
        return this.bracketInColumn;
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
    addNamedRange(spreadsheet: any, sheet: any, matchIndex: number, bracketIndex: number, cell1: any, cell2: any) {
        this.namedRangeString = `${BRACKET_NAME_PREFIX}${matchIndex}`;
        this.bracketInColumn = bracketIndex;
        let range = cell1;
        if (cell2) {
            range = sheet.getRange(cell1.getRowIndex(), cell1.getColumn(), cell2.getRowIndex() - cell1.getRowIndex() + 1);
        }
        spreadsheet.setNamedRange(this.namedRangeString, range);
        this.namedRange = spreadsheet.getRangeByName(this.namedRangeString);
    }

    /**
     * Adds existing named range to the class
     * @param {Range} range a reference to a range of named range
     * @param {number} bracketIndex index of given range (bracket) at the particular column
     */
    addNameToClass(range: any, bracketIndex: number) {
        this.namedRangeString = range.getName();
        this.bracketInColumn = bracketIndex;
        this.namedRange = range;
    }

    /**
     * Sets the border of given range or saved named range.
     * @param {Range} rng a reference to range of cells where border should be applied
     */
    setConnector(rng: any) {
        if (!rng) {
            rng = this.namedRange.offset(1, 0, this.getRowsDifference() - 1, 1);
        }
        rng.setBorder(null, null, null, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
    }

    /**
    * Sets the text in the middle of the two brackets
    * @param {string} content Te text to be inserted.
    */
    setMiddleText(content: string) {
        if (this.namedRange == null) return;
        let rng = this.middleCell();
        rng.setValue(content);
        rng.setFontWeight("bold");
        rng.setFontSize(10);
        rng.setHorizontalAlignment("center");
        rng.setVerticalAlignment("bottom");
    }

} // class