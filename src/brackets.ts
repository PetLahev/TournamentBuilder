
/**
 *
 */
enum BracketsType {
    Standard = 1,
    Qualification = 2,
    BothSides = 3
}

class Brackets {

    spreadsheet: any;
    sheet: any;
    teams: number;
    position: number;
    type: BracketsType;

    sheetBrackets = [];
    rounds: number = 0;
    colIndex: number  = 1;

    BRACKET_HEIGHT: number = 44;
    BRACKET_WIDTH: number = 222;
    CELLS_INSIDE_BRACKETS: number = 5; // total number of cells including top & bottom bracket cells
    CELLS_BETWEEN_BRACKETS: number = 6; // this + last bracket + 1 = next bracket

    constructor(spreadsheet: any, teams: number, startPosition: number, bracketType: BracketsType) {
        this.spreadsheet = spreadsheet;
        this.teams = teams;
        this.position = startPosition;
        this.type = bracketType;
    }

    /**
     * 
     * @param {Sheet} sheet
     */
    build(sheet: any) {

        this.sheet = sheet;
        this.rounds = Math.ceil(Math.log(this.teams) / Math.log(2));
        // if the number of players is not power of 2, this calculates how many matches
        // needs to played before it's power of 2
        var numOfPreBracketMatches = this.teams - Math.pow(2, Math.trunc(Math.log(this.teams) / Math.log(2)));

        if (this.type == BracketsType.Standard) {
            this.buildStandard_(numOfPreBracketMatches);
        }
        else if (this.type == BracketsType.Qualification) {
            this.buildQualification_(null);
        }
    }

    buildStandard_(numOfPreBracketMatches: number) {

        let bracketIndex = 1;
        let rowPosition = 1;
        let matchIndex = 1;
        // this will build the "pre-matches" before the main draw to get to a power of 2
        if (numOfPreBracketMatches > 0) {
            this.rounds -= 1;
            rowPosition = Math.floor(this.position + this.CELLS_INSIDE_BRACKETS / 2);
            for (let index = 0; index < numOfPreBracketMatches; index++) {
                var cell1 = this.sheet.getRange(rowPosition, this.colIndex);
                var cell2 = cell1.offset(this.CELLS_INSIDE_BRACKETS - 1, 0, 1, 1);
                bracketIndex = this.formatBracket(cell1, cell2, matchIndex++, bracketIndex);
                rowPosition += this.CELLS_BETWEEN_BRACKETS;
            }
            this.colIndex++;
        }

        // build the rest of the bracket till final match and bronze medal match
        var topCellPos = this.position;
        var bottomCellPos = topCellPos + this.CELLS_INSIDE_BRACKETS - 1;
        for (let round = 0; round < this.rounds; round++) {
            bracketIndex = 1; // reset the index of a bracket for each column
            let matchIndexInColumn = 3;
            if (round > 0) {
                ({ topCellPos, bottomCellPos } = this.getNextPosition_(round, this.colIndex, 1, topCellPos, bottomCellPos));
            }
            for (let match = 0; match < Math.pow(2, this.rounds - round - 1); match++) {
                var cell1 = this.sheet.getRange(topCellPos, this.colIndex);
                var cell2 = this.sheet.getRange(bottomCellPos, this.colIndex);
                bracketIndex = this.formatBracket(cell1, cell2, matchIndex++, bracketIndex);
                ({ topCellPos, bottomCellPos } = this.getNextPosition_(round, this.colIndex, matchIndexInColumn, topCellPos, bottomCellPos));
                matchIndexInColumn += 2;
            }
            this.colIndex++;
        }

        // build the golden and bronze medal matches
        // change the match index as the final is always the last match
        var finalBracket = this.sheetBrackets[this.sheetBrackets.length - 1];
        this.setBracketItem_(this.sheet, finalBracket.middleCell().offset(0, 1, 1, 1), null);
        this.spreadsheet.setNamedRange(`${BRACKET_NAME_PREFIX}${matchIndex}`, finalBracket.range());

        this.bronzeMedalMatch(cell1, cell2, matchIndex, bracketIndex);
    }

    /**
     * The brackets ends with given number qualified players.
     * If the number of players is not power of 2, others will get byes
     * @param {number} numOfQualified
     */
    buildQualification_(numOfQualified: number) {

        var bracketIndex = 1;
        var rowPosition = 1;
        var matchIndex = 1;

        this.rounds = Math.ceil(Math.log(this.teams) / Math.log(2));
        let numOfByes = Math.pow(2, this.rounds) - this.teams;
        let numOfRoundsToFinishBefore = Math.ceil(Math.log(numOfQualified) / Math.log(2));
        let numOfRoundsToPlay = this.rounds - numOfRoundsToFinishBefore;
        let firstRound = (this.teams + numOfByes) / 2;

        rowPosition = Math.floor(this.position + this.CELLS_INSIDE_BRACKETS / 2);
        for (let index = 0; index < firstRound; index++) {
            var cell1 = this.sheet.getRange(rowPosition, this.colIndex);
            var cell2 = cell1.offset(this.CELLS_INSIDE_BRACKETS - 1, 0, 1, 1);
            bracketIndex = this.formatBracket(cell1, cell2, matchIndex++, bracketIndex);
            rowPosition += this.CELLS_BETWEEN_BRACKETS;
        }
        this.colIndex++;
        this.rounds -= 1;

        // build the rest of the bracket till final match and bronze medal match
        var topCellPos = this.position;
        var bottomCellPos = topCellPos + this.CELLS_INSIDE_BRACKETS - 1;
        for (let round = 0; round < this.rounds; round++) {
            bracketIndex = 1; // reset the index of a bracket for each column
            let matchIndexInColumn = 1;
            ({ topCellPos, bottomCellPos } = this.getNextPosition_(round + 1, this.colIndex, matchIndexInColumn, topCellPos, bottomCellPos));

            for (let match = 0; match < Math.pow(2, this.rounds - round - 1); match++) {
                var cell1 = this.sheet.getRange(topCellPos, this.colIndex);
                var cell2 = this.sheet.getRange(bottomCellPos, this.colIndex);
                bracketIndex = this.formatBracket(cell1, cell2, matchIndex++, bracketIndex);
                matchIndexInColumn += 2;
                ({ topCellPos, bottomCellPos } = this.getNextPosition_(round + 1, this.colIndex, matchIndexInColumn, topCellPos, bottomCellPos));
            }
            this.colIndex++;
            // check if it should go to next round
            // don't forget, there was already one round created
            if (numOfRoundsToPlay <= (round + 2)) break;
        }

        // finish the middle cells
        for (let index = 0; index < numOfQualified; index++) {
            let bracket = this.sheetBrackets[this.sheetBrackets.length - index - 1];
            this.setBracketItem_(this.sheet, bracket.middleCell().offset(0, 1, 1, 1), null);
        }
    }

    bronzeMedalMatch(cell1: any, cell2: any, matchIndex: number, bracketIndex: number) {
        var semifinalBracket =this. sheetBrackets.find(br => br.columnIndex() == this.colIndex - 2
            && br.bracketIndex() == 2);
        var bottomBracketRowIndex = semifinalBracket.bottomBracket().getRowIndex();
        var cell1 = this.sheet.getRange(bottomBracketRowIndex + 2, this.colIndex - 1);
        var cell2 = this.sheet.getRange(cell1.getRowIndex() + this.CELLS_INSIDE_BRACKETS - 1, this.colIndex - 1);
        this.formatBracket(cell1, cell2, matchIndex - 1, bracketIndex);
        var bronzeBracket = this.sheetBrackets[this.sheetBrackets.length - 1];
        this.setBracketItem_(this.sheet, bronzeBracket.middleCell().offset(0, 1, 1, 1), null);
        return { cell1, cell2 };
    }

    getNextPosition_(round: number, columnIndex: number, match: number, topCellPos: number, bottomCellPos: number) {
        if (round == 0) {
            topCellPos += this.CELLS_BETWEEN_BRACKETS;
            bottomCellPos = topCellPos + this.CELLS_INSIDE_BRACKETS - 1;
        }
        else {
            // need to get the previous column matches based on the value at the 'match'
            let firstBracket = this.sheetBrackets.find(br => br.columnIndex() == columnIndex - 1
                && br.bracketIndex() == match);
            let secondBracket = this.sheetBrackets.find(br => br.columnIndex() == columnIndex - 1
                && br.bracketIndex() == match + 1);
            if (firstBracket) topCellPos = firstBracket.middleCell().getRowIndex();
            if (secondBracket) bottomCellPos = secondBracket.middleCell().getRowIndex();
        }
        return { topCellPos, bottomCellPos };
    }

    /**
     * Collects all brackets from given sheet
     * @param {Sheet} sheet a reference to a sheet where to collect brackets
     */
    collectBrackets(sheet: any) {
        var ranges = sheet.getNamedRanges();
        let names = [];
        ranges.forEach(element => {
            if (element.getName().startsWith(BRACKET_NAME_PREFIX)) {
                names.push(element);
            }
        });

        if (names.length == 0) return;
        names.sort(function (name1, name2) {
            return this.sortByColumnThenByRow_(name1.getRange(), name2.getRange());
        });

        let colIndex = names[0].getRange().getColumn();
        let bracketIndex = 0;
        names.forEach(element => {
            if (colIndex == element.getRange().getColumn()) {
                bracketIndex++;
            }
            else {
                colIndex = element.getRange().getColumn();
                bracketIndex = 1;
            }
            let bracket = new Bracket();
            this.sheetBrackets.push(bracket.addNameToClass(element, bracketIndex));
        });
    }

    /**
     * Removes all data and formatting from given sheet and also
     * removes all named ranges that belong to the brackets.
     * @param {Sheet} sheet a reference to a sheet where to remove data
     */
    removeBrackets(sheet: any) {
        var ranges = sheet.getNamedRanges();
        ranges.forEach(element => {
            if (element.getName().startsWith(BRACKET_NAME_PREFIX)) {
                element.remove();
            }
        });
        sheet.clear();
    }


    formatBracket(cell1: any, cell2: any, matchIndex: number, bracketIndex: number) {
        this.setBracketItem_(this.sheet, cell1, null);
        this. setBracketItem_(this.sheet, cell2, null);
        var bracket = new Bracket();
        bracket.addNamedRange(this.spreadsheet, this.sheet, matchIndex, bracketIndex++, cell1, cell2);
        bracket.setConnector(null);
        this.sheetBrackets.push(bracket);
        return bracketIndex;
    }

    setBracketItem_(sheet: any, rng: any, content: string) {
        if (content) {
            if (content[0] == '=') {
                rng.setFormula(content);
            }
            else {
                rng.setValue(content);
            }
        }
        rng.setBorder(null, null, true, null, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
        sheet.setRowHeight(rng.getRowIndex(), this.BRACKET_HEIGHT)
        sheet.setColumnWidth(rng.getColumnIndex(), this.BRACKET_WIDTH);
    }

    sortByColumnThenByRow_(rng1: any, rng2: any) {

        if (rng1.getColumn() > rng2.getColumn()) {
            return 1;
        } else if (rng1.getColumn() < rng2.getColumn()) {
            return -1;
        }

        // Else go to the 2nd item
        if (rng1.getRowIndex() < rng2.getRowIndex()) {
            return -1;
        } else if (rng1.getRowIndex() > rng2.getRowIndex()) {
            return 1
        } else { // nothing to split them
            return 0;
        }
    }
} // class