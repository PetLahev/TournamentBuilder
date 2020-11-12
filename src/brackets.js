/**
 * 
 * @param {Spreadsheet} spreadsheet 
 * @param {number} teams 
 * @param {number} startPosition 
 * @param {BracketsTypeEnum} bracketType 
 */
function Brackets(spreadsheet, teams, startPosition, bracketType) {

    var self = this;
    self.spreadsheet = spreadsheet;
    self.teams = teams;
    self.position = startPosition;
    self.type = bracketType;

    // private properties
    let sheetBrackets = [];
    let rounds = 0;
    let colIndex = 1;

    /**
     * 
     * @param {*} sheet
     */
    this.build = function (sheet) {

        self.sheet = sheet;
        rounds = Math.ceil(Math.log(teams) / Math.log(2));
        // if the number of players is not power of 2, this calculates how many matches needs to played
        // before it's power of 2
        var numOfPreBracketMatches = teams - Math.pow(2, Math.trunc(Math.log(teams) / Math.log(2)));

        if (self.type == BracketsTypeEnum.standard) {
            buildStandard_(numOfPreBracketMatches);
        }
    }

    function buildStandard_(numOfPreBracketMatches) {

        var bracketIndex = 1;
        var rowPosition = 1;
        var matchIndex = 1;
        // this will build the "pre-matches" before the main draw to get to a power of 2
        if (numOfPreBracketMatches > 0) {
            rounds -= 1;
            rowPosition = Math.floor(self.position + CELLS_INSIDE_BRACKETS / 2);
            for (let index = 0; index < numOfPreBracketMatches; index++) {
                var cell1 = self.sheet.getRange(rowPosition, colIndex);
                var cell2 = cell1.offset(CELLS_INSIDE_BRACKETS - 1, 0, 1, 1);
                bracketIndex = formatBracket(cell1, cell2, matchIndex++, bracketIndex);
                rowPosition += CELLS_BETWEEN_BRACKETS;
            }
            colIndex++;
        }

        // build the rest of the bracket till final match and bronze medal match
        var topCellPos = self.position;
        var bottomCellPos = topCellPos + CELLS_INSIDE_BRACKETS - 1;
        for (let round = 0; round < rounds; round++) {
            bracketIndex = 1; // reset the index of a bracket for each column
            let matchIndexInColumn = 3;
            if (round > 0) {
                ({ topCellPos, bottomCellPos } = getNextPosition_(round, colIndex, 1, topCellPos, bottomCellPos));
            }
            for (let match = 0; match < Math.pow(2, rounds - round - 1); match++) {
                var cell1 = self.sheet.getRange(topCellPos, colIndex);
                var cell2 = self.sheet.getRange(bottomCellPos, colIndex);
                bracketIndex = formatBracket(cell1, cell2, matchIndex++, bracketIndex);
                ({ topCellPos, bottomCellPos } = getNextPosition_(round, colIndex, matchIndexInColumn, topCellPos, bottomCellPos));
                matchIndexInColumn += 2;
            }
            colIndex++;
        }

        // build the golden and bronze medal matches
        // change the match index as the final is always the last match
        var finalBracket = sheetBrackets[sheetBrackets.length - 1];
        setBracketItem_(self.sheet, finalBracket.middleCell().offset(0, 1, 1, 1));
        self.spreadsheet.setNamedRange(`${NAME_PREFIX}${matchIndex}`, finalBracket.range());

        bronzeMedalMatch(cell1, cell2, matchIndex, bracketIndex);
    }

    function bronzeMedalMatch(cell1, cell2, matchIndex, bracketIndex) {
        var semifinalBracket = sheetBrackets.find(br => br.columnIndex() == colIndex - 2
            && br.bracketIndex() == 2);
        var bottomBracketRowIndex = semifinalBracket.bottomBracket().getRowIndex();
        var cell1 = self.sheet.getRange(bottomBracketRowIndex + 2, colIndex - 1);
        var cell2 = self.sheet.getRange(cell1.getRowIndex() + CELLS_INSIDE_BRACKETS - 1, colIndex - 1);
        formatBracket(cell1, cell2, matchIndex - 1, bracketIndex);
        var bronzeBracket = sheetBrackets[sheetBrackets.length - 1];
        setBracketItem_(self.sheet, bronzeBracket.middleCell().offset(0, 1, 1, 1));
        return { cell1, cell2 };
    }

    function getNextPosition_(round, columnIndex, match, topCellPos, bottomCellPos) {
        if (round == 0) {
            topCellPos += CELLS_BETWEEN_BRACKETS;
            bottomCellPos = topCellPos + CELLS_INSIDE_BRACKETS - 1;
        }
        else {
            // need to get the previous column matches based on the value at the 'match'
            var firstBracket = sheetBrackets.find(br => br.columnIndex() == columnIndex - 1
                && br.bracketIndex() == match);
            var secondBracket = sheetBrackets.find(br => br.columnIndex() == columnIndex - 1
                && br.bracketIndex() == match);
            if (firstBracket) topCellPos = firstBracket.middleCell().getRowIndex();
            if (secondBracket) bottomCellPos = secondBracket.middleCell().getRowIndex();
        }
        return { topCellPos, bottomCellPos };
    }

    /**
     * Collects all brackets from given sheet
     * @param {Sheet} sheet a reference to a sheet where to collect brackets
     */
    this.collectBrackets = function (sheet) {
        var ranges = sheet.getNamedRanges();
        ranges.forEach(element => {
            if (element.getName().startsWith(NAME_PREFIX)) {
                let bracket = new Bracket();
                // TODO: Need to add index of the bracket in the column
                sheetBrackets.push(bracket.addNameToClass(element));
            }
        });
    }

    /**
     * Removes all data and formatting from given sheet and also
     * removes all named ranges that belong to the brackets.
     * @param {Sheet} sheet a reference to a sheet where to remove data
     */
    this.removeBrackets = function (sheet) {
        var ranges = sheet.getNamedRanges();
        ranges.forEach(element => {
            if (element.getName().startsWith(NAME_PREFIX)) {
                element.remove();
            }
        });
        sheet.clear();
    }

    function formatBracket(cell1, cell2, matchIndex, bracketIndex) {
        setBracketItem_(self.sheet, cell1);
        setBracketItem_(self.sheet, cell2);
        var bracket = new Bracket();
        bracket.addNamedRange(self.spreadsheet, self.sheet, matchIndex, bracketIndex++, cell1, cell2);
        bracket.setConnector();
        sheetBrackets.push(bracket);
        return bracketIndex;
    }

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
}