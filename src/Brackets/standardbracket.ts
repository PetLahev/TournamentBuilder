
class StandardBracket extends Brackets {

    constructor(spreadsheet: any, teams: number, startPosition: number) {
        super(spreadsheet, teams, startPosition, BracketsType.Standard);
    }

    build(sheet: any) {

        this.sheet = sheet;
        this.rounds = Math.ceil(Math.log(this.teams) / Math.log(2));
        // if the number of players is not power of 2, this calculates how many matches
        // needs to played before it's power of 2
        var numOfPreBracketMatches = this.teams - Math.pow(2, Math.trunc(Math.log(this.teams) / Math.log(2)));
        this.buildStandard(numOfPreBracketMatches);
    }

    private buildStandard(numOfPreBracketMatches: number): void {

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
}