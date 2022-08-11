// cell data: [col, row]. Each starts with "1"
const stateCell = [1,1], sizeCell = [2,1], numMarkedCell = [3,1], noteCell = [2,2], firstGridCell = [2,4];
// changeGridSizeLabelStart, changeGridSizeBoxStart; columns remain constant
const cGridLStart = [2, 8], cGridBStart = [3, 8];
// newGameLabelStart, newGameBoxStart; columns remain constant
const nGameLStart = [2, 9], nGameBStart = [3, 9];
// loadingValueStart, loadingBarStart, loadingTextStart
const loadVStart = [4, 7], loadBStart = [4, 8], loadTStart = [4, 9];
const column = [" ", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
const normPixel = 21, notePixel = 42, gridPixel = 90;
const font = "Courier New";
const normFontSize = 10, noteFontSize = 24, noteFontSizeS = 16, gridEmptyFontSize = 85, gridMarkedFontSize = 48;
const oSym = "⭘", xSym = "✕";
// 0: create new game (set all cells)
// 1: determine first turn
// 2: X turn
// 3: O turn
// 4: game over (wait for a new game)
const states = [0, 1, 2, 3, 4];
const startSize = 3, startWidth = 5, startHeight = 9;
const sizes = ["three", "four"];
const sparklineOptions = '{"charttype","bar"; "max",100; "color1","lightgrey"; "rtl",false; "nan","ignore"; "empty","zero"}';

const noteMessages = ["Click board to start", "X Turn", "O Turn", "X Wins", "O Wins", "Tie"];
const winColor = "#def8d5";

/*
 * When the spreadsheet is open, check if the game needs to be set up,
 * and creates the game
 */
function onOpen(e) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getActiveSheet();
  let state = s.getRange(column[stateCell[0]] + stateCell[1]).getValue();
  let needNewBoard = true;
  for (let i = 0; i < states.length; i++) {
    if (state == states[i]) {needNewBoard = false; break;}
  }
  if (needNewBoard) {
    createNewGame(s, 0);
  }
}

/*
 * Runs the game and changes the game's states when a valid edit to the
 * sheet is made
 */
function onEdit(e) {
  // get the container spreadsheet and a sheet to play the game on
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getActiveSheet();

  // get the size of the game's grid
  let currSize = s.getRange(column[sizeCell[0]] + sizeCell[1]).getValue();
  let currSizeIdx = 0;
  for (let i = 0; i < sizes.length; i++) {
    if (currSize == sizes[i]) {currSizeIdx = i; break;}
  }

  // upon an edit, check if it is in the valid range, then go to next game state
  // (upper left corner, game grid, new game checkbox)
  if (e) {
    let thisRow = e.range.getRow();
    let thisCol = e.range.getColumn();
    let length = startSize + currSizeIdx;
    if ((thisCol >= firstGridCell[0] && thisCol < firstGridCell[0] + length && 
        thisRow >= firstGridCell[1] && thisRow < firstGridCell[1] + length) ||
        (thisCol == nGameBStart[0] && thisRow == nGameBStart[1] + currSizeIdx) || 
        (thisCol == stateCell[0] && thisRow == stateCell[1]) ) {
      runGame(s);
    }
  }
}

/*
 * Determines the states the game is on and goes to the next states.
 */
function runGame(sheet) {
  try {
    // get state and size of the sheet
    let currState = sheet.getRange(column[stateCell[0]] + stateCell[1]).getValue();
    let currSize = sheet.getRange(column[sizeCell[0]] + sizeCell[1]).getValue();
    let currSizeIdx = 0;
    for (let i = 0; i < sizes.length; i++) {
      if (currSize == sizes[i]) {currSizeIdx = i; break;}
    }

    // check if the new game checkbox was checked, and if a different size was chosen
    // and create the required game grid
    if (sheet.getRange(column[nGameBStart[0]] + (nGameBStart[1] + currSizeIdx)).isChecked()) { 
      if (sheet.getRange(column[cGridBStart[0]] + (cGridBStart[1] + currSizeIdx)).isChecked()) {
        let newSizeIdx = (currSizeIdx == sizes.length - 1) ? 0 : (currSizeIdx + 1);
        sheet.getRange(column[stateCell[0]] + stateCell[1]).setValue(states[0]);
        sheet.getRange(column[sizeCell[0]] + sizeCell[1]).setValue(sizes[newSizeIdx]);
        currState = sheet.getRange(column[stateCell[0]] + stateCell[1]).getValue();
        currSize = sheet.getRange(column[sizeCell[0]] + sizeCell[1]).getValue();
        currSizeIdx = newSizeIdx;
      }
      createNewGame(sheet, currSizeIdx);
    } else if (currState == states[0]) { // new game to be created
      createNewGame(sheet, currSizeIdx);
    } else if (currState == states[1]) { // determine the starting player
      determineStartPlay(sheet, currSizeIdx);
    } else if (currState == states[2] || currState == states[3]) { // play a player's turn
      playTurn(sheet, currSizeIdx);
    } else { // create a new sheet if no valid states
      createNewGame(sheet, currSizeIdx);
    }

    // reset the loading indicators
    sheet.getRange(column[loadTStart[0] + currSizeIdx] + (loadTStart[1] + currSizeIdx)).setValue("");
    sheet.getRange(column[loadVStart[0] + currSizeIdx] + (loadVStart[1] + currSizeIdx)).setValue("");
  } catch (err) {
    Logger.log('Failed with error %s', err.message);
  }
}

/* Checks for if a player has checked a grid box, marks it to the corresponding symbol,
 * checks if the game is over, and sets the game to over or hands the turn to the next player
 */
function playTurn(sheet, sizeIdx) {
  sheet.getRange(column[loadTStart[0] + sizeIdx] + (loadTStart[1] + sizeIdx)).setValue("Loading...");
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("33");

  let currState = sheet.getRange(column[stateCell[0]] + stateCell[1]).getValue();
  let sym = (currState == states[2]) ? xSym : oSym;

  // go through each cell in the grid, setting the first checked box to the correct symbol
  // check the surroundings, when enough cells are marked, for if the move created a win
  let played = false, won = false, tied = false;
  for (let i = 0; (i < startSize + sizeIdx) && (played == false); i++) {
    for (let j = 0; (j < startSize + sizeIdx) && (played == false); j++) {
      let currCell = sheet.getRange(column[firstGridCell[0] + i] + (firstGridCell[1] + j));
      if (currCell.isChecked()) {
        played = true;
        currCell.removeCheckboxes();
        currCell.setFontColor("#000000");
        currCell.setFontSize(gridMarkedFontSize);
        currCell.setValue(sym); 

        let numMarkedRange = sheet.getRange(column[numMarkedCell[0]] + numMarkedCell[1]);
        let numMarked = numMarkedRange.getValue() + 1;
        numMarkedRange.setValue(numMarked);

        if (numMarked >= (2 * (startSize + sizeIdx) - 1)) {
          won = checkWin(sheet, sym, sizeIdx, i, j);
          if (numMarked >= ((startSize + sizeIdx) * (startSize + sizeIdx))) {tied = !won;}
        }
      }
    }
  }
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("80");
  if (!played) {return;}

  // game not won or tied, next player's turn
  if (!won && !tied) {
    sheet.getRange(column[stateCell[0]] + stateCell[1]).setValue(states[(currState == states[2]) ? 3 : 2]);
    sheet.getRange(column[noteCell[0]] + noteCell[1]).setValue(noteMessages[(currState == states[2]) ? 2 : 1]);
  // game won or tied, set game to finished state then set the final game message
  } else {
    sheet.getRange(column[stateCell[0]] + stateCell[1]).setValue(states[4]);
    if (tied) {
      sheet.getRange(column[noteCell[0]] + noteCell[1]).setValue(noteMessages[5]);
    } else {
      sheet.getRange(column[noteCell[0]] + noteCell[1]).setValue(noteMessages[(currState == states[2]) ? 3 : 4]);
      }
  }  
}

/*
 * Parameters: sym - the symbol to be marked on the grid upon play
 *             sizeIdx - the grid's size, starting from 0
 *             col - the column to be marked
 *             row - the row to be marked
 * Returns:    (boolean) whether the game was won
 */
function checkWin(sheet, sym, sizeIdx, col, row) {
  let gridRange = sheet.getRange(column[firstGridCell[0]] + firstGridCell[1] + ":" + 
                                 column[firstGridCell[0] + startSize - 1 + sizeIdx] + 
                                 (firstGridCell[1] + startSize - 1 + sizeIdx));
  let gridVals = gridRange.getValues();

  let hor = true, ver = true, downDiag = true, upDiag = true;

  for (let i = 0; i < gridVals.length; i++) {
    if (gridVals[row][i] != sym) {hor = false;}
    if (gridVals[i][col] != sym) {ver = false;}
    if (gridVals[i][i] != sym) {downDiag = false;}
    if (gridVals[i][gridVals.length - 1 - i] != sym) {upDiag = false;}
  }

  if (hor) {
    let r = sheet.getRange(column[firstGridCell[0]] + (firstGridCell[1] + row) + ":" +
                           column[firstGridCell[0] + gridVals.length - 1] + (firstGridCell[1] + row));
    r.setBackground(winColor);
  }
  if (ver) {
    let c = sheet.getRange(column[firstGridCell[0] + col] + firstGridCell[1] + ":" + 
                           column[firstGridCell[0] + col] + (firstGridCell[1] + gridVals.length - 1));
    c.setBackground(winColor);
  }
  if (downDiag) {
    for (let i = 0; i < gridVals.length; i++) { 
      sheet.getRange(column[firstGridCell[0] + i] + (firstGridCell[1] + i)).setBackground(winColor); 
    }
  }
  if (upDiag) {
    for (let i = 0; i < gridVals.length; i++) { 
      sheet.getRange(column[firstGridCell[0] + i] + (firstGridCell[1] + gridVals.length - 1 - i)).setBackground(winColor); 
    }
  }

  let win = (hor || ver || downDiag || upDiag);
  return win;
}

/*
 * Chooses the starting player (X or O) randomly and sets the game and next states accordingly
 */
function determineStartPlay(sheet, sizeIdx) {
  sheet.getRange(column[loadTStart[0] + sizeIdx] + (loadTStart[1] + sizeIdx)).setValue("Loading...");
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("33");

  let gridRange = sheet.getRange(column[firstGridCell[0]] + firstGridCell[1] + ":" + 
                                 column[firstGridCell[0] + startSize - 1 + sizeIdx] + 
                                 (firstGridCell[1] + startSize - 1 + sizeIdx));

  let startGame = false;
  for (let i = 0; i < startSize + sizeIdx; i++) {
    for (let j = 0; j < startSize + sizeIdx; j++) {
      if (sheet.getRange(column[firstGridCell[0] + i] + (firstGridCell[1] + j)).isChecked()) {startGame = true;}
    }
  }
  if (startGame == false) {return;}
  
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("66");

  let randStart = Math.floor(Math.random() * 2); // 0 or 1

  sheet.getRange(column[noteCell[0]] + noteCell[1]).setFontSize(noteFontSize);
  sheet.getRange(column[noteCell[0]] + noteCell[1]).setValue(noteMessages[1 + randStart]);

  gridRange.setBackground("#ffffff");
  gridRange.uncheck();

  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("99");

  // set state values
  sheet.getRange(column[stateCell[0]] + stateCell[1]).setValue(states[2 + randStart]);
  sheet.getRange(column[numMarkedCell[0]] + numMarkedCell[1]).setValue(0); // number of grid cells marked
}

/*
 * Sets the number of columns and rows, cell sizes, merging, values, and borders,
 * background and text colours, text font families and sizes for a new game on the spreadsheet
 */
function createNewGame(sheet, sizeIdx) {
  let width = sheet.getMaxColumns();
  let height = sheet.getMaxRows();

  sheet.clear(); // clear sheet contents and formats

  // determine how many cols and rows are needed
  let gameWidth = startWidth + sizeIdx;
  let gameHeight = startHeight + sizeIdx + 1;

  // enforce correct number of columns and rows
  if (width > gameWidth) { sheet.deleteColumns(1, width - gameWidth); } 
  else if (width < gameWidth) { sheet.insertColumns(1, gameWidth - width); }

  if (height > gameHeight) { sheet.deleteRows(1, height - gameHeight); }
  else if (height < gameHeight) { sheet.insertRows(1, gameHeight - height); }

  /*
   * set formats
   */
  // cell ranges
  let allRange = sheet.getRange("A1:" + column[gameWidth] + gameHeight);
  let gridRange = sheet.getRange(column[firstGridCell[0]] + firstGridCell[1] + ":" + 
                                 column[firstGridCell[0] + startSize - 1 + sizeIdx] + 
                                 (firstGridCell[1] + startSize - 1 + sizeIdx));
  let noteRange = sheet.getRange(column[noteCell[0]] + noteCell[1]);
  let labelsRange = sheet.getRange(column[cGridLStart[0]] + (cGridLStart[1] + sizeIdx) + ":" +
                                   column[nGameLStart[0]] + (nGameLStart[1] + sizeIdx));
  let boxesRange = sheet.getRange(column[cGridBStart[0]] + (cGridBStart[1] + sizeIdx) + ":" +
                                  column[nGameBStart[0]] + (nGameBStart[1] + sizeIdx));
  let loadRange = sheet.getRange(column[loadBStart[0] + sizeIdx] + (loadBStart[1] + sizeIdx) + ":" + 
                                 column[loadTStart[0] + sizeIdx] + (loadTStart[1] + sizeIdx));
  let optionsRange = sheet.getRange(column[cGridLStart[0] + sizeIdx] + (cGridLStart[1] + sizeIdx) + ":" +
                                    column[loadTStart[0] + sizeIdx] + (loadTStart[1] + sizeIdx));
  // remove all checkboxes
  allRange.removeCheckboxes();
  // set column widths
  sheet.setColumnWidth(1, normPixel);
  sheet.setColumnWidth(gameWidth, normPixel);
  sheet.setColumnWidths(2, startSize + sizeIdx, gridPixel);
  // set row heights
  sheet.setRowHeightsForced(1, gameHeight, normPixel);
  sheet.setRowHeight(2, notePixel);
  sheet.setRowHeightsForced(4, startSize + sizeIdx, gridPixel);
  // set background colour
  allRange.setBackground("#ffffff");
  gridRange.setBackground("#f2f2f2");
  // set font family
  allRange.setFontFamily(font);
  // set font sizes
  noteRange.setFontSize(noteFontSize);
  gridRange.setFontSize(gridEmptyFontSize);
  optionsRange.setFontSize(normFontSize);
  // set font weights
  noteRange.setFontWeight("bold");
  labelsRange.setFontWeight("bold");
  // set font colours
  allRange.setFontColor("#ffffff");
  noteRange.setFontColor("#000000");
  gridRange.setFontColor("#f5f5f5");
  labelsRange.setFontColor("#000000");
  boxesRange.setFontColor("#999999");
  loadRange.setFontColor("#000000");
  // set loading bar
  sheet.getRange(column[loadBStart[0] + sizeIdx] + (loadBStart[1] + sizeIdx)).setValue(
      "=SPARKLINE(" + column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx) + 
      "," + sparklineOptions + ")");
  sheet.getRange(column[loadTStart[0] + sizeIdx] + (loadTStart[1] + sizeIdx)).setValue("Loading...");
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("25");
  // merge note cells
  sheet.getRange(column[noteCell[0]] + noteCell[1] + ":" + 
                 column[noteCell[0] + startSize + sizeIdx - 1] + noteCell[1]).mergeAcross();
  // set horizontal orientations
  noteRange.setHorizontalAlignment("center");
  gridRange.setHorizontalAlignment("center");
  labelsRange.setHorizontalAlignment("right");
  boxesRange.setHorizontalAlignment("center");
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("35");
  // set vertical orientations
  gridRange.setVerticalAlignment("middle");
  noteRange.setVerticalAlignment("middle");
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("45");
  // set bottom options
  let nextSize = (sizeIdx == sizes.length - 1) ? startSize : startSize + sizeIdx + 1;
  sheet.getRange(column[cGridLStart[0]] + (cGridLStart[1] + sizeIdx)).setValue(nextSize + "x" + nextSize + ": ");
  sheet.getRange(column[nGameLStart[0]] + (nGameLStart[1] + sizeIdx)).setValue("New Game: ");
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("55");
  // insert checkboxes
  gridRange.insertCheckboxes();
  sheet.getRange(column[cGridBStart[0]] + (cGridBStart[1] + sizeIdx)).insertCheckboxes();
  sheet.getRange(column[nGameBStart[0]] + (nGameBStart[1] + sizeIdx)).insertCheckboxes();
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("75");
  // set grid borders
  gridRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.DOUBLE);
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("85");
  // set note text
  noteRange.setFontSize(noteFontSizeS);
  noteRange.setValue(noteMessages[0]);
  sheet.getRange(column[loadVStart[0] + sizeIdx] + (loadVStart[1] + sizeIdx)).setValue("95");
  // set state values
  sheet.getRange(column[stateCell[0]] + stateCell[1]).setValue(states[1]);
  sheet.getRange(column[sizeCell[0]] + sizeCell[1]).setValue(sizes[sizeIdx]);
}
