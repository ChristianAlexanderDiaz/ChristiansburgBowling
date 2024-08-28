const puppeteer = require("puppeteer"); // Import Puppeteer for web scraping
const XLSX = require("xlsx"); // Import XLSX for Excel file manipulation
const fs = require("fs"); // Import fs for file system operations
const readline = require("readline"); // Able to interact with script in the terminal

/**
 * Load the excel workbook
 * @param {*} filename specified file path
 * @returns  the workbook object
 */
function loadWorkbook(filename) {
  const file = fs.readFileSync(filename); // Read the file
  return XLSX.read(file, { type: "buffer" }); // Read the file buffer
}

/**
 * Save the workbook to a specific file
 * @param {*} workbook the workbook object
 * @param {*} filename specified file path
 */
function saveWorkbook(workbook, filename) {
    const wbout = XLSX.write(workbook, {
        type: "buffer",
        bookType: "xlsx",
        cellStyles: true // Attempt to preserve cell styles
    });
    fs.writeFileSync(filename, wbout);
}

/**
 * Find the row of a specific bowler in the Excel sheet (0-index counting)
 * @param {*} sheet specified sheet
 * @param {*} bowlerName bowler name
 * @returns the row number of the bowler in the Excel sheet
 */
function findBowlerRow(sheet, bowlerName) {
  // ! Rows 2 to 51 correspond to bowlers in the Excel sheet (0 index counting)
  for (let row = 1; row <= 50; row++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: 1 })]; // Column B
    if (cell && cell.v && cell.v.trim().toUpperCase() === bowlerName.trim().toUpperCase()) {
      return row;
    }
  }
  return null;
}

function updateBowlerScores(sheet, row, newScores) {
    const scoreColumns = [3, 4, 5]; // Columns D, E, F correspond to indices 3, 4, 5 in JavaScript

    for (let score of newScores) {
        for (let col of scoreColumns) {
            const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
            const existingCell = sheet[cellRef];

            if (!existingCell || existingCell.v === '') { // Check if the cell is empty
                const newCell = { t: 'n', v: score };

                if (existingCell && existingCell.s) {
                    newCell.s = existingCell.s; // Preserve existing cell style
                }

                sheet[cellRef] = newCell; // Update the cell with new value and preserved style
                break; // Exit the inner loop after placing the score
            }
        }
    }
}

/**
 * Uses Puppeteer to scrape data from a specific URL, handling up to 3
 * attempts to load the page and extract information about players and their scores.
 * 
 * @returns The `scrapeData` function returns the scraped data from the webpage in the form of an array
 * of objects. Each object in the array represents a player and contains their name, team, an array of
 * scores, and the date of the game.
 */
async function scrapeData() {
  const browser = await puppeteer.launch({ headless: "new" });
  const page = await browser.newPage();
  const url = "https://beta.lanetalk.com/bowlingcenters/367895d6-d3f2-4d89-b4a6-9f11ea7af700/completed";
  const maxAttempts = 3;
  let attemptCount = 0;
  let data = null;

  while (attemptCount < maxAttempts) {
    attemptCount++; // Increment the attempt counter
    console.log(`Attempt ${attemptCount} to load the page...`);

    try {
      // Navigate to the target URL with a 1-minute timeout
      await page.goto(url, { waitUntil: "networkidle0", timeout: 60000 });

      // Wait for the data to load
      await page.waitForSelector(".player.ng-star-inserted");

      // Extract data from the page
      data = await page.evaluate(() => {
        const players = Array.from(document.querySelectorAll('.player.ng-star-inserted'));

        return players.map(player => {
          const nameText = player.querySelector('.nameAndDate > .name').innerText;
          const [name, team] = nameText.split('\n');
          const date = player.querySelector('.date').innerText;
          const scoreElements = player.querySelectorAll('.numbersWrapper > .numberBatch');

          // Convert NodeList to Array of scores and filter out empty strings
          const scoresArray = Array.from(scoreElements)
            .map(el => el.innerText.trim()) // Trim any whitespace around the score
            .filter(score => score !== '' && !isNaN(score)); // Filter out empty strings and non-numeric values

          return { name, team, scoresArray, date };
        });
      });

      console.log(`Successfully loaded the page on attempt ${attemptCount}.`);
      break; // Exit the loop if data was successfully scraped

    } catch (error) {
      console.error(`Attempt ${attemptCount} failed: ${error.message}`);

      if (attemptCount >= maxAttempts) {
        console.error("Maximum number of attempts reached. Unable to load the page.");
        throw error; // Re-throw the error if max attempts are reached
      }

      console.log(`Retrying... (${maxAttempts - attemptCount} attempts left)\n`);
    }
  }

  await browser.close();
  return data;
}


/**
 * The main function updates bowling scores in an Excel sheet, scraping data from LaneTalk and managing
 * active bowlers until all scores are completed.
 */
async function main() {
  // Load the workbook and get the specific sheet
  const filename = "/Users/cynical/OneDrive/Mario Kart Wii/Documents/Wednesday Night Sidepots_MacroCopy.xlsx";
  const workbook = loadWorkbook(filename);
  const sheetName = "Handicap Sidepot";
  const sheet = workbook.Sheets[sheetName];

  let activeList = [];   // Initialize an empty active list
  let hasProcessedAtLeastOneActiveBowler = false; // Flag to check if at least one active bowler has been processed to continue the script

  // Recover any bowlers that have a first score but not a second score in case the script was interrupted
  for (let row = 1; row <= 50; row++) { // Loop through B2:B51 (Rows 2 to 51 correspond to index 1 to 50)
    const bowlerNameCell = sheet[XLSX.utils.encode_cell({ r: row, c: 1 })];
    if (bowlerNameCell && bowlerNameCell.v) {
      const firstScoreCell = sheet[XLSX.utils.encode_cell({ r: row, c: 3 })]; // D column (First game)
      const secondScoreCell = sheet[XLSX.utils.encode_cell({ r: row, c: 4 })]; // E column (Second game)

      if (firstScoreCell && firstScoreCell.v && (!secondScoreCell || !secondScoreCell.v)) {
        console.log(`${bowlerNameCell.v} has a first score but not a second score, adding them to the active list.`);
        hasProcessedAtLeastOneActiveBowler = true;
        activeList.push(row); // Add the row index to the active list
      }
    }
  }

  let cycleCount = 0; // Initialize a counter for the number of cycles

  while (true) {
    cycleCount++; // Increment the cycle counter
    console.log(`Starting cycle ${cycleCount}...`);

    const scrapedData = await scrapeData(); // Scrape the data from LaneTalk

    for (const player of scrapedData) {
      const row = findBowlerRow(sheet, player.name); // Find the row of the bowler in the Excel sheet
      if (row !== null) { // If the bowler exists in the Excel sheet
        const firstScoreCell = sheet[XLSX.utils.encode_cell({ r: row, c: 3 })]; // D column (First game)
        const secondScoreCell = sheet[XLSX.utils.encode_cell({ r: row, c: 4 })]; // E column (Second game)
        const thirdScoreCell = sheet[XLSX.utils.encode_cell({ r: row, c: 5 })]; // F column (Third game)

        if (!activeList.includes(row)) {
          if (!firstScoreCell || !firstScoreCell.v) { // If the first score cell is empty
            updateBowlerScores(sheet, row, [player.scoresArray[0]]); // Update the first score
            console.log(`Updated ${player.name}'s first score to ${player.scoresArray[0]}.`);

            activeList.push(row); // Add the row index to the active list to keep searching
            hasProcessedAtLeastOneActiveBowler = true; // Set the flag to true
          }
        } else {
          if (firstScoreCell && firstScoreCell.v && (!secondScoreCell || !secondScoreCell.v)) {
            updateBowlerScores(sheet, row, [player.scoresArray[1]]); // Update the second score
            console.log(`Updated ${player.name}'s second score to ${player.scoresArray[1]}.`);
          } else if (secondScoreCell && secondScoreCell.v && (!thirdScoreCell || !thirdScoreCell.v)) {
            updateBowlerScores(sheet, row, [player.scoresArray[2]]); // Update the third score
            console.log(`Updated ${player.name}'s third score to ${player.scoresArray[2]}.`);

            activeList = activeList.filter(activeRow => activeRow !== row); // Remove the row from the active list
            console.log(`${player.name} has been removed from the active list.`);
          }
        }
      }
    }

    console.log(`Cycle ${cycleCount} completed.\n`);

    // printActiveList(activeList); // Print the active list of bowlers

    // Save the updated workbook
    saveWorkbook(workbook, filename);

    // Start the countdown for the next cycle
    const countdownSeconds = 3; // Time in seconds for the countdown
    for (let i = countdownSeconds; i >= 0; i--) {
      process.stdout.write(`\rTime left until next cycle: ${i} seconds`); // Overwrite the current line with time left
      await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for 1 second
    }

    console.log('\n'); // Move to the next line after the countdown

    // Break the loop if all active bowlers have completed their scores AND at least one bowler has been processed
    if (activeList.length === 0 && hasProcessedAtLeastOneActiveBowler) {
      console.log("All bowlers have completed all scores, stopping the script.");
      break; // Exit the loop
    }
  }

  console.log("All scores updated successfully. Exiting...");
  process.exit(0); // Exit the process
}


main();
