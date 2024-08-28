import puppeteer from "puppeteer"; // Import Puppeteer for web scraping
import readline from 'readline'; // Able to interact with script in the terminal

/**
 * Prints the names of bowlers currently in the active list.
 * 
 * @param activeList - An array of bowler IDs that are currently in the active list.
 * @param bowlers - An array of objects representing different bowlers. Each bowler object has
 * properties like `id` (the bowler's unique identifier) and `name` (the bowler's name).
 */
function printActiveBowlers(activeList, bowlers) {
  if (activeList.length === 0) {
    console.log("No bowlers are currently in the active list.");
  } else {
    console.log("Bowlers currently in the active list:");
    activeList.forEach(id => {
      const bowler = bowlers.find(b => b.id === id);
      if (bowler) {
        console.log(`- ${bowler.name}`);
      }
    });
  }
}

/**
 * Retrieves a list of all bowlers from a column in Excel sheet.
 * Either Scratch or Handicap sidepot.
 * 
 * @returns array of 'bowlers' from that column
 */
async function getCompleteBowlerList() {
  
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
 * Updates the scores of a bowler in Excel to their respective column for the set.
 * 
 * @param bowlerName - The bowlers name on screen whose scores you want to update in the database.
 * @param newScores - An array of scores that's updated for a specific bowler identified by their name.
 */
async function updateBowlerScores(bowlerName, newScores) {

}


async function updateHandicapSheet(pdf) {
  
}


async function main() {

  // Initialize maxScoreArrayLength at the beginning
  let maxScoreArrayLength = await getMaxScoreArrayLength();
  maxScoreArrayLength = await promptForMaxScoreArrayLength(maxScoreArrayLength);

  await promptForNewBowlers(maxScoreArrayLength); // Prompt to add new bowlers before starting
  let bowlers = await getCompleteBowlerList(); // Get the complete bowlers list

  // Normalize all bowler names in the database
  bowlers = bowlers.map(bowler => ({
    ...bowler,
    name: normalizeName(bowler.name),
    nicknames: bowler.nicknames ? bowler.nicknames.map(normalizeName) : []
  }));

  let activeList = []; // Initialize an empty active list
  let hasProcessedAtLeastOneActiveBowler = false; // Flag to check if at least one active bowler has been processed

  // ! Recover any bowlers that have a first score but not a second score in case the script was interrupted
  for (const bowler of bowlers) {
    const bowlerRef = doc(db, "bowlers", bowler.id);
    const bowlerSnapshot = await getDoc(bowlerRef);
    const databaseScoresArray = bowlerSnapshot.data().scores || [];

    const firstGameIndex = maxScoreArrayLength - 2; // Index for the first game
    const secondGameIndex = maxScoreArrayLength - 1; // Index for the second game

    // Check if the bowler has a first score but not a second score
    if (databaseScoresArray[firstGameIndex] && !databaseScoresArray[secondGameIndex]) {
      console.log(`${bowler.name} has a first score but not a second score, adding them to the active list.`);
      hasProcessedAtLeastOneActiveBowler = true; // Set the flag to true for the active player
      activeList.push(bowler.id); // Add the bowler to the active list
    }
  }

  let cycleCount = 0; // Initialize a counter for the number of loops

  while (true) { // Keep looping until we manually break

    cycleCount++; // Increment the counter for each loop iteration
    console.log(`Starting cycle ${cycleCount}...`); // Display the current cycle count

    const scrapedData = await scrapeData(); // Scrape data from LaneTalk

    for (const player of scrapedData) {
      // Find the bowler by name or nickname
      const bowler = findBowlerByNameOrNickname(player.name, bowlers);

      if (bowler) { // If the bowler exists in the database, begin the process
        const firstGameIndex = maxScoreArrayLength - 2; //even slot
        const secondGameIndex = maxScoreArrayLength - 1; //odd slot

        // Fetch the most recent scores from Firestore
        let bowlerRef = doc(db, "bowlers", bowler.id);
        let bowlerSnapshot = await getDoc(bowlerRef);
        let databaseScoresArray = bowlerSnapshot.data().scores || [];

        // console.log(`Initial databaseScoresArray for ${bowler.name}:`, databaseScoresArray);

        // Case 1: Bowler is not in the active list, update the first score and add them to the active list
        if (!activeList.includes(bowler.id)) {
          if (!databaseScoresArray[firstGameIndex] && player.scoresArray.length > 0) {
            databaseScoresArray[firstGameIndex] = player.scoresArray[0]; // Insert the first score

            await updateBowlerScores(bowler.id, databaseScoresArray); // Update the database with the first score
            await updateBowlerAverage(bowler.id, databaseScoresArray); // Update the bowler's average

            console.log(`${bowler.name}'s first score updated to ${player.scoresArray[0]}`);

            activeList.push(bowler.id); // Add bowler to the active list
            hasProcessedAtLeastOneActiveBowler = true; // Set the flag to true
          }
        }
        // Case 2: Bowler is already in the active list, update the second score and remove them from the active list
        else {
          if (!databaseScoresArray[secondGameIndex] && player.scoresArray.length > 1) {
            databaseScoresArray[secondGameIndex] = player.scoresArray[1]; // Insert the second score
            await updateBowlerScores(bowler.id, databaseScoresArray); // Update the database with the second score
            await updateBowlerAverage(bowler.id, databaseScoresArray); // Update the bowler's average
            console.log(`${bowler.name}'s second score updated to ${player.scoresArray[1]}`);

            activeList = activeList.filter(id => id !== bowler.id); // Remove bowler from the active list

            console.log(`${bowler.name} has completed both scores.`);
          }
        }
      }
    }

    console.log(`Cycle ${cycleCount} completed.`); // Indicate that the current cycle is complete

    printActiveBowlers(activeList, bowlers); // Print the bowlers currently in the active list

    // Start the countdown for the next cycle
    const countdownSeconds = 3; // Time in seconds for the countdown
    for (let i = countdownSeconds; i >= 0; i--) {
      process.stdout.write(`\rTime left until next cycle: ${i} seconds`); // Overwrite the current line with time left
      await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for 1 second
    }

    console.log('\n'); // Move to the next line after the countdown

    // Break the loop if all active bowlers have completed their scores AND at least one bowler has been processed
    if (activeList.length === 0 && hasProcessedAtLeastOneActiveBowler) {
      console.log("All bowlers have completed both scores, stopping the script.");
      break; // Exit the loop
    }
  }

  // After all bowlers have their scores updated, finalize the session
  await updateMaxScoreArrayLength(maxScoreArrayLength + 2);
  console.log(`Updated maxScoreArrayLength to ${maxScoreArrayLength + 2}`);

  console.log("Script completed successfully. Exiting...");
  process.exit(0); // Exit the process
}

main();