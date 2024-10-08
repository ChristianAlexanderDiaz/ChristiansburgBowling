const fs = require('fs');
const puppeteer = require("puppeteer");
const { execSync } = require('child_process');
const XLSX = require('xlsx');

async function scrapeData() {
    const browser = await puppeteer.launch({ headless: "new" });
    const page = await browser.newPage();
    const url = "https://beta.lanetalk.com/bowlingcenters/367895d6-d3f2-4d89-b4a6-9f11ea7af700/completed";
    let data = null;

    try {
        await page.goto(url, { waitUntil: "networkidle0", timeout: 60000 });
        await page.waitForSelector(".player.ng-star-inserted");

        data = await page.evaluate(() => {
            const players = Array.from(document.querySelectorAll('.player.ng-star-inserted'));
            return players.map(player => {
                const nameText = player.querySelector('.nameAndDate > .name').innerText;
                const [name, team] = nameText.split('\n');
                const date = player.querySelector('.date').innerText;
                const scoreElements = player.querySelectorAll('.numbersWrapper > .numberBatch');

                const scoresArray = Array.from(scoreElements)
                    .map(el => el.innerText.trim())
                    .filter(score => score !== '' && !isNaN(score));

                return { name, team, scoresArray, date };
            });
        });

        // Save the data to a JSON file
        fs.writeFileSync('scraped_data.json', JSON.stringify(data, null, 2));
        console.log("Data saved to scraped_data.json");

        return data; // Return the scraped data
    } catch (error) {
        console.error(`Error scraping data: ${error.message}`);
    } finally {
        await browser.close();
    }
}

function getActiveListFromExcel(excelFile) {
    const workbook = XLSX.readFile(excelFile);
    const sheet = workbook.Sheets["Handicap Sidepot"];

    const activeList = [];
    for (let row = 2; row <= 51; row++) {
        const nameCell = sheet[`B${row}`] ? sheet[`B${row}`].v : null;
        const firstGame = sheet[`D${row}`] ? sheet[`D${row}`].v : null;
        const thirdGame = sheet[`F${row}`] ? sheet[`F${row}`].v : null;

        if (nameCell && firstGame && !thirdGame) {
            // The bowler has started (has at least one score) but hasn't finished (missing third game)
            activeList.push(nameCell);
        }
    }
    return activeList;
}

async function main() {
    const excelFile = '/Users/cynical/OneDrive/Documents/Wednesday Night Sidepots_MacroCopy.xlsm';
    let cycleCount = 0;
    let hasProcessedAtLeastOnceActiveBowler = false;
    let activeList = []; // Declare activeList outside the try block

    while (true) {
        cycleCount++;
        console.log("--------------------");
        console.log(`Logic:\nStarting cycle ${cycleCount}...`);

        await scrapeData();  // Scrape and save the data to JSON

        console.log(`Cycle ${cycleCount} completed.`);
        console.log("--------------------");

        // Execute the Python script synchronously
        try {
            const pythonOutput = execSync('python3 /Users/cynical/Documents/GitHub/ChristiansburgBowling/update_excel.py').toString();
            // Split the output to differentiate between updates and active list
            const outputSections = pythonOutput.split("--------------------");
            const excelUpdates = outputSections[1]?.trim() || "No updates.";
            const activeListOutput = outputSections[2]?.trim() || "No active bowlers left.";

            console.log("\n--------------------");
            console.log(`${excelUpdates}`);
            console.log("--------------------\n");

            console.log("--------------------");
            console.log(`${activeListOutput}`);
            console.log("--------------------\n");

            // Update activeList directly from the Excel data
            activeList = getActiveListFromExcel(excelFile);
            if (activeList.length > 0) {
                hasProcessedAtLeastOnceActiveBowler = true;

                console.log("Active List:");
                activeList.forEach(bowler => {
                    console.log(`- ${bowler}`);
                });
                console.log("--------------------\n");
            }

        } catch (error) {
            console.error(`Error executing Python script: ${error.message}`);
            // Handle case where activeList might not be set if an error occurs
            activeList = []; 
        }

        if (activeList.length > 0) {
            const countdownSeconds = 15;
            for (let i = countdownSeconds; i >= 0; i--) {
                process.stdout.write(`\rTime left until next cycle: ${i} seconds`);
                await new Promise(resolve => setTimeout(resolve, 1000));
            }
            console.log('\n');
        }

        if (activeList.length === 0 && hasProcessedAtLeastOnceActiveBowler) {
            console.log("All bowlers have completed their scores. Stopping the script.");
            break;
        }
    }
}

main();