const fs = require('fs');
const puppeteer = require("puppeteer");

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

    } catch (error) {
        console.error(`Error scraping data: ${error.message}`);
    } finally {
        await browser.close();
    }
}

async function main() {
    let cycleCount = 0;

    while (true) {
        cycleCount++;
        console.log(`Starting cycle ${cycleCount}...`);

        await scrapeData();  // Scrape and save the data to JSON

        // Call the Python script to update the Excel file
        const { exec } = require('child_process');
        exec('python3 update_excel.py', (error, stdout, stderr) => {
            if (error) {
                console.error(`Error executing Python script: ${error.message}`);
                return;
            }
            if (stderr) {
                console.error(`Python script stderr: ${stderr}`);
                return;
            }
            console.log(`Python script stdout: ${stdout}`);
        });

        console.log(`Cycle ${cycleCount} completed.\n`);

        const countdownSeconds = 3;
        for (let i = countdownSeconds; i >= 0; i--) {
            process.stdout.write(`\rTime left until next cycle: ${i} seconds`);
            await new Promise(resolve => setTimeout(resolve, 1000));
        }

        console.log('\n');
    }
}

main();