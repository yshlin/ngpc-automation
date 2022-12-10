const axios = require('axios');
const {spawn} = require('child_process');
const VALID_TASKS = ['youtubeSetup', 'weeklyPub', 'mergePptx', 'hymnsDbSync']
const API_URL = 'https://script.google.com/macros/s/AKfycbyJh3C9ucRKCw4bKzTfHnVwhrvwP39yGIlFIL6D25jTYTTVlXgy0QVOuPxuBQ6s84_b/exec';

console.log('Polling started');

async function runTask(task, email) {
    console.log('Running task ' + task)
    let pyCommand = ['./main.py', `--task=${task}`, `--email=${email}`];
    // if (process.env.DRY_RUN) {
    // pyCommand.push('--dry-run')
    // }
    console.log(`Running python command "${pyCommand}"`);
    const child = spawn('python', pyCommand);
    let data = "";
    for await (const chunk of child.stdout) {
        console.log('stdout chunk: '+chunk);
        data += chunk;
    }
    let error = "";
    for await (const chunk of child.stderr) {
        console.error('stderr chunk: '+chunk);
        error += chunk;
    }
    const exitCode = await new Promise( (resolve, reject) => {
        child.on('close', resolve);
    });

    if( exitCode) {
        throw new Error( `subprocess error exit ${exitCode}, ${error}`);
    }
    return data;
}

async function markTaskDone(task) {
    console.log('Marking task ' + task + ' as done.')
    let payload = task + '=false';
    let response = await axios.post(API_URL, payload);
}

async function init() {
    while (true) {
        console.log('Polling');
        let response = await axios.post(API_URL);
        let tasks = response.data;
        let contact = response.data.contact;
        for (const [k, v] of Object.entries(tasks)) {
            if (VALID_TASKS.includes(k) && true === v) {
                try {
                    await runTask(k, contact);
                } catch (e) {
                    console.error(e);
                } finally {
                    await markTaskDone(k);
                }
            }
        }
        await sleep(15000)
    }
}

function sleep(ms) {
    return new Promise((resolve) => {
        setTimeout(resolve, ms);
    });
}

init();
// markTaskDone('weeklyPub')