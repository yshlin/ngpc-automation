/**
 * A custom client that listen to event stream, waiting for task commands.
 */
const {getEventStream, putEvents} = require('./client-util');
const {spawn} = require('child_process');

let taskState = {};     // making sure only one task is running at a time, duplicated task requests are ignored

function setState(task, state = null) {
    taskState[task] = state;
}

function runningTask() {
    for (let task in taskState) {
        if ('running' === taskState[task]) {
            return task;
        }
    }
    return null;
}

function waitingTask() {
    for (let task in taskState) {
        if ('waiting' === taskState[task]) {
            return task;
        }
    }
    return null;
}

function runTask(ev) {
    let callback = (code) => {
        console.log(`child process exited with code ${code}`);
        putEvents(ev.task, () => {})
        setState(ev.task);
        let waiting = waitingTask();
        if (waiting) {
            runTask(waiting)
        }
    };
    setState(ev.task, 'running');

    if (process.env.READ_ONLY) {
        callback(1);
        return;
    }
    let pyCommand = ['./main.py', `--task=${ev.task}`, `--email=${ev.email}`];
    if (process.env.DRY_RUN) {
        pyCommand.push('--dry-run')
    }
    console.log(`Running python command "${pyCommand}"`);
    const taskProc = spawn('python', pyCommand);
    taskProc.stdout.on('data', (data) => {
        console.log(data.toString());
    });
    taskProc.stderr.on('data', (data) => {
        console.log(data.toString());
    });
    taskProc.on('close', callback);
}

const events = getEventStream((evs) => {
    for (const ev of evs) {
        if (ev.type !== 'task' || !['hymnsDbSync', 'weeklyPub', 'mergePptx', 'youtubeSetup'].includes(ev.task)) {
            console.log('Unsupported event');
            continue;
        }
        let running = runningTask();
        if (running) {
            if (ev.task !== running) {
                setState(ev.task, 'waiting');
            }
        } else {
            runTask(ev);
        }
    }
});

process.on('SIGINT', () => {
    console.log('Event stream closed.');
    events.close();
});