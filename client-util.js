const s = (process.env.PORT === '443' ? 's': '');
const http = require('http' + s);
const EventSource = require('eventsource');
require('log-timestamp');


function postEvents(chunk, callback, url= {}) {
    const req = http.request({
        host: url.host ? url.host : process.env.HOST,
        port: url.port ? url.port : process.env.PORT,
        path: '/events',
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(chunk),
            'Api-Key': process.env.APIKEY,
        }
    }, callback);
    req.write(chunk);
    req.end();
    console.log('POST request sent')
}

function getEventStream(callback) {
    const events = new EventSource(
        `http${s}://${process.env.HOST}:${process.env.PORT}/events/stream`,
        {headers: {'Api-Key': process.env.APIKEY}}
    );

    events.onopen = () => {
        console.log('Event stream open.');
        events.confirmedEventReception = false;
    }

    events.onmessage = (e) => {
        let parsedData = JSON.parse(e.data);
        if (!Array.isArray(parsedData)) {
            parsedData = [parsedData];
        }
        if (parsedData.length > 0) {
            console.log(`Received ${parsedData.length} events`);
            callback(parsedData);
        } else if (!events.confirmedEventReception) {
            // only callback on first empty event response to confirm reception
            console.log('Event reception confirmed.');
            callback(parsedData);
        }
        events.confirmedEventReception = true;
    };

    events.onerror = function (e) {
        if (e) {
            if (e.status === 401 || e.status === 403) {
                console.log('not authorized');
            }
        }
    };
    console.log('Listen to event stream.');
    return events;
}

function putEvents(task, callback, url={}) {
    let chunk = JSON.stringify({
        type: 'task',
        task: task,
    });
    const req = http.request({
        host: url.host ? url.host : process.env.HOST,
        port: url.port ? url.port : process.env.PORT,
        path: '/events',
        method: 'PUT',
        headers: {
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(chunk),
            'Api-Key': process.env.APIKEY,
        },
    }, callback);
    req.write(chunk);
    req.end();
    console.log('PUT request sent')
}

function getClientCount(callback, url={}) {
    const req = http.request({
        host: url.host ? url.host : process.env.HOST,
        port: url.port ? url.port : process.env.PORT,
        path: '/clients/count',
        method: 'GET',
        headers: {
            'Api-Key': process.env.APIKEY,
        },
    }, callback);
    req.end();
    console.log('GET request sent')
}

module.exports = {postEvents, getEventStream, putEvents, getClientCount}