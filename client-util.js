const http = require('http');
const EventSource = require('eventsource');


function postEvents(chunk, callback, url= {}) {
    const req = http.request({
        host: url.host ? url.host : process.env.HOST,
        port: url.port ? url.port : process.env.PORT,
        path: '/api/events',
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
        `http://${process.env.HOST}:${process.env.PORT}/api/events/stream`,
        {headers: {'Api-Key': process.env.APIKEY}}
    );

    events.onmessage = (e) => {
        let parsedData = JSON.parse(e.data);
        if (!Array.isArray(parsedData)) {
            parsedData = [parsedData];
        }
        console.log(`Received ${parsedData.length} events`);
        callback(parsedData, events);
    };

    events.onerror = function (e) {
        if (e) {
            if (e.status === 401 || e.status === 403) {
                console.log('not authorized');
            }
        }
    };
    console.log('Listen to event stream.')
}

function putEvents(task, callback, url={}) {
    let chunk = JSON.stringify({
        type: 'task',
        task: task,
    });
    const req = http.request({
        host: url.host ? url.host : process.env.HOST,
        port: url.port ? url.port : process.env.PORT,
        path: '/api/events',
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

module.exports = {postEvents, getEventStream, putEvents}