const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const {v4: uuid} = require('uuid');
const {response} = require("express");
const app = express();

const auth = (req, res, next) => {
    if (req.get('Api-Key') === process.env.APIKEY) {
        next();
    } else {
        res.sendStatus(403);
    }
}

app.use(auth);
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: false}));

let clients = [];
let events = [];

app.get('/clients/count', (request, response) => {
    response.send({count: clients.length});
});

app.get('/events/stream', (request, response) => {
    const headers = {
        'Content-Type': 'text/event-stream',
        'Connection': 'keep-alive',
        'Cache-Control': 'no-cache'
    };
    response.writeHead(200, headers);

    const data = `data: ${JSON.stringify(events)}\n\n`;
    response.write(data);

    const clientId = uuid().toString();

    const newClient = {
        id: clientId,
        response
    };

    clients.push(newClient);

    request.on('close', () => {
        // console.log(`${clientId} events stream closed`);
        clients = clients.filter(client => client.id !== clientId);
    });
    console.log(`${clientId} listening to event stream from ${request.header('Origin')}`);
});

app.post('/events', (request, respsonse) => {
    const newEvent = request.body;
    if (newEvent.hasOwnProperty('type') && newEvent.hasOwnProperty('task') && newEvent.hasOwnProperty('email') && newEvent.type === 'task') {
        console.log(`Adding event of type ${newEvent.task} from ${request.header('Origin')}`);
        events.push(newEvent);
        clients.forEach(client => client.response.write(`data: ${JSON.stringify(newEvent)}\n\n`));
        respsonse.sendStatus(200);
    } else {
        respsonse.sendStatus(403);
    }
});

app.put('/events', (request, response) => {
    const event = request.body;
    if (event.type && event.task && event.type === 'task') {
        console.log(`Clearing events of task ${event.task} from ${request.header('Origin')}`);
        events = events.filter(event => event.task !== event.task);
        response.sendStatus(200);
    } else {
        response.sendStatus(403);
    }
});

/**
 * keeping server alive
 */
let intervalId = setInterval(() => {
    clients.forEach(client => client.response.write(`data: []\n\n`));
}, 50000);

app.stopPinging = () => {
    clearInterval(intervalId)
}

module.exports = app
