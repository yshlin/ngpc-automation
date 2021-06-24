const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const {v4: uuid} = require('uuid');
const app = express();

app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: false}));
//TODO: auth with API key

let clients = [];
let events = [];

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
    console.log(`${clientId} listening to event stream`);
});

app.post('/events', (request, respsonse) => {
    const newEvent = request.body;
    if (newEvent.hasOwnProperty('type') && newEvent.hasOwnProperty('task') && newEvent.hasOwnProperty('email') && newEvent.type === 'task') {
        console.log(`Adding event of type ${newEvent.task}`);
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
        console.log(`Clearing events of task ${event.task}`);
        events = events.filter(event => event.task !== event.task);
        response.sendStatus(200);
    } else {
        response.sendStatus(403);
    }
});



module.exports = app
