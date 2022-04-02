const app = require('./app');
const {postEvents, getEventStream, putEvents, clientReset, getClientCount} = require('./client-util');
const {testData, chunks} = require('./test-data');


let server;

beforeAll((done) => {
    server = app.listen(process.env.PORT, () => {
        console.log(`Events service listening at http://${process.env.HOST}:${process.env.PORT}`);
        done();
    });
});

afterAll((done) => {
    console.log('Events service stopped.');
    app.stopPinging();
    server.close(() => done());
});

describe('Event stream API', () => {
    it('POST /events', (done) => {
        postEvents(chunks[0], (res) => {
            res.resume();
            res.on('end', () => {
                expect(res.complete).toBe(true);
                done();
            });
        });
    });
    it('GET /events/stream', (done) => {
        let receivedData = [];
        const events = getEventStream((data) => {
            receivedData = receivedData.concat(data);
            if (3 === receivedData.length) {
                getClientCount((res) => {
                    res.resume();
                    res.on('data', (data) => {
                        expect(JSON.parse(data).count).toEqual(1)
                        events.close();
                        expect(receivedData).toEqual(testData)
                        done();
                    });
                });
            }
        });
        postEvents(chunks[1], () => {
        });
        postEvents(chunks[2], () => {
        });
    });

    it('PUT /events', (done) => {
        let p1 = new Promise((resolve) => putEvents(testData[0].task, resolve));
        let p2 = new Promise((resolve) => putEvents(testData[1].task, resolve));
        let p3 = new Promise((resolve) => putEvents(testData[2].task, resolve));
        Promise.all([p1, p2, p3]).then(() => {
            const events = getEventStream((data) => {
                events.close();
                expect(data).toEqual([])
                done();
            });
        })
    });
});


