const {postEvents, putEvents} = require('./client-util');
const {testData, chunks} = require('./test-data');


test('POST /events (live)', (done) => {
    postEvents(chunks[0], (res) => {
        res.resume();
        res.on('end', () => {
            expect(res.complete).toBe(true);
            done();
        });
    });
});

test('PUT /events (live)', (done) => {
    let p1 = new Promise((resolve) => putEvents(testData[0].task, resolve));
    let p2 = new Promise((resolve) => putEvents(testData[1].task, resolve));
    let p3 = new Promise((resolve) => putEvents(testData[2].task, resolve));
    Promise.all([p1, p2, p3]).then(() => {
        done();
    });
});