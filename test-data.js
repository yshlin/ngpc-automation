let testData = [{
    "type": "task",
    "task": "weeklyPub",
    "email": "test@example.com",
}, {
    "type": "task",
    "task": "mergePptx",
    "email": "test@example.com",
}, {
    "type": "task",
    "task": "youtubeSetup",
    "email": "test@example.com",
},]
let chunks = [JSON.stringify(testData[0]), JSON.stringify(testData[1]), JSON.stringify(testData[2])];

module.exports = {testData, chunks}