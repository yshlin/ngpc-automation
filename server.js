/**
 * A custom SSE-based (server-sent events) event streaming service.
 */
const app = require('./app');

server = app.listen(process.env.PORT, () => {
    console.log(`Events service listening at http://localhost:${process.env.PORT}`)
});
