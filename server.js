const server = require("./app");
const PORT = 3000;

server.listen(PORT, () => {
    console.log(`server listing on port ${PORT}`);
});
