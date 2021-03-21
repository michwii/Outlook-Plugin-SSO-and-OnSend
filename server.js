const express = require('express');
const https = require('https');
const fs = require('fs');
const port = 443;
var AAD_HELPER = require("./api/authenticate/AAD_HELPER");

var key = fs.readFileSync(__dirname + '/certificates/private.key');
var cert = fs.readFileSync(__dirname + '/certificates/public.cert');
var options = {
  key: key,
  cert: cert
};

app = express()
app.use(function (req, res, next) {
  console.log('Time:', Date.now());
  next();
});
app.use(express.static('public'));



app.get('/', (req, res) => {
   res.send('Now using https..');
});

var server = https.createServer(options, app);

server.listen(port, () => {
  console.log("server starting on port : " + port)
});