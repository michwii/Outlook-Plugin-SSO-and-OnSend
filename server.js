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

app.post('/api/authenticate', async (req, res) => {
  extractToken = function(header){
    if (header){
        const split = header.split(' ');
        return split[1];
    }else{
        return "";
    }
  }
  const bootstrapToken = extractToken(req.headers['authorization']);
  const tenantId = '323d9c5b-c193-4c4f-8fb6-a029b2a10ca3';

  var ADD_Helper = new AAD_HELPER(tenantId);
  const accessToken = await ADD_Helper.exchangeBootstrapToken(bootstrapToken);

  res.json(accessToken);
});

var server = https.createServer(options, app);

server.listen(port, () => {
  console.log("server starting on port : " + port)
});