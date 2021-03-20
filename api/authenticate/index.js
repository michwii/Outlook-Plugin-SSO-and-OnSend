
var AAD_HELPER = require("./AAD_HELPER");
var console;

module.exports = async function (context, req) {
    const bootstrapToken = extractToken(req.headers['authorization']);
    const tenantId = '323d9c5b-c193-4c4f-8fb6-a029b2a10ca3';
    console = context;

    var AAD_Helper = new AAD_HELPER(tenantId);
    context.res = await AAD_Helper.exchangeBootstrapToken(bootstrapToken);
}

extractToken = function(header){
    if (header){
        const split = header.split(' ');
        return split[1];
    }else{
        return "";
    }
}
