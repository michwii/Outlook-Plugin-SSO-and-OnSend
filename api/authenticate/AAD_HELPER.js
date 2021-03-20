const got = require("got");
var FormData = require("form-data");

class AAD_HELPER {

    constructor(tenantId){
        this._tenantId = tenantId;
    }

    exchangeBootstrapToken = async function(bootstrapToken){
        var bodyFormData = new FormData();
        bodyFormData.append('grant_type', 'urn:ietf:params:oauth:grant-type:jwt-bearer');
        bodyFormData.append('client_id', '08bbbb87-f561-49c5-b8b7-f9cbd2a096ed');
        bodyFormData.append('client_secret', 'B.ZhLFFQViSF22-3nH_kpM8Lmx1SZIZ~~8');
        bodyFormData.append('assertion', bootstrapToken);
        bodyFormData.append('scope', 'openid');
        bodyFormData.append('requested_token_use', 'on_behalf_of');
        const url = 'https://login.microsoftonline.com/'+this._tenantId+'/oauth2/v2.0/token';
        
        try{
            const {body} = await got(url, {
                method: 'POST',
                headers: bodyFormData.getHeaders(),
                body: bodyFormData,
                responseType: 'json'
            });
            return body;
        }catch(e){
            console.log('Erreur détectée');
            return e.response.body;
        }
    };
}

module.exports = AAD_HELPER;