/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var mailboxItem;
var bootstrapToken ;
var accessToken ;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

Office.onReady(async info => {
    console.log("Office ready");
    
});

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
async function validateBody(event) {
    try{
        bootstrapToken = await OfficeRuntime.auth.getAccessToken({ 
            allowConsentPrompt: true, 
            allowSignInPrompt: true, 
            forMSGraphAccess: true
        });    
        console.log(bootstrapToken);
        event.completed({ allowEvent: true });
    }catch(exeption){
        console.log(exeption);
        event.completed({ allowEvent: false });
    }
    
}
  