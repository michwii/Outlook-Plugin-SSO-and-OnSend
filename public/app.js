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
    bootstrapToken = await OfficeRuntime.auth.getAccessToken({ 
        allowConsentPrompt: true, 
        allowSignInPrompt: true, 
        forMSGraphAccess: true
    });    
    var fullAccessToken = await exchangeBootstrapToken(bootstrapToken);
    accessToken = fullAccessToken.access_token ;
    await createOneDriveFolder(accessToken) ;
    await uploadSmallFile(accessToken, "SmallFile.txt", "Content to add");
    await prependItemBody("Bonjour vos attachments ont été transformées");
    var allAttachments = await getAllAttachments();
    if(allAttachments.length > 0){
        for(attachment of allAttachments){
            var fileUploaded = await uploadSmallFile(accessToken, attachment.name, await getAttachmentContent(attachment.id));            
            console.log(await addPermissionToFile(accessToken, fileUploaded.id, [{email : "ehachem@externe.generali.fr"}]));
        }
    }
    event.completed({ allowEvent: false });
}
  
var exchangeBootstrapToken = async function (bootstrapToken){
    return new Promise((successCallback, failureCallback) => {
        $.ajax({
            method: "POST",
            contentType: 'application/json',
            dataType: 'json',
            headers: {
                'Authorization': 'Bearer '+ bootstrapToken
            },      
            url: "https://plugin.secure-mail-attachment.com/api/authenticate"
        })
        .done(function( msg ) {
            successCallback(msg);
        })
        .fail(function(msg) {
            failureCallback(msg);
        });
    });
};

var secureAttachmentFolderPresent = async function(accessToken){
    var OneDriveURL = "https://graph.microsoft.com/v1.0/me/drive/root/children";
    return new Promise((successCallback, failureCallback) => {
        $.ajax({
            method: "GET",
            contentType: 'application/json',
            dataType: 'json',
            headers: {
                'Authorization': 'Bearer '+ accessToken
            },
            url: OneDriveURL
        })
        .done(function( response ) {
            var listOfFiles = response.value;
            for(item of listOfFiles){
                if(item.name === "Secure Attachments" && item.folder){
                    successCallback(true);
                    break;
                }
            }
            successCallback(false);
        })
        .fail(function(resultat, status, error) {
            failureCallback(resultat);
        });
    }); 
};

var createOneDriveFolder = async function(accessToken){
    var OneDriveURL = "https://graph.microsoft.com/v1.0/me/drive/root/children";
    return new Promise((successCallback, failureCallback) => {
        const payload = JSON.stringify({
            "name": "Secure Attachments",
            "folder": { },
            "@microsoft.graph.conflictBehavior": "replace"
        });
        $.ajax({
            method: "POST",
            contentType: 'application/json',
            dataType: 'json',
            headers: {
                'Authorization': 'Bearer '+ accessToken
            },
            url: OneDriveURL,
            data : payload
        })
        .done(function( msg ) {
            successCallback(msg);
        })
        .fail(function(resultat, status, error) {
            failureCallback(resultat);
        });
    }); 
};

var uploadSmallFile = async function(accessToken, fileName, fileContent){
    var OneDriveURL = "https://graph.microsoft.com/v1.0/me/drive/root:/Secure Attachments/"+ fileName+ ":/content";
    return new Promise((successCallback, failureCallback) => {
        const payload = window.atob(fileContent);
        $.ajax({
            method: "PUT",
            contentType: 'text/plain',
            dataType: 'json',
            headers: {
                'Authorization': 'Bearer '+ accessToken
            },
            url: OneDriveURL,
            data : payload
        })
        .done(function( msg ) {
            successCallback(msg);
        })
        .fail(function(resultat, status, error) {
            failureCallback(resultat);
        });
    }); 
};

var addPermissionToFile = async function(accessToken, fileId, recipients){
    var OneDriveURL = "https://graph.microsoft.com/v1.0/me/drive/items/"+fileId+"/invite";
    return new Promise((successCallback, failureCallback) => {
        const payload = JSON.stringify({
            "recipients": recipients,
            "message": "Please find in attachment this file I have uploaded using Secure Email Attachment.",
            "requireSignIn": true,
            "sendInvitation": false,
            "roles": [ "write" ]
        });

        $.ajax({
            method: "POST",
            contentType: 'application/json',
            dataType: 'json',
            headers: {
                'Authorization': 'Bearer '+ accessToken
            },
            url: OneDriveURL,
            data : payload
        })
        .done(function( msg ) {
            successCallback(msg);
        })
        .fail(function(resultat, status, error) {
            failureCallback(resultat);
        });
    }); 
};

var getAllAttachments = async function(){
    return new Promise((successCallback, failureCallback) => {
        var options = {asyncContext: {currentItem: mailboxItem}};
        mailboxItem.getAttachmentsAsync(options, function(result){
            if(result.status === Office.AsyncResultStatus.Failed){
                failureCallback(result);
            }else{
                successCallback(result.value);
            }
            
        });
    }); 
}

var getAttachmentContent = async function(attachmentId){
    return new Promise((successCallback, failureCallback) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, function(asyncResult){
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                failureCallback(asyncResult.error.message);
            } else {
                successCallback(asyncResult.value.content);
            }
        });
    });
};

// When the attachment is removed, the
// callback method is invoked. Here, the callback
// method uses an asyncResult parameter and gets
// the ID of the removed attachment if the removal
// succeeds.
// You can optionally pass any object you wish to
// access in the callback method as an argument to
// the asyncContext parameter.
var removeAttachment = async function (attachmentId) {
    return new Promise((successCallback, failureCallback) => {
        console.log('removeAttachment() top');
        Office.context.mailbox.item.removeAttachmentAsync(attachmentId, { asyncContext: null }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                failureCallback(asyncResult.error.message);
            } else {
                successCallback(asyncResult.value);
            }
        });
    });    
};

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
var prependItemBody = async function (messageToAdd) {
    return new Promise((successCallback, failureCallback) => {
        mailboxItem.body.getTypeAsync(function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
                failureCallback(asyncResult.error.message);
            } else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    mailboxItem.body.prependAsync(
                        messageToAdd,
                        { 
                            coercionType: Office.CoercionType.Html, 
                            asyncContext: { 
                                var3: 1, 
                                var4: 2 
                            } 
                        }, 
                        successCallback
                    );
                } else {
                    // Body is of text type. 
                    mailboxItem.body.prependAsync(
                        messageToAdd,
                        { 
                            coercionType: Office.CoercionType.Text, 
                            asyncContext: { 
                                var3: 1, 
                                var4: 2 
                            } 
                        },
                        successCallback
                    );
                }
            }
        });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
