/**
* @NotOnlyCurrentDoc
**/

function getPipedriveService() {
    // Create a new service with the given name. The name will be used when
    // persisting the authorized token, so ensure it is unique within the
    // scope of the property store.
    return OAuth2.createService('pipedrive')
  
        // Set the endpoint URLs, which are the same for all Google services.
        .setAuthorizationBaseUrl('https://oauth.pipedrive.com/oauth/authorize')
        .setTokenUrl('https://oauth.pipedrive.com/oauth/token')
        
        // Set the client ID and secret, from the Google Developers Console.
        .setClientId('0ed37b47151c07f0')
        .setClientSecret('77866127873a8e503d6de8187c1aa1c7216a6b63')
  
        // Set the name of the callback function in the script referenced
        // above that should be invoked to complete the OAuth flow.
        .setCallbackFunction('authCallback')
  
        // Set the property store where authorized tokens should be persisted.
        .setPropertyStore(PropertiesService.getUserProperties())
        .setCache(CacheService.getUserCache())
        .setLock(LockService.getUserLock())
  
        // Set the scopes to request
        .setScope('deals:read')
  
        // Below are Google-specific OAuth2 parameters.
  
        // Sets the login hint, which will prevent the account chooser screen
        // from being shown to users logged in with multiple accounts.
        .setParam('login_hint', Session.getActiveUser().getEmail())
  
        // Requests offline access.
        .setParam('access_type', 'offline')
  
        // Forces the approval prompt every time. This is useful for testing,
        // but not desirable in a production application.
        .setParam('approval_prompt', 'force');
  }
  
  function authCallback(request) {
    let pipedriveService = getPipedriveService();
    let isAuthorized = pipedriveService.handleCallback(request);
    if (isAuthorized) {
      return HtmlService.createHtmlOutput('Success! You can close this tab.');
    } else {
      return HtmlService.createHtmlOutput('Denied. You can close this tab');
    }
  }
  
  function getPipedriveAccounts() {
    let pipedriveService = getPipedriveService();
    let response = UrlFetchApp.fetch('https://api-proxy.pipedrive.com/deals', {
      headers: {
        Authorization: 'Bearer ' + pipedriveService.getAccessToken(),
        'Content-Type': 'application/json'
      }
    });
    let parsedResponse = JSON.parse(response);
    
    let collectedData = [];
    parsedResponse["data"].forEach(function(element) {
        collectedData.push({
          title: element['title'],
          org_name: element['org_name']
        });
    });
    return collectedData;
  }