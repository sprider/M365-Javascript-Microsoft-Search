const myMSALObj = new msal.PublicClientApplication(msalConfig);

let accessToken;
let username = "";

myMSALObj
  .handleRedirectPromise()
  .then(handleResponse)
  .catch((err) => {
    console.error(err);
  });

function handleResponse(resp) {
  if (resp !== null) {
    username = resp.account.username;
    showWelcomeMessage(resp.account);
  } else {
    const currentAccounts = myMSALObj.getAllAccounts();
    if (currentAccounts === null) {
      return;
    } else if (currentAccounts.length > 1) {
      console.warn("Multiple accounts detected.");
    } else if (currentAccounts.length === 1) {
      username = currentAccounts[0].username;
      showWelcomeMessage(currentAccounts[0]);
    }
  }
}

function signIn() {
  myMSALObj.loginRedirect(loginRequest);
}

function signOut() {
  const logoutRequest = {
    account: myMSALObj.getAccountByUsername(username),
  };

  myMSALObj.logout(logoutRequest);
}

function getTokenRedirect(request) {
  request.account = myMSALObj.getAccountByUsername(username);
  return myMSALObj.acquireTokenSilent(request).catch((error) => {
    console.warn(
      "silent token acquisition fails. acquiring token using redirect"
    );
    if (error instanceof msal.InteractionRequiredAuthError) {
      return myMSALObj.acquireTokenRedirect(request);
    } else {
      console.warn(error);
    }
  });
}

function searchM365(searchText) {
  getTokenRedirect(loginRequest)
    .then((response) => {
      callMSSearchGraph(
        graphConfig.graphMicrosoftSearchEndpoint,
        response.accessToken,
        searchText,
        updateSearchUI
      );
    })
    .catch((error) => {
      console.error(error);
    });
}
