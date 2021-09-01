// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new Msal.UserAgentApplication(msalConfig);

var teamsContext = {};

function setTeamsContext(context) {
  teamsContext = context;
  loginRequest.loginHint = teamsContext.loginHint;
  tokenRequest.loginHint = teamsContext.loginHint;
  
  if (myMSALObj.getAccount()) {
    showWelcomeMessage(myMSALObj.getAccount());
  }
}

function signIn() {
  myMSALObj.loginPopup(loginRequest)
    .then(loginResponse => {
      console.log("id_token acquired at: " + new Date().toString());
      console.log(loginResponse);
      
      if (myMSALObj.getAccount()) {
        showWelcomeMessage(myMSALObj.getAccount());
        seeProfile();
      }
    }).catch(error => {
      console.log(error);
    });
}

function signOut() {
  myMSALObj.logout();
}

function getTokenPopup(request) {
  return myMSALObj.acquireTokenSilent(request)
    .catch(error => {
      console.log(error);
      console.log("silent token acquisition fails. acquiring token using popup");
          
      // fallback to interaction when silent call fails
        return myMSALObj.acquireTokenPopup(request)
          .then(tokenResponse => {
            return tokenResponse;
          }).catch(error => {
            console.log(error);
          });
    });
}

function seeProfile() {
  if (myMSALObj.getAccount()) {
    getTokenPopup(loginRequest)
      .then(response => {
        callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, updateUI);
        profileButton.classList.add("d-none");
        mailButton.classList.remove("d-none");
        membersButton.classList.remove("d-none");
      }).catch(error => {
        console.log(error);
      });
  } else {    
    signIn();
  }
}

function readMail() {
  if (myMSALObj.getAccount()) {
    getTokenPopup(tokenRequest)
      .then(response => {
        callMSGraph(graphConfig.graphMailEndpoint, response.accessToken, updateUI);
      }).catch(error => {
        console.log(error);
      });
  }
}

function readMembers() {
  var endpoint = graphConfig.graphEndpoint + "/groups/" + teamsContext.groupId + "/members";
  if (myMSALObj.getAccount()) {
    getTokenPopup(tokenRequest)
      .then(response => {
        callMSGraph(endpoint, response.accessToken, updateUI);
      }).catch(error => {
        console.log(error);
      });
  }
}
