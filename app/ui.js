// Select DOM elements to work with
const welcomeDiv = document.getElementById("WelcomeMessage");
const signInButton = document.getElementById("SignIn");
const cardDiv = document.getElementById("card-div");
const searchM365Button = document.getElementById("searchM365");
const profileDiv = document.getElementById("profile-div");

function showWelcomeMessage(account) {
  // Reconfiguring DOM elements
  cardDiv.style.display = "initial";
  welcomeDiv.innerHTML = `Welcome ${account.username}`;
  signInButton.setAttribute("onclick", "signOut();");
  signInButton.setAttribute("class", "btn btn-success");
  signInButton.innerHTML = "Sign Out";
}

function updateSearchUI(data, endpoint) {
  console.log("Graph API responded at: " + new Date().toString());

  if (endpoint === graphConfig.graphMicrosoftSearchEndpoint) {
    if (data.value[0].hitsContainers[0].hits.length < 1) {
      alert("Search result is empty!");
    } else {
      const tabList = document.getElementById("list-tab");
      tabList.innerHTML = ""; // clear tabList at each readMail call
      const tabContent = document.getElementById("nav-tabContent");
      
      data.value[0].hitsContainers[0].hits.map((d, i) => {
        // Keeping it simple
        if (i < 10) {
          const listItem = document.createElement("a");
          listItem.setAttribute(
            "class",
            "list-group-item list-group-item-action"
          );
          listItem.setAttribute("id", "list" + i + "list");
          listItem.setAttribute("data-toggle", "list");
          listItem.setAttribute("href", "#list" + i);
          listItem.setAttribute("role", "tab");
          listItem.setAttribute("aria-controls", i);
          listItem.innerHTML = d._source.webUrl;
          tabList.appendChild(listItem);

          const contentItem = document.createElement("div");
          contentItem.setAttribute("class", "tab-pane fade");
          contentItem.setAttribute("id", "list" + i);
          contentItem.setAttribute("role", "tabpanel");
          contentItem.setAttribute("aria-labelledby", "list" + i + "list");
          contentItem.innerHTML = d._source.webUrl;
          tabContent.appendChild(contentItem);
        }
      });
    }
  }
}
