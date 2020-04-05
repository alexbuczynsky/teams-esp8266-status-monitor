/**
 * @type {IAppOptions}
 */
const DEFAULT_OPTIONS = {
  esp_base_url: "http://10.0.0.30",
  status_red: "Busy,Do not disturb",
  status_yellow: "Be right back",
  status_off: "Available,Appear away",
  teams_status: "Away"
}

const OPTIONS_STATE = {...DEFAULT_OPTIONS};

let teams_status_state = "";

/**
 * @returns {Promise<IAppOptions>}
 */
function loadOptions() {
  console.debug("loading options...")
  // return new Promise(resolve => chrome.storage.sync.get(DEFAULT_OPTIONS, (options) => {
  //   console.debug("options loaded", options)
  //   resolve(options);
  // }))

  return new Promise(resolve => resolve(OPTIONS_STATE))
}

/**
 * @param {IAppOptions} options
 * @returns {Promise<void>}
 */
function saveOptions(options) {
  return new Promise(resolve => {
    Object.assign(OPTIONS_STATE, options)
    resolve();
  })
}

// ON INSTALLED
chrome.runtime.onInstalled.addListener(async () => {
  console.log('[onInstalled] event received');

  const options = await loadOptions();
  await saveOptions({...DEFAULT_OPTIONS, ...options});
  console.log('Options initialized');
  onStartup();
});


// fetch and save data when chrome restarted, alarm will continue running when chrome is restarted
chrome.runtime.onStartup.addListener(() => {
  console.log('[onStartup] event received');
  onStartup();
});

function onStartup() {
  setInterval(checkConnection, 1000)
  setInterval(fetchTeamsStatus, 1000)
  setInterval(reloadTeamsTab, 10000)
}

const getTeamsStatus = () => document.querySelectorAll(`[mri='appHeaderBar.authenticatedUserMri'`)[0].children[0].title
const reloadPage = () => window.location.reload()


async function checkConnection(){
  const options = await loadOptions();
  const client = new ESPClient(options.esp_base_url);
  const isOnline = await client.ping();

  console.log(isOnline ? "ONLINE" : "OFFLINE")

  if(isOnline){
    const {teams_status} = await loadOptions();
    console.log({teams_status});

      switch(options.teams_status){
        case "Away":
        case "Appear away":
        case "Be right back":
          await client.sendColors({ red: false, yellow: true})
          break;
        case "Busy":
        case "Do not disturb":
          await client.sendColors({ red: true, yellow: false})
          break;
        case "Available":
          await client.sendColors({ red: false, yellow: false})
          break;
        // default:
        //   await client.sendColors({ red: false, yellow: false})
      }
  }
}

function getTeamsTab() {
  return new Promise(resolve => {
    chrome.tabs.query({ url: "https://teams.microsoft.com/*"}, (results) => {
      if (results.length == 0) {
        chrome.tabs.create({url: 'https://teams.microsoft.com/'});
        return getTeamsTab();
      }else{
        return resolve(results[0]);
      }
    })
  })
}

async function executeScriptOnTeamsTab(script) {
  const teamsTab = await getTeamsTab();
  return new Promise(resolve => {
    chrome.tabs.executeScript(teamsTab.id, {
      code: '(' + script + ')();' //argument here is a string but function.toString() returns function's code
    }, resolve);
  })
  
}



async function reloadTeamsTab() {
  console.log("starting reload...")
  const teamsTab = await getTeamsTab();
  console.log({teamsTab})
  console.log("sending reload command...")
  chrome.tabs.reload(teamsTab.id)
}


console.log("Popup DOM fully loaded and parsed");
async function fetchTeamsStatus() {
  const [teams_status] = await executeScriptOnTeamsTab(getTeamsStatus)
  console.log('Received Teams Status: ' + teams_status)
  await saveOptions({ teams_status: teams_status })
}




class ESPClient {
  static timeout(ms, promise) {
    return new Promise(function(resolve, reject) {
      setTimeout(function() {
        reject(new Error("timeout"))
      }, ms)
      promise.then(resolve, reject)
    })
  }

  /**
   * @param {string} baseurl
   */
  constructor(baseurl, timeout = 5000) {
    this.baseurl = baseurl;
    this.timeout = timeout;
  }

  async ping() {
    try {
      await ESPClient.timeout(this.timeout, fetch(this.baseurl));
      return true;
    }catch(err) {
      console.error(err)
      return false;
    }
  }

  sendLEDSignal(gpio, isOn = false) {
    return fetch(`${this.baseurl}/${gpio}/${isOn ? "on" : "off"}`);
  }

  sendColors({red = false, yellow = false}){
    return Promise.all([
      this.sendLEDSignal(4, red),
      this.sendLEDSignal(5, yellow)
    ])
  }
}