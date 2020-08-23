/* 
  Copyright 2020. Jefferson "jscher2000" Scher. License: MPL-2.0.
  version 0.1 - initial concept
  version 0.2 - open new tab next to current page
  version 0.5 - options page to specify which OWA to compose on (uses storage)
  version 0.6 - customize menu depending on which OWA will be used
  version 0.7 - handle multi-argument mailto links
*/

/**** Create and populate data structure ****/
// Default starting values
var oPrefs = {
	domain: "live.com"
}
// Update oPrefs from storage
let getPrefs = browser.storage.local.get("prefs").then((results) => {
	if (results.prefs != undefined){
		if (JSON.stringify(results.prefs) != '{}'){
			var arrSavedPrefs = Object.keys(results.prefs)
			for (var j=0; j<arrSavedPrefs.length; j++){
				oPrefs[arrSavedPrefs[j]] = results.prefs[arrSavedPrefs[j]];
			}
		}
	}
}).catch((err) => {console.log('Error retrieving "prefs" from storage: '+err.message);});

/**** Context menu items ****/
let linkcontext = browser.menus.create({
  id: "outlookmail_mailtolink",
  title: "New Outlook Mail message",
  contexts: ["link"],
  icons: {
	"16": "icons/mailto-outlook-16.png",
	"32": "icons/mailto-outlook-32.png"
  }
});

let pagecontext = browser.menus.create({
  id: "outlookmail_page",
  title: "New Outlook Mail message (page title, URL)",
  contexts: ["page"],
  icons: {
	"16": "icons/mailto-outlook-16.png",
	"32": "icons/mailto-outlook-32.png"
  }
});

let selcontext = browser.menus.create({
  id: "outlookmail_selection",
  title: "New Outlook Mail message (with selected text)",
  contexts: ["selection"],
  icons: {
	"16": "icons/mailto-outlook-16.png",
	"32": "icons/mailto-outlook-32.png"
  }
});

/* Customize menu wording for Office 365 */
var txtBase = 'New Outlook Mail message';
function updateMenu(){
	if (oPrefs.domain == 'live.com'){
		txtBase = 'New Outlook Mail message';
	} else {
		txtBase = 'New Office 365 message';
	}
	browser.menus.update("outlookmail_mailtolink", {
		title: txtBase
	});
	browser.menus.update("outlookmail_page", {
		title: txtBase + " (page title, URL)"
	});
	browser.menus.update("outlookmail_selection", {
		title: txtBase + " (with selected text)"
	});
}
getPrefs.then(() => {
	if (oPrefs.domain != 'live.com'){
		updateMenu();
	}
});

browser.menus.onClicked.addListener((menuInfo, currTab) => {
	var baseUrl = 'https://outlook.' + oPrefs.domain + '/mail/deeplink/compose?';
	var OWAUrl = '';
	switch (menuInfo.menuItemId) {
		case 'outlookmail_mailtolink':
			// Check the link protocol
			var linkSplit = menuInfo.linkUrl.split(':');
			if (linkSplit[0].toLowerCase() === 'mailto'){ // this is a mailto link
				linkSplit = linkSplit[1].split('?');
				OWAUrl = baseUrl + 'to=' + encodeURIComponent(linkSplit[0]);
				// handle additional mailto parameters (0.7)
				if (linkSplit.length > 1){
					for (var i=1; i<linkSplit.length; i++){
						var argSplit = linkSplit[i].split('&');
						for (var j=0; j<argSplit.length; j++){
							var param = argSplit[j].split('=');
							switch (param[0].toLowerCase()) {
								case 'cc':
									OWAUrl += '&cc=' + encodeURIComponent(param[1]);
									break;
								case 'bcc':
									OWAUrl += '&bcc=' + encodeURIComponent(param[1]);
									break;
								case 'subject':
									OWAUrl += '&subject=' + encodeURIComponent(param[1]);
									break;
								case 'body':
									OWAUrl += '&body=' + encodeURIComponent(param[1]);
									break;
								default:
									// ignore other parameters for now
							}
						}
					}
				}

			} else {	// this is not a mailto link
				// Share link
				console.log('What do to with this: '+menuInfo.linkUrl+'; text:'+menuInfo.linkText);
				OWAUrl = baseUrl + 'subject=Sharing%20a%20link' + '&body=Link:%20' + encodeURIComponent(menuInfo.linkUrl) + 
						' \n\n(From) ' + encodeURIComponent(currTab.url);
			}
			break;
		case 'outlookmail_page':
			// this is Email Link for the page
			OWAUrl = baseUrl + 'subject=' + encodeURIComponent(currTab.title) + '&body=Link:%20' + encodeURIComponent(currTab.url);
			break;
		case 'outlookmail_selection':
			OWAUrl = baseUrl;
			// This could be selected on a link... Try to find an address
			if (menuInfo.linkUrl){
				var linkSplit = menuInfo.linkUrl.split(':');
				if (linkSplit[0].toLowerCase() === 'mailto'){ // this is a mailto link
					linkSplit = linkSplit[1].split('?');
					OWAUrl += 'to=' + encodeURIComponent(linkSplit[0]);
					// do not add mailto parameters because are going to get those from the page
					OWAUrl += '&';
				}
			}
			// this is Email Link for the page when text is selected
			var seltext = menuInfo.selectionText;
			seltext = seltext.replace(/^\s*$/,'').replace(/\r/g,'\r').replace(/\n/g,'\n').replace(/^\s+|\s+$/g,' ');
			seltext = ' \n\n(Selection) ' + seltext.replace(new RegExp(/\u2019/g),'\'').replace(new RegExp(/\u201A/g),',').replace(new RegExp(/\u201B/g),'\'');
			OWAUrl += 'subject=' + encodeURIComponent(currTab.title) + '&body=Link:%20' + encodeURIComponent(currTab.url) +
					encodeURIComponent(seltext);
			break;
		default:
			// WTF?
	}
	if (OWAUrl.length > 0){
		browser.tabs.create({
			url: OWAUrl,
			index: currTab.index + 1,
			active: true
		}).then((newTab) => {
			/* No action */;
		}).catch((err) => {
			console.log('Error in tabs.create with "' + OWAUrl + '": ' + err.message);
		});
	}
});

/**** MESSAGING ****/
function handleMessage(request, sender, sendResponse) {
	if ("get" in request) {
		// Send oPrefs to Options page
		sendResponse({
			prefs: oPrefs
		});
	} else if ("update" in request) {
		// Receive pref update from Options page, store to oPrefs, and commit to storage
		var oSettings = request["update"];
		oPrefs.domain = oSettings.domain;
		updateMenu();
		browser.storage.local.set({prefs: oPrefs})
			.catch((err) => {console.log('Error on browser.storage.local.set(): '+err.message);});
	}
}
browser.runtime.onMessage.addListener(handleMessage);