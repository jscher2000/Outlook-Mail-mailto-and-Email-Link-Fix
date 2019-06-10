/* 
  Copyright 2018. Jefferson "jscher2000" Scher. License: MPL-2.0.
  version 0.1 - initial concept
  version 0.2 - open new tab next to current page
*/

// this URL launches the Inbox first, then a message on the side, very slow:
// const baseUrl = 'http://outlook.live.com/default.aspx?rru=compose&';

// This URL creates a stand-alone tab that gets "stranded" but it's faster:
const baseUrl = 'https://outlook.live.com/mail/deeplink/compose?';

/**** Context menu item ****/

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

browser.menus.onClicked.addListener((menuInfo, currTab) => {
	var OWAUrl = '';
	switch (menuInfo.menuItemId) {
		case 'outlookmail_mailtolink':
			// Check the link protocol
			var linkSplit = menuInfo.linkUrl.split(':');
			if (linkSplit[0].toLowerCase() === 'mailto'){ // this is a mailto link
				OWAUrl = baseUrl + 'to=' + encodeURIComponent(linkSplit[1]);
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
					OWAUrl += 'to=' + encodeURIComponent(linkSplit[1]) + '&';
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
