// Select current option
browser.runtime.sendMessage({
	get: "oPrefs"
}).then((oSettings) => {
	var SLD = oSettings.prefs.domain.replace('.com', '');
	var myRads = document.querySelectorAll('input[name="siteRads"]');
	for (var i=0; i<myRads.length; i++){
		if (myRads[i].value == SLD){
			myRads[i].checked = true;
			myRads[i].setAttribute('checked', 'checked');
			// we're done with this for loop
			break;
		}
	}
}).catch((err) => {
	console.log('Problem getting settings: '+err.message);
});
function updatePref(evt){
	// Only if one of our radio button was clicked
	if (evt.target.nodeName != 'INPUT') return;
	if (evt.target.getAttribute('name') != 'siteRads') return;
	// Check value
	var myRads = document.querySelectorAll('input[name="siteRads"]');
	for (var i=0; i<myRads.length; i++){
		if (myRads[i].checked){
			var domainValue = myRads[i].value + '.com';
			// Send preference update
			var oPrefsUpdate = {domain: domainValue};
			// send to background script
			browser.runtime.sendMessage({
				update: oPrefsUpdate
			});
			// we're done with this for loop
			break;
		}
	}
}
// Attach event handler to the select control
document.getElementById('pSiteRads').addEventListener('click', updatePref, false);