
/*
 * Global Variables
 */
var LOGGING_ENABLED = true;

//Set a start date to filter out older emails
// var date = new Date('August 22, 2015 00:00:00');
// var START_TIME = date.getTime();
var START_DATE = new Date('August 22, 2015 00:00:00');
var MAX_THREADS_TO_PROCESS = 2;

//Set which label we are using to filter which emails are being retrieved
var RSVP_LABEL = 'Wedding RSVPs';
var SPREADSHEET_ID = '1Tqa1SaUywdLx_IQS__4DU-D8B9g905KIdn4cC3suql4';

if(LOGGING_ENABLED){
	//Clear out previous logs
	Logger.clear();
	Logger.log("Starting script: Parse Wedding RSVPs");
}

/*
 * Function that retrieves a list of gmail email threads
 * with a provided label and parses them.
 * @param - {string} rsvp_label - The label applied to emails that should be parsed.
 */
function processRSVPs(rsvp_label) {

	var rsvp_threads = getGmailThreadsWithLabelName(rsvp_label);

	if(rsvp_threads.length <= 0){
		Logger.log("No Wedding RSVP Threads found");
		return false;
	}

	messages = [];
	for(var i = 0; i < rsvp_threads.length; i++) {

		messages = messages.concat(rsvp_threads[i].getMessages());

		//While debugging, only try to process the first few email threads
		if(MAX_THREADS_TO_PROCESS > 0 && (i + 1) >= MAX_THREADS_TO_PROCESS){
			break;
		}

	}

	//Filter out messages we don't want to parse (based on date or other criteria)
	messages = filterMessages(messages);

	for(i = 0; i < messages.length; i++){
		Logger.log(messages[i].getBody());

		var rsvp = parseRSVPEmail(messages[i].getBody());
		rsvp.date = formatDate(messages[i].getDate());

		//Add rsvp entry to spreadsheet
		if(addRSVPToSpreadsheet(rsvp)) {
			Logger.log("RSVP Successfully added to spreadsheet: " + JSON.stringify(rsvp));
		}else{
			Logger.log("RSVP could not be added to spreadsheet: " + JSON.stringify(rsvp));
		}

	}

}

/*
 * Loops through an array of email threads and parses each thread.
 * @param - {Object} rsvp - the rsvp object to add to the spreadsheet.
 * @returns - {boolean} - true if added to the spreadsheet,
 *  false if already exists or there was an error.
 */
function addRSVPToSpreadsheet(rsvp){

	//Check if the spreadsheet exists or create a new one
	spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

	if(!spreadsheet){
		return false;
	}

	return false;

}


/*
 * Filters an array of messages and returns a modified version of the original array
 * @param - {GmailMessage[]} messages - list of messages to be filtered.
 */
function filterMessages(messages) {

	if(messages.length <= 0){
		return false;
	}

	for (var i = 0; i < messages.length; i++) {

		//Filter out messages if they are older than our start date.
		if(!checkMessageDate(messages[i], START_DATE)) {
			messages.splice(i, 1);
			continue;
		}

	}

	return messages;

}

/*
 * Returns true if a message occurred after a specified date, otherwise false.
 * This allows emails to be skipped that have already been processed.
 * @param {GmailMessage} message - The GmailMessage object to be checked.
 * @returns {boolean}
 */
function checkMessageDate(message, start_date) {
	return (message.getDate().getTime() > start_date.getTime()) ? true : false;
}

/*
 * Parses an email of a specific format to retrieve its component parts
 * The email must be formatted with as: <string>: <string>, representing a key/value pair.
 * @param - {string} body - the contents of an rsvp email.
 * @returns - {Object} - rsvp details object.
 */
function parseRSVPEmail(body) {

	body = body.replaceAll('<br />', '', body);

	var message_lines = body.split("\n");

	//Initialize our RSVP Object
	rsvp = {};

	for(i = 0; i < message_lines.length; i++) {

		var line_parts = message_lines[i].split(':');

		//Make the line type lower case, replace spaces with underscores, and remove any other unexpected character
		// var line_type = line_parts[0].toLowerCase().replaceAll(' ', '_').replaceAll('/[^a-z0-9_]/', '');
		var line_type = line_parts[0].toLowerCase().replaceAll(' ', '_', line_type);
		line_type = replaceAll('/[^a-z0-9_]/', '', line_type);
		
		//In case there's a ":" in the message, reassemble the rest of the line
		var line_value = "";

		for(var j = 1; j < line_parts.length; j++) {
			line_value += line_parts[j];
		}
		
		line_value = line_value.trim();

		rsvp[line_type] = line_value;

		if(line_type == 'can_attend'){
			//can_attend is the last line of the email that needs to be parsed
			//Break out of all loops or return
			break;
		}

	}

	return rsvp;

}

/*
 * Returns an array of GmailThreads that match a provided label.
 * @param {string} label_name - Gmail Label name being searched for.
 */
function getGmailThreadsWithLabelName(label_name) {

	var threads = [];
	var new_threads = [];
	var start = 0;
	var loop_count = 30;

	//First see if we can find the label
	var label = GmailApp.getUserLabelByName(label_name);

	//We could not retrieve a label, fail out.
	if(!label){
		return threads;
	}

	do {
		new_threads = label.getThreads(start,loop_count);
		threads = threads.concat(new_threads);
		start += loop_count;
	}
	while(new_threads.length > 0);

	if(LOGGING_ENABLED){
		for (var i = 0; i < threads.length; i++) {
			Logger.log(threads[i].getFirstMessageSubject());
		}
	}
	
	return threads;

}

/*
 * Additional String method expanding on the existing replace method,
 * allowing all matches of a regular expression to be replaced.
 * @param {string} find - Regular expression string
 * @param {string} replace - String to replace the matches that are found.
 */
// String.prototype.replaceAll = function(find, replace) {
// 	return this.replace(new RegExp(escapeRegExp(find), 'g'), replace);
// }

function replaceAll(find, replace, subject){
	return subject.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}

/*
 * Helper function that escapes a regular expression provided as a string
 * @param {string} regex - Regular expression string.
 */
function escapeRegExp(regex) {
	return regex.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}

/*
 * Helper function to format a javascript datetime value to be stored in a spreadsheet
 */
function formatDate(date) {
	var formatted_date = (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear();
	var formatted_time = date.getHours() + ':' + date.getMinutes() + ':' + date.getSeconds();
	return formatted_date + ' ' + formatted_time;
}

processRSVPs(RSVP_LABEL);