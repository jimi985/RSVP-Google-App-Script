
var LOGGING_ENABLED = true;

//Set a start date to filter out older emails
var START_DATE = new Date('August 22, 2015 00:00:00');

//Set which label we are using to filter which emails are being retrieved
var rsvp_label = 'Wedding RSVPs';

if(LOGGING_ENABLED){
	Logger.log("Starting script: Parse Wedding RSVPs");
}

processRSVPs(RSVP_LABEL);

/*
 * Function that retrieves a list of gmail email threads
 * with a provided label and parses them.
 */
function processRSVPs(rsvp_label){

	var rsvp_threads = getGmailThreadsWithLabelName(rsvp_label);
	processThreads(rsvp_threads);

}

/*
 * Loops through an array of email threads and parses each thread.
 * @param - {string[]} - messages
 */
function processThreads(threads){

	if(threads.length <= 0){
		Logger.log("No Wedding RSVPs Threads found");
		return false;
	}

	for (var i = 0; i < threads.length; i++) {
		Logger.log(threads[i].getFirstMessageSubject());
		processThreadMessages(threads[i].getMessages());
	}

}

/*
 * Loops through a thread containing an array of messages and parses them.
 * @param - {string[]} - messages
 */
function processThreadMessages(messages){

	if(messages.length <= 0){
		return false;
	}

	for (var i = 0; i < messages.length; i++) {
		
		if(!checkMessageDate(messages[i])){
			Logger.log("Skipping message: " + messages[i].getMessage() + " because it is too old.");
			continue;
		}

		Logger.log(messages[i].getBody());
	}

}

/*
 * Returns boolean if a message occurred after a specified date.
 * This allows emails to be skipped that have already been processed.
 * @param {GmailMessage} message - The GmailMessage object to be checked.
 */
function checkMessageDate(message, start_date){
	return (message.getDate().getTime() > start_date.getTime()) ? true : false;
}

/*
 * Parses an email of a specific format to retrieve its component parts
 * The email must be formatted with as: <string>: <string>, representing a key/value pair.
 * @param - {string} body - the contents of an rsvp email.
 */
function parseRSVPEmail(body){
	msg = msg.replaceAll('<br />', '');

	var message_lines = msg.split("\n");

	//Name the loop so we can break out of it when needed
	message_loop:
	for(i = 0; i < message_lines.length; i++){

		var line_parts = message_lines[i].split(':');

		//Make the line type lower case, replace spaces with underscores, and remove any other unexpected character
		var line_type = line_parts[0].toLowerCase().replaceAll(' ', '_').replaceAll('/[^a-z0-9_]/', '');
		
		//In case there's a ":" in the message, reassemble the rest of the line
		var line_value = "";
		for(var j = 1; j < line_parts.length; j++){
			line_value += line_parts[j];
		}
		
		line_value = line_value.trim();

		switch(line_type){

			case 'from':
				break;

			case 'subject':
				break;

			case 'name':
				break;

			case 'number_of_guests':
				break;

			case 'events':
				break;

			case 'can_attend':
				//can_attend is the last line of the email that needs to be parsed
				//Break out of all loops or return
				// return;
				break message_loop;

		}

	}
}

/*
 * Returns an array of GmailThreads that match a provided label.
 * @param {string} label_name - Gmail Label name being searched for.
 */
function getGmailThreadsWithLabelName(label_name){

	var threads = [];
	var new_threads = [];
	var start = 0;
	var loop_count = 30;

	//First see if we can find the label
	var label = GmailApp.getUserLabelByName(label_name);

	do {
		new_threads = label.getThreads(start,loop_count);
		threads = threads.concat(new_threads);
		start += loop_count;
	}
	while(new_threads.length > 0);

	if(LOGGING_ENABLED){
		for (var i = 0; i < threads.length; i++) {
			// Logger.log(threads[i].getFirstMessageSubject());
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
String.prototype.replaceAll = function(find, replace){
	return this.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}

/*
 * Helper function that escapes a regular expression provided as a string
 * @param {string} regex - Regular expression string.
 */
function escapeRegExp(regex) {
    return regex.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}
 
