/*
 *
 * Script that pulls data in from Gmail with a specific label.
 * All RSVP emails for my wedding are automatically filtered to this label.
 * Then the emails are parsed and appended to the spreadhseet if they aren't already in the sheet.
 *
 * Example Email:
 
From: John Doe & Jane Doe <jdoe123@gmail.com>
Subject: John Doe & Jane Doe Fitzgerald has RSVP

Name: John Doe & Jane Doe
Email: jdoe123@gmail.com
Number of Guest(s): 1
Events:
Can Attend: No

 *
 */


/*
 * Global Variables
 */
var LOGGING_ENABLED = true;

//Set a start date to filter out older emails
// var date = new Date('August 22, 2015 00:00:00');
// var START_TIME = date.getTime();
var START_DATE = new Date('August 22, 2015 00:00:00');
var MAX_THREADS_TO_PROCESS = 0;

//Set which label we are using to filter which emails are being retrieved
var RSVP_LABEL = 'Wedding RSVPs';
var SPREADSHEET_ID = '1UTJ1Bs33uF0nbEEHYGIYUL4cv5ftV_W1j-jdhadw9po';

var spreadsheet_field_names = [];
var sheet;

if(LOGGING_ENABLED){
	//Clear out previous logs
	Logger.clear();
	log("Starting script: Parse Wedding RSVPs");
}

/*
 * RSVP class that stores the details of the rsvp
 * Also has a toArray method that is useful for passing the data
 * into a Gmail Spreadsheet.
 */
function RSVP(){

	this.property_order = ['date', 'name', 'email', 'number_of_guests', 'events', 'can_attend'];

	this.toArray = function(){
		var arr = [];
		for(i = 0; i < this.property_order.length; i++){
			if(this.hasOwnProperty(this.property_order[i])){
				arr.push(this[this.property_order[i]]);
			}
		}
		return arr;
	}

}

/*
 * Loads the spreadsheet based on the predefined spreadsheet,
 * Then populates the list of header fields within the spreadsheet.
 * @returns - {boolean} - true if successful, false on failure.
 */
function initializeSpreadsheet(){

	var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

	if(!spreadsheet){
		return false;
	}

	sheet = spreadsheet.getSheets()[0];

	if(!sheet){
		return false;
	}

	spreadsheet_field_names = getSpreadsheetFieldNames(sheet);

	if(!spreadsheet_field_names || spreadsheet_field_names.length <= 0){
		return false;
	}

	return true;

}

/*
 * Function that retrieves a list of gmail email threads
 * with a provided label and parses them.
 * @param - {string} rsvp_label - The label applied to emails that should be parsed.
 */
function processRSVPs(rsvp_label) {

	var rsvp_threads = getGmailThreadsWithLabelName(rsvp_label);

	if(rsvp_threads.length <= 0){
		log("No Wedding RSVP Threads found");
		return false;
	}

	//We found some emails to parse, open the spreadsheet
	if(!initializeSpreadsheet()){
		log("Failed to initialize spreadsheet.");
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

		var rsvp = parseRSVPEmail(messages[i].getBody());
		rsvp.date = formatDate(messages[i].getDate());

		//Add rsvp entry to spreadsheet
		if(addRSVPToSpreadsheet(rsvp)) {
			log("RSVP Successfully added to spreadsheet: " + JSON.stringify(rsvp));
		}else{
			log("RSVP could not be added to spreadsheet: " + JSON.stringify(rsvp));
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

	//Convert the rsvp object into an array
	rsvp_array = rsvp.toArray();	

	if(RSVPEntryExists(rsvp, sheet)){
		log('Entry already found in spreadsheet: ' + JSON.stringify(rsvp));
		return false;
	}

	sheet = sheet.appendRow(rsvp_array);

	return true;

}

/*
 * Returns the list of field names from the first row of a spreadsheet.
 * @param {Sheet} sheet - Gmail Sheet object.
 * @param {string[]} - array of field names.
 */
function getSpreadsheetFieldNames(sheet){

	//Find the column index of the last row
	var last_column = sheet.getLastColumn();

	//Sheet.getRange(row, column, numRows, numColumns)
	//Get all of the values from the first row
	var range = sheet.getRange(1, 1, 1, last_column);
	var values = range.getValues();

	var header = [];
	for (var row in values) {
		for (var col in values[row]) {
			header_column_name = values[row][col].toLowerCase();
			header_column_name = replaceAll(' ', '_', header_column_name);
			header_column_name = replaceAll('[^a-z0-9_]', '', header_column_name);
			header.push(header_column_name);
		}
	}

	return header;

}

/*
 * Checks if the RSVP entry already exists in the Spreadsheet.
 * @param {RSVP} rsvp - RSVP object to be checked.
 * @param {Sheet} sheet - Gmail Sheet to be checked.
 * @returns {boolean} - true if the entry exists or if there was an error, otherwise false.
 */
function RSVPEntryExists(rsvp, sheet){

	//Find the column that we need to be searching
	search_column = spreadsheet_field_names.indexOf('name');

	if(search_column == -1){
		//Flag the entry as existing since there was an error.
		
		return true;
	}

	var last_row = sheet.getLastRow();

	// getRange(row, column, numRows)
	//Get all of the values from the search column for all rows
	//(Add 1 to the search_column as the Spreadsheet is 1-indexed instead of 0-indexed)
	//Start on row 2 as we don't want the header row
	var range = sheet.getRange(2, search_column + 1, last_row);

	var values = range.getValues();

	for (var row in values) {
		for (var col in values[row]) {
			if(values[row][col] == rsvp.name){
				return true;
			}
		}
	}

	return false;

}

/*
 * Filters an array of messages and returns a modified version of the original array
 * @param - {GmailMessage[]} messages - list of messages to be filtered.
 * @returns - {GmailMessage[]} - filtered list of messages.
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
	log('message time: ' + message.getDate().getTime());
	log('start time: ' + start_date.getTime());
	return (message.getDate().getTime() > start_date.getTime()) ? true : false;
}

/*
 * Parses an email of a specific format to retrieve its component parts
 * The email must be formatted with as: <string>: <string>, representing a key/value pair.
 * @param - {string} body - the contents of an rsvp email.
 * @returns - {Object} - rsvp details object.
 */
function parseRSVPEmail(body) {

	body = replaceAll('<br />', '', body);

	var message_lines = body.split("\n");

	//Initialize our RSVP Object
	rsvp = new RSVP();

	var skip_fields = ['from', 'subject']

	for(i = 0; i < message_lines.length; i++) {

		var line_parts = message_lines[i].split(':');

		//Make the line type lower case, replace spaces with underscores, and remove any other unexpected character
		// var line_type = line_parts[0].toLowerCase().replaceAll(' ', '_').replaceAll('/[^a-z0-9_]/', '');
		var line_type = line_parts[0].toLowerCase();
		line_type = replaceAll(' ', '_', line_type);
		line_type = replaceAll('[^a-z0-9_]', '', line_type);
		
		//Skip over fields we don't need
		if(skip_fields.indexOf(line_type) != -1){
			continue;
		}

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
	return subject.replace(new RegExp(find, 'g'), replace);
}

/*
 * Helper function that escapes a regular expression provided as a string
 * @param {string} regex - Regular expression string.
 * This escape does not work if you are trying to use metacharacters in your regular expression.
 */
function escapeRegExp(regex) {
	return regex.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}

/*
 * Helper function to format a javascript datetime value to be stored in a spreadsheet
 * @param {Date} date - Javascript date object to be formatted
 * @returns {string} - date formatted as "MM/DD/YYYY HH:MM:SS"
 */
function formatDate(date) {
	var formatted_date = (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear();
	var formatted_time = date.getHours() + ':' + date.getMinutes() + ':' + date.getSeconds();
	return formatted_date + ' ' + formatted_time;
}

/*
 * Custom logging function that checks if logging has been enabled
 * @param {String} message - message to be logged.
 */
function log(message){
	if(LOGGING_ENABLED){
		Logger.log(message);
	}
}

processRSVPs(RSVP_LABEL);