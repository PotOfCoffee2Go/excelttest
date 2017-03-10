"use strict";

var auth = {key: 'yourKey', token: 'yourToken'};

// Using an event emitter to throttle updates to trello
//   so does not over run the Trello rate limit
const EventEmitter = require('events');
class MyEmitter extends EventEmitter {}
const myEmitter = new MyEmitter();

var simplyTrello = require('simply-trello');

function sendToTrello(list, card, label) {
    label = label || 'green';
    var trello =  {
            path: {
                board: 'Seating Chart',
                list: list,
                card: card
            },
            content: {
            cardLabelColors: label,
            cardRemove: false
            }
        };

    simplyTrello.send (auth, trello, function(err, result) {
        if (err) { // Stop on error
            console.log(err.message);
        }
        else { // fire another request to trello
            myEmitter.emit('trelloEvent'); // Send next card to trello
        }
    });
}

// Card entries that we are going to add (or update existing) cards
var trelloEntries = [];

// Open the workbook/sheet with data used to create cards
var XLSX = require('xlsx');
var workbook = XLSX.readFile('./2016 O&B Guest List.xlsx');
var first_sheet_name = workbook.SheetNames[0];
/* Get worksheet */
var worksheet = workbook.Sheets[first_sheet_name];

// Walk thru the rows of the worksheet
for (var z in worksheet) {
    /* all keys that do not begin with "!" correspond to cell addresses */
    if(z[0] === '!') continue;

    // Data begins in row 4 of the sheet
    // If column C and row > 3
    if (z[0] === 'C' && Number(z.substr(1)) > 3 ) {
        // Name of the list (ie: table #) is in column 'B'
        var table = 'Table ' + worksheet[('B' + z.substr(1))].v;

        // Add first (column D) & last (column C) name of person to that table
        var person = worksheet[('D' + z.substr(1))].v + ' ' + worksheet[z].v;
        trelloEntries.push({list:table, card:person});
    }
}

// When we get a 'trelloEvent' add that list/card to the board
myEmitter.on('trelloEvent', () => {
    if (trelloEntries.length) { // there are people left to send to trello
        sendToTrello(trelloEntries[0].list, trelloEntries[0].card);
        trelloEntries.shift(); // done with this person - so remove from array
        console.log(trelloEntries.length); // display countdown
    }
});

// Start sending updates to Trello
myEmitter.emit('trelloEvent');
