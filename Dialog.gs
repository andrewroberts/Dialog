// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 13 March 2015 18:45 GMT
// JSCS (Google, 2sp) - 13 March 2015 18:50 GMT

// Dialog.gs
// =========
//
// This library provides an easy way to open a custom dialog 
// (Ui.showModalDialog()). 
//
// Library Key: MWPmswuaTtvxxYA71VTxu7B8_L47d2MW6

/*
 * Copyright (C) 2015-2018 Andrew Roberts
 * 
 * This program is free software: you can redistribute it and/or modify it under
 * the terms of the GNU General Public License as published by the Free Software
 * Foundation, either version 3 of the License, or (at your option) any later 
 * version.
 * 
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License along with 
 * this program. If not, see http://www.gnu.org/licenses/.
 */

// Private Properties
// ------------------

// The function to use for logging. Let the script using the library 
// manage that, it'll need these methods
var Log_ = {
  init: function() {},
  finer: function() {},
  finest: function() {},
  info: function() {},
  warning: function() {},
};

// Public Methods
// --------------

var DIALOG_HEIGHT = 100;
var DIALOG_WIDTH = 400;

/**
 * Initialise the library.
 *
 * This mainly needs to be called to pass in a object to use for logging.
 *
 * @param {object} Logging library object
 */
 
function init(logLibrary) {
  if (logLibrary) Log_ = logLibrary;
}

/**
 * Alert or prompt the user. If no response is required from the user
 * this version uses a modal dialog rather than an Ui.alert() as 
 * long messages get truncated otherwise.
 *
 * The buttons parameter is optional, but if used the function
 * returns the response from the user.
 *
 * Will use log object if initialised via init().
 *
 * @param {string} title
 * @param {string} message
 * @param {number} height, defaults to ALERT_HEIGHT
 * @param {number} width, defaults to ALERT_WIDTH 
 * @param {Ui.ButtonSet} buttons, used if response needed
 *
 * @return {Ui.Button} response of button press if buttons arg
 *   defined, otherwise null 
 */

function show(title, message, height, width, buttons) {

  title = setDefault(title, '');
  message = setDefault(message, '');
  width = setDefault(width, DIALOG_WIDTH);  
  height = getHeight();
  buttons = setDefault(buttons, null);

  // TODO - Dynamically work out the dialog height.

  // TODO - Could be faster to pass the HTML as a string rather
  // than us template

  var template = HtmlService.createTemplateFromFile('DialogTemplate');
  template.message = message;
  
  var htmlOutput = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setHeight(height)
    .setWidth(width);
        
  var response = null;
  var ui;
  
  if (gotActiveDoc()) {
  
    ui = getUi();
  
    if (buttons) {

      Log_.info('Waiting for response from user');
      response = ui.alert(title, message, buttons);
      Log_.info('Got response: ' + response);
      
    } else {

      Log_.info('Displaying message to user');
      ui.showModalDialog(htmlOutput, title);
    }
    
  } else {
  
    Log_.warning('No UI available. title: ' + title + ', message: ' + message);
  }
  
  return response;

  // Private Functions
  // -----------------

  function getHeight() {
    if (height === undefined) {
      var numberOfChars = message.length
      var charsPerLine = width / 10
      height = ((numberOfChars / charsPerLine) + 2) * 22
    }
    return height
  }

  function gotActiveDoc() {    
    return DocumentApp.getActiveDocument() || SpreadsheetApp.getActive()
  }

  function getUi() {
  
    var ui
  
    try {
    
      ui = SpreadsheetApp.getUi();
      
    } catch (error) {
    
      try {
    
        ui = DocumentApp.getUi();
        
      } catch (error) {
      
        Log_.warning('Dialog: No UI available. title: ' + title + ', message: ' + message);
        return
      }
    }
    
    return ui
  }

  function setDefault(value, defaultValue) {  
    return (value === undefined) ? defaultValue : value;
  }
  
} // show()
