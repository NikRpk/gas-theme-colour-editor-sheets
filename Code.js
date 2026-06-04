// Set the default colour values. Currently set to the HelloFresh colour scheme
const DEFAULT_SETTINGS = {
  "text": "000000",
  "background": "ffffff",
  "accent1": "50c846",
  "accent2": "009646",
  "accent3": "ff5f64",
  "accent4": "d9d9d9",
  "accent5": "ff941a",
  "accent6": "1464ff",
  "hyperlink": "009646",
  "extra1": "FF941A",
  "extra2": "FFE900",
  "extra3": "FF63AA",
  "extra4": "FEF8F0",
  "extra5": "232323"
};

// Default number of extra colour slots shown on first open
const DEFAULT_EXTRA_COUNT = 5;

var PROPERTIES = ["text", "background", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink", "extra1", "extra2", "extra3", "extra4", "extra5"];

function onInstall(e) {
  onOpen(e);
};

function onOpen(e) {
  var ui = SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Set theme colours', 'setTheme')
    .addItem('Edit theme colours', 'showSidebar')
    .addItem('Reset to default', 'resetToDefault')
    .addToUi();
};


/* ---------------------------------------------------------- */
/* ------------------- MENU FUNCTIONS ----------------------- */
/* ---------------------------------------------------------- */


function showSidebar() {
  setUpProperties_();
  var html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Edit colours');
  SpreadsheetApp.getUi().showSidebar(html);
};

function setTheme() {
  setUpProperties_();
  var ui = SpreadsheetApp.getUi();

  var userProperties = PropertiesService.getUserProperties().getProperties();
  var extraHexValues = [];

  for (var key in userProperties) {
    if (key === 'extraCount') continue;
    if (key.startsWith('extra')) {
      extraHexValues.push(userProperties[key]);
    } else {
      setThemeColour_(key, userProperties[key]);
    }
  }

  // Inject all extras in a single temp-sheet operation
  if (extraHexValues.length > 0) {
    injectExtraColours_(extraHexValues);
  };

  ui.alert('Succesfully changed the theme colours!', ui.ButtonSet.OK);
};

function resetToDefault(sidebar) {
  var userProperties = PropertiesService.getUserProperties();
  var ui = SpreadsheetApp.getUi();

  var confirm = ui.alert("Confirm action", "Are you sure you want to reset your default colours?", ui.ButtonSet.YES_NO);

  if (confirm === ui.Button.NO) {
    return;
  };

  PROPERTIES.forEach(function(prop) {
    var defaultValue = DEFAULT_SETTINGS[prop];
    if (defaultValue) {
      userProperties.setProperty(prop, defaultValue);
      if (prop.startsWith('extra')) {
        // handled in bulk injection below
      } else {
        setThemeColour_(prop, defaultValue);
      }
    } else {
      userProperties.setProperty(prop, "ffffff");
    }
  });

  // Re-inject all default extra colours in one temp-sheet operation
  var defaultExtras = PROPERTIES
    .filter(function(p) { return p.startsWith('extra'); })
    .map(function(p) { return DEFAULT_SETTINGS[p]; })
    .filter(Boolean);
  if (defaultExtras.length > 0) {
    injectExtraColours_(defaultExtras);
  }

  userProperties.setProperty('extraCount', DEFAULT_EXTRA_COUNT.toString());

  if (sidebar) {
    showSidebar();
  }
};


/* ------------------------------------------------------------- */
/* ------------------- Sidebar FUNCTIONS ----------------------- */
/* ------------------------------------------------------------- */


function saveColour(dict) {
  var ui = SpreadsheetApp.getUi();
  var counter = 0;

  // Save the extra colour count so we restore the right number of rows next open
  if (dict.extraCount !== undefined) {
    PropertiesService.getUserProperties().setProperty('extraCount', dict.extraCount.toString());
    delete dict.extraCount;
  }

  var extraHexValues = [];

  for (var key in dict) {
    if (dict[key] === '') continue;

    if (key.startsWith('extra')) {
      extraHexValues.push(dict[key]);
      setColourProperty_(key, dict[key]);
    } else {
      setThemeColour_(key, dict[key]);
      setColourProperty_(key, dict[key]);
    }
    counter++;
  }

  // Inject all extras in a single temp-sheet operation
  if (extraHexValues.length > 0) {
    injectExtraColours_(extraHexValues);
  };

  return counter;
};

// Returns all saved extra colour values as a JSON string for use in the sidebar template
function getExtraColoursJson() {
  var userProperties = PropertiesService.getUserProperties().getProperties();
  var extras = {};
  for (var key in userProperties) {
    if (key.startsWith('extra') && key !== 'extraCount') {
      extras[key] = userProperties[key];
    }
  }
  return JSON.stringify(extras);
};

// Returns the number of extra colour rows that were last saved (minimum DEFAULT_EXTRA_COUNT)
function getExtraCount() {
  var saved = PropertiesService.getUserProperties().getProperty('extraCount');
  var count = saved ? parseInt(saved, 10) : DEFAULT_EXTRA_COUNT;
  return Math.max(count, DEFAULT_EXTRA_COUNT);
};


/* ------------------------------------------------------------- */
/* ------------------- SUPPORT FUNCTIONS ----------------------- */
/* ------------------------------------------------------------- */


function setUpProperties_() {
  var userProperties = PropertiesService.getUserProperties().getKeys();

  for (var property of PROPERTIES) {
    if (!userProperties.includes(property)) {
      var defaultPropertyValue = DEFAULT_SETTINGS[property];
      if (defaultPropertyValue) {
        setColourProperty_(property, defaultPropertyValue);
      };
    };
  };

  if (!userProperties.includes('extraCount')) {
    setColourProperty_('extraCount', DEFAULT_EXTRA_COUNT.toString());
  }
};

function setThemeColour_(dtype, colour) {
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var color = app.newColor().setRgbColor('#' + colour).build();

  var themeColorType = app.ThemeColorType[dtype.toUpperCase()];

  if (themeColorType) {
    ss.getSpreadsheetTheme().setConcreteColor(themeColorType, color);
  } else {
    Logger.log(`Error. Unknown input: ${dtype} for ${colour} colour.`);
  };
};

// Injects multiple colours into Sheets' "CUSTOM" palette in a single temp sheet operation.
// All colours are painted into one row of cells, flushed, then the sheet is deleted — one tab flicker total.
function injectExtraColours_(hexArray) {
  var colours = hexArray.filter(function(h) { return h && h.trim() !== ''; });
  if (colours.length === 0) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Temp_Colours_' + new Date().getTime();
  var tempSheet = ss.insertSheet(sheetName);

  colours.forEach(function(hex, i) {
    var fullHex = '#' + hex.replace(/^#/, '');
    tempSheet.getRange(1, i + 1).setBackground(fullHex);
  });

  SpreadsheetApp.flush();
  ss.deleteSheet(tempSheet);
};

function setColourProperty_(type, colour) {
  PropertiesService.getUserProperties().setProperty(type, colour);
};

function getColour_(type) {
  var property = PropertiesService.getUserProperties().getProperty(type);
  if (!property) {
    return "none";
  }
  return property;
};
