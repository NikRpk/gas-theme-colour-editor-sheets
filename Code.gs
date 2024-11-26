// Set the default colour values. Currently set to the HelloFresh colour scheme
const DEFAULT_SETTINGS = {
  "text": "000000",
  "background": "ffffff",
  "accent1": "96dc14",
  "accent2": "009646",
  "accent3": "ff5f64",
  "accent4": "d9d9d9",
  "accent5": "00a0e6",
  "accent6": "1464ff",
  "hyperlink": "009646"
};

var PROPERTIES = ["text", "background", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink"];

function onInstall(e) {
  onOpen(e);
};

//So that the add-on runs on each open
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
  //Runs the sidebar and builds this from the html template. 

  setUpProperties_();
  var html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Edit colours');
  SpreadsheetApp.getUi().showSidebar(html);
};

function setTheme() {
  // Pulls the user properties as an input. It expects this as a dictionary with each colour type as the key and the colour codes (hexidecimal) as the values
  setUpProperties_();
  var ui = SpreadsheetApp.getUi();

  var userProperties = PropertiesService.getUserProperties().getProperties();
  for (var key in userProperties) {
    setThemeColour_(key, userProperties[key])
  };

  ui.alert('Succesfully changed the theme colours!',ui.ButtonSet.OK);
};

function resetToDefault(sidebar) {
  // Reset the user properties and the document theme to the default (defined by the script properties)
  var userProperties = PropertiesService.getUserProperties();
  var ui = SpreadsheetApp.getUi();

  var confirm = ui.alert("Confirm action", "Are you sure you want to reset your default colours?", ui.ButtonSet.YES_NO)

  if (confirm === ui.Button.NO) {
    return;
  };

  // Iterate over each property and reset it to the default value from script properties
  PROPERTIES.forEach(function(prop) {
    var defaultValue = DEFAULT_SETTINGS[prop];

    if (defaultValue) {
      userProperties.setProperty(prop, defaultValue);  // Reset user property to script property value
      setThemeColour_(prop, defaultValue);
    } else {
      // Set default to black if script property is not found
      userProperties.setProperty(prop, "ffffff");
    }
  });

  // Optionally alert the user that the properties have been reset
  //ui.alert("Update", 'All settings have been reset to default values.', ui.ButtonSet.OK);

  // Show sidebar if the reset is triggered from the sidebar 
  if (sidebar) {
    showSidebar();
  }
};


/* ------------------------------------------------------------- */
/* ------------------- Sidebar FUNCTIONS ----------------------- */
/* ------------------------------------------------------------- */


//Upon saving the colours in the sidebar, the colours are set as the theme colours and then the user properties are updated so that they can be accessed the next time as well.
function saveColour(dict) {
  var ui = SpreadsheetApp.getUi();
  var counter = 0;

  for (var key in dict) {
    if (dict[key] !== '') {  // Only process if there is a non-empty color value
      // Update theme colors
      setThemeColour_(key, dict[key]);
      // Update user properties
      setColourProperty_(key, dict[key]);
      counter += 1;
    };
  };

  // Display an alert to the user about the save success
  ui.alert("Update", 'Successfully saved ' + counter + ' colour(s) and changed the theme!', ui.ButtonSet.OK);
  showSidebar();
};


/* ------------------------------------------------------------- */
/* ------------------- SUPPORT FUNCTIONS ----------------------- */
/* ------------------------------------------------------------- */


//Checks if the relevant user properties have been created. If not, take the default ones from the script properties
function setUpProperties_() {
  var userProperties = PropertiesService.getUserProperties().getKeys();

  for (property of PROPERTIES) {
    // If the user property does not exist, set it to the corresponding value from default properties
    if (!userProperties.includes(property)) {
      var defaultPropertyValue = DEFAULT_SETTINGS[property];
      Logger.log(defaultPropertyValue)
      if (defaultPropertyValue) {
        setColourProperty_(property, defaultPropertyValue); // Create the user property and set it to the default property value
      };
    };
  };
};

//Sets the colour for each theme. It only updates the passed colour type e.g. 'accent2' with the hexidecimal colour
function setThemeColour_(dtype, colour) {
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var color = app.newColor().setRgbColor(colour).build();

  // Dynamically access the theme color type using the dtype
  var themeColorType = app.ThemeColorType[dtype.toUpperCase()];

  // Check if the color type exists (to avoid errors)
  if (themeColorType) {
    ss.getSpreadsheetTheme().setConcreteColor(themeColorType, color);
  } else {
    Logger.log(`Error. Unknown input: ${dtype} for ${colour} colour.`);
  };
};

//Sets the value for a certain user property (specified by the type and then hexideciaml code)
function setColourProperty_(type,colour) {
  sp = PropertiesService.getUserProperties();
  sp.setProperty(type, colour);
};

//Pulls the value of a certain property
function getColour_(type) {
  var property = PropertiesService.getUserProperties().getProperty(type);
  
  if(!property) {
    return "none"
  }
  return property;
};
