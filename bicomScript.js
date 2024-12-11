/**
 * bicomScript.js - STEAMolus BiCom (Bi-Communcation Tool)
 * 
 * @copyright  Copyright (c) 2022 Jordi Campos Miralles (at STEAMolus)
 * @license    GNU General Public License v3.0
 * @author     Jordi Campos Miralles
 * @author     Gemma Garcia Caceres
 */

// Version
const BICOM_VERSION                 = "1.1.3-2024.12.11-12:45h"

// Sheet names
const README_SHEET_NAME             = "ReadMe"
const USR_SHEET_NAME                = "Usr" 
const COORD_SHEET_NAME              = "Coord"
const CFG_SHEET_NAME                = "Cfg"
const SHEETS_TO_KEEP_IN_USR_COPY    = [USR_SHEET_NAME, CFG_SHEET_NAME]

const USR_MASTER                    = "Master" // Name of the special user used by Coord

// Script inputs
const CFG_USERS_NAMES_1_5           = "B04"
const CFG_USERS_EMAILS_1_5          = "C04"
const CFG_USERS_URLS_1_5            = "D04"
const CFG_CLEAN_CONFIG_1_4          = "E04"
const CFG_CLEAN_META_1_4            = "F04"
const CFG_HIDE_META_1_4             = "G04"
const CFG_HIDE_TECH_1_4             = "H04"
const CFG_HIDE_TOP_1_4              = "I04"
const CFG_HIDE_BOTTOM_1_4           = "J04"
const CFG_TPL_USR_ROW_IN_COORD_1_4  = "K04"
const CFG_ROW_HIDDEN_TAGS_1_4       = "P04"
const BLOCKED_QUESTIONIDS1_4cfg     = "L04"
const BLOCKED_CHECKBOXES1_4cfg      = "M04"
const ANSWERS_RANGE1_4cfg           = "N04"


// Script outputs
const CFG_COORD_URL                 = "C05" // (this is also an input)
const CFG_PARTICULAR_NAME           = "C06"
const EDITION_STATUS_CELL           = "E06"
const EDITION_STATUS_LOCKED         = "Locked"
const EDITION_STATUS_CHANGING       = "Changing"
const EDITION_STATUS_UNLOCKED       = "Unlocked"



// User list in Coord's View
const NUM_ROWS_PER_USER             = 4
const ROW_HIDDEN_TAG                = "h"



// General stuff
const NOT_FOUND                     = -1
const NOERROR                       =  0
const ERROR_GENERIC                 =  1



// =======================================================================================
// MAIN MENU functions
// =======================================================================================

/**
* Main Menu
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi()

  // Check: if not Coord => EXIT
  var ss        = SpreadsheetApp.getActiveSpreadsheet()
  var sheetCfg  =  ss.getSheetByName(CFG_SHEET_NAME)
  if (sheetCfg.getRange(CFG_PARTICULAR_NAME).getValue() != USR_MASTER) return

  sheetCfg.getRange(CFG_COORD_URL).setValue( bicom_getSpreadsheetUrl() )

  // If Coord => create Menu
  ui.createMenu("BiCom")
      .addItem("1st Activate script", "menu_Activation")
      .addSeparator()
      .addSubMenu( ui.createMenu("a. Define CONTENTS (Usr tab)")
          .addItem("Clear ALL meta-info", "menu_ClearALLmeta")
          .addItem("Add meta rows BEFORE a certain one", "menu_AddMetaBEFORE")
          // FUTURE entries to add
          // .addItem("Add meta rows AFTER  a certain one", "menu_AddMetaAFTER")
          // .addItem("Delete meta rows BEFORE a certain one", "menu_DelMetaBEFORE")
          // .addItem("Delete meta rows AFTER  a certain one", "menu_DelMetaAFTER")
      )
      .addSubMenu( ui.createMenu("b. Manage USERS (Coord tab)")
        .addSubMenu( ui.createMenu("Add rows")
          .addItem("Add users BEFORE a certain one", "menu_AddUsersBEFORE")
          .addItem("Add users AFTER  a certain one", "menu_AddUsersAFTER")
        )
        .addSubMenu( ui.createMenu("Show/Hide rows (if desired)")
          .addItem("Show Mixt",    "menu_ShowMixtRowForALLusers")
          .addItem("Hide Mixt",    "menu_HideMixtRowForALLusers")
          .addItem("Show Coord",   "menu_ShowCoordRowForALLusers")
          .addItem("Hide Coord",   "menu_HideCoordRowForALLusers")
          .addItem("Show User",    "menu_ShowUserRowForALLusers")
          .addItem("Hide User",    "menu_HideUserRowForALLusers")
          .addItem("Hide InCoord", "menu_HideInCoordRowForALLusers")
          .addSubMenu( ui.createMenu("Collapse/Expand Groups")
            .addItem("Collapse SOME users", "menu_CollapseGroupsForSOMEusers")
            .addItem("Collapse ALL users",  "menu_CollapseGroupsForALLusers")
            .addItem("Expand SOME users",   "menu_ExpandGroupsForSOMEusers")
            .addItem("Expand ALL users",    "menu_ExpandGroupsForALLusers")
          )
        )
        .addSubMenu( ui.createMenu("Delete rows (and sheets)")
          .addItem("Delete SOME users!", "menu_DelSOMEusers")
          .addItem("Delete almost ALL users!", "menu_DelALLusers")
        )
      )
      .addSubMenu( ui.createMenu("c. Create FILES for users (Coord tab)")
        .addSubMenu( ui.createMenu("Create files")
          .addItem("Create files for SOME users", "menu_CreateSpreadsheetsForSOMEusers")
          .addItem("Create files for ALL users", "menu_CreateSpreadsheetsForALLusers")
        )
        .addSubMenu( ui.createMenu("Lock editing (optional before granting access)")
          .addItem("Lock edition", "menu_LockEdition")
          .addItem("UNlock edition", "menu_UnlockEdition")
        )
        .addSubMenu( ui.createMenu("Grant access (and send emails to users)")
          .addItem("Grant access to SOME users", "menu_GrantAccessForSOMEusers")
          .addItem("Grant access to ALL users", "menu_GrantAccessForALLusers")
        )
      )
      .addSubMenu( ui.createMenu("d. Manage INTERACTION (Coord tab)")
        .addSubMenu( ui.createMenu("Lock/Unlock any EDIT for all user")
          .addItem("Lock edition", "menu_LockEdition")
          .addItem("UNlock edition", "menu_UnlockEdition")
        )
        .addSubMenu( ui.createMenu("Lock/Unlock particular QUESTIONS for all users")
          .addItem("Lock a question", "menu_BlockQuestion")
          .addItem("UNlock a question", "menu_UnblockQuestion")
        )
      )
      .addSubMenu( ui.createMenu("e. Finish and ARCHIVE data (Coord tab)")
        .addItem("Lock edition", "menu_LockEdition")
        .addItem("Archive (to be done)", "to_be_done")
      )
      .addSeparator()
      .addItem("About STEAMolus BiCom", "menu_About")
      .addToUi()
}


/**
* Menu entry to promote a first execution that asks for script execution permission
*/
function menu_Activation() {
  // This function is empty, it simply to make that user accepts Script execution permissions
}



/**
 * Menu entry for Delete ALL Users
 */
function menu_ClearALLmeta()  {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "CLEARING ALL meta information from the Usr's view (all contents and questions will be lost)."+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling clearing"); return; }
 

  ss.toast("Clearing ALL meta-info...", "Clearing");


  // Create the spreadsheets
  var result = usrSheet_Meta_Clear()
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast("All meta cleared.", "Done")
}




/**
 * Menu entry for Add meta rows BEFORE a certain row in the Usr sheet's list
 */
function menu_AddMetaBEFORE()  {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

Browser.msgBox("TO-DO menu_AddMetaBEFORE")
return
  // Ask the row of the first user => refUsrRow_int
  var result = ui.prompt("Reference",
                         "Number of the row of the reference user (the one having its name):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling addition"); return; }
  var refUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(refUsrRow_int) ) { ss.toast("Not an integer number"); return }

  // Ask the row of the last user => numOfNewUsers
  var result = ui.prompt("Last", "Number of new users to add BEFORE:",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling addition"); return; }
  var numOfNewUsers = parseInt(result.getResponseText())
  if (!Number.isInteger(numOfNewUsers) ) { ss.toast("Not an integer number"); return }


  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "ADDING "+numOfNewUsers+" users to the Coord's list BEFORE row "+refUsrRow_int+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling addition"); return; }
 

  ss.toast("Adding "+numOfNewUsers+" users...", "Adding");


  // Adding rows
  var result = coordSheet_AddUsers(refUsrRow_int, numOfNewUsers, "before")
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast(numOfNewUsers+" users added. Wait some seconds to see the updates", "Done")
}






/**
 * Menu entry for Add user rows BEFORE a certain row in the Coord sheet's list
 */
function menu_AddUsersBEFORE()  {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Ask the row of the first user => refUsrRow_int
  var result = ui.prompt("Reference",
                         "Number of the row of the reference user (the one having its name):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling addition"); return; }
  var refUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(refUsrRow_int) ) { ss.toast("Not an integer number"); return }

  // Ask the row of the last user => numOfNewUsers
  var result = ui.prompt("Last", "Number of new users to add BEFORE:",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling addition"); return; }
  var numOfNewUsers = parseInt(result.getResponseText())
  if (!Number.isInteger(numOfNewUsers) ) { ss.toast("Not an integer number"); return }


  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "ADDING "+numOfNewUsers+" users to the Coord's list BEFORE row "+refUsrRow_int+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling addition"); return; }
 

  ss.toast("Adding "+numOfNewUsers+" users...", "Adding");


  // Adding rows
  var result = coordSheet_AddUsers(refUsrRow_int, numOfNewUsers, "before")
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast(numOfNewUsers+" users added. Wait some seconds to see the updates", "Done")
}


/**
 * Menu entry for Add user rows AFTER a certain row in the Coord sheet's list
 */
function menu_AddUsersAFTER()  {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Ask the row of the first user => refUsrRow_int
  var result = ui.prompt("Reference",
                         "Number of the row of the reference user (the one having its name):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling addition"); return; }
  var refUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(refUsrRow_int) ) { ss.toast("Not an integer number"); return }

  // Ask the row of the last user => numOfNewUsers
  var result = ui.prompt("Last", "Number of new users to add AFTER:",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling addition"); return; }
  var numOfNewUsers = parseInt(result.getResponseText())
  if (!Number.isInteger(numOfNewUsers) ) { ss.toast("Not an integer number"); return }


  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "ADDING "+numOfNewUsers+" users to the Coord's list AFTER row "+refUsrRow_int+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling addition"); return; }
 

  ss.toast("Adding "+numOfNewUsers+" users...", "Adding");


  // Adding rows
  var result = coordSheet_AddUsers(refUsrRow_int, numOfNewUsers, "after")
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast(numOfNewUsers+" users added. Wait some seconds to see the updates", "Done")
}





/**
 * Menu entry for Delete some Users
 */
function menu_DelSOMEusers()  {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Ask the row of the first user => firstUsrRow_int
  var result = ui.prompt("First", "Number of the row of the FIRST user (the one having its name):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling deletion"); return; }
  var firstUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(firstUsrRow_int) ) { ss.toast("Not an integer number"); return }

  // Ask the row of the last user => lastUsrRow_int
  var result = ui.prompt("Last", "Number of the row of the LAST user (the one having its name):"+
                          "\n\n(enter the same number than before to choose a single 1 user)",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling deletion"); return; }
  var lastUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(lastUsrRow_int) ) { ss.toast("Not an integer number"); return }


  // Confirm execution
  var totalSelectedUsers = (lastUsrRow_int+NUM_ROWS_PER_USER-firstUsrRow_int)/NUM_ROWS_PER_USER
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "DELETING "+totalSelectedUsers+" users (from the Coord's list and their Spreadsheets)."+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling deletion"); return; }
 

  ss.toast("Deleting "+totalSelectedUsers+" users...", "Deleting");


  // Create the spreadsheets
  var result = coordSheet_DelUsers(firstUsrRow_int, lastUsrRow_int)
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast(totalSelectedUsers+" users deleted.", "Done")
}


/**
 * Menu entry for Delete ALL Users
 */
function menu_DelALLusers()  {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "DELETING ALL users (from the Coord's list and their Spreadsheets)."+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling deletion"); return; }
 

  ss.toast("Deleting ALL users...", "Deleting");


  // Create the spreadsheets
  var result = coordSheet_DelUsers()
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast("All users deleted.", "Done")
}





/**
* Menu entry to show the MIXT row for ALL users
*/
function menu_ShowMixtRowForALLusers() {
  // Execute
  var result = coordSheet_ShowMixtRows()
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }
}


/**
* Menu entry to show the MIXT row for ALL users
*/
function menu_HideMixtRowForALLusers() {
  // Execute
  var result = coordSheet_HideMixtRows()
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }
}

/** TODO
menu_ShowCoordRowForALLusers
menu_HideCoordRowForALLusers
menu_ShowUserRowForALLusers
menu_HideUserRowForALLusers
menu_HideInCoordRowForALLusers
*/







/**
* Menu entry to collapse the groups of SOME users
*/
function menu_CollapseGroupsForSOMEusers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Ask the row of the first user => firstUsrRow_int
  var result = ui.prompt("First", "Number of the row of the FIRST user (the one having its name):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling collapsing groups"); return; }
  var firstUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(firstUsrRow_int) ) { ss.toast("Not an integer number"); return }

  // Ask the row of the last user => lastUsrRow_int
  var result = ui.prompt("Last", "Number of the row of the LAST user (the one having its name):"+
                          "\n\n(enter the same number than before to choose a single 1 user)",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling collapsing groups"); return; }
  var lastUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(lastUsrRow_int) ) { ss.toast("Not an integer number"); return }


  // Inform about execution
  var totalSelectedUsers = (lastUsrRow_int+NUM_ROWS_PER_USER-firstUsrRow_int)/NUM_ROWS_PER_USER

  // Execute
  var result = coordSheet_CollapseUserGroups(firstUsrRow_int, lastUsrRow_int)
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }
}



/**
* Menu entry to collapse the groups of ALL users
*/
function menu_CollapseGroupsForALLusers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Execute
  var result = coordSheet_CollapseUserGroups()
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }
}



/**
* Menu entry to expand the groups of SOME users
*/
function menu_ExpandGroupsForSOMEusers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Ask the row of the first user => firstUsrRow_int
  var result = ui.prompt("First", "Number of the row of the FIRST user (the one having its name):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling collapsing groups"); return; }
  var firstUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(firstUsrRow_int) ) { ss.toast("Not an integer number"); return }

  // Ask the row of the last user => lastUsrRow_int
  var result = ui.prompt("Last", "Number of the row of the LAST user (the one having its name):"+
                          "\n\n(enter the same number than before to choose a single 1 user)",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling collapsing groups"); return; }
  var lastUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(lastUsrRow_int) ) { ss.toast("Not an integer number"); return }


  // Inform about execution
  var totalSelectedUsers = (lastUsrRow_int+NUM_ROWS_PER_USER-firstUsrRow_int)/NUM_ROWS_PER_USER

  // Execute
  var result = coordSheet_ExpandUserGroups(firstUsrRow_int, lastUsrRow_int)
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }
}


/**
* Menu entry to expand the groups of ALL users
*/
function menu_ExpandGroupsForALLusers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Execute
  var result = coordSheet_ExpandUserGroups()
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }
}





/**
* Menu entry to create a separated Spreadsheet for SOME users
*/
function menu_CreateSpreadsheetsForSOMEusers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()


  // Ask the row of the first user => firstUsrRow_int
  var result = ui.prompt("First", "Number of the row of the FIRST user (the one having its name):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling creation"); return; }
  var firstUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(firstUsrRow_int) ) { ss.toast("Not an integer number"); return }

  // Ask the row of the last user => lastUsrRow_int
  var result = ui.prompt("Last", "Number of the row of the LAST user (the one having its name):"+
                          "\n\n(enter the same number than before to choose a single user)",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling creation"); return; }
  var lastUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(lastUsrRow_int) ) { ss.toast("Not an integer number"); return }


  // Default file related values
  var filePrefix               = ss.getName()+"-"
  var file     = DriveApp.getFileById(ss.getId());
  var folders  = file.getParents();
  var folderId = folders.next().getId()

  // Ask the filePrefix
  var result = ui.prompt("File prefix", "The user files will have this\n\n"+
    "default PREFIX '"+filePrefix+"'\n\n"+
    "or the one you enter here:\n\n"+
    "(leave blank to get the default value)",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling creation"); return; }
  if (result.getResponseText()) { filePrefix = result.getResponseText() }


  // Confirm execution
  var totalSelectedUsers = (lastUsrRow_int+NUM_ROWS_PER_USER-firstUsrRow_int)/NUM_ROWS_PER_USER
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "- Creating a separated Google Spreadsheet for "+totalSelectedUsers+" users.\n"+
    " with prefix '"+filePrefix+"'\n"+
    " in the same Drive folder than this file\n\n"+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling creation"); return; }
 
  ss.toast("Creating "+totalSelectedUsers+" separated spreadsheets...", "Creating");


  // Create the spreadsheets
  var result = coordSheet_CreateSpreadsheetsForUsers(folderId,filePrefix,firstUsrRow_int,lastUsrRow_int)
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast(totalSelectedUsers+" users' spreadsheets created.", "Done")
}




/**
* Menu entry to create a separated Spreadsheet for ALL users
*/
function menu_CreateSpreadsheetsForALLusers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Default file related values
  var filePrefix               = ss.getName()+"-"
  var file     = DriveApp.getFileById(ss.getId());
  var folders  = file.getParents();
  var folderId = folders.next().getId()

  // Ask the filePrefix
  var result = ui.prompt("File prefix", "The user files will have this\n\n"+
    "default PREFIX '"+filePrefix+"'\n\n"+
    "or the one you enter here:\n\n"+
    "(leave blank to get the default value)",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling creation"); return; }
  if (result.getResponseText()) { filePrefix = result.getResponseText() }

  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "- Creating a separated Google Spreadsheet for ALL users.\n"+
    " with prefix '"+filePrefix+"'\n"+
    " in the same Drive folder than this file\n\n"+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling creation"); return; }


  ss.toast("Creating separated spreadsheets", "Creating...");

  // Create the spreadsheets
  var result = coordSheet_CreateSpreadsheetsForUsers(folderId, filePrefix)
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast("All users' spreadsheets created.", "Done")
}



/**
* Menu entry to grant access to SOME users
*/
function menu_GrantAccessForSOMEusers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()


  // Ask the row of the first user => firstUsrRow_int
  var result = ui.prompt("First", "Number of the row of the FIRST user (the one having its name):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling granting"); return; }
  var firstUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(firstUsrRow_int) ) { ss.toast("Not an integer number"); return }

  // Ask the row of the last user => lastUsrRow_int
  var result = ui.prompt("Last", "Number of the row of the LAST user (the one having its name):"+
                          "\n\n(enter the same number than before to choose a single user)",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling granting"); return; }
  var lastUsrRow_int = parseInt(result.getResponseText())
  if (!Number.isInteger(lastUsrRow_int) ) { ss.toast("Not an integer number"); return }



  // Confirm execution
  var totalSelectedUsers = (lastUsrRow_int+NUM_ROWS_PER_USER-firstUsrRow_int)/NUM_ROWS_PER_USER
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "- Granting access to "+totalSelectedUsers+" users.\n"+
    "- including sending an email message sharing their individual document"+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling granting access"); return; }


  ss.toast("Granting access to ALL users", "Granting access...");

  // Grant access
  var result = coordSheet_GrantAccessForUsers(firstUsrRow_int,lastUsrRow_int)
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast("All users have access to their spreadsheets through a received mail.", "Done")
}





/**
* Menu entry to grant access to ALL users
*/
function menu_GrantAccessForALLusers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "- Granting access to ALL users.\n"+
    "- including sending an email message sharing their individual document"+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling granting access"); return; }


  ss.toast("Granting access to ALL users", "Granting access...");

  // Grant access
  var result = coordSheet_GrantAccessForUsers()
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast("All users have access to their spreadsheets through a received mail.", "Done")
}




/**
* Menu entry to lock edition for users (so they'll be able to see the sheets but they won't
* be able to edit their answers)
*/
function menu_LockEdition() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "- Locking edition for ALL users.\n"+
    "(so they'll be able to see the sheets but they won't be able to edit their answers)"+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling locking edition"); return; }


  ss.toast("Locking edition for ALL users", "Locking edition...");

  // Grant access
  var result = coordSheet_LockEditionForUsers()
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast("All users have access to their spreadsheets through a received mail.", "Done")
}


/**
* Menu entry to unlock edition for users (so they'll be able to edit their answers)
*/
function menu_UnlockEdition() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "- Unlocking edition for ALL users.\n"+
    "(so they'll be able to edit their answers)"+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling unlocking edition"); return; }


  ss.toast("Unlocking edition for ALL users", "Unlocking edition...");

  // Grant access
  var result = coordSheet_UnlockEditionForUsers()
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast("All users can edit their answers.", "Done")
}






/**
 * Menu entry to block a certain question updating all users' sheets
 */
function menu_BlockQuestion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Ask the column of the question => checkboxCol_int
  var result = ui.prompt("Question", "ID of question (the number below it in the 'Blocked' row):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling blocking"); return; }
  var questionId = parseInt(result.getResponseText())
  if (!Number.isInteger(questionId) ) { ss.toast("Not an integer number"); return }


  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "BLOCKING question ID("+questionId+") for ALL users"+
    " (each user sheet will be updated so they will no longer be able answer it)."+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling blocking"); return; }
 

  ss.toast("Blocking the question for ALL users...", "Blocking");

  // Mark the question as Blocked in the Coord Sheet
  var sheetCfg   = ss.getSheetByName(CFG_SHEET_NAME)
  var blockedCheckboxes2_4ref   = sheetCfg.getRange(BLOCKED_CHECKBOXES1_4cfg ).getValue()  
  var blockedCheckboxes3_4addr  = sheetCfg.getRange(blockedCheckboxes2_4ref  ).getValue()
  var sheetCoord = ss.getSheetByName(COORD_SHEET_NAME)
  var blockedCheckboxes4_4range = sheetCoord.getRange(blockedCheckboxes3_4addr )  
  var checkboxesRow = blockedCheckboxes4_4range.getRow()
  ss.getSheetByName(COORD_SHEET_NAME).getRange(checkboxesRow, questionId).check()

  // Blocking the Question
  var result = coordSheet_BlockQuestion(questionId)
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast("Question blocked for all users.", "Done")
}


/**
 * Menu entry to UNblock a certain question updating all users' sheets
 */
function menu_UnblockQuestion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // Ask the column of the question => checkboxCol_int
  var result = ui.prompt("Question", "ID of question (the number below it in the 'Blocked' row):",
                          ui.ButtonSet.OK);
  if (result.getSelectedButton() != ui.Button.OK) { ss.toast("Cancelling unblocking"); return; }
  var questionId = parseInt(result.getResponseText())
  if (!Number.isInteger(questionId) ) { ss.toast("Not an integer number"); return }


  // Confirm execution
  var result = ui.alert("Please confirm","Do you confirm:\n\n"+
    "UNblocking question ID("+questionId+") for ALL users"+
    " (each user sheet will be updated so they will no longer be able answer it)."+
    "\n\n?", ui.ButtonSet.YES_NO_CANCEL);
  if(result != ui.Button.YES) { ss.toast("Cancelling unblocking"); return; }
 

  ss.toast("Unblocking the question for ALL users...", "Unblocking");

  // Mark the question as UNblocked in the Coord Sheet
  var sheetCfg   = ss.getSheetByName(CFG_SHEET_NAME)
  var blockedCheckboxes2_4ref   = sheetCfg.getRange(BLOCKED_CHECKBOXES1_4cfg ).getValue()  
  var blockedCheckboxes3_4addr  = sheetCfg.getRange(blockedCheckboxes2_4ref  ).getValue()
  var sheetCoord = ss.getSheetByName(COORD_SHEET_NAME)
  var blockedCheckboxes4_4range = sheetCoord.getRange(blockedCheckboxes3_4addr )  
  var checkboxesRow = blockedCheckboxes4_4range.getRow()
  ss.getSheetByName(COORD_SHEET_NAME).getRange(checkboxesRow, questionId).uncheck()

  // Blocking the Question
  var result = coordSheet_UnblockQuestion(questionId)
  if (result != NOERROR) { Browser.msgBox("ERROR("+result+")") ; return }

  ss.toast("Question unblocked for all users.", "Done")
}



/**
 * About function (mainly to test running the script and to check its version)
 */
function menu_About() {
  var ui = SpreadsheetApp.getUi();

  Browser.msgBox("STEAMolus BiCom version=v."+BICOM_VERSION)
}

































// =======================================================================================
// 2n LEVEL functions
// =======================================================================================

/**
* Clear some meta-info (from Usr's tab)
*
* @param {Number} iniAbsoluteRow - [OPTIONAL] the initial meta row
* @param {Number} endAbsoluteRow - [OPTIONAL] the last meta row
* @return {Number} constNOERROR if success of the corresponding error code/text
*/
function usrSheet_Meta_Clear(iniAbsoluteRow=0, endAbsoluteRow=0) {

  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  var cleanMeta2_4ref     = sheetCfg.getRange(CFG_CLEAN_META_1_4  ).getValue()
  var cleanMeta3_4addr    = sheetCfg.getRange(cleanMeta2_4ref).getValue()

  // Locate meta range ==========================================================
  var sheetUsr            = ss.getSheetByName(USR_SHEET_NAME)
  
  var cleanMeta4_4range   = sheetUsr.getRange(cleanMeta3_4addr  )


  // Compute ini / end
  if (iniAbsoluteRow == 0) { // Default ini row = first in meta
    iniAbsoluteRow = cleanMeta4_4range.getRowIndex()+2
  }
  if (endAbsoluteRow == 0) { // Default end row = last in list
    endAbsoluteRow = cleanMeta4_4range.getRowIndex()+cleanMeta4_4range.getNumRows()-2
  }

  var iniRelativeRow = iniAbsoluteRow - (cleanMeta4_4range.getRowIndex()+2)
  var endRelativeRow = endAbsoluteRow - (cleanMeta4_4range.getRowIndex()+2)
  if (   (iniRelativeRow <0)
      || ( (endRelativeRow-iniRelativeRow) < 0 )
      || ( (endRelativeRow-iniRelativeRow) > cleanMeta4_4range.getNumRows() ) ) {
        return "Invalid row numbers iniRelativeRow("+iniRelativeRow+") "+
                "endRelativeRow("+endRelativeRow+") from "+
                "iniAbsoluteRow("+iniAbsoluteRow+") endAbsoluteRow("+endAbsoluteRow+") "
  }

  var rangeToClean = cleanMeta4_4range.offset( iniRelativeRow+2, 0, endRelativeRow-iniRelativeRow+1)

  ss.toast("Cleaning meta from row "+iniAbsoluteRow+" to row "+endAbsoluteRow+"...")

  rangeToClean.clearContent()

  return NOERROR
}










/**
 * Add user rows to the Coord sheet's list
 * @param {Number} refRow the reference user row (the one that contains its name).
 * @param {Number} numOfUsers number of new users.
 * @param {string} sense "before" or "after".
 * @return {Number} constNOERROR if success of the corresponding error code/text.
 */
function coordSheet_AddUsers(refRow, numOfUsers, sense)  {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  var usersNames2_5ref         = sheetCfg.getRange(CFG_USERS_NAMES_1_5).getValue()
  var tplUsrRowInCoord2_4ref   = sheetCfg.getRange(CFG_TPL_USR_ROW_IN_COORD_1_4 ).getValue()  
     

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // Get Template range
  var tplUsrRowInCoord3_4addr      = sheetCfg.getRange(tplUsrRowInCoord2_4ref   ).getValue()
  var tplUsrRowInCoord4_4range     = sheetCfg.getRange(tplUsrRowInCoord3_4addr )  
  var tplNumOfRows_int = tplUsrRowInCoord4_4range.getNumRows()

  numOfAddedRows = tplNumOfRows_int * numOfUsers

  // Get UsrNames to obtain the reference row
  var usersNames3_5addr      = sheetCfg.getRange(usersNames2_5ref).getValue()
  var usersNames4_5range     = sheetCoord.getRange(usersNames3_5addr)

  // Check valid refRow
  if (    (refRow < usersNames4_5range.getRow()) 
       || (refRow > usersNames4_5range.getLastRow()) ) {
    return "Row "+refRow+" is not in users list "+
           "( usersNamesIni="+usersNames4_5range.getRow()+
           ", usersNamesEnd="+usersNames4_5range.getLastRow()+")"
  }

  // Create rows
  var initialRow
  var finalRow

  if (sense == "before") {
    initialRow = refRow
    sheetCoord.insertRowsBefore(initialRow, numOfAddedRows)    
  } else if (sense == "after") {
    initialRow = refRow + tplNumOfRows_int
    sheetCoord.insertRowsAfter((initialRow-1), numOfAddedRows)
  } else { return "Unknown sense("+sense+")" }

  finalRow   = initialRow + numOfAddedRows-1

  // Set default height for all new rows
  sheetCoord.setRowHeightsForced(initialRow, numOfAddedRows, 21)

  // Remove grouping on the new rows
  var allAddedRowsRange = sheetCoord.getRange( initialRow, 1, numOfAddedRows)
  allAddedRowsRange.shiftRowGroupDepth(-1)

  // For each user: copy template, hide "InCoord" row and clean name
  for ( rowId = initialRow; rowId <= finalRow ; rowId+= tplNumOfRows_int) {
    ss.toast("Formatting new row "+rowId+"...", "Adding users")

    // Fill rows
    var destinationRange = sheetCoord.getRange( rowId, 1)
    tplUsrRowInCoord4_4range.copyTo( destinationRange,
                                     SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false )
    destinationRange.getCell(1,1).clearContent() // Clear name cell
    

    // Set last row (the "InCoord") height to 2 pixels to "hide" it (even if the group is opened)
    sheetCoord.setRowHeightsForced(rowId+(NUM_ROWS_PER_USER-1), 1, 2) // Alternative: sheetCoord.hideRows( rowId+NUM_ROWS_PER_USER-1 )

    // Group all rows belonging to a single user
    var rowsOneUserRange = sheetCoord.getRange( rowId+1, 1, NUM_ROWS_PER_USER-1)
    rowsOneUserRange.shiftRowGroupDepth(1)
  }

  return NOERROR
}



/**
 * Delete some users (from Coord's list and their separated spreadsheets)
 *
 * @param {Number} iniAbsoluteRow [OPTIONAL] the initial user row (the one that contains its name)
 * @param {Number} endAbsoluteRow [OPTIONAL] the last user row (the one that contains its name)
 * @return {Number} constNOERROR if success of the corresponding error code/text
 */
function coordSheet_DelUsers(iniAbsoluteRow=0, endAbsoluteRow=0) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  // Related to: user list -------------------------------
  var usersNames2_5ref         = sheetCfg.getRange(CFG_USERS_NAMES_1_5).getValue()
  var usersUrls2_5ref          = sheetCfg.getRange(CFG_USERS_URLS_1_5).getValue()

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // In order to: fetch data (references and values) ---------------------------
  var usersNames3_5addr      = sheetCfg.getRange(usersNames2_5ref).getValue()
  var usersNames4_5range     = sheetCoord.getRange(usersNames3_5addr)
  var usersNames5_5values    = usersNames4_5range.getValues()

  // In order to: write data (just references) ---------------------------------
  var usersUrls3_5addr       = sheetCfg.getRange(usersUrls2_5ref).getValue()
  var usersUrls4_5range      = sheetCoord.getRange(usersUrls3_5addr)

  // Compute row numbers relative to user s name list
  var usersNamesIniAbsoluteRow = usersNames4_5range.getRowIndex()

  // Check if we are going to delete all rows (we'll have to keep one)
  var deletingALLrows = false
  if (   ( iniAbsoluteRow == 0 && endAbsoluteRow == 0 )
      || ( (endAbsoluteRow-iniAbsoluteRow) == (usersNames4_5range.getNumRows()-NUM_ROWS_PER_USER) ) ) {
    deletingALLrows = true
  }

  // Set boundaries
  if (iniAbsoluteRow == 0) { // Default ini row = first in list
    iniAbsoluteRow = usersNamesIniAbsoluteRow
  }
  if (endAbsoluteRow == 0) { // Default end row = last in list
    endAbsoluteRow = usersNamesIniAbsoluteRow + usersNames4_5range.getNumRows()-NUM_ROWS_PER_USER
  }

  if (deletingALLrows) {
    iniAbsoluteRow += NUM_ROWS_PER_USER // We must leave 1 row (which we'll clear)
  }

  // Check other boundaries
  var iniRelativeRow = iniAbsoluteRow - usersNamesIniAbsoluteRow
  var endRelativeRow = endAbsoluteRow - usersNamesIniAbsoluteRow
  if (   (iniRelativeRow <0)
      || ((iniRelativeRow%NUM_ROWS_PER_USER) != 0)
      || ( (endRelativeRow-iniRelativeRow) < 0 )
      || ( (endRelativeRow-iniRelativeRow) > usersNames4_5range.getNumRows() )
      || ((endRelativeRow%NUM_ROWS_PER_USER) != 0) ) {
        return "Delete: Invalid row numbers iniRelativeRow("+iniRelativeRow+") "+
                "endRelativeRow("+endRelativeRow+") from "+
                "iniAbsoluteRow("+iniAbsoluteRow+") endAbsoluteRow("+endAbsoluteRow+") "+
                "usersNamesIniAbsoluteRow("+usersNamesIniAbsoluteRow+")"
  }

  // Iterate for each user ==========================================================
  for ( rowId = iniRelativeRow; rowId <= endRelativeRow; rowId+=NUM_ROWS_PER_USER) {
    
    // Get user data  ------------------------------------------
    var usrName     = usersNames5_5values[rowId][0]
    ss.toast("Processing user "+(((rowId-iniRelativeRow)+NUM_ROWS_PER_USER)/NUM_ROWS_PER_USER)+": "+usrName+"...")

    var cellURLRange = "R"+(usersNames4_5range.getRow()+rowId)+"C"+usersUrls4_5range.getColumn()
    var usrURL = sheetCoord.getRange(cellURLRange).getValue()

    // Trash file (if the user row has an associated file)
    if (usrURL!="") {
      // Get user s FILE --------------------------------------------
      var usrFileId = getIdFromSpreadsheetURL(usrURL)
      const file = DriveApp.getFileById(usrFileId)

      // Delete user s FILE ------------------------------------------
      file.setTrashed(true)
      sheetCoord.getRange(cellURLRange).clearContent()
    }
  }

  if (deletingALLrows) {
    coordSheet_AddUsers(usersNamesIniAbsoluteRow, 1, "after")
    coordSheet_DelUsers(usersNamesIniAbsoluteRow, usersNamesIniAbsoluteRow)
  }


  // Delete selected user's ROW -------------------------------------------
  ss.toast("Deleting rows in Coord's view...", "Deleting")
  sheetCoord.deleteRows(iniAbsoluteRow, (endAbsoluteRow+(NUM_ROWS_PER_USER-1)-iniAbsoluteRow)+1 )

  return NOERROR
}



/**
 * Hide MIXT row for some users in Coord tab
 *
 * @param {Number} iniAbsoluteRow [OPTIONAL] the initial user row (the one that contains its name)
 * @param {Number} endAbsoluteRow [OPTIONAL] the last user row (the one that contains its name)
 * @return {Number} constNOERROR if success of the corresponding error code/text
 */
function coordSheet_HideMixtRows(iniAbsoluteRow=0, endAbsoluteRow=0) {
JORDI
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  // Related to: user list -------------------------------
  var usersNames2_5ref         = sheetCfg.getRange(CFG_USERS_NAMES_1_5).getValue()
  var usersUrls2_5ref          = sheetCfg.getRange(CFG_USERS_URLS_1_5).getValue()
  var usersEmails2_5ref        = sheetCfg.getRange(CFG_USERS_EMAILS_1_5).getValue()

  // Related to: tidy up ---------------------------------
  var cleanConfig2_4ref        = sheetCfg.getRange(CFG_CLEAN_CONFIG_1_4).getValue()
  var cleanMeta2_4ref          = sheetCfg.getRange(CFG_CLEAN_META_1_4  ).getValue()
  var hideMeta2_4ref           = sheetCfg.getRange(CFG_HIDE_META_1_4   ).getValue()
  var hideTech2_4ref           = sheetCfg.getRange(CFG_HIDE_TECH_1_4   ).getValue()
  var hideTop2_4ref            = sheetCfg.getRange(CFG_HIDE_TOP_1_4    ).getValue()
  var hideBottom2_4ref         = sheetCfg.getRange(CFG_HIDE_BOTTOM_1_4 ).getValue()  
     

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // In order to: fetch data (references and values) ---------------------------
  var usersNames3_5addr      = sheetCfg.getRange(usersNames2_5ref).getValue()
  var usersNames4_5range     = sheetCoord.getRange(usersNames3_5addr)
  var usersNames5_5values    = usersNames4_5range.getValues()

  var usersEmails3_5addr     = sheetCfg.getRange(usersEmails2_5ref).getValue()
  var usersEmails4_5range    = sheetCoord.getRange(usersEmails3_5addr)
  var usersEmails5_5values   = usersEmails4_5range.getValues()
    

  // In order to: write data (just references) ---------------------------------
  var usersUrls3_5addr       = sheetCfg.getRange(usersUrls2_5ref).getValue()
  var usersUrls4_5range      = sheetCoord.getRange(usersUrls3_5addr)


  // Compute row numbers relative to user s name list
  var usersNamesIniAbsoluteRow = usersNames4_5range.getRowIndex()

  if (iniAbsoluteRow == 0) { // Default ini row = first in list
    iniAbsoluteRow = usersNamesIniAbsoluteRow
  }
  if (endAbsoluteRow == 0) { // Default end row = last in list
    endAbsoluteRow = usersNamesIniAbsoluteRow + usersNames4_5range.getNumRows()-NUM_ROWS_PER_USER
  }

  var iniRelativeRow = iniAbsoluteRow - usersNamesIniAbsoluteRow
  var endRelativeRow = endAbsoluteRow - usersNamesIniAbsoluteRow
  if (   (iniRelativeRow <0)
      || ((iniRelativeRow%NUM_ROWS_PER_USER) != 0)
      || ( (endRelativeRow-iniRelativeRow) < 0 )
      || ( (endRelativeRow-iniRelativeRow) > usersNames4_5range.getNumRows() )
      || ((endRelativeRow%NUM_ROWS_PER_USER) != 0) ) {
        return "Create: Invalid row numbers iniRelativeRow("+iniRelativeRow+") "+
                "endRelativeRow("+endRelativeRow+") from "+
                "iniAbsoluteRow("+iniAbsoluteRow+") endAbsoluteRow("+endAbsoluteRow+") "+
                "usersNamesIniAbsoluteRow("+usersNamesIniAbsoluteRow+")"
  }


  ss.toast("TODO...", "Hide MIXT rows")

  // template file: hide technical stuff
  const hideMeta3_4addr       = sheetCfg.getRange(hideMeta2_4ref   ).getValue()
  const hideMeta4_4range      = templateSheetUsr.getRange(hideMeta3_4addr  )
  templateSheetUsr.hideColumn(hideMeta4_4range) // hide Usr s clockwork


  // Iterate for each user ==========================================================
  for ( rowId = iniRelativeRow; rowId <= endRelativeRow; rowId+=NUM_ROWS_PER_USER) {
    
    // Get user data  ------------------------------------------
    var usrName     = usersNames5_5values[rowId][0]

    ss.toast("Hiding Mixt row for user "+(((rowId-iniRelativeRow)+NUM_ROWS_PER_USER)/NUM_ROWS_PER_USER)+": "+usrName+"...", "Hide MIXT rows")

    // Write hidden tag in User's Mixt row
    var cellURLRange = "R"+(usersNames4_5range.getRow()+rowId)+"C"+usersUrls4_5range.getColumn()
    sheetCoord.getRange(cellURLRange).setValue(ROW_HIDDEN_TAG)

  }

  return NOERROR
}



/**
 * Collapse groups for some users in Coord tab
 *
 * @param {Number} iniAbsoluteRow [OPTIONAL] the initial user row (the one that contains its name)
 * @param {Number} endAbsoluteRow [OPTIONAL] the last user row (the one that contains its name)
 * @return {Number} constNOERROR if success of the corresponding error code/text
 */
function coordSheet_CollapseUserGroups(iniAbsoluteRow=0, endAbsoluteRow=0) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  // Related to: user list -------------------------------
  var usersNames2_5ref         = sheetCfg.getRange(CFG_USERS_NAMES_1_5).getValue()
  var usersUrls2_5ref          = sheetCfg.getRange(CFG_USERS_URLS_1_5).getValue()

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // In order to: fetch data (references and values) ---------------------------
  var usersNames3_5addr      = sheetCfg.getRange(usersNames2_5ref).getValue()
  var usersNames4_5range     = sheetCoord.getRange(usersNames3_5addr)

  // Compute row numbers relative to user s name list
  var usersNamesIniAbsoluteRow = usersNames4_5range.getRowIndex()

  if (iniAbsoluteRow == 0) { // Default ini row = first in list
    iniAbsoluteRow = usersNamesIniAbsoluteRow
  }
  if (endAbsoluteRow == 0) { // Default end row = last in list
    endAbsoluteRow = usersNamesIniAbsoluteRow + usersNames4_5range.getNumRows()-NUM_ROWS_PER_USER
  }

  // Check other boundaries
  var iniRelativeRow = iniAbsoluteRow - usersNamesIniAbsoluteRow
  var endRelativeRow = endAbsoluteRow - usersNamesIniAbsoluteRow
  if (   (iniRelativeRow <0)
      || ((iniRelativeRow%NUM_ROWS_PER_USER) != 0)
      || ( (endRelativeRow-iniRelativeRow) < 0 )
      || ( (endRelativeRow-iniRelativeRow) > usersNames4_5range.getNumRows() )
      || ((endRelativeRow%NUM_ROWS_PER_USER) != 0) ) {
        return "Delete: Invalid row numbers iniRelativeRow("+iniRelativeRow+") "+
                "endRelativeRow("+endRelativeRow+") from "+
                "iniAbsoluteRow("+iniAbsoluteRow+") endAbsoluteRow("+endAbsoluteRow+") "+
                "usersNamesIniAbsoluteRow("+usersNamesIniAbsoluteRow+")"
  }

  // Locate range
  var targetUserRowsRange = sheetCoord.getRange( iniAbsoluteRow, 1, (endAbsoluteRow-iniAbsoluteRow)+NUM_ROWS_PER_USER )

  // Collapse range
  targetUserRowsRange.collapseGroups()

  return NOERROR
}



/**
 * Expand groups for some users in Coord tab
 *
 * @param {Number} iniAbsoluteRow [OPTIONAL] the initial user row (the one that contains its name)
 * @param {Number} endAbsoluteRow [OPTIONAL] the last user row (the one that contains its name)
 * @return {Number} constNOERROR if success of the corresponding error code/text
 */
function coordSheet_ExpandUserGroups(iniAbsoluteRow=0, endAbsoluteRow=0) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  // Related to: user list -------------------------------
  var usersNames2_5ref         = sheetCfg.getRange(CFG_USERS_NAMES_1_5).getValue()
  var usersUrls2_5ref          = sheetCfg.getRange(CFG_USERS_URLS_1_5).getValue()

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // In order to: fetch data (references and values) ---------------------------
  var usersNames3_5addr      = sheetCfg.getRange(usersNames2_5ref).getValue()
  var usersNames4_5range     = sheetCoord.getRange(usersNames3_5addr)

  // Compute row numbers relative to user s name list
  var usersNamesIniAbsoluteRow = usersNames4_5range.getRowIndex()

  if (iniAbsoluteRow == 0) { // Default ini row = first in list
    iniAbsoluteRow = usersNamesIniAbsoluteRow
  }
  if (endAbsoluteRow == 0) { // Default end row = last in list
    endAbsoluteRow = usersNamesIniAbsoluteRow + usersNames4_5range.getNumRows()-NUM_ROWS_PER_USER
  }

  // Check other boundaries
  var iniRelativeRow = iniAbsoluteRow - usersNamesIniAbsoluteRow
  var endRelativeRow = endAbsoluteRow - usersNamesIniAbsoluteRow
  if (   (iniRelativeRow <0)
      || ((iniRelativeRow%NUM_ROWS_PER_USER) != 0)
      || ( (endRelativeRow-iniRelativeRow) < 0 )
      || ( (endRelativeRow-iniRelativeRow) > usersNames4_5range.getNumRows() )
      || ((endRelativeRow%NUM_ROWS_PER_USER) != 0) ) {
        return "Delete: Invalid row numbers iniRelativeRow("+iniRelativeRow+") "+
                "endRelativeRow("+endRelativeRow+") from "+
                "iniAbsoluteRow("+iniAbsoluteRow+") endAbsoluteRow("+endAbsoluteRow+") "+
                "usersNamesIniAbsoluteRow("+usersNamesIniAbsoluteRow+")"
  }

  // Locate range
  var targetUserRowsRange = sheetCoord.getRange( iniAbsoluteRow, 1, (endAbsoluteRow-iniAbsoluteRow)+NUM_ROWS_PER_USER )

  // Expand range
  targetUserRowsRange.expandGroups()

  return NOERROR
}










/**
 * Create a separated Google Spreadsheet for some users (and sends them a sharing email)
 *
 * @param {Text}   folderId the destination Drive folder ID
 * @param {Text}   filePrefix the prefix used for each user (it will be followed by user's name)
 * @param {Number} iniAbsoluteRow [OPTIONAL] the initial user row (the one that contains its name)
 * @param {Number} endAbsoluteRow [OPTIONAL] the last user row (the one that contains its name)
 * @return {Number} constNOERROR if success of the corresponding error code/text
 */
function coordSheet_CreateSpreadsheetsForUsers(folderId, filePrefix, iniAbsoluteRow=0, endAbsoluteRow=0) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  // Related to: user list -------------------------------
  var usersNames2_5ref         = sheetCfg.getRange(CFG_USERS_NAMES_1_5).getValue()
  var usersUrls2_5ref          = sheetCfg.getRange(CFG_USERS_URLS_1_5).getValue()
  var usersEmails2_5ref        = sheetCfg.getRange(CFG_USERS_EMAILS_1_5).getValue()

  // Related to: tidy up ---------------------------------
  var cleanConfig2_4ref        = sheetCfg.getRange(CFG_CLEAN_CONFIG_1_4).getValue()
  var cleanMeta2_4ref          = sheetCfg.getRange(CFG_CLEAN_META_1_4  ).getValue()
  var hideMeta2_4ref           = sheetCfg.getRange(CFG_HIDE_META_1_4   ).getValue()
  var hideTech2_4ref           = sheetCfg.getRange(CFG_HIDE_TECH_1_4   ).getValue()
  var hideTop2_4ref            = sheetCfg.getRange(CFG_HIDE_TOP_1_4    ).getValue()
  var hideBottom2_4ref         = sheetCfg.getRange(CFG_HIDE_BOTTOM_1_4 ).getValue()  
     

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // In order to: fetch data (references and values) ---------------------------
  var usersNames3_5addr      = sheetCfg.getRange(usersNames2_5ref).getValue()
  var usersNames4_5range     = sheetCoord.getRange(usersNames3_5addr)
  var usersNames5_5values    = usersNames4_5range.getValues()

  var usersEmails3_5addr     = sheetCfg.getRange(usersEmails2_5ref).getValue()
  var usersEmails4_5range    = sheetCoord.getRange(usersEmails3_5addr)
  var usersEmails5_5values   = usersEmails4_5range.getValues()
    

  // In order to: write data (just references) ---------------------------------
  var usersUrls3_5addr       = sheetCfg.getRange(usersUrls2_5ref).getValue()
  var usersUrls4_5range      = sheetCoord.getRange(usersUrls3_5addr)


  // Locate files and folders ==========================================================
  const file = DriveApp.getFileById(ss.getId())
  var folder = DriveApp.getFolderById(folderId)


  // Compute row numbers relative to user s name list
  var usersNamesIniAbsoluteRow = usersNames4_5range.getRowIndex()

  if (iniAbsoluteRow == 0) { // Default ini row = first in list
    iniAbsoluteRow = usersNamesIniAbsoluteRow
  }
  if (endAbsoluteRow == 0) { // Default end row = last in list
    endAbsoluteRow = usersNamesIniAbsoluteRow + usersNames4_5range.getNumRows()-NUM_ROWS_PER_USER
  }

  var iniRelativeRow = iniAbsoluteRow - usersNamesIniAbsoluteRow
  var endRelativeRow = endAbsoluteRow - usersNamesIniAbsoluteRow
  if (   (iniRelativeRow <0)
      || ((iniRelativeRow%NUM_ROWS_PER_USER) != 0)
      || ( (endRelativeRow-iniRelativeRow) < 0 )
      || ( (endRelativeRow-iniRelativeRow) > usersNames4_5range.getNumRows() )
      || ((endRelativeRow%NUM_ROWS_PER_USER) != 0) ) {
        return "Create: Invalid row numbers iniRelativeRow("+iniRelativeRow+") "+
                "endRelativeRow("+endRelativeRow+") from "+
                "iniAbsoluteRow("+iniAbsoluteRow+") endAbsoluteRow("+endAbsoluteRow+") "+
                "usersNamesIniAbsoluteRow("+usersNamesIniAbsoluteRow+")"
  }


  // Create TEMPLATE file from current Coord file ===================================
  ss.toast("Creating template...", "Create User's files")

  // Duplicate file ------------------------------------------
  const templateFileName = filePrefix + "-Template"
  const templateFile = file.makeCopy(templateFileName, folder)
  
  /* TODO: check if this code needs to be implemented
    // Copy editor permissions to template sheet from Coord sheet
    // TODO: get emails from users returned by getEditors
    //      templateFile.addEditors( file.getEditors() )
    // template file: Grant permissions to the corresponding emails
    var emailsArray = [{}]
    emailsArray = usrEmail.split(",")
    emailsArray.forEach( function( email ) { templateFile.addEditor( email ) })
    // TODO: test call addEditor whith the array
    // templateFile.addEditor( emailsArray )
  */

  // template file: write basic data
  const templateSpreadsheet = SpreadsheetApp.open(templateFile)
  const templateSheetCfg    = templateSpreadsheet.getSheetByName(CFG_SHEET_NAME)
  const templateSheetUsr    = templateSpreadsheet.getSheetByName(USR_SHEET_NAME)

  // => write CoodUrl
  const coordUrl       = sheetCfg.getRange(CFG_COORD_URL).getValue()
  templateSheetCfg.getRange(CFG_COORD_URL).setValue(coordUrl) // into template file

  // template file: clean cells to let ImportsRange fill them
  // => clean Config second section (CFG)
  const cleanConfig3_4addr    = sheetCfg.getRange(cleanConfig2_4ref).getValue()
  const cleanConfig4_4range   = templateSheetCfg.getRange(cleanConfig3_4addr  ) // Cfg sheet!
  cleanConfig4_4range.clearContent()

  ss.toast("Cleaning template...", "Create User's files")
  // => clean MetaData  (USR)
  const cleanMeta3_4addr    = sheetCfg.getRange(cleanMeta2_4ref).getValue()
  const cleanMeta4_4range   = templateSheetUsr.getRange(cleanMeta3_4addr  ) // Usr sheet!
  cleanMeta4_4range.clearContent()

  // template file: hide technical stuff
  const hideMeta3_4addr       = sheetCfg.getRange(hideMeta2_4ref   ).getValue()
  const hideMeta4_4range      = templateSheetUsr.getRange(hideMeta3_4addr  )
  templateSheetUsr.hideColumn(hideMeta4_4range) // hide Usr s clockwork

  const hideTech3_4addr       = sheetCfg.getRange(hideTech2_4ref   ).getValue()
  const hideTech4_4range      = templateSheetUsr.getRange(hideTech3_4addr   )
  templateSheetUsr.hideColumn(hideTech4_4range) // hide Usr s right tech columns

  const hideTop3_4addr        = sheetCfg.getRange(hideTop2_4ref   ).getValue()
  const hideTop4_4range       = templateSheetUsr.getRange(hideTop3_4addr    )
  templateSheetUsr.hideRow(hideTop4_4range) // hide Usr s top tech rows

  const hideBottom3_4addr      = sheetCfg.getRange(hideBottom2_4ref   ).getValue()
  const hideBottom4_4range     = templateSheetUsr.getRange(hideBottom3_4addr )  
  templateSheetUsr.hideRow(hideBottom4_4range) // hide Usr s bottom tech rows

  // template file: remove undesired sheets
  templateSheetCfg.hideSheet() // hide Cfg sheet
  for (var idx=0; idx < templateSpreadsheet.getNumSheets() ;) { // remove undesired sheets
    var sheet = templateSpreadsheet.getSheets()[idx]
    
    if (SHEETS_TO_KEEP_IN_USR_COPY.indexOf(sheet.getName()) == NOT_FOUND ) {
      templateSpreadsheet.deleteSheet(sheet)
    } else {
      idx++
    }
  }  


  // Iterate for each user ==========================================================
  for ( rowId = iniRelativeRow; rowId <= endRelativeRow; rowId+=NUM_ROWS_PER_USER) {
    
    // Get user data  ------------------------------------------
    var usrName     = usersNames5_5values[rowId][0]
    var usrEmail    = usersEmails5_5values[rowId][0]

    if (usrName == "" || usrEmail =="") {
      return "Empty name or email on row "+(rowId + usersNamesIniAbsoluteRow)
    }
    ss.toast("Creating user "+(((rowId-iniRelativeRow)+NUM_ROWS_PER_USER)/NUM_ROWS_PER_USER)+": "+usrName+" file...", "Create User's files")

    // Duplicate FILE ------------------------------------------
    var newFileName = filePrefix + usrName
    
    if(!folder.getFilesByName(newFileName).hasNext()) {      
      // Duplicate file
      const newFile = templateFile.makeCopy(newFileName, folder)
      
      // NEW file: write usrName
      const newSpreadsheet = SpreadsheetApp.open(newFile)
      const newSheetCfg    = newSpreadsheet.getSheetByName(CFG_SHEET_NAME)
      newSheetCfg.getRange(CFG_PARTICULAR_NAME).setValue(usrName)

      // COORD file: storing URL pointing to new file
      var cellURLRange = "R"+(usersNames4_5range.getRow()+rowId)+"C"+usersUrls4_5range.getColumn()
      sheetCoord.getRange(cellURLRange).setValue(newFile.getUrl())
    } 

  }

  // Delete TEMPLATE file ------------------------------------------
  templateFile.setTrashed(true)

  return NOERROR
}



/**
 * Grants access to some users to their separated Google Spreadsheet & sends them a sharing email
 *
 * @param {Number} iniAbsoluteRow [OPTIONAL] the initial user row (the one that contains its name)
 * @param {Number} endAbsoluteRow [OPTIONAL] the last user row (the one that contains its name)
 * @return {Number} constNOERROR if success of the corresponding error code/text
 */
function coordSheet_GrantAccessForUsers(iniAbsoluteRow=0, endAbsoluteRow=0) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  // Related to: user list -------------------------------
  var usersNames2_5ref         = sheetCfg.getRange(CFG_USERS_NAMES_1_5).getValue()
  var usersUrls2_5ref          = sheetCfg.getRange(CFG_USERS_URLS_1_5).getValue()
  var usersEmails2_5ref        = sheetCfg.getRange(CFG_USERS_EMAILS_1_5).getValue()

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // In order to: fetch data (references and values) ---------------------------
  var usersNames3_5addr      = sheetCfg.getRange(usersNames2_5ref).getValue()
  var usersNames4_5range     = sheetCoord.getRange(usersNames3_5addr)
  var usersNames5_5values    = usersNames4_5range.getValues()

  var usersUrls3_5addr       = sheetCfg.getRange(usersUrls2_5ref).getValue()
  var usersUrls4_5range      = sheetCoord.getRange(usersUrls3_5addr)
  var usersUrls5_5values     = usersUrls4_5range.getValues()

  var usersEmails3_5addr     = sheetCfg.getRange(usersEmails2_5ref).getValue()
  var usersEmails4_5range    = sheetCoord.getRange(usersEmails3_5addr)
  var usersEmails5_5values   = usersEmails4_5range.getValues()
    

  // Compute row numbers relative to user s name list
  var usersNamesIniAbsoluteRow = usersNames4_5range.getRowIndex()

  if (iniAbsoluteRow == 0) { // Default ini row = first in list
    iniAbsoluteRow = usersNamesIniAbsoluteRow
  }
  if (endAbsoluteRow == 0) { // Default end row = last in list
    endAbsoluteRow = usersNamesIniAbsoluteRow + usersNames4_5range.getNumRows()-NUM_ROWS_PER_USER
  }

  var iniRelativeRow = iniAbsoluteRow - usersNamesIniAbsoluteRow
  var endRelativeRow = endAbsoluteRow - usersNamesIniAbsoluteRow
  if (   (iniRelativeRow <0)
      || ((iniRelativeRow%NUM_ROWS_PER_USER) != 0)
      || ( (endRelativeRow-iniRelativeRow) < 0 )
      || ( (endRelativeRow-iniRelativeRow) > usersNames4_5range.getNumRows() )
      || ((endRelativeRow%NUM_ROWS_PER_USER) != 0) ) {
        return "Grant access: Invalid row numbers iniRelativeRow("+iniRelativeRow+") "+
                "endRelativeRow("+endRelativeRow+") from "+
                "iniAbsoluteRow("+iniAbsoluteRow+") endAbsoluteRow("+endAbsoluteRow+") "+
                "usersNamesIniAbsoluteRow("+usersNamesIniAbsoluteRow+")"
  }

  // Iterate for each user ==========================================================
  for ( rowId = iniRelativeRow; rowId <= endRelativeRow; rowId+=NUM_ROWS_PER_USER) {
    
    // Get user data  ------------------------------------------
    var usrName     = usersNames5_5values[rowId][0]
    var usrUrl      = usersUrls5_5values[rowId][0]
    var usrEmail    = usersEmails5_5values[rowId][0]

    if (usrName == "" || usrUrl =="" || usrEmail =="") {
      return "Empty name or email or url on row "+(rowId + usersNamesIniAbsoluteRow)
    }
    ss.toast("Processing user "+(((rowId-iniRelativeRow)+NUM_ROWS_PER_USER)/NUM_ROWS_PER_USER)+": "+usrName+"...")

    // Duplicate FILE ------------------------------------------
    usrFile = DriveApp.getFileById(getIdFromUrl(usrUrl))

    if(!usrFile) {
      return "Cannot open url on row "+(rowId + usersNamesIniAbsoluteRow)
    }

    // Grant permissions to the corresponding user's emails
    var emailsArray = [{}]
    emailsArray = usrEmail.split(",")
    emailsArray.forEach( function( email ) { usrFile.addEditor( email ) })
    //TODO: use     usrFile.addEditors( emailsArray )
  }

  return NOERROR
}



/**
 * Lock edition to some users to their separated Google Spreadsheet 
 * (so they'll be able to see the sheets but they won't be able to edit their answers)
 *
 * @param {Number} iniAbsoluteRow [OPTIONAL] the initial user row (the one that contains its name)
 * @param {Number} endAbsoluteRow [OPTIONAL] the last user row (the one that contains its name)
 * @return {Number} constNOERROR if success of the corresponding error code/text
 */
function coordSheet_LockEditionForUsers(iniAbsoluteRow=0, endAbsoluteRow=0) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  // Related to: user list -------------------------------
  var usersNames2_5ref         = sheetCfg.getRange(CFG_USERS_NAMES_1_5).getValue()
  var usersUrls2_5ref          = sheetCfg.getRange(CFG_USERS_URLS_1_5).getValue()
  var usersEmails2_5ref        = sheetCfg.getRange(CFG_USERS_EMAILS_1_5).getValue()

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // In order to: fetch data (references and values) ---------------------------
  var usersNames3_5addr      = sheetCfg.getRange(usersNames2_5ref).getValue()
  var usersNames4_5range     = sheetCoord.getRange(usersNames3_5addr)
  var usersNames5_5values    = usersNames4_5range.getValues()

  var usersUrls3_5addr       = sheetCfg.getRange(usersUrls2_5ref).getValue()
  var usersUrls4_5range      = sheetCoord.getRange(usersUrls3_5addr)
  var usersUrls5_5values     = usersUrls4_5range.getValues()

  var usersEmails3_5addr     = sheetCfg.getRange(usersEmails2_5ref).getValue()
  var usersEmails4_5range    = sheetCoord.getRange(usersEmails3_5addr)
  var usersEmails5_5values   = usersEmails4_5range.getValues()
    

  // Compute row numbers relative to user s name list
  var usersNamesIniAbsoluteRow = usersNames4_5range.getRowIndex()

  if (iniAbsoluteRow == 0) { // Default ini row = first in list
    iniAbsoluteRow = usersNamesIniAbsoluteRow
  }
  if (endAbsoluteRow == 0) { // Default end row = last in list
    endAbsoluteRow = usersNamesIniAbsoluteRow + usersNames4_5range.getNumRows()-NUM_ROWS_PER_USER
  }

  var iniRelativeRow = iniAbsoluteRow - usersNamesIniAbsoluteRow
  var endRelativeRow = endAbsoluteRow - usersNamesIniAbsoluteRow
  if (   (iniRelativeRow <0)
      || ((iniRelativeRow%NUM_ROWS_PER_USER) != 0)
      || ( (endRelativeRow-iniRelativeRow) < 0 )
      || ( (endRelativeRow-iniRelativeRow) > usersNames4_5range.getNumRows() )
      || ((endRelativeRow%NUM_ROWS_PER_USER) != 0) ) {
        return "Lock: Invalid row numbers iniRelativeRow("+iniRelativeRow+") "+
                "endRelativeRow("+endRelativeRow+") from "+
                "iniAbsoluteRow("+iniAbsoluteRow+") endAbsoluteRow("+endAbsoluteRow+") "+
                "usersNamesIniAbsoluteRow("+usersNamesIniAbsoluteRow+")"
  }

  
  // Iterate for each user ==========================================================
  sheetCfg.getRange(EDITION_STATUS_CELL).setValue(EDITION_STATUS_CHANGING)
  for ( rowId = iniRelativeRow; rowId <= endRelativeRow; rowId+=NUM_ROWS_PER_USER) {
    
    // Get user data  ------------------------------------------
    var usrName     = usersNames5_5values[rowId][0]
    var usrUrl      = usersUrls5_5values[rowId][0]
    var usrEmail    = usersEmails5_5values[rowId][0]

    if (usrName == "" || usrUrl =="" || usrEmail =="") {
      return "Empty name or email or url on row "+(rowId + usersNamesIniAbsoluteRow)
    }
    ss.toast("Processing user "+(((rowId-iniRelativeRow)+NUM_ROWS_PER_USER)/NUM_ROWS_PER_USER)+": "+usrName+"...")

    // Duplicate FILE ------------------------------------------
    usrFile = DriveApp.getFileById(getIdFromUrl(usrUrl))

    if(!usrFile) {
      return "Cannot open url on row "+(rowId + usersNamesIniAbsoluteRow)
    }

    // Grant permissions to the corresponding user's emails
    var emailsArray = [{}]
    emailsArray = usrEmail.split(",")
    emailsArray.forEach( function( email ) { usrFile.removeEditor( email ) })
    emailsArray.forEach( function( email ) { usrFile.addViewer( email ) })
    //TODO: use     usrFile.addViewers( emailsArray )
  }

  sheetCfg.getRange(EDITION_STATUS_CELL).setValue(EDITION_STATUS_LOCKED)

  return NOERROR
}




/**
 * Unlock edition to some users to their separated Google Spreadsheet 
 * (so they'll be able to edit their answers)
 *
 * @param {Number} iniAbsoluteRow [OPTIONAL] the initial user row (the one that contains its name)
 * @param {Number} endAbsoluteRow [OPTIONAL] the last user row (the one that contains its name)
 * @return {Number} constNOERROR if success of the corresponding error code/text
 */
function coordSheet_UnlockEditionForUsers(iniAbsoluteRow=0, endAbsoluteRow=0) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  
  // Related to: user list -------------------------------
  var usersNames2_5ref         = sheetCfg.getRange(CFG_USERS_NAMES_1_5).getValue()
  var usersUrls2_5ref          = sheetCfg.getRange(CFG_USERS_URLS_1_5).getValue()
  var usersEmails2_5ref        = sheetCfg.getRange(CFG_USERS_EMAILS_1_5).getValue()

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // In order to: fetch data (references and values) ---------------------------
  var usersNames3_5addr      = sheetCfg.getRange(usersNames2_5ref).getValue()
  var usersNames4_5range     = sheetCoord.getRange(usersNames3_5addr)
  var usersNames5_5values    = usersNames4_5range.getValues()

  var usersUrls3_5addr       = sheetCfg.getRange(usersUrls2_5ref).getValue()
  var usersUrls4_5range      = sheetCoord.getRange(usersUrls3_5addr)
  var usersUrls5_5values     = usersUrls4_5range.getValues()

  var usersEmails3_5addr     = sheetCfg.getRange(usersEmails2_5ref).getValue()
  var usersEmails4_5range    = sheetCoord.getRange(usersEmails3_5addr)
  var usersEmails5_5values   = usersEmails4_5range.getValues()
    

  // Compute row numbers relative to user s name list
  var usersNamesIniAbsoluteRow = usersNames4_5range.getRowIndex()

  if (iniAbsoluteRow == 0) { // Default ini row = first in list
    iniAbsoluteRow = usersNamesIniAbsoluteRow
  }
  if (endAbsoluteRow == 0) { // Default end row = last in list
    endAbsoluteRow = usersNamesIniAbsoluteRow + usersNames4_5range.getNumRows()-NUM_ROWS_PER_USER
  }

  var iniRelativeRow = iniAbsoluteRow - usersNamesIniAbsoluteRow
  var endRelativeRow = endAbsoluteRow - usersNamesIniAbsoluteRow
  if (   (iniRelativeRow <0)
      || ((iniRelativeRow%NUM_ROWS_PER_USER) != 0)
      || ( (endRelativeRow-iniRelativeRow) < 0 )
      || ( (endRelativeRow-iniRelativeRow) > usersNames4_5range.getNumRows() )
      || ((endRelativeRow%NUM_ROWS_PER_USER) != 0) ) {
        return "Unlock: Invalid row numbers iniRelativeRow("+iniRelativeRow+") "+
                "endRelativeRow("+endRelativeRow+") from "+
                "iniAbsoluteRow("+iniAbsoluteRow+") endAbsoluteRow("+endAbsoluteRow+") "+
                "usersNamesIniAbsoluteRow("+usersNamesIniAbsoluteRow+")"
  }

  // Iterate for each user ==========================================================
  sheetCfg.getRange(EDITION_STATUS_CELL).setValue(EDITION_STATUS_CHANGING)
  for ( rowId = iniRelativeRow; rowId <= endRelativeRow; rowId+=NUM_ROWS_PER_USER) {
    
    // Get user data  ------------------------------------------
    var usrName     = usersNames5_5values[rowId][0]
    var usrUrl      = usersUrls5_5values[rowId][0]
    var usrEmail    = usersEmails5_5values[rowId][0]

    if (usrName == "" || usrUrl =="" || usrEmail =="") {
      return "Empty name or email or url on row "+(rowId + usersNamesIniAbsoluteRow)
    }
    ss.toast("Processing user "+(((rowId-iniRelativeRow)+NUM_ROWS_PER_USER)/NUM_ROWS_PER_USER)+": "+usrName+"...")

    // Duplicate FILE ------------------------------------------
    usrFile = DriveApp.getFileById(getIdFromUrl(usrUrl))

    if(!usrFile) {
      return "Cannot open url on row "+(rowId + usersNamesIniAbsoluteRow)
    }

    // Grant permissions to the corresponding user's emails
    var emailsArray = [{}]
    emailsArray = usrEmail.split(",")
    usrFile.addEditors( emailsArray )
// TODELETE:    emailsArray.forEach( function( email ) { usrFile.removeViewer( email ) })
  }

  sheetCfg.getRange(EDITION_STATUS_CELL).setValue(EDITION_STATUS_UNLOCKED)

  return NOERROR
}







/**
 * Blocking a certain question updating all users' sheets
 * 
 * @param {Number} checkboxCol the column number of the blocking check box
 */
function coordSheet_BlockQuestion(checkboxCol) {
  return coordSheet_BlockOrUnblockQuestion(checkboxCol, true)
}

/**
 * Unblocking a certain question updating all users' sheets
 * 
 * @param {Number} checkboxCol the column number of the blocking check box
 */
function coordSheet_UnblockQuestion(checkboxCol) {
  return coordSheet_BlockOrUnblockQuestion(checkboxCol, false)
}

/**
 * Blocking or unblocking a certain question updating all users' sheets
 * 
 * @param {Number} checkboxCol the column number of the blocking check box
 * @param {Number} block       true means blocking, false means unblocking
 */
function coordSheet_BlockOrUnblockQuestion(checkboxCol, block) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // Get the range to block
  var answerA1Not = getAnswerRowFromQuestionId(checkboxCol)

  // Load config values ==========================================================
  var sheetCfg                 = ss.getSheetByName(CFG_SHEET_NAME)
  var usersUrls2_5ref          = sheetCfg.getRange(CFG_USERS_URLS_1_5).getValue()

  // Load data from Coord Sheet ======================================================
  var sheetCoord    = ss.getSheetByName(COORD_SHEET_NAME);

  // In order to: write data (just references) ---------------------------------
  var usersUrls3_5addr       = sheetCfg.getRange(usersUrls2_5ref).getValue()
  var usersUrls4_5range      = sheetCoord.getRange(usersUrls3_5addr)
  var usersUrls5_5values     = usersUrls4_5range.getValues()


  // Iterate for each user ==========================================================
  var endRelativeRow = usersUrls4_5range.getNumRows()-NUM_ROWS_PER_USER
  for ( rowId = 0; rowId <= endRelativeRow; rowId+=NUM_ROWS_PER_USER) {
    
    // Get user data  ------------------------------------------
    var usrURL = usersUrls5_5values[rowId][0]
    if (usrURL == "") {
      ss.toast("Empty URL in row: "+(rowId + usersUrls4_5range.getRowIndex()))
    } else {
      ss.toast("Processing user "+((rowId+NUM_ROWS_PER_USER)/NUM_ROWS_PER_USER))


      // Get user s FILE --------------------------------------------
      var usrFileId = getIdFromSpreadsheetURL(usrURL)
      const usrFile = DriveApp.getFileById(usrFileId)

      // Get user s "Usr Sheet"
      var usrSpreadsheet = SpreadsheetApp.open(usrFile)
      var usrSheetUsr    = usrSpreadsheet.getSheetByName(USR_SHEET_NAME)

      if ( block ) {
        // Add protection ------------------------------------------
        var protection = usrSheetUsr.getRange(answerA1Not).protect()
                        .setDescription("B-"+answerA1Not)

        protection.removeEditors(protection.getEditors())
      } else {
        // Del protection ------------------------------------------
        var protections = usrSheetUsr.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      
        for (var i = 0; i < protections.length; i++) {
          var protection = protections[i]
        
          if (protection.getDescription() == "B-"+answerA1Not) {
            protection.remove()
          }
        }
      }
    }
  }

  return NOERROR
}










































// =======================================================================================
// 3d LEVEL functions
// =======================================================================================


/**
 * Get Google SpreadSheet s ID from its URL
 *
 * @param {string} url the spreadsheet's url
 * @return spreadsheet's ID
 */
function getIdFromSpreadsheetURL(url) {
   return SpreadsheetApp.openByUrl(url).getId();
}


/**
 * Computes the answer row A1Notation depending on the blocking checkbox selected
 * 
 * @param {Number} questionId the column number of the question
 * @param {Number} firstQuestionId_int [OPTIONAL] the initial checkbox col in CoordSheet
 * @param {Number} firstAnswerRow_int  [OPTIONAL] the initial answer row in UsrSheet
 */
function getAnswerRowFromQuestionId(questionId, firstQuestionId_int=0, firstAnswerRow_int=0) {

  // If some parameters are missing, let's get them from Config
  if ( firstQuestionId_int==0 || ! firstAnswerRow_int==0 ) {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var sheetCfg                   = ss.getSheetByName(CFG_SHEET_NAME)
    var sheetCoord                 = ss.getSheetByName(COORD_SHEET_NAME)

    // Load first checkbox col number from config
    if ( firstQuestionId_int==0 ) {
      var blockedQuestionIds2_4ref   = sheetCfg.getRange(BLOCKED_QUESTIONIDS1_4cfg).getValue()  
      var blockedQuestionIds3_4addr  = sheetCfg.getRange(blockedQuestionIds2_4ref   ).getValue()
      var blockedQuestionIds4_4range = sheetCoord.getRange(blockedQuestionIds3_4addr )  

      firstQuestionId_int = blockedQuestionIds4_4range.getColumn()
    }

    // Load first answer row number from config
    if ( firstAnswerRow_int==0 ) {
      var answersRange2_4ref   = sheetCfg.getRange(ANSWERS_RANGE1_4cfg ).getValue()  
      var answersRange3_4addr  = sheetCfg.getRange(answersRange2_4ref   ).getValue()
      var answersRange4_4range = sheetCoord.getRange(answersRange3_4addr )  

      firstAnswerRow_int = answersRange4_4range.getRow()
    }  
  }

  // Compute answer in A1 notation
  var answerRow   = (questionId - firstQuestionId_int)  + firstAnswerRow_int
  var answerA1Not = 'V'+answerRow+':W'+answerRow

  return answerA1Not
}









































// =======================================================================================
// Toolkit functions
// =======================================================================================

/**
 * Get Drive ID from a Drive item URL
 *
 * @param {Text} url the Drive item's URL
 * @return drive item's ID
 */
function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/)[0];
}


/**
 * Get the ChatGPT response to a given prompt
 *
 * Webography:
 * https://platform.openai.com/api-keys
 * https://www.freecodecamp.org/espanol/news/como-integrar-chatgpt-con-google-sheets-usando-google-apps-script/
 * https://rows.com/blog/post/chatgpt-google-sheets
 *
 * @param {Text} prompt for ChatGPT
 * @return ChatGPT response
 * 
*/
function promptChatGPT(prompt) {
  const apiKey = '<your-api-key>'; 
  const apiUrl = 'https://api.openai.com/v1/chat/completions';

  if (!prompt) {
    return "Error: Please provide a prompt.";
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${apiKey}`,
    },
    payload: JSON.stringify({
      model: "gpt-4o",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 1000,
    }),
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());
    const answer = jsonResponse.choices[0].message.content.trim();
    return answer;
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

































// =======================================================================================
// CUSTOM functions (available from Spreadsheet s cells)
// =======================================================================================


/**
 * Get current SpreadSheet s URL
 *
 * @return spreadsheet's url
 * @customfunction
 */
function bicom_getSpreadsheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}




/**
 * Get the ChatGPT response to a given prompt
 *
 * @param {Text} prompt for ChatGPT
 * @return ChatGPT response
 * @customfunction
 */
function bicom_GPT(prompt) {
  return promptChatGPT(prompt);
}































// =======================================================================================
// TO-DO functions (for further development)
// =======================================================================================


// TODO compare strings without diacritics

// LOWER(CLEAN(TRIM(AY22)))=LOWER(CLEAN(TRIM(AY$15)))

// TODO: have a look to https://www.labnol.org/replace-accented-characters-210709

// var ACCENTED = 'ÀÁÂÃÄÅàáâãäåÒÓÔÕÕÖØòóôõöøÈÉÊËèéêëðÇçÐÌÍÎÏìíîïÙÚÛÜùúûüÑñŠšŸÿýŽž';
// var REGULAR = 'AAAAAAaaaaaaOOOOOOOooooooEEEEeeeeeCcDIIIIiiiiUUUUuuuuNnSsYyyZz';
// var REGEXP = new RegExp('[' + ACCENTED + ']', 'g');

// function replaceDiacritics(str) {
//   function replace(match) {
//     var p = ACCENTED.indexOf(match);
//     return REGULAR[p];
//   }
//   return str.replace(REGEXP, replace);
// }




// /* DO NOT USE TO UPDATE USER'S SPREADSHEETS BECAUSE IT EXCEEDS THE TIME LIMIT

// // =======================================================================================
// // ON EDIT functions
// // =======================================================================================

// /**
//  * Trigger on edit to call other functions depending on the context
//  * It has a MAXIMUM execution time of 30 seconds!!!
//  * 
//  * @param {Event} e event handler (includes edited range...)
// */
// function onEdit(e) {
//   const range = e.range;

//   // Work ONLY in BLOCKING ROW
//   if (range.rowStart == CHECKBOXES_BLOCKING_ROW) {

//     if (range.isChecked()) {

//       onEditBlockQuestion(range.columnStart)

//     } else if (range.isChecked() == false) { 

//       onEditUnblockQuestion(range.columnStart)

//     } // else: it was not a checkbox (isChecked returns null instead of true or false)
//   }
// }





