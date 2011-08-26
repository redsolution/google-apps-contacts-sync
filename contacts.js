/* * * * *
 * Ivan Gromov (c) 2011 Redsolution LLC
 *
 * * * * * * * * * * * * * * * */


/* *
 * Months to Enum convertor
 */
var months = {
    0: ContactsApp.Month.JANUARY,
    1: ContactsApp.Month.FEBRUARY,
    2: ContactsApp.Month.MARCH,
    3: ContactsApp.Month.APRIL,
    4: ContactsApp.Month.MAY,
    5: ContactsApp.Month.JUNE,
    6: ContactsApp.Month.JULY,
    7: ContactsApp.Month.AUGUST,
    8: ContactsApp.Month.SEPTEMBER,
    9: ContactsApp.Month.OCTOBER,
    10: ContactsApp.Month.NOVEMBER,
    11: ContactsApp.Month.DECEMBER
};

/* *
 * Column config. YOu can edit column numbers to your taste
 */
var COL = {
    LAST_NAME: 1,
    FIRST_NAME: 2,
    MIDDLE_NAME: 3,
    EMAIL: 7,
    W_PHONE: 4,
    C_PHONE: 5,
    H_PHONE: 6,
    BIRTH: 8,
    GROUPS: 9
};

/* *
 * Utility to trim whitespaces. Used in group names
 */
function trim(string) {
    return string.replace(/(^\s+)|(\s+$)/g, "");
}

/* *
 * Hightlight cell with incorrect or missing data with red
 */
function markError(row, col) {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(row, col, 1, 1).setBackgroundColor('#DB9378');
}

/* *
 * Clear highlighted errors. Highlight header and restore striped row colors.
 */
function clearErrors() {
    var sheet = SpreadsheetApp.getActiveSheet(),
        headerColor = '#333333',
        evenColor = '#FFFFFF',
        oddCOlor = '#E3E3E3';

    // set header color
    sheet.getRange(1, 1, 1, 25).setBackgroundColor(headerColor).setFontColor('#FFFFFF').setFontWeight('bold');

    for (var i = 2; i <= sheet.getLastRow(); i++) {
        if (i % 2) {
            // set even color
            sheet.getRange(i, 1, 1, 25).setBackgroundColor(evenColor);
        } else {
            // set odd
            sheet.getRange(i, 1, 1, 25).setBackgroundColor(oddCOlor);
        }
    }
}

/* *
 * If sheet has no data filled, this utility automatically creates
 * header fot table. You may rename columns later, they filled 
 * automatically only when 1st row is empty.
 * Script runs every time at synchronization.
 */
function initHeader() {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(1, COL.LAST_NAME, 1, 1).setValue(COL.LAST_NAME + '. Last name');
    sheet.getRange(1, COL.FIRST_NAME, 1, 1).setValue(COL.FIRST_NAME + '. First name');
    sheet.getRange(1, COL.MIDDLE_NAME, 1, 1).setValue(COL.MIDDLE_NAME + '. Middle name');
    sheet.getRange(1, COL.EMAIL, 1, 1).setValue(COL.EMAIL + '. Email');
    sheet.getRange(1, COL.W_PHONE, 1, 1).setValue(COL.W_PHONE + '. Work phone');
    sheet.getRange(1, COL.C_PHONE, 1, 1).setValue(COL.C_PHONE + '. Cell phone');
    sheet.getRange(1, COL.H_PHONE, 1, 1).setValue(COL.H_PHONE + '. Home phone');
    sheet.getRange(1, COL.BIRTH, 1, 1).setValue(COL.BIRTH + '. Birthday');
    sheet.getRange(1, COL.GROUPS, 1, 1).setValue(COL.GROUPS + '. Groups');

}

/* *
 * More complicated function to create or update phone nubmer
 *
 * @param contact - ContactApp.Contact instance
 * @param phone - String phone number
 * @param type - ContactsApp.Field type
 */
function updatePhone(contact, phone, type) {
    if (phone.length < 3) {
        return;
    }
    // Check whether user already has phone of this type
    if (contact.getPhones(type).length == 0) {
        contact.addPhone(type, phone);
        Logger.log('Set phone for: ' + contact.getEmails());
    } else {
        var existing = contact.getPhones(type),
            repeated = false;
        for (k = 0; k < existing.length; k++) {
            if (existing[k].getPhoneNumber() == phone) {
                repeated = true;
            }
        }
        if (!repeated) {
            var phoneField = contact.addPhone(type, phone);
            phoneField.setAsPrimary();
        }
    }
}

function syncQuiet() {
    try {
        sync(false);
    } catch (e) {
        // Do nothing in quiet mode
    }
}

function syncVerbose() {
    try {
        sync(true);
    } catch (e) {
        Browser.msgBox('Ошибка импорта: ' + e);
    }
}

/* *
 * Main synchronization utility. Does almost all job.
 *
 * @param verbose - Boolean verbose level. If false, no msgBox appears.
 */
function sync(verbose) {
    var sheet = SpreadsheetApp.getActiveSheet(),
        created = 0,
        updated = 0;

    if (sheet.getLastRow() == 0) {
        initHeader();
    }

    for (var i = 2; i <= sheet.getLastRow(); i++) {
        var contact = null;
        var data = {};
        // hightlight
        var prevColor = sheet.getRange(i, COL.LAST_NAME, 1, 1).getBackgroundColor();
        sheet.getRange(i, COL.LAST_NAME, 1, 1).setBackgroundColor('#FFFFBF');


        data.lastName = sheet.getRange(i, COL.LAST_NAME, 1, 1).getValue();
        data.firstName = sheet.getRange(i, COL.FIRST_NAME, 1, 1).getValue();
        data.midName = sheet.getRange(i, COL.MIDDLE_NAME, 1, 1).getValue();
        data.email = sheet.getRange(i, COL.EMAIL, 1, 1).getValue();
        data.cell_phone = sheet.getRange(i, COL.C_PHONE, 1, 1).getValue();
        data.work_phone = sheet.getRange(i, COL.W_PHONE, 1, 1).getValue();
        data.home_phone = sheet.getRange(i, COL.H_PHONE, 1, 1).getValue();
        data.birthday = sheet.getRange(i, COL.BIRTH, 1, 1).getValue();
        data.groups = sheet.getRange(i, COL.GROUPS, 1, 1).getValue().split(",");

        if (data.email) {
            if (ContactsApp.getContact(data.email)) {
                // Update
                contact = ContactsApp.getContact(data.email);
                Logger.log('\nUpdate contact: ' + data.email);
                updated++;
            } else {
                // Create
                contact = ContactsApp.createContact(data.firstName, data.lastName, data.email);
                Logger.log('\nCreate contact: ' + data.email);
                created++;
            }

            if (data.lastName) {
                contact.setFamilyName(data.lastName);
            } else {
                markError(i, COL.LAST_NAME);
            }
            if (data.firstName) {
                contact.setGivenName(data.firstName);
            } else {
                markError(i, COL.FIRST_NAME);
            }
            if ((data.lastName + data.firstName)) {
                contact.setShortName(data.firstName + ' ' + data.lastName);
                contact.setNickname(data.firstName + ' ' + data.lastName);
            }
            if (data.midName) {
                contact.setMiddleName(data.midName);
            } else {
                markError(i, COL.MIDDLE_NAME);
            }

            // Birthday workaround
            if (data.birthday instanceof Date) {
                if (contact.getDates(ContactsApp.Field.BIRTHDAY).length == 0) {
                    contact.addDate(ContactsApp.Field.BIRTHDAY, months[data.birthday.getMonth()], data.birthday.getDate() + 1, data.birthday.getYear());
                    Logger.log('Set birthday for: ' + data.email);
                }
/*else {
// Something went wrong here, I tired to messing around an "Unexpected exception upon serializing continuation"
var birthday = contact.getDates();
birthday.setDate(months[data.birthday.getMonth()],
data.birthday.getDate(),
data.birthday.getYear());
}*/
            } else {
                markError(i, COL.BIRTH);
            }

            // Phones workaround
            updatePhone(contact, data.cell_phone, ContactsApp.Field.MOBILE_PHONE);
            updatePhone(contact, data.work_phone, ContactsApp.Field.WORK_PHONE);
            updatePhone(contact, data.home_phone, ContactsApp.Field.HOME_PHONE);

            // Groups workaround
            for (var j = 0; j < data.groups.length; j++) {
                var group = null,
                    groupName = trim(data.groups[j]);

                if (!ContactsApp.getContactGroup(groupName)) {
                    // Create
                    group = ContactsApp.createContactGroup(groupName);
                    Logger.log('\nCreate group: ' + groupName);
                } else {
                    // Update
                    group = ContactsApp.getContactGroup(groupName);
                    Logger.log('\nUpdate group: ' + groupName);
                }
                contact.addToGroup(group);
            }
        } else {
            markError(i, COL.EMAIL);
        }

        // Unhighlight
        sheet.getRange(i, COL.LAST_NAME, 1, 1).setBackgroundColor(prevColor);
    }

    if (verbose) {
        // Browser.msgBox(Logger.getLog());
        Browser.msgBox('Контакты обновлены!\n ' + created + ' создано и ' + updated + 'обновлено');
    }
}

/* *
 * Register menu item and reset errors when sheet is opening.
 */
function onOpen() {
    var menuItems = [{
        name: "Синхронизировать",
        functionName: "syncVerbose"
    }];
    clearErrors();
    SpreadsheetApp.getActiveSpreadsheet().addMenu("Контакты", menuItems);
}
