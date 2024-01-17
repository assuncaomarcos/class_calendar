/**
 * @OnlyCurrentDoc Limits the script to only accessing the current presentation.
 */

/*
 * Copyright (c) 2023 Marcos Dias de Assuncao
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

/**
 * Represents the layout configuration of the master view. Adjust these
 * accordingly if you change the master slides.
 * This variable contains the following properties:
 *
 * - `TITLE`: Represents the position of the title slide in the used master view.
 * - `CALENDAR`: Represents the position of the calendar slide in the master view.
 * - `END`: Represents the position of the end calendar in the master view.
 * - `BLANK`: Represents a blank slide in the master view.
 *
 * @typedef {Object} MasterLayout
 * @property {number} TITLE - The position of the title component.
 * @property {number} CALENDAR - The position of the calendar component.
 * @property {number} END - The position of the end component.
 * @property {number} BLANK - A blank position.
 */
var MasterLayout = {
    TITLE: 1,
    CALENDAR: 0,
    END: 2,
    BLANK: 3
};

/**
 * Opens the sidebar when the script is run.
 *
 * @param {Event} event - The event object that triggered the function.
 *
 * @return {void}
 */
function onOpen(event) {
    SlidesApp.getUi().createAddonMenu()
        .addItem('Créer un calendrier', 'showSidebar')
        .addToUi();
}

/**
 * Executes when the add-on is installed.
 *
 * @param {Event} event - The installation event object.
 *
 * @return {void}
 */
function onInstall(event) {
    onOpen(event);
}

/**
 * Show the sidebar.
 *
 * @returns {void}
 */
function showSidebar() {
    const ui = HtmlService
        .createTemplateFromFile('Sidebar')
        .evaluate()
        .setTitle('Paramètres du calendrier');
    SlidesApp.getUi().showSidebar(ui);
}

/**
 * Retrieves the layout at the specified index from the active presentation's masters.
 *
 * @param {number} layoutIndex - The index of the layout to retrieve.
 * @returns {Layout} - The layout object at the specified index, or null if none exists.
 * @private
 */
function getLayout_(layoutIndex) {
    const pres = SlidesApp.getActivePresentation();
    const masters = pres.getMasters();
    if (masters.length) {
        // Always use the first master
        const layouts = masters.pop().getLayouts();
        return layouts.length > layoutIndex ? layouts[layoutIndex] : null;
    }
    return null;
}

/**
 * Executes a batch update on a Google Slides presentation.
 *
 * @param {string} presentationId - The ID of the presentation to update.
 * @param {object[]} requests - An array of update requests to apply to the presentation.
 */
function executeBatchUpdate(presentationId, requests) {
    try {
        Slides.Presentations.batchUpdate({'requests': requests}, presentationId);
    } catch (e) {
        Logger.log('Failed with error %s', e.message);
    }
}

/**
 * Creates a calendar in Google Slides based on the provided calendar information.
 *
 * @param {Object} calendarInfo - The calendar information object.
 * @param {string} calendarInfo.semester - The semester of the calendar.
 * @param {string} calendarInfo.year - The year of the calendar.
 * @param {Array} calendarInfo.classes - The array of dates for classes.
 * @param {Array} calendarInfo.labs - The array of dates for labs.
 * @param {Array} calendarInfo.exams - The array of dates for exams.
 * @param {Array} calendarInfo.quizzes - The array of dates for quizzes.
 * @param {Array} calendarInfo.holidays - The array of dates for holidays.
 *
 * @returns {boolean} - Returns true if the calendar was successfully created.
 */
function createCalendar(calendarInfo) {
    var pres = SlidesApp.getActivePresentation();
    var months = MonthsInSemester[calendarInfo.semester];
    var year = parseInt(calendarInfo.year);
    var classes = convertToDateArray(calendarInfo.classes);
    var labs = convertToDateArray(calendarInfo.labs);
    var exams = convertToDateArray(calendarInfo.exams);
    var quizzes = convertToDateArray(calendarInfo.quizzes);
    var holidays = convertToDateArray(calendarInfo.holidays);

    for (var month = months[0]; month <= months[1]; month++) {
        createMonthCalendar(pres.getId(), year, month, classes, labs, exams, holidays, quizzes);
    }

    return true;
}

function convertToDateArray(dateString) {
    var dateArray = dateString.split(', ');
    return dateArray.map(function(date) {
        return new Date(date.trim());
    });
}

/**
 * Includes another HTML file in the current HTML file.
 *
 * @param {string} filename - The name of the HTML file to be included.
 *
 * @returns {*} - The content of the included HTML file.
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ------------------- For debugging purposes -------------------

/**
 * Retrieves the details of an element on a specific slide of the active presentation.
 *
 * @returns {void}
 */
function elementDetails(slideIndex) {
    var pres = SlidesApp.getActivePresentation();
    var slide = pres.getSlides()[slideIndex];
    readPageElementIds(pres.getId(), slide.getObjectId());
}

/**
 * Retrieves the page element ids for a specific page in a presentation.
 *
 * @param {string} presentationId - The ID of the presentation.
 * @param {string} pageId - The ID of the page.
 * @returns {*} - The page element ids response.
 */
function readPageElementIds(presentationId, pageId) {
    try {
        const response = Slides.Presentations.Pages.get(presentationId, pageId);
        Logger.log(response);
        return response;
    } catch (e) {
        Logger.log('Failed with error %s', e.message);
    }
}

/**
 * Creates a test calendar for a specific year and month, with specified class dates,
 * lab dates, exam dates, holidays, and quiz dates.
 *
 * @returns {void}
 *
 * @example
 * createTestCalendar();
 */
function createTestCalendar() {
    const pres = SlidesApp.getActivePresentation();
    var classDates = [new Date("2024-01-10"), new Date("2024-01-17"), new Date("2024-01-24"), new Date("2024-01-31")];
    var labDates = [new Date("2024-01-11"), new Date("2024-01-18"), new Date("2024-01-25")];
    var holidays = [new Date("2024-01-01")];
    var examDates = [new Date("2024-01-26")];
    var quizDates = [new Date("2024-01-15"), new Date("2024-01-27")];
    createMonthCalendar(pres.getId(), 2024, 0, classDates, labDates, examDates, holidays, quizDates);
}