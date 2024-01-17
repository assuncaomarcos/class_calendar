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
 * Represents the months in a semester.
 *
 * @typedef {Object} MonthsInSemester
 * @property {Array<number>} W - The months in the Winter semester. Index 0 represents January.
 * @property {Array<number>} S - The months in the Spring semester.
 * @property {Array<number>} F - The months in the Fall semester.
 */
var MonthsInSemester = {
    "W" : [0, 3],
    "S" : [4, 7],
    "F" : [8, 11]
}

/**
 * Enumeration of date types.
 *
 * @typedef {Number} DateTypes
 *
 * @property {Number} CLASS - Represents a class date.
 * @property {Number} LAB - Represents a lab date.
 * @property {Number} QUIZ - Represents a quiz date.
 * @property {Number} EXAM - Represents an exam date.
 * @property {Number} HOLIDAY - Represents a holiday date.
 */
var DateTypes = {
    CLASS: 0,
    LAB: 1,
    QUIZ: 2,
    EXAM: 3,
    HOLIDAY: 4
}

/**
 * Creates the requests for rendering the weekdays labels in the calendar.
 *
 * @param {string} pageId - The ID of the page to make the requests for.
 * @returns {Array<*>} - The array of requests for rendering the weekdays labels.
 */
function getWeekdaysRequest(pageId) {
    var requests = [];
    var translateY = CalendarMargins.DAY_LABELS_TOP;
    var translateX = CalendarMargins.LEFT;
    for (var i=0; i < WeekDayNames.length; i++) {
        requests = requests.concat(getLabelRequest(pageId, WeekDayNames[i],
            WeekdayLabelScale.SCALE_X, WeekdayLabelScale.SCALE_Y,
            translateX, translateY, "CENTER"));
        translateX += WeekdayLabelScale.SCALE_X * DefaultMagnitude + CalendarMargins.COLUMN_GAP;
    }
    return requests;
}

/**
 * Retrieves the requests for creating a month box element in a specified slide.
 *
 * @param {string} pageId - The ID of the target slide.
 * @param {string} content - The content of the month box.
 * @param {number} scaleX - The horizontal scale of the month box.
 * @param {number} scaleY - The vertical scale of the month box.
 * @param {number} translateX - The horizontal translation of the month box.
 * @param {number} translateY - The vertical translation of the month box.
 *
 * @returns {Array} - An array of requests for creating a month box element.
 */
function getMonthBoxRequest(pageId, content, scaleX, scaleY, translateX, translateY) {
    const pageElementId = Utilities.getUuid();
    return [
        getCreateShapeRequest(pageId, pageElementId,
            CalendarElementsDefaultShape, scaleX, scaleY,
            translateX, translateY),
        getUpdateShapeRequest(pageElementId,
            MonthYearShapeStyle.BACKGROUND_FILL,
            MonthYearShapeStyle.OUTLINE,
            MonthYearShapeStyle.CONTENT_ALIGNMENT),
        getInsertTextRequest(pageElementId, content),
        getUpdateTextStyleRequest(pageElementId, MonthYearTextStyle.FONT_FAMILY, MonthYearTextStyle.FONT_SIZE),
        getUpdateParagraphStyle(pageElementId)
    ];
}

/**
 * Retrieves the request object for creating a day element of a calendar.
 *
 * @param {string} pageId - The ID of the page where the day element will be created.
 * @param {string} content - The content/text to be inserted in the day element.
 * @param {number} scaleX - The scale factor for the X-axis transformation of the day element.
 * @param {number} scaleY - The scale factor for the Y-axis transformation of the day element.
 * @param {number} translateX - The translation value for the X-axis transformation of the day element.
 * @param {number} translateY - The translation value for the Y-axis transformation of the day element.
 * @returns {object[]} - An array of requests for creating the day element on the page.
 */
function getDayElementRequest(pageId, content, scaleX, scaleY, translateX, translateY) {
    var pageElementId = Utilities.getUuid();
    var requests = [
        getCreateShapeRequest(pageId, pageElementId,
            CalendarElementsDefaultShape, scaleX, scaleY,
            translateX, translateY),
        getUpdateShapeRequest(pageElementId,
            DayShapeStyle.NORMAL.BACKGROUND_FILL,
            DayShapeStyle.NORMAL.OUTLINE,
            DayShapeStyle.NORMAL.CONTENT_ALIGNMENT),
        getInsertTextRequest(pageElementId, content),
        getUpdateTextStyleRequest(pageElementId, DayTextStyle.FONT_FAMILY, DayTextStyle.FONT_SIZE),
        getUpdateParagraphStyle(pageElementId)
    ];
    return { elementId: pageElementId, requests: requests };
}

/**
 * Creates requests for creating the day elements of a month in a calendar.
 *
 * @param {string} pageId - The ID of the slide where the calendar is located.
 * @param {number} year - The year of the month.
 * @param {number} month - The month (0-based index) for which to create the day elements.
 *
 * @returns {object} An object containing the requests and element IDs for the day elements of the month.
 *                   - requests: An array of requests to create the day elements.
 *                   - elementIds: An array of element IDs for the created day elements.
 */
function createMonthDaysRequest(pageId, year, month) {
    const DAYS_IN_WEEK = 7;
    const WEEKS_IN_MONTH = 6; // Maximum number of weeks in a month

    const scaleX = DayShapeScale.SCALE_X;
    const scaleY =  DayShapeScale.SCALE_Y;
    const left = CalendarMargins.LEFT;
    const shiftX = DayShapeScale.SCALE_X * DefaultMagnitude + CalendarMargins.COLUMN_GAP;
    const shiftY = DayShapeScale.SCALE_Y * DefaultMagnitude + CalendarMargins.LINE_GAP;
    const top = CalendarMargins.DAY_ELEMENTS_TOP; // Start from below day labels

    const startDate = new Date(year, month, 1);
    const firstDayOfWeek = startDate.getDay();
    const requests = [];
    var date = new Date(startDate);
    var dayElementIds = [];

    for (var week = 0; week < WEEKS_IN_MONTH; week++) {
        for (var dayOfWeek = 0; dayOfWeek < DAYS_IN_WEEK; dayOfWeek++) {
            // Starts drawing elements at the proper day-of-week column
            if ((week === 0 && dayOfWeek < firstDayOfWeek) || date.getUTCMonth() !== month) {
                continue;
            }

            var translateX = left + (shiftX * dayOfWeek);
            var translateY = top + (shiftY * week);
            var content = (date.getUTCDate() < 10 ? '0' : '') + date.getDate();
            var response = getDayElementRequest(pageId, content, scaleX, scaleY, translateX, translateY);
            requests.push(response.requests);
            dayElementIds.push(response.elementId);

            date.setDate(date.getDate() + 1);
            if (date.getMonth() !== month) {
                break;
            }
        }
    }
    return { requests: requests, elementIds: dayElementIds };
}

/**
 * Get the requests for the month and year boxes for a given slide.
 *
 * @param {string} slideId - The ID of the slide.
 * @param {number} year - The year of the boxes.
 * @param {number} month - The month of the boxes.
 * @returns {Array} - An array of requests for the month and year boxes.
 */
function getMonthYearRequest(slideId, year, month) {
    return getMonthBoxRequest(slideId, MonthNames[month],
        MonthBoxScale.SCALE_X, MonthBoxScale.SCALE_Y, CalendarMargins.LEFT, CalendarMargins.TOP)
        .concat(getMonthBoxRequest(slideId, year.toString(),
            YearBoxScale.SCALE_X, YearBoxScale.SCALE_Y, 4634893.87, CalendarMargins.TOP));
}

/**
 * Returns an array of update day shape requests.
 *
 * @param {Date[]} dates - An array of dates.
 * @param {string[]} dayElementIds - An array of day element IDs.
 * @param {object} dayShapeStyle - The style object for the day shape.
 * @param {string} dayShapeStyle.BACKGROUND_FILL - The background fill of the day shape.
 * @param {string} dayShapeStyle.OUTLINE - The outline of the day shape.
 * @param {string} dayShapeStyle.CONTENT_ALIGNMENT - The content alignment of the day shape.
 * @returns {object[]} - An array of update day shape requests.
 */
function getUpdateDayShapeRequest(dates, dayElementIds, dayShapeStyle) {
    var requests = [];
    for (var i = 0; i < dates.length; i++) {
        var date = dates[i];
        var dayIndex = date.getUTCDate() - 1;
        requests.push(
            getUpdateShapeRequest(dayElementIds[dayIndex],
                dayShapeStyle.BACKGROUND_FILL,
                dayShapeStyle.OUTLINE,
                dayShapeStyle.CONTENT_ALIGNMENT
            )
        );
    }
    return requests;
}

/**
 * Creates a legend item with given parameters.
 * @param {string} pageId - The ID of the slide where the legend item will be created.
 * @param {string} shapeType - The type of shape for the legend item.
 * @param {string} label - The label text for the legend item.
 * @param {number} marginLeft - The left margin for the legend item.
 * @param {number} marginTop - The top margin for the legend item.
 * @param {object} itemShapeStyle - The style settings for the shape of the legend item.
 * @param {object} itemShapeDimensions - The dimensions of the shape for the legend item.
 * @returns {Array} An array of requests for creating and updating the legend item in Google Slides API.
 */
function legendItem(pageId, shapeType, label, marginLeft, marginTop, itemShapeStyle, itemShapeDimensions) {
    var pageElementId = Utilities.getUuid();
    var labelMarginLeft = marginLeft + (itemShapeDimensions.SCALE_X + 0.05) * DefaultMagnitude;
    var labelMarginTop = marginTop - (LegendLabelScale.SCALE_Y * DefaultMagnitude * 0.25);

    // Deconstruction doesn't work when v8 not activated
    return [
        getCreateShapeRequest(pageId, pageElementId,
            shapeType, itemShapeDimensions.SCALE_X, itemShapeDimensions.SCALE_Y,
            marginLeft, marginTop),
        getUpdateShapeRequest(
            pageElementId,
            itemShapeStyle.BACKGROUND_FILL,
            itemShapeStyle.OUTLINE,
            itemShapeStyle.CONTENT_ALIGNMENT
        ),
        getLabelRequest(
            pageId, label,
            LegendLabelScale.SCALE_X, LegendLabelScale.SCALE_Y,
            labelMarginLeft, labelMarginTop, "START"
        )
    ];
}

/**
 * Creates an array of quiz element requests based on the provided parameters.
 *
 * @param {string} pageId - The ID of the slide.
 * @param {number} translateX - The X-axis translation value.
 * @param {number} translateY - The Y-axis translation value.
 * @returns {Array} An array of quiz element requests.
 */
function getQuizElementRequest(pageId, translateX, translateY) {
    var pageElementId = Utilities.getUuid();
    return requests = [
        getCreateShapeRequest(pageId, pageElementId,
            QuizShapeStyle.SHAPE, QuizShapeScale.SCALE_X, QuizShapeScale.SCALE_Y,
            translateX, translateY),
        getUpdateShapeRequest(pageElementId,
            QuizShapeStyle.BACKGROUND_FILL,
            QuizShapeStyle.OUTLINE)
    ];
}

/**
 * Creates the requests for displaying quizzes on a calendar page.
 *
 * @param {number} pageId - The ID of the calendar page.
 * @param {Date[]} quizDates - An array of quiz dates.
 * @returns {Object[]} An array of requests for displaying quiz elements on the calendar page.
 */
function getQuizzesRequest(pageId, quizDates) {
    if (!quizDates || quizDates.length === 0) {
        return [];
    }

    var firstDate = quizDates[0];
    var firstDayOfMonth = new Date(firstDate.getUTCFullYear(), firstDate.getUTCMonth(), 1);
    var firstDayOfWeek = firstDayOfMonth.getUTCDay();

    var baseShiftX = (DayShapeScale.SCALE_X - QuizShapeScale.SCALE_X - 0.05) * DefaultMagnitude;
    var baseShiftY = ((DayShapeScale.SCALE_Y - QuizShapeScale.SCALE_Y) / 2) * DefaultMagnitude;

    var requests = [];
    for (var i = 0; i < quizDates.length; i++) {
        date = quizDates[i];
        var offsetDate = date.getUTCDate() + firstDayOfWeek - 1;
        var weekOfMonth = Math.floor(offsetDate / 7);
        var dayOfWeek = date.getUTCDay();

        var dayOffsetX = dayOfWeek * (DayShapeScale.SCALE_X * DefaultMagnitude + CalendarMargins.COLUMN_GAP);
        var translateX = CalendarMargins.LEFT + baseShiftX + dayOffsetX;
        var weekOffsetY = weekOfMonth * (DayShapeScale.SCALE_Y * DefaultMagnitude + CalendarMargins.LINE_GAP);
        var translateY = CalendarMargins.DAY_ELEMENTS_TOP + baseShiftY + weekOffsetY;
        requests = requests.concat(getQuizElementRequest(pageId, translateX, translateY));
    }
    return requests;
}

/**
 * Filters an array of dates based on a given year and month.
 *
 * @param {Date[]} dates - An array of Date objects.
 * @param {number} year - The year to filter by.
 * @param {number} month - The month to filter by (0-11, where 0 represents January and 11 represents December).
 * @returns {Date[]} - An array of filtered dates. If the input array is null, it returns null.
 */
// array.filter() does not seem to work when v8 is deactivated
function filterDates(dates, year, month) {
    if (dates == null) {
        return null;
    }
    var newDates = [];
    for (var i = 0; i < dates.length; i++) {
        var date = dates[i];
        if (date.getUTCMonth() === month && date.getUTCFullYear() === year) {
            newDates.push(date);
        }
    }
    return newDates;
}

/**
 * Creates a month calendar slide on a Google Slide presentation.
 *
 * @param {string} presentationId - The ID of the Google Slide presentation.
 * @param {number} year - The year of the calendar.
 * @param {number} month - The month of the calendar.
 * @param {Array} classDates - An array of dates for class events.
 * @param {Array} labDates - An array of dates for lab events.
 * @param {Array} examDates - An array of dates for exam events.
 * @param {Array} holidays - An array of dates for holiday events.
 * @param {Array} quizDates - An array of dates for quiz events.
 *
 * @return {void}
 */
function createMonthCalendar(presentationId, year, month, classDates, labDates, examDates, holidays, quizDates) {
    var slideId = createSlide(presentationId, getLayout_(MasterLayout.CALENDAR).getObjectId());

    var requests = getMonthYearRequest(slideId, year, month);
    requests = requests.concat(getWeekdaysRequest(slideId));

    var moResponse = createMonthDaysRequest(slideId, year, month);
    requests = requests.concat(moResponse.requests);
    var dayElementIds = moResponse.elementIds;

    var dateMap = {};
    dateMap[DateTypes.CLASS] = {
        dates: filterDates(classDates, year, month),
        style: DayShapeStyle.CLASS
    };
    dateMap[DateTypes.LAB] = {
        dates: filterDates(labDates, year, month),
        style: DayShapeStyle.LAB
    };
    dateMap[DateTypes.EXAM] = {
        dates: filterDates(examDates, year, month),
        style: DayShapeStyle.EXAM
    };
    dateMap[DateTypes.HOLIDAY] = {
        dates: filterDates(holidays, year, month),
        style: DayShapeStyle.HOLIDAY
    };
    dateMap[DateTypes.QUIZ] = {
        dates: filterDates(quizDates, year, month)
    };

    var dateTypes = [DateTypes.CLASS, DateTypes.LAB, DateTypes.EXAM, DateTypes.HOLIDAY, DateTypes.QUIZ];
    var dateTypeStatus = {};

    for (var i = 0; i < dateTypes.length; i++) {
        var type = dateTypes[i];
        var dateType = dateMap[type];
        if (dateType.dates && dateType.dates.length > 0) {
            switch (type) {
                case DateTypes.CLASS:
                case DateTypes.LAB:
                case DateTypes.EXAM:
                case DateTypes.HOLIDAY:
                    requests = requests.concat(getUpdateDayShapeRequest(dateType.dates, dayElementIds, dateType.style));
                    dateTypeStatus[type] = true;
                    break;
                case DateTypes.QUIZ:
                    requests = requests.concat(getQuizzesRequest(slideId, dateType.dates));
                    dateTypeStatus[type] = true;
                    break;
            }
        }
    }

    requests = requests.concat(
        createLegendRequests(slideId, dateTypeStatus[DateTypes.CLASS],
            dateTypeStatus[DateTypes.LAB], dateTypeStatus[DateTypes.EXAM],
            dateTypeStatus[DateTypes.HOLIDAY], dateTypeStatus[DateTypes.QUIZ])
    );
    executeBatchUpdate(presentationId, requests);
}

/**
 * Creates an array of requests to generate legend items based on the specified parameters.
 *
 * @param {number} pageId - The ID of the page to create legend items for.
 * @param {boolean} hasClass - Indicates whether the legend should include a class item.
 * @param {boolean} hasLab - Indicates whether the legend should include a lab item.
 * @param {boolean} hasExam - Indicates whether the legend should include an exam item.
 * @param {boolean} hasHoliday - Indicates whether the legend should include a holiday item.
 * @param {boolean} hasQuiz - Indicates whether the legend should include a quiz item.
 * @returns {Array} - An array of requests to generate the legend items.
 */
function createLegendRequests(pageId, hasClass, hasLab, hasExam, hasHoliday, hasQuiz) {
    const legendSlotWidth = (LegendItemScale.SCALE_X + LegendLabelScale.SCALE_X) * DefaultMagnitude + CalendarMargins.COLUMN_GAP;
    const legendSlotHeight = (LegendLabelScale.SCALE_Y) * DefaultMagnitude + CalendarMargins.LEGEND_LINE_GAP;
    const slotsPerLine = 3;
    var translateX = 0;
    var translateY = 0;

    var slot = 0;
    const requests = [];
    const legendTypes = [
        { flag: hasClass, label: LegendLabel.CLASS, style: DayShapeStyle.CLASS },
        { flag: hasLab, label: LegendLabel.LAB, style: DayShapeStyle.LAB },
        { flag: hasExam, label: LegendLabel.EXAM, style: DayShapeStyle.EXAM },
        { flag: hasHoliday, label: LegendLabel.HOLIDAY, style: DayShapeStyle.HOLIDAY },
        { flag: hasQuiz, label: LegendLabel.QUIZ, style: QuizShapeStyle }
    ];

    for (var i = 0; i < legendTypes.length; i++) {
        var legendType = legendTypes[i];
        if (legendType.flag) {
            translateX = CalendarMargins.LEGEND_LEFT + ((slot % slotsPerLine) * legendSlotWidth);
            translateY = CalendarMargins.LEGEND_TOP + Math.floor(slot / slotsPerLine) * legendSlotHeight;
            if (legendType.label === LegendLabel.QUIZ) {
                var baseShiftX = (LegendItemScale.SCALE_X - QuizShapeScale.SCALE_X) * DefaultMagnitude;
                var baseShiftY = ((LegendItemScale.SCALE_Y - QuizShapeScale.SCALE_Y) / 2) * DefaultMagnitude;
                requests.push(legendItem(pageId, QuizShapeStyle.SHAPE, legendType.label, baseShiftX + translateX, translateY + baseShiftY, legendType.style, QuizShapeScale));
            } else {
                requests.push(legendItem(pageId, CalendarElementsDefaultShape, legendType.label, translateX, translateY, legendType.style, LegendItemScale));
            }
            slot++;
        }
    }

    return requests;
}