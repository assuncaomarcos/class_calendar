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
 * Represents the default magnitude in EMUs of all elements of a calendar.
 * All slide elements are scaled based on DefaultMagnitude.
 */
var DefaultMagnitude = 3000000;

/**
 * Represents the scale values for weekday labels.
 */
var WeekdayLabelScale = {
    SCALE_X: 0.85,
    SCALE_Y: 0.2766
}

/**
 * DayShapeScale represents the scale values for DayShape components.
 */
var DayShapeScale = {
    SCALE_X: 0.85,
    SCALE_Y: 0.2766
}

/**
 * Represents a scale factor for a quiz shape.
 */
var QuizShapeScale = {
    SCALE_X: 0.17,
    SCALE_Y: 0.17
}

/**
 * Represents a legend item scale.
 */
var LegendItemScale = {
    SCALE_X: 0.5774,
    SCALE_Y: 0.1548
}

/**
 * Represents the scale values for the legend labels.
 */
var LegendLabelScale = {
    SCALE_X: 0.9,
    SCALE_Y: 0.2
}

/**
 * Represents the scaling of the month box.
 */
var MonthBoxScale = {
    SCALE_X: 1.1099,
    SCALE_Y: 0.3376
}

/**
 * The YearBoxScale object represents the scale values for the year box.
 */
var YearBoxScale = {
    SCALE_X: 0.8817,
    SCALE_Y: 0.3376
}

/**
 * Object representing the margins for a calendar layout.
 * @property {number} TOP - The top margin value.
 * @property {number} LEFT - The left margin value.
 * @property {number} COLUMN_GAP - The gap between calendar columns.
 * @property {number} LINE_GAP - The gap between calendar lines.
 * @property {number} DAY_LABELS_TOP - The top position of day labels.
 * @property {number} DAY_ELEMENTS_TOP - The top position of day elements.
 * @property {number} LEGEND_LEFT - The left position of the legend.
 * @property {number} LEGEND_TOP - The top position of the legend.
 * @property {number} LEGEND_LINE_GAP - The gap between legend lines.
 */
var CalendarMargins = {
    TOP: 0.25 * DefaultMagnitude,
    LEFT: 0.25 * DefaultMagnitude,
    COLUMN_GAP: 0.16 * DefaultMagnitude,
    LINE_GAP: 0.16 * DefaultMagnitude,
    DAY_LABELS_TOP: 0.9 * DefaultMagnitude,
    DAY_ELEMENTS_TOP: 1.25 * DefaultMagnitude,
    LEGEND_LEFT: 1.03 * DefaultMagnitude,
    LEGEND_TOP: 3.84 * DefaultMagnitude,
    LEGEND_LINE_GAP: 0.05 * DefaultMagnitude
}
