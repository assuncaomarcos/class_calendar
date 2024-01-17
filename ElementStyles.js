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
 * Represents the default shape for calendar elements.
 */
var CalendarElementsDefaultShape = "FLOW_CHART_TERMINATOR";

/**
 * Object representing the text style for a day element.
 */
var DayTextStyle = {
    FONT_FAMILY: "Oswald",
    FONT_SIZE: 37
}

/**
 * Represents the text style for displaying the month and year.
 */
var MonthYearTextStyle = {
    FONT_FAMILY: "Oswald",
    FONT_SIZE: 48
}

/**
 * Represents the different types of label text properties.
 *
 * @typedef {Object} LabelTextType
 * @property {string} FONT_FAMILY - The font family for the label text.
 * @property {number} FONT_SIZE - The font size for the label text.
 */
var LabelTextType = {
    FONT_FAMILY: "Oswald",
    FONT_SIZE: 37
}

/**
 * @typedef {Object} QuizShapeStyle - Represents the style configuration for a quiz shape.
 *
 * @property {string} SHAPE - The shape of the quiz.
 * @property {Object} BACKGROUND_FILL - The background fill of the quiz shape.
 * @property {Object} OUTLINE - The outline of the quiz shape.
 */
var QuizShapeStyle = {
    SHAPE: "ELLIPSE",
    BACKGROUND_FILL: {
        "solidFill": {
            "color": {
                "rgbColor": {"red": 0.84705883, "green": 0.13725491, "blue": 0.2}
            },
            "alpha": 1
        }
    },
    OUTLINE: {
        "outlineFill": {
            "solidFill": {
                "color": {"rgbColor": {"red": 0, "green": 0, "blue": 0}},
                "alpha": 1
            }
        },
        "weight": {"magnitude": 28575, "unit": "EMU"},
        "dashStyle": "SOLID"
    },
    CONTENT_ALIGNMENT: "MIDDLE"
}

/**
 * Represents the style configuration for a month-year shape.
 * @type {object}
 * @property {object} BACKGROUND_FILL - The background fill properties.
 * @property {object} OUTLINE - The outline properties.
 * @property {string} CONTENT_ALIGNMENT - The content alignment for the shape. Possible values: "MIDDLE".
 */
var MonthYearShapeStyle = {
    BACKGROUND_FILL: {
        "propertyState": "NOT_RENDERED"
    },
    OUTLINE: {
        "outlineFill": {
            "solidFill": {
                "color": {"rgbColor": {"red": 0.4, "green": 0.4, "blue": 0.4}},
                "alpha": 1
            }
        },
        "weight": {"magnitude": 38100, "unit": "EMU"},
        "dashStyle": "SOLID"
    },
    CONTENT_ALIGNMENT: "MIDDLE"
}

/**
 * Represents the visual style of a day shape in a calendar.
 * Each property of the DayShapeStyle object represents a different aspect of the style.
 *
 * @typedef {Object} DayShapeStyle
 *
 * @property {Object} HOLIDAY - The style for a holiday.
 * @property {Object} HOLIDAY.BACKGROUND_FILL - The background fill style for a holiday.
 * @property {Object} HOLIDAY.OUTLINE - The outline style for a holiday.
 * @property {string} HOLIDAY.CONTENT_ALIGNMENT - The content alignment for a holiday.
 * @property {Object} HOLIDAY.TEXT_COLOR - The text color style for a holiday.
 *
 * @property {Object} CLASS - The style for a class.
 * (Similar structure as the HOLIDAY style)
 *
 * @property {Object} LAB - The style for a lab.
 * (Similar structure as the HOLIDAY style)
 *
 * @property {Object} EXAM - The style for an exam.
 * (Similar structure as the HOLIDAY style)
 *
 * @property {Object} NORMAL - The default style for a normal day.
 * (Similar structure as the HOLIDAY style)
 */
var DayShapeStyle= {
    HOLIDAY: {
        BACKGROUND_FILL: {
            "solidFill": {
                "color": {
                    "rgbColor": {"red": 0.74509805, "green": 0.8784314, "blue": 0.6901961}
                },
                "alpha": 1
            }
        },
        OUTLINE: {
            "outlineFill": {
                "solidFill": {
                    "color": {"rgbColor": {"red": 0, "green": 0, "blue": 0}},
                    "alpha": 1
                }
            },
            "weight": {"magnitude": 38100, "unit": "EMU"},
            "dashStyle": "LONG_DASH_DOT"
        },
        CONTENT_ALIGNMENT: "MIDDLE",
        TEXT_COLOR: {"rgbColor": {"red": 0, "green": 0, "blue": 0}}
    },
    CLASS: {
        BACKGROUND_FILL: {
            "solidFill": {
                "color": {
                    "rgbColor": {"red": 0.9411765, "green": 0.7019608, "blue": 0.70980394}
                },
                "alpha": 1
            }
        },
        OUTLINE: {
            "outlineFill": {
                "solidFill": {
                    "color": {"rgbColor": {"red": 0, "green": 0, "blue": 0}},
                    "alpha": 1
                }
            },
            "weight": {"magnitude": 38100, "unit": "EMU"},
            "dashStyle": "DASH"
        },
        CONTENT_ALIGNMENT: "MIDDLE",
        TEXT_COLOR: {"rgbColor": {"red": 0, "green": 0, "blue": 0}}
    },
    LAB: {
        BACKGROUND_FILL: {
            "solidFill": {
                "color": {
                    "rgbColor": {"red": 1, "green": 0.90588236, "blue": 0.63529414}
                },
                "alpha": 1
            }
        },
        OUTLINE: {
            "outlineFill": {
                "solidFill": {
                    "color": {"rgbColor": {"red": 0, "green": 0, "blue": 0}},
                    "alpha": 1
                }
            },
            "weight": {"magnitude": 38100, "unit": "EMU"},
            "dashStyle": "DOT"
        },
        CONTENT_ALIGNMENT: "MIDDLE",
        TEXT_COLOR: {"rgbColor": {"red": 0, "green": 0, "blue": 0}}
    },
    EXAM: {
        BACKGROUND_FILL: {
            "solidFill": {
                "color": {
                    "rgbColor": {"red": 0.6, "green": 0.7019608, "blue": 0.85882354}
                },
                "alpha": 1
            }
        },
        OUTLINE: {
            "outlineFill": {
                "solidFill": {
                    "color": {"rgbColor": {"red": 0, "green": 0, "blue": 0}},
                    "alpha": 1
                }
            },
            "weight": {"magnitude": 76200, "unit": "EMU"},
            "dashStyle": "SOLID"
        },
        CONTENT_ALIGNMENT: "MIDDLE",
        TEXT_COLOR: {"rgbColor": {"red": 0, "green": 0, "blue": 0}}
    },
    NORMAL: {
        BACKGROUND_FILL: {
            "solidFill": {
                "color": {
                    "rgbColor": {"red": 0.8784314, "green": 0.8784314, "blue": 0.8784314}
                },
                "alpha": 1
            }
        },
        OUTLINE: {
            "outlineFill": {
                "solidFill": {
                    "color": {"rgbColor": {"red": 0.8784314, "green": 0.8784314, "blue": 0.8784314}},
                    "alpha": 1
                }
            },
            "weight": {"magnitude": 9525, "unit": "EMU"},
            "dashStyle": "SOLID"
        },
        CONTENT_ALIGNMENT: "MIDDLE",
        TEXT_COLOR: {"rgbColor": {"red": 0, "green": 0, "blue": 0}}
    },
}
