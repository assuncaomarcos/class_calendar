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
 * Creates a slide in a presentation with the specified layout.
 *
 * @param {string} presentationId - The ID of the presentation.
 * @param {string} layoutId - The ID of the slide layout to apply to the slide.
 * @returns {string} - The ID of the created slide object.
 */
function createSlide(presentationId, layoutId) {
    const pageId = Utilities.getUuid();

    const requests = [{
        'createSlide': {
            'objectId': pageId,
            'slideLayoutReference': {
                'layoutId': layoutId
            }
        }
    }];

    try {
        const slide = Slides.Presentations.batchUpdate({'requests': requests}, presentationId);
        return slide.replies[0].createSlide.objectId;
    } catch (e) {
        Logger.log('Failed with error %s', e.message);
    }
}

/**
 * Returns a request object for creating a shape in a specified page.
 *
 * @param {string} pageId - The ID of the page where the shape will be created.
 * @param {string} pageElementId - The ID of the new shape element.
 * @param {string} shapeType - The type of shape to create.
 * @param {number} scaleX - The horizontal scale factor of the shape.
 * @param {number} scaleY - The vertical scale factor of the shape.
 * @param {number} translateX - The horizontal translation of the shape.
 * @param {number} translateY - The vertical translation of the shape.
 * @returns {Object} The request object for creating the shape.
 * @example
 *
 * const pageId = "examplePageId";
 * const pageElementId = "exampleShapeId";
 * const shapeType = "RECTANGLE";
 * const scaleX = 1;
 * const scaleY = 1;
 * const translateX = 0;
 * const translateY = 0;
 *
 * const request = getCreateShapeRequest(pageId, pageElementId, shapeType, scaleX, scaleY, translateX, translateY);
 * console.log(request);
 */
function getCreateShapeRequest(pageId, pageElementId, shapeType, scaleX, scaleY, translateX, translateY) {
    return  {
        "createShape": {
            "objectId": pageElementId,
            "shapeType": shapeType,
            "elementProperties": {
                "pageObjectId": pageId,
                "size": {
                    "height": {"magnitude": DefaultMagnitude, "unit": "EMU"},
                    "width": {"magnitude": DefaultMagnitude, "unit": "EMU"}
                },
                "transform": {
                    "scaleX": scaleX,
                    "scaleY": scaleY,
                    "translateX": translateX,
                    "translateY": translateY,
                    "unit": "EMU"
                }
            }
        }
    }
}

/**
 * Creates an update shape request object.
 *
 * @param {string} pageElementId - The ID of the page element to update.
 * @param {object} backgroundFill - The background fill object.
 * @param {object} outline - The outline object.
 * @param {string} contentAlignment - The content alignment value.
 * @returns {object} - The update shape request object.
 */
function getUpdateShapeRequest(pageElementId, backgroundFill, outline, contentAlignment) {
    return {
        "updateShapeProperties": {
            "objectId": pageElementId,
            "shapeProperties":  {
                "shapeBackgroundFill": backgroundFill,
                "outline": outline,
                "contentAlignment": contentAlignment || "MIDDLE"
            },
            "fields": "shapeBackgroundFill,outline,contentAlignment"
        }
    }
}

/**
 * Generates a request object for inserting text into a page element.
 *
 * @param {string} pageElementId - The ID of the page element to insert the text into.
 * @param {string} textContent - The text content to be inserted.
 * @returns {Object} The request object for inserting text.
 * @throws {TypeError} If either pageElementId or textContent is not a string.
 */
function getInsertTextRequest(pageElementId, textContent) {
    return {
        "insertText": {
            "objectId": pageElementId,
            "text": textContent
        }
    }
}

/**
 * Generates an update text style request object for a given page element.
 *
 * @param {string} pageElementId - The ID of the page element to update.
 * @param {string} fontFamily - The font family to apply to the text style.
 * @param {number} fontSize - The font size to apply to the text style.
 * @param {*} color - The color to apply to the text style. If not provided, default black color is used.
 * @param {boolean} bold - Whether to apply bold style to the text. Default is false.
 * @param {boolean} italic - Whether to apply italic style to the text. Default is false.
 * @param {boolean} underline - Whether to apply underline style to the text. Default is false.
 * @param {boolean} strikethrough - Whether to apply strikethrough style to the text. Default is false.
 * @returns {Object} - The update text style request object.
 */
function getUpdateTextStyleRequest(pageElementId, fontFamily,
                                   fontSize, color, bold, italic,
                                   underline, strikethrough) {
    return {
        "updateTextStyle": {
            "objectId": pageElementId,
            "style": {
                "fontFamily": fontFamily,
                "fontSize": {"magnitude": fontSize, "unit": "PT"},
                "foregroundColor": {
                    "opaqueColor": color ? color : {"rgbColor": {"red": 0, "green": 0, "blue": 0}}
                },
                "bold": bold ? bold : false,
                "italic": italic ? italic : false,
                "underline": underline ? underline : false,
                "strikethrough": strikethrough ? strikethrough : false
            },
            "textRange": {
                "type": "ALL"
            },
            "fields": "fontFamily,fontSize,foregroundColor,bold,italic,underline,strikethrough"
        }
    }
}

/**
 * Returns an object with the parameters for updating the paragraph style of a page element.
 *
 * @param {string} pageElementId - The ID of the page element to update.
 * @param {string=} alignment - The alignment of the text. Defaults to "CENTER" if not provided.
 * @param {string=} spacingMode - The spacing mode of the text. Defaults to "COLLAPSE_LISTS" if not provided.
 * @param {string=} direction - The direction of the text. Defaults to "LEFT_TO_RIGHT" if not provided.
 * @returns {Object} - An object with the parameters for updating the paragraph style.
 *   - updateParagraphStyle: {
 *       objectId: string,
 *       style: {
 *         alignment: string,
 *         spacingMode: string,
 *         direction: string
 *       },
 *       textRange: {
 *         type: string
 *       },
 *       fields: string
 *     }
 */
function getUpdateParagraphStyle(pageElementId, alignment, spacingMode, direction) {
    return {
        "updateParagraphStyle": {
            "objectId": pageElementId,
            "style": {
                "alignment": alignment ? alignment : "CENTER",
                "spacingMode": spacingMode ? spacingMode : "COLLAPSE_LISTS",
                "direction": direction ? direction : "LEFT_TO_RIGHT"
            },
            "textRange": {
                "type": "ALL"
            },
            "fields": "alignment,spacingMode,direction"
        }
    }
}

/**
 * Retrieves the label request objects for creating a label on a Google Slides page.
 *
 * @param {string} pageId - The ID of the page where the label will be created.
 * @param {string} content - The text content of the label.
 * @param {number} scaleX - The horizontal scaling factor of the label.
 * @param {number} scaleY - The vertical scaling factor of the label.
 * @param {number} translateX - The horizontal translation of the label.
 * @param {number} translateY - The vertical translation of the label.
 * @param {string} alignment - The alignment of the label text.
 * @returns {Object[]} - An array of label request objects:
 *   - `createShape` request object contains the shape type, element properties, and object ID.
 *   - `insertText` request object contains the text content and object ID.
 *   - `updateTextStyle` request object contains the text range, style properties, and object ID.
 *   - `updateParagraphStyle` request object contains the text range, style properties, and object ID.
 */
function getLabelRequest(pageId, content, scaleX, scaleY, translateX, translateY, alignment) {
    const pageElementId = Utilities.getUuid();

    return [
        getCreateShapeRequest(pageId, pageElementId,
            "TEXT_BOX", scaleX, scaleY,
            translateX, translateY),
        getInsertTextRequest(pageElementId, content),
        getUpdateTextStyleRequest(pageElementId, LabelTextType.FONT_FAMILY, LabelTextType.FONT_SIZE),
        getUpdateParagraphStyle(pageElementId, alignment)
    ];
}

/**
 * Converts a measurement in emus (English Metric Units) to points.
 *
 * @param {number} emu - The value in emus to be converted.
 * @returns {number} - The converted value in points.
 */
function emuToPoints(emu) {
    return (emu / 914400) * 72;
}

/**
 * Converts points to emu (English Metric Unit) based on the input points value.
 *
 * @param {number} points - The number of points to be converted to emu.
 * @returns {number} - The converted value in emu units.
 */
function pointsToEmu(points) {
    return (points / 72) * 914400;
}
