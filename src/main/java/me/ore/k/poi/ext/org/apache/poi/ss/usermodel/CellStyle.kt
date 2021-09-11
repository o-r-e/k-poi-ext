/*
 *    Copyright 2021 Obuhov R.
 *
 *    Licensed under the Apache License, Version 2.0 (the "License");
 *    you may not use this file except in compliance with the License.
 *    You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 */

@file:Suppress("unused")

package me.ore.k.poi.ext.org.apache.poi.ss.usermodel

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.IndexedColors


// region Border
/**
 * Applies the specified [style] and [color] to the desired borders
 *
 * @param style Border style
 * @param color Border color
 * @param left If `true`, passes
 *   [style] to [setBorderLeft()] [CellStyle.setBorderLeft] and
 *   [color].[index][IndexedColors.index] to [setLeftBorderColor()][CellStyle.setLeftBorderColor]
 * @param top If `true`, passes
 *   [style] to [setBorderTop()][CellStyle.setBorderTop] and
 *   [color].[index][IndexedColors.index] to [setTopBorderColor()][CellStyle.setTopBorderColor]
 * @param right If `true`, passes
 *   [style] to [setBorderRight()][CellStyle.setBorderRight] and
 *   [color].[index][IndexedColors.index] to [setRightBorderColor()][CellStyle.setRightBorderColor]
 * @param bottom If `true`, passes
 *   [style] to [setBorderBottom()][CellStyle.setBorderBottom] and
 *   [color].[index][IndexedColors.index] to [setBottomBorderColor()][CellStyle.setBottomBorderColor]
 */
fun CellStyle.border(
    style: BorderStyle = BorderStyle.THIN,
    color: IndexedColors = IndexedColors.AUTOMATIC,
    left: Boolean = true,
    top: Boolean = true,
    right: Boolean = true,
    bottom: Boolean = true
) {
    if (left) {
        this.borderLeft = style
        this.leftBorderColor = color.index
    }
    if (top) {
        this.borderTop = style
        this.topBorderColor = color.index
    }
    if (right) {
        this.borderRight = style
        this.rightBorderColor = color.index
    }
    if (bottom) {
        this.borderBottom = style
        this.bottomBorderColor = color.index
    }
}

/**
 * Applies the specified [style] and [color] to the left border (using [border()][border])
 *
 * @param style Border style
 * @param color Border color
 */
fun CellStyle.borderLeft(style: BorderStyle = BorderStyle.THIN, color: IndexedColors = IndexedColors.AUTOMATIC) {
    this.border(style, color, top = false, right = false, bottom = false)
}

/**
 * Applies the specified [style] and [color] to the top border (using [border()][border])
 *
 * @param style Border style
 * @param color Border color
 */
fun CellStyle.borderTop(style: BorderStyle = BorderStyle.THIN, color: IndexedColors = IndexedColors.AUTOMATIC) {
    this.border(style, color, left = false, right = false, bottom = false)
}

/**
 * Applies the specified [style] and [color] to the right border (using [border()][border])
 *
 * @param style Border style
 * @param color Border color
 */
fun CellStyle.borderRight(style: BorderStyle = BorderStyle.THIN, color: IndexedColors = IndexedColors.AUTOMATIC) {
    this.border(style, color, left = false, top = false, bottom = false)
}

/**
 * Applies the specified [style] and [color] to the bottom border (using [border()][border])
 *
 * @param style Border style
 * @param color Border color
 */
fun CellStyle.borderBottom(style: BorderStyle = BorderStyle.THIN, color: IndexedColors = IndexedColors.AUTOMATIC) {
    this.border(style, color, left = false, top = false, right = false)
}
// endregion
