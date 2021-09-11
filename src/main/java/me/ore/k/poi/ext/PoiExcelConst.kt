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

package me.ore.k.poi.ext


/**
 * Constants used when working with MS Excel workbooks
 */
interface PoiExcelConst {
    companion object {
        /**
         * Multiplier for column width (column width is measured in units of "1/256 character width")
         */
        const val WIDTH_UNIT_MULTIPLIER = 256

        /**
         * Default text line height
         */
        const val TEXT_LINE_HEIGHT_MULTIPLIER = 1.4

        /**
         * Multiplier for calculating vertical margins depending on the font size
         */
        const val CELL_VERTICAL_MARGIN_MULTIPLIER = 0.2
    }
}
