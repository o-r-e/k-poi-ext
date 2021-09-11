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

package me.ore.k.poi.ext.org.apache.poi.ss.usermodel

import org.apache.poi.ss.usermodel.Font
import org.apache.poi.xssf.usermodel.XSSFFont


/**
 * Copies the settings from [source] to the current font
 *
 * @param source Settings source
 */
fun Font.clone(source: Font) {
    this.fontName = source.fontName
    this.fontHeight = source.fontHeight
    this.italic = source.italic
    this.strikeout = source.strikeout
    this.color = source.color
    this.typeOffset = source.typeOffset
    this.underline = source.underline
    this.charSet = source.charSet
    this.bold = source.bold

    if ((this is XSSFFont) && (source is XSSFFont)) {
        this.scheme = source.scheme
        this.xssfColor.argbHex = source.xssfColor.argbHex
    }
}