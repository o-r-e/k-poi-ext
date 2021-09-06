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

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.util.CellRangeAddress
import me.ore.k.poi.ext.PoiExcelConst
import kotlin.math.ceil
import kotlin.math.floor
import kotlin.math.max


fun Cell.getVisibleText(): String = DataFormatter().formatCellValue(this, this.sheet.workbook.creationHelper.createFormulaEvaluator())

fun Cell.calculateTextLines(columnWidthInChars: Double): List<String> {
    val text = this.getVisibleText()
    if (text.isEmpty()) {
        return listOf("")
    }

    val lines = text.lines()

    if (!this.cellStyle.wrapText) {
        return lines
    }

    val columnWidthInCharsInt = floor(columnWidthInChars).toInt()

    val result = ArrayList<String>()

    lines.forEach { line ->
        if (line.length < columnWidthInChars) {
            result.add(line)
        } else {
            val buffer = StringBuilder()
            val flushBuffer = {
                if (buffer.isNotEmpty()) {
                    result.add(buffer.trim().toString())
                    buffer.clear()
                }
            }
            val addPart = { start: Int, end: Int ->
                val partLength = end - start
                if (partLength > columnWidthInChars) {
                    flushBuffer()

                    val fullLineCount = floor(partLength / columnWidthInChars).toInt()
                    for (i in 1 .. fullLineCount) {
                        val fullLineStart = start + ((i - 1) * fullLineCount)
                        val fullLineEnd = start + (i * fullLineCount)
                        result.add(line.substring(fullLineStart, fullLineEnd))
                    }

                    val restLength = partLength - (fullLineCount * columnWidthInCharsInt)
                    if (restLength > 0) {
                        buffer.append(line, partLength - restLength, partLength)
                    }
                } else if (buffer.length + partLength > columnWidthInChars) {
                    flushBuffer()
                    buffer.append(line, start, end)
                } else {
                    buffer.append(line, start, end)
                }
                Unit
            }

            var charIndex = 0
            while (charIndex < line.length) {
                val char = line[charIndex]
                if (char.isLetter() || char.isDigit() || (char == '_')) {
                    var lastCharIndex = charIndex + 1

                    while (lastCharIndex < line.length) {
                        val lastChar = line[lastCharIndex]
                        if (lastChar.isLetter() || lastChar.isDigit() || (lastChar == '_')) {
                            lastCharIndex++
                        } else {
                            break
                        }
                    }

                    addPart(charIndex, lastCharIndex)

                    charIndex = lastCharIndex
                } else if (char.isWhitespace()) {
                    buffer.append(char)
                    charIndex++
                } else {
                    addPart(charIndex, charIndex + 1)
                    charIndex++
                }
            }

            flushBuffer()
        }
    }

    return result
}

fun Cell.calculateSuitableHeight(ignoreMergedRegions: Boolean = false): Double {
    val row = this.row!!
    val sheet = row.sheet!!

    val mergedRegion = if (ignoreMergedRegions) {
        null
    } else {
        sheet.mergedRegions.firstOrNull { it.containsRow(this.rowIndex) && it.containsColumn(this.columnIndex) }
    }

    return this.calculateSuitableHeight(mergedRegion)
}

fun Cell.calculateSuitableHeight(mergedRegion: CellRangeAddress? = null): Double {
    val row = this.row!!
    val sheet = row.sheet!!
    val workbook = sheet.workbook!!
    val defaultFont = this.sheet.workbook.defaultFont

    val columnWidthInChars = if (mergedRegion != null) {
        var columnWidthInChars = 0.0
        for (mergedColumnIndex in mergedRegion.firstColumn .. mergedRegion.lastColumn) {
            columnWidthInChars += sheet.getColumnWidth(mergedColumnIndex).toDouble() / PoiExcelConst.WIDTH_UNIT_MULTIPLIER
        }
        columnWidthInChars
    } else {
        sheet.getColumnWidth(this.columnIndex).toDouble() / PoiExcelConst.WIDTH_UNIT_MULTIPLIER
    }


    val cellFont = workbook.getFontAt(this.cellStyle.fontIndex)
    val fontScale = defaultFont.fontHeightInPoints.toDouble() / cellFont.fontHeightInPoints
    val cellWidthInChars = columnWidthInChars * fontScale * 1.1

    val textLines = this.calculateTextLines(ceil(cellWidthInChars))
    val textLineCount = textLines.size

    val textHeight = cellFont.fontHeightInPoints * textLineCount * PoiExcelConst.TEXT_LINE_HEIGHT_MULTIPLIER
    val marginHeight = defaultFont.fontHeightInPoints * PoiExcelConst.CELL_VERTICAL_MARGIN_MULTIPLIER
    return textHeight + marginHeight
}


fun Cell.merge(rowSpan: Int = 1, columnSpan: Int = 1) {
    val firstRow = this.rowIndex
    val lastRow = max(firstRow, firstRow + rowSpan - 1)
    val firstColumn = this.columnIndex
    val lastColumn = max(firstColumn, firstColumn + columnSpan - 1)

    val address = CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn)

    this.sheet.addMergedRegion(address)
}


fun Cell.setCellValue(value: Number) {
    this.setCellValue(value.toDouble())
}
