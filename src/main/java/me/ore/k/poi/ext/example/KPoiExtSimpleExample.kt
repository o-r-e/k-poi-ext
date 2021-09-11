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

package me.ore.k.poi.ext.example

import me.ore.k.poi.ext.CellInit
import me.ore.k.poi.ext.org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.nio.file.Files
import java.nio.file.Paths


object KPoiExtSimpleExample {
    /**
     * This code will create example XLSX file in user home directory
     *
     * Code is not elegant (sometimes dangerous and ugly) - it is just example of usage of this lib
     */
    @JvmStatic
    fun main(args: Array<String>) {
        val outputFile = run {
            val userHomeDirectory = Paths.get(System.getProperty("user.home"))
            val outputFileName = "${this.javaClass.name}.output.xlsx"
            userHomeDirectory.resolve(outputFileName)
        }

        XSSFWorkbook().use { workbook ->
            // Initializing fonts
            val fonts = run {
                val fonts = object {
                    lateinit var default: Font
                    lateinit var bold: Font
                }

                // Saving default font in variable
                fonts.default = workbook.defaultFont {
                    // Setting height of default font as 10 points (pt)
                    it.fontHeightInPoints = 10
                }

                // Cloning settings of default font in new font and saving this new font in variable
                fonts.bold = workbook.cloneFont {
                    // Set new font as bold
                    it.bold = true
                }

                fonts
            }
            println("Fonts initialized")

            // Initializing styles
            val styles = run {
                val styles = object {
                    lateinit var default: CellStyle
                    lateinit var boldCenter: CellStyle
                    lateinit var boldCenterBorder: CellStyle
                }

                // Saving default cell style in variable
                styles.default = workbook.defaultCellStyle {
                    // Configuring default cell style

                    it.setFont(fonts.default)
                    it.alignment = HorizontalAlignment.GENERAL
                    it.verticalAlignment = VerticalAlignment.CENTER
                }

                // Cloning settings of default cell style into new cell style and saving this new cell style in variable
                styles.boldCenter = workbook.cloneCellStyle {
                    // Configuring new cell style

                    it.setFont(fonts.bold)
                    it.alignment = HorizontalAlignment.CENTER
                }

                // Cloning settings of `styles.boldCenter` into new cell style and saving this new cell style in variable
                styles.boldCenterBorder = workbook.cloneCellStyle(styles.boldCenter) {
                    // Setting all borders as thin and applying "automatic" color to them
                    it.border()

                    // Configuring left and right borders
                    it.border(BorderStyle.THICK, IndexedColors.BLUE_GREY, top = false, bottom = false)

                    // Configuring bottom border
                    it.borderBottom(BorderStyle.DASHED, IndexedColors.RED)
                }

                styles
            }
            println("Styles initialized")

            // Creating and filling sheet
            workbook.createSheet("Example") { sheet ->
                // Setting widths of columns
                sheet.setColumnWidthInPixels("A", 80)
                sheet.setColumnWidthInPixels("B", 160)
                sheet.setColumnWidthInPixels("C", 240)

                // A1 - default style
                sheet.cell("A1", "Cell A1") { it.cellStyle = styles.default }

                // A2 - B4 - merged cell (3 rows, 2 columns), style "boldCenter"
                sheet.cell("A2", "Cell A2 - B4") {
                    // Set style to cell
                    it.cellStyle = styles.boldCenter

                    // Merge cell, 3 rows, 2 columns
                    it.merge(rowSpan = 3, columnSpan = 2)
                }

                // C1 - C5 - cells with same style; used "cell initializer"
                sheet.cells(
                    afterCellInit = { it.cellStyle = styles.boldCenterBorder }
                ) { cells: CellInit ->
                    cells.cell("C1", "Cell C1")
                    cells.cell("C3", "Cell C3")
                    cells.cell("C5", "Cell C5")

                    cells.cell("C2") { it.setCellValue("Cell C2 !") }
                    cells.cell("C4") { it.setCellValue("Cell C4 !") }
                }
            }
            println("Workbook filled")

            // Saving workbook to file
            Files.newOutputStream(outputFile).use { stream ->
                workbook.write(stream)
                stream.flush()
            }
            println("Workbook saved to file \"$outputFile\"")
        }
    }
}