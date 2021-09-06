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

package r.obuhov.k.poi.ext

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import r.obuhov.k.poi.ext.org.apache.poi.ss.usermodel.*
import java.nio.file.Files
import java.nio.file.Paths


object Test {
    @JvmStatic
    fun main(args: Array<String>) {
        "".lines()
        println("Started")

        if (args.isEmpty())
            error("Program arguments are empty; output file path is required")

        val outputFilePath = run {
            val pathString = args.first()
            if (pathString.isEmpty())
                error("First program argument is path of output file; this argument is empty")

            try {
                Paths.get(pathString)
                    ?: error("Parsed path is null")
            } catch (e: Exception) {
                throw IllegalStateException(
                    "First program argument is path of output file; this argument is not valid path",
                    e
                )
            }
        }
        println("Output file path - $outputFilePath")

        SXSSFWorkbook(15).use { workbook ->
            try {
                var timeMs: Long

                // region Filling
                println("Filling")
                timeMs = System.currentTimeMillis()

                val fonts = run {
                    val fonts = object {
                        lateinit var default: Font
                        lateinit var bold: Font
                    }

                    fonts.default = workbook.defaultFont { it.fontHeightInPoints = 10 }
                    fonts.bold = workbook.cloneFont { it.bold = true }

                    fonts
                }

                val styles = run {
                    val styles = object {
                        lateinit var default: CellStyle
                        lateinit var preTitle: CellStyle
                        lateinit var mainTitle: CellStyle
                        val titleTable = object {
                            lateinit var codesText: CellStyle
                            lateinit var header: CellStyle
                            lateinit var names: CellStyle
                            lateinit var codes: CellStyle
                        }
                    }

                    styles.default = workbook.defaultCellStyle {
                        it.setFont(fonts.default)
                        it.verticalAlignment = VerticalAlignment.CENTER
                        it.wrapText = true
                    }
                    styles.preTitle = workbook.cloneCellStyle {
                        it.alignment = HorizontalAlignment.RIGHT
                    }
                    styles.mainTitle = workbook.cloneCellStyle {
                        it.setFont(fonts.bold)
                        it.alignment = HorizontalAlignment.CENTER
                    }
                    styles.titleTable.codesText = workbook.cloneCellStyle {
                        it.setFont(fonts.bold)
                        it.alignment = HorizontalAlignment.CENTER
                    }
                    styles.titleTable.header = workbook.cloneCellStyle {
                        it.setFont(fonts.bold)
                    }
                    styles.titleTable.names = workbook.cloneCellStyle()
                    styles.titleTable.codes = workbook.cloneCellStyle {
                        it.alignment = HorizontalAlignment.CENTER
                        it.border()
                    }

                    styles
                }

                workbook.createSheet("Data") { sheet ->
                    sheet.setColumnWidthInPixels("A", 112)
                    sheet.setColumnWidthInPixels("B", 288)
                    sheet.setColumnWidthInPixels("C", 112)
                    sheet.setColumnWidthInPixels("D", 80)
                    for (columnTitle in 'E' .. 'J') {
                        sheet.setColumnWidthInPixels(columnTitle.toString(), 97)
                    }

                    // region Приложение 2 к Правилам составления и представления бюджетной заявки, Форма 01-111
                    sheet.cells({ it.cellStyle = styles.preTitle }) preTitleCells@ { cells ->
                        cells.cell("J1", "Приложение 2")
                        cells.cell("G2", "к Правилам составления и представления бюджетной заявки") {
                            it.merge(columnSpan = 4)
                        }
                        cells.cell("J3", "Форма 01-111")
                    }
                    // endregion

                    // region Расчет расходов на оплату труда административных государственных служащих
                    sheet.cells({ it.cellStyle = styles.mainTitle }) { cells ->
                        cells.cell("A5", "Расчет расходов на оплату труда административных государственных служащих") {
                            it.merge(columnSpan = 10)
                        }
                    }
                    // endregion

                    // region Таблица-заголовок - заголовок
                    sheet.cells({
                        it.merge(columnSpan = 2)
                        it.cellStyle = styles.titleTable.header
                    }) { cells ->
                        cells.cell("A8", "Год")
                        cells.cell("A9", "Вид данных (прогноз, план, отчет)")
                        cells.cell("A10", "Функциональная группа")
                        cells.cell("A11", "Администратор программ")
                        cells.cell("A12", "Государственное учреждение")
                        cells.cell("A13", "Программа")
                        cells.cell("A14", "Подпрограмма")
                        cells.cell("A15", "Специфика")
                    }
                    // endregion

                    // region Таблица-заголовок - названия
                    sheet.cells({
                        it.merge(columnSpan = 7)
                        it.cellStyle = styles.titleTable.names
                    }) { cells ->
                        cells.cell("C10", "Гос. услуги общего характера")
                        cells.cell("C11", "Аппарат акима области")
                        cells.cell("C12", "ГУ Аппарат акима Павлодарской области")
                        cells.cell("C13", "Услуги по обеспечению деятельности акима области ".repeat(3).trim())
                        cells.cell("C14", "За счет средств местного бюджета")
                        cells.cell("C15", "Оплата труда")
                    }
                    // endregion

                    // region Таблица-заголовок - коды
                    sheet.cell("J7", "Коды") { it.cellStyle = styles.titleTable.codesText }

                    sheet.cells({ it.cellStyle = styles.titleTable.codes }) { cells ->
                        cells.cell("J8", 2021)
                        cells.cell("J9", "Отчет")
                        cells.cell("J10", "01")
                        cells.cell("J11", "120")
                        cells.cell("J12", "1203001")
                        cells.cell("J13", "001")
                        cells.cell("J14", "015")
                        cells.cell("J15", "111")
                    }
                    // endregion

                    // region Таблица-заголовок - коды 2
                    sheet.getRow(7).let { row ->
                        row.createCell(9).let { cell ->
                            cell.setCellValue(2021.toDouble())
                            cell.cellStyle = styles.titleTable.codes
                        }
                    }
                    sheet.getRow(8).let { row ->
                        row.createCell(9).let { cell ->
                            cell.setCellValue("Отчет")
                            cell.cellStyle = styles.titleTable.codes
                        }
                    }
                    sheet.getRow(9).let { row ->
                        row.createCell(9).let { cell ->
                            cell.setCellValue("01")
                            cell.cellStyle = styles.titleTable.codes
                        }
                    }
                    sheet.getRow(10).let { row ->
                        row.createCell(9).let { cell ->
                            cell.setCellValue("120")
                            cell.cellStyle = styles.titleTable.codes
                        }
                    }
                    sheet.getRow(11).let { row ->
                        row.createCell(9).let { cell ->
                            cell.setCellValue("1203001")
                            cell.cellStyle = styles.titleTable.codes
                        }
                    }
                    sheet.getRow(12).let { row ->
                        row.createCell(9).let { cell ->
                            cell.setCellValue("001")
                            cell.cellStyle = styles.titleTable.codes
                        }
                    }
                    sheet.getRow(13).let { row ->
                        row.createCell(9).let { cell ->
                            cell.setCellValue("015")
                            cell.cellStyle = styles.titleTable.codes
                        }
                    }
                    sheet.getRow(14).let { row ->
                        row.createCell(9).let { cell ->
                            cell.setCellValue("111")
                            cell.cellStyle = styles.titleTable.codes
                        }
                    }
                    // endregion

                    // region Высоты строк таблицы-заголовка
                    for (rowIndex in 10 .. 15) {
                        sheet.row(rowIndex - 1).stretchToMaxContent()
                    }
                    // endregion
                }
                println("Filled, ${System.currentTimeMillis() - timeMs} ms")
                // endregion

                // region Saving
                println("Saving")
                timeMs = System.currentTimeMillis()
                Files.newOutputStream(outputFilePath).use { stream ->
                    workbook.write(stream)
                    stream.flush()
                }
                println("Saved, ${System.currentTimeMillis() - timeMs} ms")
                // endregion
            } finally {
                workbook.dispose()
            }
        }

        println("Finished")
    }
}
