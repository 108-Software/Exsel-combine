package main

import (
	"fmt"
	"log"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Папка с исходными файлами
	sourceDir := "./files"
	outputFile := "merged_result.xlsx"

	// Находим все Excel файлы
	files, err := filepath.Glob(filepath.Join(sourceDir, "*.xlsx"))
	if err != nil {
		log.Fatal("Ошибка при поиске файлов:", err)
	}

	if len(files) == 0 {
		log.Fatal("Не найдено Excel файлов в папке:", sourceDir)
	}

	fmt.Printf("Найдено файлов: %d\n", len(files))

	// Открываем первый файл для анализа структуры
	template, err := excelize.OpenFile(files[0])
	if err != nil {
		log.Fatal("Ошибка при открытии первого файла:", err)
	}
	defer template.Close()

	templateSheet := template.GetSheetName(0)

	// Получаем все строки первого файла для анализа
	firstRows, err := template.GetRows(templateSheet)
	if err != nil {
		log.Fatal("Ошибка при чтении первого файла:", err)
	}

	// Создаем результирующий файл
	result := excelize.NewFile()
	defer result.Close()

	// Удаляем стандартный лист
	result.DeleteSheet("Sheet1")

	// Создаем новый лист
	resultSheet := "Объединенные данные"
	_, err = result.NewSheet(resultSheet)
	if err != nil {
		log.Fatal("Ошибка создания листа:", err)
	}

	// Копируем ширину колонок из шаблона
	fmt.Println("Копирование структуры колонок...")
	for col := 1; col <= len(firstRows[0]); col++ {
		colLetter, _ := excelize.ColumnNumberToName(col)

		// Получаем ширину колонки из шаблона
		if width, err := template.GetColWidth(templateSheet, colLetter); err == nil && width > 0 {
			if err := result.SetColWidth(resultSheet, colLetter, colLetter, width); err != nil {
				log.Printf("Ошибка установки ширины колонки %s: %v\n", colLetter, err)
			}
		}

		// Копируем видимость колонок
		if visible, err := template.GetColVisible(templateSheet, colLetter); err == nil {
			if err := result.SetColVisible(resultSheet, colLetter, visible); err != nil {
				log.Printf("Ошибка установки видимости колонки %s: %v\n", colLetter, err)
			}
		}
	}

	// Карта для соответствия стилей (оригинальный стиль -> новый стиль)
	styleMap := make(map[int]int)

	// Функция для копирования стиля
	copyStyle := func(originalStyleID int) int {
		if originalStyleID == 0 {
			return 0
		}

		// Проверяем, не копировали ли уже этот стиль
		if newID, exists := styleMap[originalStyleID]; exists {
			return newID
		}

		// Получаем определение стиля
		style, err := template.GetStyle(originalStyleID)
		if err != nil || style == nil {
			return 0
		}

		// Создаем новый стиль в результирующем файле
		newStyleID, err := result.NewStyle(style)
		if err != nil {
			log.Printf("Ошибка создания стиля: %v\n", err)
			return 0
		}

		styleMap[originalStyleID] = newStyleID
		return newStyleID
	}

	// Копируем данные из всех файлов
	currentRow := 1
	headerCopied := false
	totalRowsCopied := 0

	for i, filePath := range files {
		fmt.Printf("Обработка %d/%d: %s\n", i+1, len(files), filepath.Base(filePath))

		f, err := excelize.OpenFile(filePath)
		if err != nil {
			log.Printf("  ⚠️ Ошибка открытия: %v\n", err)
			continue
		}

		sheetName := f.GetSheetName(0)
		rows, err := f.GetRows(sheetName)
		if err != nil {
			log.Printf("  ⚠️ Ошибка чтения: %v\n", err)
			f.Close()
			continue
		}

		if len(rows) == 0 {
			f.Close()
			continue
		}

		// Определяем, с какой строки начинать копирование
		startRow := 0
		if headerCopied {
			startRow = 1 // Пропускаем заголовок в остальных файлах
		} else {
			headerCopied = true
		}

		rowsCopied := 0

		// Копируем строки
		for rowIdx := startRow; rowIdx < len(rows); rowIdx++ {
			for colIdx := 0; colIdx < len(rows[rowIdx]); colIdx++ {
				colLetter, _ := excelize.ColumnNumberToName(colIdx + 1)
				cellRef := fmt.Sprintf("%s%d", colLetter, currentRow)
				originalCellRef := fmt.Sprintf("%s%d", colLetter, rowIdx+1)

				// Копируем значение ячейки
				value := rows[rowIdx][colIdx]
				if err := result.SetCellStr(resultSheet, cellRef, value); err != nil {
					log.Printf("  ⚠️ Ошибка записи %s: %v\n", cellRef, err)
				}

				// Копируем стиль (только для заголовка из первого файла)
				var styleID int
				if !headerCopied && rowIdx == 0 {
					// Для заголовка используем стиль из шаблона
					styleID, _ = template.GetCellStyle(templateSheet, originalCellRef)
				} else if i == 0 && rowIdx > 0 {
					// Для данных первого файла копируем их стили
					styleID, _ = f.GetCellStyle(sheetName, originalCellRef)
				}

				if styleID != 0 {
					newStyleID := copyStyle(styleID)
					if newStyleID != 0 {
						if err := result.SetCellStyle(resultSheet, cellRef, cellRef, newStyleID); err != nil {
							log.Printf("  ⚠️ Ошибка применения стиля: %v\n", err)
						}
					}
				}

				// Копируем гиперссылки - исправлено: метод возвращает 3 значения
				hasHyperlink, link, err := f.GetCellHyperLink(sheetName, originalCellRef)
				if err != nil {
					log.Printf("  ⚠️ Ошибка получения гиперссылки для %s: %v\n", originalCellRef, err)
				} else if hasHyperlink && link != "" {
					// Определяем тип ссылки
					linkType := "External" // По умолчанию внешняя ссылка
					if len(link) > 0 && link[0] == '#' {
						linkType = "Location" // Внутренняя ссылка
					}
					if err := result.SetCellHyperLink(resultSheet, cellRef, link, linkType); err != nil {
						log.Printf("  ⚠️ Ошибка установки гиперссылки: %v\n", err)
					}
				}

				// Копируем формулы (только из первого файла для сохранения целостности)
				if i == 0 {
					if formula, err := f.GetCellFormula(sheetName, originalCellRef); err == nil && formula != "" {
						if err := result.SetCellFormula(resultSheet, cellRef, formula); err != nil {
							log.Printf("  ⚠️ Ошибка установки формулы: %v\n", err)
						}
					}
				}
			}
			currentRow++
			rowsCopied++
		}

		f.Close()
		fmt.Printf("  ✅ Скопировано строк: %d\n", rowsCopied)
		totalRowsCopied += rowsCopied
	}

	// Применяем объединение ячеек из шаблона (только для заголовка)
	if len(files) > 0 {
		mergeCells, err := template.GetMergeCells(templateSheet)
		if err == nil {
			for _, mergeCell := range mergeCells {
				// Корректируем диапазон для заголовка (только первая строка)
				if mergeCell.GetStartAxis()[1:] == "1" {
					if err := result.MergeCell(resultSheet, mergeCell.GetStartAxis(), mergeCell.GetEndAxis()); err != nil {
						log.Printf("Ошибка объединения ячеек: %v\n", err)
					}
				}
			}
		}
	}

	// Применяем автофильтр
	if currentRow > 1 && len(firstRows) > 0 {
		lastCol := len(firstRows[0])
		lastColLetter, _ := excelize.ColumnNumberToName(lastCol)
		rangeRef := fmt.Sprintf("%s1:%s%d", "A", lastColLetter, currentRow-1)
		if err := result.AutoFilter(resultSheet, rangeRef, []excelize.AutoFilterOptions{}); err != nil {
			log.Printf("Ошибка применения автофильтра: %v\n", err)
		}
	}

	// Закрепляем заголовок (фиксируем первую строку)
	if err := result.SetPanes(resultSheet, &excelize.Panes{
		Freeze:      true,
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
	}); err != nil {
		log.Printf("Ошибка закрепления заголовка: %v\n", err)
	}

	// Сохраняем результат
	err = result.SaveAs(outputFile)
	if err != nil {
		log.Fatal("Ошибка сохранения файла:", err)
	}

	fmt.Println("\n" + "================================================")
	fmt.Printf("✅ ОБЪЕДИНЕНИЕ ЗАВЕРШЕНО УСПЕШНО!\n")
	fmt.Printf("📁 Обработано файлов: %d\n", len(files))
	fmt.Printf("📊 Скопировано строк данных: %d\n", totalRowsCopied)
	fmt.Printf("🎨 Стили и структура унаследованы из: %s\n", filepath.Base(files[0]))
	fmt.Printf("💾 Результат сохранен в: %s\n", outputFile)
	fmt.Println("================================================")
}
