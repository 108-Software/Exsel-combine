package main

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// selectFilesWithDialog открывает стандартное окно выбора файлов Windows
func selectFilesWithDialog() ([]string, error) {
	// Создаем временный PowerShell скрипт
	psScript := `Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Выберите Excel файлы для объединения"
$dialog.Filter = "Excel файлы (*.xlsx;*.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*"
$dialog.Multiselect = $true
$dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

$result = $dialog.ShowDialog()
if ($result -eq "OK") {
    $files = $dialog.FileNames
    # Выводим файлы в UTF-8
    $utf8 = New-Object System.Text.UTF8Encoding
    foreach ($file in $files) {
        $bytes = $utf8.GetBytes($file)
        [Console]::OpenStandardOutput().Write($bytes, 0, $bytes.Length)
        [Console]::OpenStandardOutput().WriteByte(0)  # Разделитель null
    }
}`
	
	cmd := exec.Command("powershell", "-NoProfile", "-Command", psScript)
	cmd.Stderr = os.Stderr
	
	// Запускаем команду и читаем вывод
	output, err := cmd.Output()
	if err != nil {
		return nil, fmt.Errorf("ошибка при выборе файлов: %v", err)
	}
	
	// Разбираем вывод (строки разделены null байтами)
	if len(output) == 0 {
		return nil, fmt.Errorf("файлы не выбраны")
	}
	
	files := strings.Split(string(output), "\x00")
	// Удаляем последний пустой элемент если есть
	if len(files) > 0 && files[len(files)-1] == "" {
		files = files[:len(files)-1]
	}
	
	if len(files) == 0 {
		return nil, fmt.Errorf("файлы не выбраны")
	}
	
	return files, nil
}

// mergeFiles объединяет выбранные Excel файлы
func mergeFiles(files []string, outputFile string) error {
	if len(files) == 0 {
		return fmt.Errorf("нет файлов для объединения")
	}
	
	
	// Проверяем существование файлов
	for _, file := range files {
		if _, err := os.Stat(file); os.IsNotExist(err) {
			return fmt.Errorf("файл не существует: %s", file)
		}
	}
	fmt.Println()
	
	// Открываем первый файл для анализа структуры
	template, err := excelize.OpenFile(files[0])
	if err != nil {
		return fmt.Errorf("ошибка при открытии первого файла '%s': %v", filepath.Base(files[0]), err)
	}
	defer template.Close()
	
	templateSheet := template.GetSheetName(0)
	
	// Получаем все строки первого файла для анализа
	firstRows, err := template.GetRows(templateSheet)
	if err != nil {
		return fmt.Errorf("ошибка при чтении первого файла: %v", err)
	}
	
	if len(firstRows) == 0 {
		return fmt.Errorf("первый файл не содержит данных")
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
		return fmt.Errorf("ошибка создания листа: %v", err)
	}
	
	
	// Копируем ширину колонок из шаблона
	if len(firstRows) > 0 && len(firstRows[0]) > 0 {
		for col := 1; col <= len(firstRows[0]); col++ {
			colLetter, _ := excelize.ColumnNumberToName(col)
			
			if width, err := template.GetColWidth(templateSheet, colLetter); err == nil && width > 0 {
				result.SetColWidth(resultSheet, colLetter, colLetter, width)
			}
			
			if visible, err := template.GetColVisible(templateSheet, colLetter); err == nil {
				result.SetColVisible(resultSheet, colLetter, visible)
			}
		}
	}
	
	// Карта для соответствия стилей
	styleMap := make(map[int]int)
	
	// Функция для копирования стиля
	copyStyle := func(originalStyleID int) int {
		if originalStyleID == 0 {
			return 0
		}
		
		if newID, exists := styleMap[originalStyleID]; exists {
			return newID
		}
		
		style, err := template.GetStyle(originalStyleID)
		if err != nil || style == nil {
			return 0
		}
		
		newStyleID, err := result.NewStyle(style)
		if err != nil {
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
			fmt.Printf("  ⚠️ Файл пуст, пропускаем\n")
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
				result.SetCellStr(resultSheet, cellRef, value)
				
				// Копируем стиль
				var styleID int
				if !headerCopied && rowIdx == 0 {
					styleID, _ = template.GetCellStyle(templateSheet, originalCellRef)
				} else if i == 0 && rowIdx > 0 {
					styleID, _ = f.GetCellStyle(sheetName, originalCellRef)
				}
				
				if styleID != 0 {
					newStyleID := copyStyle(styleID)
					if newStyleID != 0 {
						result.SetCellStyle(resultSheet, cellRef, cellRef, newStyleID)
					}
				}
				
				// Копируем гиперссылки
				hasHyperlink, link, err := f.GetCellHyperLink(sheetName, originalCellRef)
				if err == nil && hasHyperlink && link != "" {
					linkType := "External"
					if len(link) > 0 && link[0] == '#' {
						linkType = "Location"
					}
					result.SetCellHyperLink(resultSheet, cellRef, link, linkType)
				}
				
				// Копируем формулы (только из первого файла)
				if i == 0 {
					if formula, err := f.GetCellFormula(sheetName, originalCellRef); err == nil && formula != "" {
						result.SetCellFormula(resultSheet, cellRef, formula)
					}
				}
			}
			currentRow++
			rowsCopied++
		}
		
		f.Close()
		totalRowsCopied += rowsCopied
	}
	
	// Применяем объединение ячеек из шаблона
	if len(files) > 0 {
		mergeCells, err := template.GetMergeCells(templateSheet)
		if err == nil {
			for _, mergeCell := range mergeCells {
				if len(mergeCell.GetStartAxis()) > 1 && mergeCell.GetStartAxis()[1:] == "1" {
					result.MergeCell(resultSheet, mergeCell.GetStartAxis(), mergeCell.GetEndAxis())
				}
			}
		}
	}
	
	// Применяем автофильтр
	if currentRow > 1 && len(firstRows) > 0 && len(firstRows[0]) > 0 {
		lastCol := len(firstRows[0])
		lastColLetter, _ := excelize.ColumnNumberToName(lastCol)
		rangeRef := fmt.Sprintf("%s1:%s%d", "A", lastColLetter, currentRow-1)
		result.AutoFilter(resultSheet, rangeRef, []excelize.AutoFilterOptions{})
	}
	
	// Закрепляем заголовок
	result.SetPanes(resultSheet, &excelize.Panes{
		Freeze:      true,
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
	})
	
	err = result.SaveAs(outputFile)
	if err != nil {
		return fmt.Errorf("ошибка сохранения файла: %v", err)
	}
	
	return nil
}

func main() {
	fmt.Println("================================================")
	fmt.Println("   ОБЪЕДИНИТЕЛЬ EXCEL ФАЙЛОВ")
	fmt.Println("================================================")
	fmt.Println()
	
	// Открываем окно выбора файлов
	fmt.Println("🖱️ Открывается окно выбора файлов...")
	fmt.Println("(Выберите нужные Excel файлы и нажмите 'Открыть')")
	fmt.Println()
	
	files, err := selectFilesWithDialog()
	if err != nil {
		fmt.Printf("❌ Ошибка: %v\n", err)
		fmt.Println("\nНажмите Enter для выхода...")
		fmt.Scanln()
		return
	}
	
	if len(files) == 0 {
		fmt.Println("❌ Файлы не выбраны")
		fmt.Println("\nНажмите Enter для выхода...")
		fmt.Scanln()
		return
	}
	
	// Показываем выбранные файлы
	fmt.Printf("\n✅ Выбрано файлов: %d\n\n", len(files))
	for i, file := range files {
		fmt.Printf("%d. %s\n", i+1, filepath.Base(file))
	}

	
	reader := bufio.NewReader(os.Stdin)
	
	var selectedFiles []string
	selectedFiles = files
	
	if len(selectedFiles) == 0 {
		fmt.Println("❌ Файлы не выбраны")
		fmt.Println("\nНажмите Enter для выхода...")
		fmt.Scanln()
		return
	}
	
	// Запрашиваем имя выходного файла
	fmt.Print("\n💾 Введите имя выходного файла (по умолчанию merged_result.xlsx): ")
	outputName, _ := reader.ReadString('\n')
	outputName = strings.TrimSpace(outputName)
	if outputName == "" {
		outputName = "merged_result.xlsx"
	}
	
	// Добавляем расширение .xlsx если его нет
	if !strings.HasSuffix(strings.ToLower(outputName), ".xlsx") &&
		!strings.HasSuffix(strings.ToLower(outputName), ".xls") {
		outputName += ".xlsx"
	}
	
	// Выполняем объединение
	err = mergeFiles(selectedFiles, outputName)
	if err != nil {
		fmt.Printf("\n❌ Ошибка: %v\n", err)
	} else {
		exec.Command("explorer", "/select,", outputName).Start()
	}
	
	fmt.Println("\nНажмите Enter для выхода...")
	fmt.Scanln()
}

// parseIndices парсит ввод пользователя (например "1,3,5" или "1-5" или "1 3 5")
func parseIndices(input string, maxCount int) []int {
	// Заменяем разделители на пробелы
	replacer := strings.NewReplacer(",", " ", ";", " ")
	input = replacer.Replace(input)
	
	parts := strings.Fields(input)
	indicesMap := make(map[int]bool)
	
	for _, part := range parts {
		if strings.Contains(part, "-") {
			// Диапазон
			rangeParts := strings.Split(part, "-")
			if len(rangeParts) == 2 {
				var start, end int
				fmt.Sscanf(rangeParts[0], "%d", &start)
				fmt.Sscanf(rangeParts[1], "%d", &end)
				
				if start < 1 {
					start = 1
				}
				if end > maxCount {
					end = maxCount
				}
				
				for i := start; i <= end; i++ {
					indicesMap[i] = true
				}
			}
		} else {
			// Одиночный номер
			var num int
			fmt.Sscanf(part, "%d", &num)
			if num >= 1 && num <= maxCount {
				indicesMap[num] = true
			}
		}
	}
	
	// Преобразуем map в slice
	result := make([]int, 0, len(indicesMap))
	for i := range indicesMap {
		result = append(result, i)
	}
	
	// Сортируем
	for i := 0; i < len(result)-1; i++ {
		for j := i + 1; j < len(result); j++ {
			if result[i] > result[j] {
				result[i], result[j] = result[j], result[i]
			}
		}
	}
	
	return result
}