package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
	"gopkg.in/yaml.v2"
)

// Config представляет структуру конфигурационного файла
type Config struct {
	Table1 struct {
		Name   string `yaml:"name"`
		Sheet  string `yaml:"sheet"`
		Column string `yaml:"column"`
		Header string `yaml:"header"`
	} `yaml:"table1"`
	Table2 struct {
		Name   string `yaml:"name"`
		Sheet  string `yaml:"sheet"`
		Column string `yaml:"column"`
		Header string `yaml:"header"`
	} `yaml:"table2"`
	Output struct {
		Name string `yaml:"name"`
	} `yaml:"output"`
}

func main() {
	log.Println("Начало выполнения программы")

	// Открываем и читаем конфигурационный файл
	log.Println("Открытие конфигурационного файла")
	configFile, err := os.Open("config.yaml")
	if err != nil {
		log.Fatalf("Не удалось открыть config.yaml: %v", err)
	}
	defer configFile.Close()

	var config Config
	decoder := yaml.NewDecoder(configFile)
	if err := decoder.Decode(&config); err != nil {
		log.Fatalf("Ошибка декодирования config.yaml: %v", err)
	}
	log.Println("Конфигурационный файл успешно прочитан")

	// Открываем Excel-файлы
	log.Printf("Открытие файла: %s", config.Table1.Name)
	file1, err := excelize.OpenFile(config.Table1.Name)
	if err != nil {
		log.Fatalf("Не удалось открыть %s: %v", config.Table1.Name, err)
	}
	defer file1.Close()
	log.Printf("Файл %s успешно открыт", config.Table1.Name)

	log.Printf("Открытие файла: %s", config.Table2.Name)
	file2, err := excelize.OpenFile(config.Table2.Name)
	if err != nil {
		log.Fatalf("Не удалось открыть %s: %v", config.Table2.Name, err)
	}
	defer file2.Close()
	log.Printf("Файл %s успешно открыт", config.Table2.Name)

	// Получаем индексы столбцов для сравнения
	log.Println("Получение индексов столбцов для сравнения")
	columnIndex1, err := excelize.ColumnNameToNumber(config.Table1.Column)
	if err != nil {
		log.Fatalf("Ошибка преобразования столбца %s: %v", config.Table1.Column, err)
	}
	columnIndex1-- // Преобразуем к 0-индексации

	columnIndex2, err := excelize.ColumnNameToNumber(config.Table2.Column)
	if err != nil {
		log.Fatalf("Ошибка преобразования столбца %s: %v", config.Table2.Column, err)
	}
	columnIndex2-- // Преобразуем к 0-индексации

	// Читаем строки из таблиц
	log.Printf("Чтение строк из файла: %s", config.Table1.Name)
	rows1, err := file1.GetRows(config.Table1.Sheet)
	if err != nil {
		log.Fatalf("Ошибка чтения строк из %s: %v", config.Table1.Name, err)
	}

	log.Printf("Чтение строк из файла: %s", config.Table2.Name)
	rows2, err := file2.GetRows(config.Table2.Sheet)
	if err != nil {
		log.Fatalf("Ошибка чтения строк из %s: %v", config.Table2.Name, err)
	}

	// Проверяем заголовки
	log.Println("Проверка заголовков столбцов")
	if len(rows1) == 0 || len(rows2) == 0 {
		log.Fatalf("Одна из таблиц пуста")
	}
	if rows1[0][columnIndex1] != config.Table1.Header {
		log.Fatalf("Заголовок столбца в первой таблице не совпадает с ожидаемым: %s", config.Table1.Header)
	}
	if rows2[0][columnIndex2] != config.Table2.Header {
		log.Fatalf("Заголовок столбца во второй таблице не совпадает с ожидаемым: %s", config.Table2.Header)
	}
	log.Println("Заголовки столбцов успешно проверены")

	// Создаем CSV-файл
	log.Printf("Создание выходного файла: %s", config.Output.Name)
	outputFile, err := os.Create(config.Output.Name)
	if err != nil {
		log.Fatalf("Ошибка создания выходного файла: %v", err)
	}
	defer outputFile.Close()

	// Записываем BOM для корректного отображения в Excel
	_, err = outputFile.WriteString("\xEF\xBB\xBF")
	if err != nil {
		log.Fatalf("Ошибка записи BOM в выходной файл: %v", err)
	}

	writer := csv.NewWriter(outputFile)
	defer writer.Flush()

	// Сравниваем строки
	log.Println("Начало сравнения строк")
	for _, row1 := range rows1[1:] { // Пропускаем заголовок
		if len(row1) <= columnIndex1 {
			continue
		}
		value1 := row1[columnIndex1]

		// Проверяем, есть ли совпадение во второй таблице
		found := false
		for _, row2 := range rows2[1:] { // Пропускаем заголовок
			if len(row2) <= columnIndex2 {
				continue
			}
			value2 := row2[columnIndex2]

			if value1 == value2 {
				found = true
				break
			}
		}

		// Если совпадение не найдено, записываем строку в CSV-файл
		if !found {
			if err := writer.Write(row1); err != nil {
				log.Fatalf("Ошибка записи строки в CSV: %v", err)
			}
		}
	}
	log.Println("Сравнение строк завершено")

	log.Println("Программа успешно завершена")
	fmt.Println("Сравнение завершено. Результаты записаны в", config.Output.Name)
}
