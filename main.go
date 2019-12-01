package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/fsnotify/fsnotify"
	"github.com/unidoc/unioffice/document"
)

var watcher *fsnotify.Watcher
var openedFile string

func main() {
	var err error
	watcher, err = fsnotify.NewWatcher()
	if err != nil {
		fmt.Println("ERROR", err)
	}
	defer watcher.Close()

	/*go func() {
		if err := filepath.Walk("//172.26.142.161/обмен", watchDir); err != nil {
			fmt.Println("ERROR", err)
		}
	}()*/

	done := make(chan bool)
	if err := watcher.Add(`E:\FWatcher`); err != nil {
		fmt.Println("ERROR", err)
	}
	//тикер на проверку чтения файла раз в 1 сек
	ticker := time.NewTicker(time.Second)
	defer ticker.Stop()
	//проверяем можно ли открыть файл
	go func() {
		for {
			select {
			case <-ticker.C:
				if openedFile != "" {
					XlWatch(openedFile)
				}
			}
		}
	}()
	//Обрабатываем события
	go func() {
		for {
			select {
			case event, ok := <-watcher.Events:
				if !ok {
					return
				}
				fmt.Printf("EVENT! %#v\n", event)
				if event.Op&fsnotify.Create == fsnotify.Create {
					fileName := filepath.Base(event.Name)
					if strings.Contains(fileName, "Григорьев") &&
						fileName[strings.LastIndex(fileName, "."):] == ".xlsx" {
						if fileName[:2] != "~$" { //Исключаем открытые файлы из обработки
							time.Sleep(2 * time.Second) // пауза, чтобы успел скопироваться
							XlWatch(event.Name)
						}
					}
					if strings.Contains(fileName, "Отч") &&
						fileName[strings.LastIndex(fileName, "."):] == ".docx" {
						time.Sleep(2 * time.Second) // пауза, чтобы успел скопироваться
						DocWatch(event.Name)
					}
				}

			case err, ok := <-watcher.Errors:
				if !ok {
					return
				}
				fmt.Println("ERORR", err)
			}
		}
	}()

	<-done
}

//Рутина открытия и записи в XLSX файл
func XlWatch(fileName string) error {
	var i0, j0 = 1, 1
	//чтение
	xlFile, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Printf("Error Open File: %s", err)
		if openedFile != fileName {
			openedFile = fileName
		}
		return err
	}

	defer func() {
		if r := recover(); r != nil {
			fmt.Println(r)
		}
	}()
	var alph = "ABCDEFGHIJKLMNOPRSTUVWXYZ"

	for _, sheet := range xlFile.GetSheetMap() {
		rows, err := xlFile.GetRows(sheet)
		if err != nil {
			fmt.Println(err)
		}

		for i, row := range rows {
			for j, cell := range row {
				if cell == "Наименование работ" {
					j0 = j
					break
				}
			}
			val, err := xlFile.GetCellValue(sheet, string(alph[j0])+strconv.Itoa(i+1))
			if err != nil {
				fmt.Println(err)
			}
			if val == "Наименование работ" {
				i0 = i + 1
			}
		}
	}

	//запись

	err = xlFile.SetCellStr(xlFile.GetSheetName(1), string(alph[j0])+strconv.Itoa(i0+1), "Смотрел Аниме JoJo")
	if err != nil {
		fmt.Printf("Error SetCell%s", err)
	}
	err = xlFile.SetCellStr(xlFile.GetSheetName(1), string(alph[j0])+strconv.Itoa(i0+5), "Слушал Gachi музыку")
	if err != nil {
		fmt.Printf("Error SetCell%s", err)
	}
	err = xlFile.SetCellStr(xlFile.GetSheetName(1), string(alph[j0])+strconv.Itoa(i0+6), "Молилъся за доброе здравие начальниковъ своих: Алексия, Виталия, Дениса Юрьевича и Дениса Валентиновича")
	if err != nil {
		fmt.Printf("Error SetCell%s", err)
	}
	err = xlFile.Save()
	if err != nil {
		fmt.Printf("Error Save File: %s", err)
		if openedFile != fileName {
			openedFile = fileName
		}
		return err
	}
	openedFile = ""
	return nil
}

func watchDir(path string, fi os.FileInfo, err error) error {

	if fi.Mode().IsDir() {
		return watcher.Add(path)
	}

	return err
}

//Рутина чтения и записи docx файла
func DocWatch(fileName string) {
	doc, err := document.Open(fileName)
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}

	paragraphs := []document.Paragraph{}
	for _, p := range doc.Paragraphs() {
		paragraphs = append(paragraphs, p)
	}

	for _, sdt := range doc.StructuredDocumentTags() {
		for _, p := range sdt.Paragraphs() {
			paragraphs = append(paragraphs, p)
		}
	}

	for _, p := range paragraphs {
		for _, r := range p.Runs() {
			r.ClearContent()
			r.AddText("ANIME")
		}
	}
	doc.SaveToFile(fileName)
}
