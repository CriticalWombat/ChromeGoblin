package main

import (
	"database/sql"
	"errors"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"slices"
	"time"

	"github.com/xuri/excelize/v2"
	_ "modernc.org/sqlite"
)

type HistoryFile struct {
	Browser string
	User    string
	Path    string
}

func copyToTemp(srcPath string) (string, error) {
	input, err := os.ReadFile(srcPath)
	if err != nil {
		return "", fmt.Errorf("failed to read source DB: %w", err)
	}
	dstPath := filepath.Join(os.TempDir(), filepath.Base(srcPath)+"_copy.db")
	err = os.WriteFile(dstPath, input, 0644)
	if err != nil {
		return "", fmt.Errorf("failed to write temp DB: %w", err)
	}
	return dstPath, nil
}

func webkitToTime(webkitTimestamp int64) time.Time {
	const webkitEpoch = 11644473600000000
	unixMicro := webkitTimestamp - webkitEpoch
	return time.UnixMicro(unixMicro)
}

func fileExists(fp string) bool {
	if _, err := os.Stat(fp); errors.Is(err, os.ErrNotExist) {
		return false
	}
	return true
}

func fetchAllHistories() []HistoryFile {
	entries, err := os.ReadDir("C:\\Users")
	if err != nil {
		log.Fatalf("Error reading profiles: %v", err)
	}

	skip := []string{"Administrator", "All Users", "Default", "Public", "desktop.ini"}
	var files []HistoryFile

	for _, u := range entries {
		name := u.Name()
		if slices.Contains(skip, name) {
			continue
		}
		base := filepath.Join("C:\\Users", name)

		// Chrome
		chrome := filepath.Join(base, "AppData", "Local", "Google", "Chrome", "User Data", "Default", "History")
		if fileExists(chrome) {
			files = append(files, HistoryFile{"Chrome", name, chrome})
		}

		// Edge
		edge := filepath.Join(base, "AppData", "Local", "Microsoft", "Edge", "User Data", "Default", "History")
		if fileExists(edge) {
			files = append(files, HistoryFile{"Edge", name, edge})
		}

		// Firefox
		ffProfileDir := filepath.Join(base, "AppData", "Roaming", "Mozilla", "Firefox", "Profiles")
		if entries, err := os.ReadDir(ffProfileDir); err == nil {
			for _, p := range entries {
				path := filepath.Join(ffProfileDir, p.Name(), "places.sqlite")
				if fileExists(path) {
					files = append(files, HistoryFile{"Firefox", name, path})
				}
			}
		}

		// Brave
		brave := filepath.Join(base, "AppData", "Local", "BraveSoftware", "Brave-Browser", "User Data", "Default", "History")
		if fileExists(brave) {
			files = append(files, HistoryFile{"Brave", name, brave})
		}
	}
	return files
}

func initDatabase(DSN string) (*sql.DB, error) {
	db, err := sql.Open("sqlite", DSN)
	if err != nil {
		return nil, err
	}
	return db, nil
}

func exportRowsToSheet(f *excelize.File, sheetName string, rows *sql.Rows) {
	f.NewSheet(sheetName)
	f.SetCellValue(sheetName, "A1", "Title")
	f.SetCellValue(sheetName, "B1", "URL")
	f.SetCellValue(sheetName, "C1", "Visit Count")
	f.SetCellValue(sheetName, "D1", "Last Visit")

	rowIndex := 2
	for rows.Next() {
		var url, title string
		var visitCount int
		var lastVisitRaw int64

		if err := rows.Scan(&url, &title, &visitCount, &lastVisitRaw); err != nil {
			log.Printf("Scan error: %v", err)
			continue
		}

		var lastVisitStr string
		if lastVisitRaw == 0 {
			lastVisitStr = "MALFORMED"
		} else {
			lastVisit := webkitToTime(lastVisitRaw)
			lastVisitStr = lastVisit.Format(time.RFC3339)
		}

		f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), title)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), url)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowIndex), visitCount)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowIndex), lastVisitStr)
		rowIndex++
	}
}

func main() {
	histories := fetchAllHistories()
	if len(histories) == 0 {
		log.Fatal("No browser history files found.")
	}

	// Collect unique users
	userSet := make(map[string]struct{})
	for _, h := range histories {
		userSet[h.User] = struct{}{}
	}
	users := make([]string, 0, len(userSet))
	for u := range userSet {
		users = append(users, u)
	}

	fmt.Println("Select a user to analyze:")
	for i, u := range users {
		fmt.Printf("  %d) %s\n", i+1, u)
	}
	fmt.Printf("  %d) ALL\n", len(users)+1)

	var userChoice int
	fmt.Print("Enter choice: ")
	_, err := fmt.Scan(&userChoice)
	if err != nil || userChoice < 1 || userChoice > len(users)+1 {
		log.Fatal("Invalid selection")
	}

	var selected []HistoryFile
	if userChoice == len(users)+1 {
		selected = histories
	} else {
		chosenUser := users[userChoice-1]
		fmt.Printf("Select a browser for user '%s':\n", chosenUser)
		browsers := []HistoryFile{}
		for _, h := range histories {
			if h.User == chosenUser {
				browsers = append(browsers, h)
			}
		}
		for i, b := range browsers {
			fmt.Printf("  %d) %s\n", i+1, b.Browser)
		}
		fmt.Printf("  %d) ALL\n", len(browsers)+1)

		var browserChoice int
		fmt.Print("Enter browser choice: ")
		_, err := fmt.Scan(&browserChoice)
		if err != nil || browserChoice < 1 || browserChoice > len(browsers)+1 {
			log.Fatal("Invalid browser selection")
		}
		if browserChoice == len(browsers)+1 {
			selected = browsers
		} else {
			selected = append(selected, browsers[browserChoice-1])
		}
	}

	excelFile := excelize.NewFile()
	first := true

	for _, h := range selected {
		copiedDBPath, err := copyToTemp(h.Path)
		if err != nil {
			log.Printf("Failed to copy %s DB: %v", h.Browser, err)
			continue
		}
		defer os.Remove(copiedDBPath)

		db, err := initDatabase(copiedDBPath)
		if err != nil {
			log.Printf("Failed to open DB for %s: %v", h.Browser, err)
			continue
		}
		defer db.Close()

		query := "SELECT url, title, visit_count, last_visit_time FROM urls"
		if h.Browser == "Firefox" {
			query = "SELECT url, title, visit_count, last_visit_date FROM moz_places"
		}

		rows, err := db.Query(query)
		if err != nil {
			log.Printf("Query failed for %s: %v", h.Browser, err)
			continue
		}
		defer rows.Close()

		sheetName := fmt.Sprintf("%s_%s", h.Browser, h.User)
		if first {
			excelFile.SetSheetName("Sheet1", sheetName)
			first = false
		} else {
			excelFile.NewSheet(sheetName)
		}
		exportRowsToSheet(excelFile, sheetName, rows)
	}

	outPath := filepath.Join(os.TempDir(), "browser_history_export.xlsx")
	if err := excelFile.SaveAs(outPath); err != nil {
		log.Fatalf("Excel save failed: %v", err)
	}
	fmt.Println("Excel written to:", outPath)
}
