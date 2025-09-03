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

	"github.com/xuri/excelize/v2" // Excel file generation
	_ "modernc.org/sqlite"        // SQLite driver
)

// HistoryFile represents a browser history file path associated with a user.
type HistoryFile struct {
	Browser string
	User    string
	Path    string
}

// copyToTemp duplicates the specified SQLite file to a temporary location
// to avoid potential locks and corruption.
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

// webkitToTime converts a WebKit timestamp (microseconds since Windows epoch)
// to Go's time.Time.
func webkitToTime(webkitTimestamp int64) time.Time {
	const webkitEpoch = 11644473600000000
	unixMicro := webkitTimestamp - webkitEpoch
	return time.UnixMicro(unixMicro)
}

// fileExists checks if the given file path exists.
func fileExists(fp string) bool {
	if _, err := os.Stat(fp); errors.Is(err, os.ErrNotExist) {
		return false
	}
	return true
}

// fetchAllHistories searches user directories for browser history files and
// validates that they contain URL visit data.
func fetchAllHistories() []HistoryFile {
	userBase := "C:\\Users"
	entries, err := os.ReadDir(userBase)
	if err != nil {
		log.Fatalf("Failed to read user directory: %v", err)
	}

	// Skip system and irrelevant users
	skip := []string{"Administrator", "All Users", "Default", "Default User", "Public", "desktop.ini"}
	var validHistories []HistoryFile

	// For each user profile
	for _, entry := range entries {
		user := entry.Name()
		if slices.Contains(skip, user) {
			continue
		}

		base := filepath.Join(userBase, user)

		// Define browsers and corresponding history DB paths
		browserPaths := []struct {
			Browser string
			Path    string
			Query   string
		}{
			{
				"Chrome",
				filepath.Join(base, "AppData", "Local", "Google", "Chrome", "User Data", "Default", "History"),
				"SELECT COUNT(1) FROM urls",
			},
			{
				"Edge",
				filepath.Join(base, "AppData", "Local", "Microsoft", "Edge", "User Data", "Default", "History"),
				"SELECT COUNT(1) FROM urls",
			},
			{
				"Brave",
				filepath.Join(base, "AppData", "Local", "BraveSoftware", "Brave-Browser", "User Data", "Default", "History"),
				"SELECT COUNT(1) FROM urls",
			},
		}

		// Process Chrome-like browsers
		for _, b := range browserPaths {
			if fileExists(b.Path) {
				copied, err := copyToTemp(b.Path)
				if err != nil {
					log.Printf("[%s][%s] Copy failed: %v", b.Browser, user, err)
					continue
				}
				if hasRows(copied, b.Query) {
					validHistories = append(validHistories, HistoryFile{b.Browser, user, b.Path})
				} else {
					log.Printf("[%s][%s] Query returned 0 rows", b.Browser, user)
				}
				os.Remove(copied) // cleanup
			} else {
				log.Printf("[%s][%s] File not found", b.Browser, user)
			}
		}

		// Process Firefox separately (profile subdirs)
		ffProfileDir := filepath.Join(base, "AppData", "Roaming", "Mozilla", "Firefox", "Profiles")
		if ffEntries, err := os.ReadDir(ffProfileDir); err == nil {
			for _, prof := range ffEntries {
				ffPath := filepath.Join(ffProfileDir, prof.Name(), "places.sqlite")
				if fileExists(ffPath) {
					copied, err := copyToTemp(ffPath)
					if err != nil {
						log.Printf("[Firefox][%s] Copy failed: %v", user, err)
						continue
					}
					if hasRows(copied, "SELECT COUNT(1) FROM moz_places") {
						validHistories = append(validHistories, HistoryFile{"Firefox", user, ffPath})
					} else {
						log.Printf("[Firefox][%s] Query returned 0 rows", user)
					}
					os.Remove(copied)
				} else {
					log.Printf("[Firefox][%s] places.sqlite not found", user)
				}
			}
		} else {
			log.Printf("[Firefox][%s] Profiles dir not found: %v", user, err)
		}
	}

	return validHistories
}

// hasRows checks if the given SQLite database file has at least one row
// for the specified query (typically SELECT COUNT).
func hasRows(path string, query string) bool {
	db, err := sql.Open("sqlite", path)
	if err != nil {
		return false
	}
	defer db.Close()

	row := db.QueryRow(query)
	var count int
	if err := row.Scan(&count); err != nil {
		return false
	}
	return count > 0
}

// initDatabase initializes a connection to an SQLite database.
func initDatabase(DSN string) (*sql.DB, error) {
	db, err := sql.Open("sqlite", DSN)
	if err != nil {
		return nil, err
	}
	return db, nil
}

// exportRowsToSheet writes query results from a browser history database
// to a specific Excel sheet.
func exportRowsToSheet(f *excelize.File, sheetName string, rows *sql.Rows) {
	f.NewSheet(sheetName)
	f.SetCellValue(sheetName, "A1", "Title")
	f.SetCellValue(sheetName, "B1", "URL")
	f.SetCellValue(sheetName, "C1", "Visit Count")
	f.SetCellValue(sheetName, "D1", "Last Visit")

	rowIndex := 2
	for rows.Next() {
		var urlNS, titleNS sql.NullString
		var visitCount int
		var lastVisitRaw sql.NullInt64

		if err := rows.Scan(&urlNS, &titleNS, &visitCount, &lastVisitRaw); err != nil {
			log.Printf("Scan error: %v", err)
			continue
		}

		url := ""
		if urlNS.Valid {
			url = urlNS.String
		}
		title := ""
		if titleNS.Valid {
			title = titleNS.String
		}

		var lastVisitStr string
		if !lastVisitRaw.Valid || lastVisitRaw.Int64 == 0 {
			lastVisitStr = "MALFORMED"
		} else {
			lastVisit := webkitToTime(lastVisitRaw.Int64)
			lastVisitStr = lastVisit.Format(time.RFC3339)
		}

		f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), title)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), url)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowIndex), visitCount)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowIndex), lastVisitStr)
		rowIndex++
	}
}

// main is the program's entry point. It orchestrates the discovery,
// selection, extraction, and export of browser history data.
func main() {
	histories := fetchAllHistories()
	if len(histories) == 0 {
		log.Fatal("No browser history files found.")
	}

	// Collect distinct users found in history records
	userSet := make(map[string]struct{})
	for _, h := range histories {
		userSet[h.User] = struct{}{}
	}
	users := make([]string, 0, len(userSet))
	for u := range userSet {
		users = append(users, u)
	}

	// Prompt user to select a specific user or "ALL"
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
		selected = histories // All users
	} else {
		chosenUser := users[userChoice-1]

		// Prompt for browser selection for selected user
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
			selected = browsers // All browsers for selected user
		} else {
			selected = append(selected, browsers[browserChoice-1])
		}
	}

	// Create new Excel file for results
	excelFile := excelize.NewFile()
	first := true

	// Process each selected history database
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

		// Select appropriate query based on browser
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

	// Write the final Excel file to the temp directory
	outPath := filepath.Join(os.TempDir(), "browser_history_export.xlsx")
	if err := excelFile.SaveAs(outPath); err != nil {
		log.Fatalf("Excel save failed: %v", err)
	}
	fmt.Println("Excel written to:", outPath)
}
