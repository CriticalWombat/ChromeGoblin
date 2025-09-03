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

// Copy SQLite history file to avoid file locks
func copyToTemp(srcPath string) (string, error) {
	input, err := os.ReadFile(srcPath)
	if err != nil {
		return "", fmt.Errorf("failed to read source DB: %w", err)
	}

	dstPath := filepath.Join(os.TempDir(), "History_temp_copy.db")
	err = os.WriteFile(dstPath, input, 0644)
	if err != nil {
		return "", fmt.Errorf("failed to write temp DB: %w", err)
	}

	return dstPath, nil
}

// Convert webkit time to time.Time
func webkitToTime(webkitTimestamp int64) time.Time {
	const webkitEpoch = 11644473600000000
	unixMicro := webkitTimestamp - webkitEpoch
	return time.UnixMicro(unixMicro)
}

// Check if file path is valid
func fileExists(fp string) bool {
	if _, err := os.Stat(fp); errors.Is(err, os.ErrNotExist) {
		return false
	}
	return true
}

// Check local C:\Users dir for usernames, Check for chrome history file within each, and return a list of matches.
func fetch_files() map[string]string {
	entries, err := os.ReadDir("C:\\Users")
	if err != nil {
		log.Fatalf("Error reading profiles: %v", err)
	}

	historyPaths := make(map[string]string)
	skippedUsers := []string{"Administrator", "All Users", "Default", "Default User", "Public", "desktop.ini"}

	for _, u := range entries {
		if slices.Contains(skippedUsers, u.Name()) {
			continue
		}
		user := u.Name()
		dbpath := filepath.Join("C:\\Users", user, "AppData\\Local\\Google\\Chrome\\User Data\\Default\\History")
		if fileExists(dbpath) {
			log.Printf("Found Chrome history for user: %s", user)
			historyPaths[user] = dbpath
		}
	}
	return historyPaths
}

// Initialize sqlite and return a db connection
func initDatabase(DSN string) (*sql.DB, error) {
	db, err := sql.Open("sqlite", DSN)
	if err != nil {
		return nil, err
	}
	return db, nil
}

// Export history rows to Excel
func exportToExcel(rows *sql.Rows) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Headers
	f.SetCellValue(sheet, "A1", "Title")
	f.SetCellValue(sheet, "B1", "URL")
	f.SetCellValue(sheet, "C1", "Visit Count")
	f.SetCellValue(sheet, "D1", "Last Visit")

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

		f.SetCellValue(sheet, fmt.Sprintf("A%d", rowIndex), title)
		f.SetCellValue(sheet, fmt.Sprintf("B%d", rowIndex), url)
		f.SetCellValue(sheet, fmt.Sprintf("C%d", rowIndex), visitCount)
		f.SetCellValue(sheet, fmt.Sprintf("D%d", rowIndex), lastVisitStr)

		rowIndex++
	}

	tempFilePath := filepath.Join(os.TempDir(), "chrome_history_export.xlsx")
	if err := f.SaveAs(tempFilePath); err != nil {
		log.Fatalf("Failed to save Excel file: %v", err)
	}

	fmt.Println("Excel file written to:", tempFilePath)
}

func main() {

	historyMap := fetch_files()

	if len(historyMap) == 0 {
		log.Fatal("No Chrome history files found.")
	}

	// Display options
	usernames := make([]string, 0, len(historyMap))
	fmt.Println("Select a user profile to analyze:")
	i := 1
	for user := range historyMap {
		fmt.Printf("  %d) %s\n", i, user)
		usernames = append(usernames, user)
		i++
	}

	// Prompt
	var selection int
	fmt.Print("Enter choice (number): ")
	_, err := fmt.Scan(&selection)
	if err != nil || selection < 1 || selection > len(usernames) {
		log.Fatalf("Invalid selection.")
	}

	selectedUser := usernames[selection-1]
	DSN := historyMap[selectedUser]

	// Copy the DB to avoid file lock issues on the chrome history file
	copiedDBPath, err := copyToTemp(DSN)
	if err != nil {
		log.Fatalf("Failed to copy DB: %v", err)
	}

	// Attempt removal of copied DB file at the end of main()
	defer func() {
		if err := os.Remove(copiedDBPath); err != nil {
			log.Printf("Warning: failed to clean up temp DB: %v", err)
		}
	}()

	// Instantiate DB on copied DB
	db, err := initDatabase(copiedDBPath)
	if err != nil {
		log.Fatalf("Failed to open DB: %v", err)
	}

	// Attempt closing the DB connection at the end of main()
	defer db.Close()

	// Test DB connectivity prior to running a query
	if err := db.Ping(); err != nil {
		log.Fatalf("Failed to ping DB: %v", err)
	}
	log.Println("Database connection successful.")

	// Fetch url table data from SQLite chrome history file
	rows, err := db.Query("SELECT url, title, visit_count, last_visit_time FROM urls")
	if err != nil {
		log.Fatalf("Query failed: %v", err)
	}

	// Close rows at the end of main()
	defer rows.Close()

	// Export url data to excel
	exportToExcel(rows)
}
