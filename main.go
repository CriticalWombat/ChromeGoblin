package main

import (
	"database/sql"
	"log"
	"os"
	"slices"

	_ "modernc.org/sqlite"
)

// Return list of user profiles found on a windows machine
func fetch_files() []string {
	entries, err := os.ReadDir("C:\\Users")
	if err != nil {
		log.Fatal(err)
	}

	var DSN []string

	skipped_users := []string{"Administrator", "All Users", "Default", "Default User", "Public", "desktop.ini"}

	for _, u := range entries {
		if slices.Contains(skipped_users, u.Name()) {
			continue
		}
		dbpath := "C:\\Users\\" + u.Name() + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\History.bak"
		DSN = append(DSN, dbpath)
	}
	return DSN
}

// Initialize sqlite and return a db connection
func initDatabase(DSN string) (*sql.DB, error) {
	db, err := sql.Open("sqlite", DSN)
	if err != nil {
		return nil, err
	}
	return db, nil
}

func main() {
	fp := fetch_files()

	//static assignment of my userprofile for testing purposes
	DSN := fp[2]

	db, err := initDatabase(DSN)
	if err != nil {
		log.Fatalf("Failed to open DB: %v", err)
	}
	defer db.Close()

	if err := db.Ping(); err != nil {
		log.Fatalf("Failed to ping DB: %v", err)
	}

	log.Println("Database connection successful.")
}
