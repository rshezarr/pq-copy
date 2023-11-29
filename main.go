package main

import (
	"database/sql"
	"fmt"
	"github.com/tealeg/xlsx"
	"log"
)

const (
	host     = ""
	port     = ""
	user     = ""
	password = ""
	dbname   = ""
)

func main() {
	db := ConnDB()

	file, err := xlsx.OpenFile("filename.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	copyCommand := `COPY 'table' FROM STDIN;`

	tx, err := db.Begin()
	if err != nil {
		log.Fatal(err)
	}
	defer tx.Rollback()

	copyStmt, err := tx.Prepare(copyCommand)
	if err != nil {
		log.Fatal(err)
	}

	for _, sheet := range file.Sheets {
		for _, row := range sheet.Rows {
			var values []interface{}
			for _, cell := range row.Cells {
				values = append(values, cell.String())
			}
			_, err = copyStmt.Exec(values...)
			if err != nil {
				log.Fatal(err)
			}
		}
	}

	_, err = copyStmt.Exec()
	if err != nil {
		log.Fatal(err)
	}

	// Commit the transaction
	err = tx.Commit()
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Data copied successfully.")

	db.Close()
}

func ConnDB() *sql.DB {
	connStr := fmt.Sprintf("host=%s port=%s user=%s password=%s dbname=%s sslmode=disable", host, port, user, password, dbname)
	db, err := sql.Open("postgres", connStr)
	if err != nil {
		log.Fatal("Error connecting to the database:", err)
	}
	return db
}
