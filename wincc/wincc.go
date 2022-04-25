package wincc

import (
	"database/sql"

	"github.com/mattn/go-adodb"
	_ "github.com/mattn/go-adodb"
)

type identity struct {
	provider,
	dataSource,
	catalog string
}

func (i identity) GiveIdentity(P, D, C string) {
	i.provider = P
	i.dataSource = D
	i.catalog = C
}

func (i identity) CreateConnection() interface{} {
	db_type := "adodb_with_cursorlocation"
	conn_string := i.provider + i.dataSource + i.catalog

	sql.Register(db_type, &adodb.AdodbDriver{
		CursorLocation: 3,
	})

	db, err := sql.Open(db_type, conn_string)
	if err != nil {
		return err
	}
	return db
}

func ExecuteQuery(q string, db *sql.DB) interface{} {
	row, err := db.Query(q)
	if err != nil {
		return err
	}
	return row
}
