package wincc

import (
	"database/sql"
	"time"

	"github.com/mattn/go-adodb"
	_ "github.com/mattn/go-adodb"
)

type identity struct {
	provider,
	dataSource,
	catalog string
}

type data struct {
	valueId   int
	timeStamp *time.Time
	realValue float32
	quality   int
	flags     string
}

func (i *identity) CreateIdentity(p, d, c string) {
	i.provider = p
	i.dataSource = d
	i.catalog = c
}

func (i *identity) CreateConnection() (*sql.DB, error) {
	db_type := "adodb_with_cursorlocation"
	conn_string := i.provider + i.dataSource + i.catalog

	sql.Register(db_type, &adodb.AdodbDriver{
		CursorLocation: 3,
	})

	db, err := sql.Open(db_type, conn_string)
	return db, err
}

func ExecuteQuery(db *sql.DB, output string, startTime, endTime *time.Time) (*sql.Rows, error) {
	q := "TAG:R '" + output + "', '" + startTime.String() + "', '" + endTime.String() + "' "
	row, err := db.Query(q)
	return row, err
}

// func ReadData() []data {

// }
