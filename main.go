package main

import (
	"database/sql"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/BurntSushi/toml"
	"github.com/jinzhu/gorm"
)

var allFlag bool

type DBConfig struct {
	Host     string `toml:"host"`
	User     string `toml:"user"`
	Password string `toml:"password"`
	Port     string `toml:"port"`
	Name     string `toml:"name"`
}

type Config struct {
	DB DBConfig `toml:"database"`
}

type ColumnInfo struct {
	TableName       string         `gorm:"column:tablename"`
	NumCount        int            `gorm:"column:numcount"`
	ColumnName      string         `gorm:"column:columnname"`
	PKey            sql.NullString `gorm:"column:primarykey"`
	ColumnType      string         `gorm:"column:columntype"`
	Length          string         `gorm:"column:length"`
	NumberPrecision string         `gorm:"column:numberprecision"`
	DefaultValue    sql.NullString `gorm:"column:defaultvalue"`
	NotNull         sql.NullString `gorm:"column:notnull"`
}

type TableName struct {
	Name string `gorm:"column:name"`
}

func GormConnect() *gorm.DB {
	var config Config
	_, err := toml.DecodeFile("./dbconfig.toml", &config)
	if err != nil {
		fmt.Println(err)
	}
	gormConn := fmt.Sprintf(
		"server=%s;user id=%s;password=%s;port=%v;database=%s",
		config.DB.Host,
		config.DB.User,
		config.DB.Password,
		config.DB.Port,
		config.DB.Name,
	)
	db, err := gorm.Open("mssql", gormConn)
	if err != nil {
		log.Fatalf("Open Connection failed  of err : %s\n", err.Error())
	} else {
		fmt.Printf("DB Connected\n")
	}
	return db
}

func GetStruct(tableName string) (columns []ColumnInfo, err error) {
	db := GormConnect()
	defer db.Close()
	text, err := ioutil.ReadFile("./sql/query.sql")
	if err != nil {
		fmt.Println(err)
	}
	sqlText := strings.ReplaceAll(string(text), "#{tableName}", "'"+tableName+"'")
	sqlText = string(sqlText)
	db.Raw(sqlText).Scan(&columns)
	return columns, err
}

func main() {
	flag.BoolVar(&allFlag, "allFlag", false, "Help Message")
	flag.Parse()
	args := flag.Args()
	// if flag.NArg() != RightArgs {
	// 	log.Fatalf("error argument: want to argument %v", RightArgs)
	// }
	tableNameList := make([]TableName, 0)
	db := GormConnect()
	defer db.Close()
	if allFlag {
		text, err := ioutil.ReadFile("./sql/tableList.sql")
		if err != nil {
			fmt.Println(err)
		}
		sqlText := string(text)
		db.Raw(sqlText).Scan(&tableNameList)
		//	fmt.Printf("%+v", tableNameList)
	} else {
		argsTable := args[1:]
		tableNameList2 := make([]TableName, len(argsTable))
		for i := range argsTable {
			tableNameList2[i].Name = argsTable[i]
		}
		tableNameList = append(tableNameList, tableNameList2...)
	}

	header := []string{
		"No.", "Name", "Type", "Length", "Precision", "Not Null", "Primary Key", "Default Value", "Content", "Description",
	}
	//	fmt.Printf("args=%s,args[0]=%s",args,args[0])
	for _, val := range tableNameList {
		switch args[0] {
		case "excel":
			excelOutput(val.Name, header)
			fmt.Printf("ExcelOutput : %s\n", val.Name)
		case "markdown":
			markdownOutput(val.Name, header)
			fmt.Printf("MarkdownOutput : %s\n", val.Name)
		default:
			fmt.Print("Input args Error!!")
		}
	}
}

func markdownOutput(tableName string, header []string) {

	file, err := os.Create("./out/TableDoc_" + tableName + ".md")
	if err != nil {
		fmt.Println(err)
	}
	_, err = file.WriteString("| " + strings.Join(header, " | ") + " |\n")
	var separator string
	separator = "| "
	for i := 0; i < len(header); i++ {
		separator = separator + "--- |"
		if i == len(header)-1 {
			separator = separator + "\n"
		}
	}
	_, err = file.WriteString(separator)
	columns, _ := GetStruct(tableName)
	for _, val := range columns {
		row := []string{
			strconv.Itoa(val.NumCount),
			val.ColumnName,
			val.ColumnType,
			val.Length,
			val.NumberPrecision,
			val.NotNull.String,
			val.PKey.String,
			val.DefaultValue.String,
			"",
			"",
		}
		_, err = file.WriteString("| " + strings.Join(row, " | ") + " |\n")
	}

	defer file.Close()
}

func excelOutput(tableName string, header []string) {
	columns, _ := GetStruct(tableName)
	//	fmt.Println("print data :")
	var columnsmap []map[string]interface{}
	for i := 0; i < len(columns); i++ {
		columndic := make(map[string]interface{})
		// columndic["tablename"] = columns[i].TableName
		columndic["numcount"] = columns[i].NumCount
		columndic["columnname"] = columns[i].ColumnName
		columndic["primarykey"] = columns[i].PKey.String
		columndic["columntype"] = columns[i].ColumnType
		columndic["length"] = columns[i].Length
		columndic["numberprecision"] = columns[i].NumberPrecision
		columndic["defaultvalue"] = columns[i].DefaultValue.String
		columndic["notnull"] = columns[i].NotNull.String
		columnsmap = append(columnsmap, columndic)
	}
	// fmt.Println(columns)
	// fmt.Println(columnsmap)

	//Create New Excel
	f := excelize.NewFile()
	index := f.NewSheet("Sheet1")

	for colNum, v := range header {
		sheetPosition := Div(colNum+1) + "1"
		// fmt.Print(sheetPosition + "\n")
		f.SetCellValue("Sheet1", sheetPosition, v) //nolint
	}

	//Set Cell Style : 1st line
	style, err := f.NewStyle(`{
		"alignment":
		{"horizontal": "center","Vertical": "center"},
		"font":{"bold":true},
		"border": [
			{"type": "left","color": "000000","style" : 2},
			{"type": "top","color": "000000","style" : 2},
			{"type": "bottom","color": "000000","style" : 2},
			{"type": "right","color": "000000","style" : 2}],
		"fill":{"type":"pattern","color":["#CCFFFF"],"pattern":1}
		}`)
	if err != nil {
		fmt.Println(err)
	}
	err = f.SetCellStyle("Sheet1", "A1", "J1", style)
	if err != nil {
		fmt.Println(err)
	}

	//Set Cell Value By line then traverse columns
	for lineNum, dic := range columnsmap {
		colNum := 0
		// fmt.Println(lineNum, dic)
		for k := range dic { //nolint
			colNum++
			Anewposition := "A" + strconv.Itoa(lineNum+2)
			Bnewposition := "B" + strconv.Itoa(lineNum+2)
			Cnewposition := "C" + strconv.Itoa(lineNum+2)
			Dnewposition := "D" + strconv.Itoa(lineNum+2)
			Enewposition := "E" + strconv.Itoa(lineNum+2)
			Fnewposition := "F" + strconv.Itoa(lineNum+2)
			Gnewposition := "G" + strconv.Itoa(lineNum+2)
			Hnewposition := "H" + strconv.Itoa(lineNum+2)
			// Inewposition := "I" + strconv.Itoa(lineNum+2)
			switch k {
			//				fmt.Println("name :" ,dic["name"])
			// case "tablename":
			// 	f.SetCellValue("Sheet1", Anewposition, dic["tablename"]) //nolint
			case "numcount":
				f.SetCellValue("Sheet1", Anewposition, dic["numcount"]) //nolint
			case "columnname":
				f.SetCellValue("Sheet1", Bnewposition, dic["columnname"]) //nolint
			case "columntype":
				f.SetCellValue("Sheet1", Cnewposition, dic["columntype"]) //nolint
			case "length":
				f.SetCellValue("Sheet1", Dnewposition, dic["length"]) //nolint
			case "numberprecision":
				f.SetCellValue("Sheet1", Enewposition, dic["numberprecision"]) //nolint
			case "notnull":
				f.SetCellValue("Sheet1", Fnewposition, dic["notnull"]) //nolint
			case "primarykey":
				f.SetCellValue("Sheet1", Gnewposition, dic["primarykey"]) //nolint
			case "defaultvalue":
				f.SetCellValue("Sheet1", Hnewposition, dic["defaultvalue"]) //nolint

			}
		}

	}

	// width := getFitColWidth

	//set active sheet of the workbook
	f.SetActiveSheet(index)
	//save xlsx file by the given path
	if err := f.SaveAs("./out/TableDoc_" + tableName + ".xlsx"); err != nil {
		println(err.Error())
	} else {
		// fmt.Print("Excelized already , now go check : TableDoc_.xlsx ")
	}
}

func Div(Num int) string {
	var (
		Str  string = ""
		k    int
		temp []int
	)
	Slice := []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	if Num > 26 {
		for {
			k = Num % 26
			if k == 0 {
				temp = append(temp, 26)
				k = 26
			} else {
				temp = append(temp, k)
			}
			Num = (Num - k) / 26
			if Num <= 26 {
				temp = append(temp, Num)
				break
			}
		}
	} else {
		return Slice[Num]
	}
	for _, value := range temp {
		Str = Slice[value] + Str
	}
	return Str
}

// //Auto-fitting columns width by searching the string size
// func searchCount(src string) int {
// 	letters := "abcdefghijklmnopqrstuvwxyz"
// 	letters = letters + strings.ToUpper(letters)
// 	nums := "0123456789"
// 	chars := "(/#)"

// 	numCount := 0
// 	letterCount := 0
// 	othersCount := 0
// 	charsCount := 0

// 	for _, i := range src {
// 		switch {
// 		case strings.ContainsRune(letters, i) == true:
// 			letterCount += 1
// 		case strings.ContainsRune(nums, i) == true:
// 			numCount += 1
// 		case strings.ContainsRune(chars, i) == true:
// 			charsCount += 1
// 		default:
// 			othersCount += 1
// 		}
// 	}
// 	return numCount*1 + letterCount*1 + charsCount*1 + othersCount*2
// }

// //Algorithm : compute columns width
// func getFitColWidth(sheet string, columns []*ColumnInfo) map[string]float64{
// 	var rate float64 = 1.2
// 	maxFix :=make(map[string]float64)
// 	for _, value := range columns {
// 		reg := regexp.MustCompile(`[[:upper:]]+`)
// 		lettersStrs :=reg.FindAllString(value., -1)
// 		col := ""
// 		if len(lettersStrs) > 0{
// 			col = lettersStrs[0]
// 		}
// 		split := strings.Split(value.Value, "\r\n")
// 		maxLength :=0
// 		for _, s := range split{
// 			length :=searchCount(s)
// 			if maxLength <length{
// 				maxLength = length
// 			}
// 		}
// 		width := float64(maxLength) *rate
// 		if vv, err :=maxFix[col]; err{
// 			if vv< width {
// 				maxFix[col] = width
// 			}
// 		} else {
// 			maxFix[col] = width
// 		}
// 	}
// 	return maxFix
// }
