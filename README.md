# DB-Tool

A database tool for you to ***batch generating of datebase table documentation*** by  `.xlsx` or `.md` format in a single run. ( ***MSSQL*** supported only so far )

You can choose to generate documentations of ***all tables in a certain database or several selected tables*** in the database. And also the format of it. 

## Quick Start

***1. Set up configuration file in***  `dbconfig.toml` ***like this :*** 

```toml
[database]
host = "10.10.10.10"
user = "yourname"
password = "yourpassword"
port = 5050
name = "DB Name"
```



***2. Run command line like this :***

Case when ***all tables*** in a certain database by `.md` format : 

```go
go run main.go -allFlag=true markdown
```

Or by `.xlsx` format : 

```go
go run main.go -allFlag=true excel
```

Case when ***tables whose name provided*** : 

```go
go run main.go -allFlag=false markdown YourTableName1 YourTableName2 YourTableName3 
```

Or by `.xlsx` format : 

```go
go run main.go -allFlag=false excel YourTableName1 YourTableName2 YourTableName3 
```

