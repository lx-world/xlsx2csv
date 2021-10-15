win-build:
	@CGO_ENABLED=1 GOOS=windows GOARCH=amd64 go build --buildmode=c-shared -o output/xlsx2csv.dll main.go export.go