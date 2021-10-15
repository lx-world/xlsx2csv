package main

/*
#include<stdlib.h>

*/
import "C"

import (
	"errors"
	"fmt"
	"os"
	"strings"
	"unsafe"

	"github.com/tealeg/xlsx/v3"
)

//export OpenExcelFile
func OpenExcelFile(fileName *C.char) (fName *C.char, sheetList *C.char, errRes *C.char) {
	fStr := C.GoString(fileName)
	xlFile, err := xlsx.OpenFile(fStr)
	if err != nil {
		return C.CString(""), C.CString(""), C.CString(err.Error())
	}
	sheetLen := len(xlFile.Sheets)
	if sheetLen == 0 {
		return C.CString(""), C.CString(""), C.CString("no sheet")
	}

	var sheetSc []string
	for _, v := range xlFile.Sheets {
		sheetSc = append(sheetSc, v.Name)
	}
	return C.CString(fStr), C.CString(strings.Join(sheetSc, ";")), C.CString("")
}

//export ExportCsv
func ExportCsv(file *C.char, sheetList *C.char, csvPath *C.char) (bool, *C.char) {
	fileStr := C.GoString(file)
	sheetStr := C.GoString(sheetList)
	csvPathStr := C.GoString(csvPath)
	if fileStr == "" || sheetStr == "" || csvPathStr == "" {
		fmt.Printf("GO ====> %s\n", fmt.Sprintf("某些参数为空, xlsx:%v, sheet:%v, csv:%v", fileStr, sheetStr, csvPathStr))
		return false, C.CString(fmt.Sprintf("某些参数为空, xlsx:%v, sheet:%v, csv:%v", fileStr, sheetStr, csvPathStr))
	}

	sheetSc := strings.Split(sheetStr, ";")
	csvPathSc := strings.Split(csvPathStr, ";")

	if len(sheetSc) != len(csvPathSc) {
		fmt.Printf("GO ====> %s\n", "sheet and csv length not same")
		return false, C.CString("sheet and csv length not same")
	}

	xlFile, err := xlsx.OpenFile(fileStr)
	if err != nil {
		return false, C.CString(err.Error())
	}
	for idx := range csvPathSc {
		if err = genCsv(xlFile, sheetSc[idx], csvPathSc[idx]); err != nil {
			fmt.Printf("GO ====> %v\n", err)
			return false, C.CString(err.Error())
		}
	}

	return true, C.CString("")
}

func genCsv(xlFile *xlsx.File, sheetName string, csvPath string) error {
	var sheet *xlsx.Sheet
	for _, v := range xlFile.Sheets {
		if v.Name == sheetName {
			sheet = v
			break
		}
	}
	if sheet == nil {
		fmt.Printf("GO ====> %s\n", "sheet name error")
		return errors.New("sheet name error")
	}
	out, err := os.Create(csvPath)
	if err != nil {
		fmt.Printf("GO ====> %v\n", err)
		return err
	}
	defer out.Close()
	if err = generateCSVFromXLSXFile(out, sheet); err != nil {
		fmt.Printf("GO ====> %v\n", err)
		return err
	}
	return nil
}

//export Free
func Free(p unsafe.Pointer) {
	C.free(p)
}

// EchoString test
//export EchoString
func EchoString(str *C.char) *C.char {
	res := C.GoString(str)
	return C.CString(res)
}

// EchoMulti test
//export EchoMulti
func EchoMulti() (int, int) {
	return 1, 2
}
