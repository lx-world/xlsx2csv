// Copyright 2011-2015, The xlsx2csv Authors.
// All rights reserved.
// For details, see the LICENSE file.

package main

import (
	"encoding/csv"
	"fmt"
	"io"

	"github.com/tealeg/xlsx/v3"
)

func generateCSVFromXLSXFile(w io.Writer,sheet *xlsx.Sheet) error {
	cw := csv.NewWriter(w)

	var vals []string
	err := sheet.ForEachRow(func(row *xlsx.Row) error {
		if row != nil {
			vals = vals[:0]
			err := row.ForEachCell(func(cell *xlsx.Cell) error {
				if cell.NumFmt == "mm-dd-yy"{
					cell.NumFmt = "m/d/yyyy"
				}
				str, err := cell.FormattedValue()
				if err != nil {
					fmt.Printf("GO ====> %v\n",err)
					return err
				}
				vals = append(vals, str)
				return nil
			})
			if err != nil {
				fmt.Printf("GO ====> %v\n",err)
				return err
			}
		}
		if err := cw.Write(vals); err != nil{
			fmt.Printf("GO ====> %v\n",err)
			return err
		}
		return nil
	})
	if err != nil {
		fmt.Printf("GO ====> %v\n",err)
		return err
	}
	cw.Flush()
	return cw.Error()
}

func main() {
}
