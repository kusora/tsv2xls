package main

import (
	"bufio"
	"flag"
	xlsx "github.com/kusora/xlsx"
	"os"
	"strings"
	"io"
	"fmt"
)

/*
	source 为源文件完整地址, dest为目标文件地址
	需要测试一下go里面对文件读写时候的相对路径和绝对路径的区别
*/
func Convert(source string, dest string) error {
	fmt.Printf("convert %s to %s \n", source, dest)
	fi, err := os.Open(source)
	if err != nil {
		return err
	}
	defer fi.Close()

	var file *xlsx.File
	var sheet *xlsx.Sheet

	file = xlsx.NewFile()
	sheet = file.AddSheet("Sheet1")

	r := bufio.NewReader(fi)
	for {
		line, err := r.ReadString('\n')
		if err != nil {
			if err == io.EOF {
				if strings.Contains(line, "\t") {
					colValues := strings.Split(line, "\t")
					dataRow := sheet.AddRow()
					dataRow.WriteSlice(&colValues, len(colValues))
				}
				file.Save(dest)
				fmt.Printf("successfully finish reading %s \n", source)
				return nil
			}
			return err
		}
		if strings.Contains(line, "\t") {
			value := line[:len(line)-1]
			colValues := strings.Split(value, "\t")
			dataRow := sheet.AddRow()
			dataRow.WriteSlice(&colValues, len(colValues))
		}
	}
	return nil
}

func main() {
	source := flag.String("s", "", "source file path")
	dest := flag.String("d", "", "destination file path")
	flag.Parse()
	if *source != "" && *dest != "" {
		err := Convert(*source, *dest)
		if err != nil {
			fmt.Println(err)
			return
		}
		fmt.Printf("successfully convert file %s to %s \n", *source, *dest)
	}
}
