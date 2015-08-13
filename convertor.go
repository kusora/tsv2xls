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
func ConvertToXls(source string, dest string) error {
	fmt.Printf("convert %s to %s \n", source, dest)
	fi, err := os.Open(source)
	if err != nil {
		return err
	}
	defer func() {
		if err := fi.Close(); err != nil {
			panic(err)
		}
	}()

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

func ConvertToHtml(source string, dest string) error {
	fmt.Printf("convert %s to %s \n", source, dest)
	fi, err := os.Open(source)
	if err != nil {
		return err
	}
	defer func() {
		if err := fi.Close(); err != nil {
			panic(err)
		}
	}()

	fo, err := os.Create(dest)
	if err != nil {
		return err
	}

	defer func() {
		if err := fo.Close(); err != nil {
			panic(err)
		}
	}()

	r := bufio.NewReader(fi)
	w := bufio.NewWriter(fo)
	w.WriteString("<html><head><meta charset=\"utf-8\"/></head><table>")
	w.Flush()
	for {
		line, err := r.ReadString('\n')
		if err != nil {
			if err == io.EOF {
				if strings.Contains(line, "\t") {
					value := line[:len(line)-1]
					colValues := strings.Split(value, "\t")
					w.WriteString("<tr>")
					for _, colValue := range colValues {
						w.WriteString("<td>" + colValue + "</td>")
					}
					w.WriteString("</tr>")
				}
				w.WriteString("</table></html>")
				fmt.Printf("successfully finish reading %s \n", source)
				w.Flush()
				return nil
			}
			return err
		}
		if strings.Contains(line, "\t") {
			value := line[:len(line)-1]
			colValues := strings.Split(value, "\t")
			w.WriteString("<tr>")
			for _, colValue := range colValues {
				w.WriteString("<td>" + colValue + "</td>")
			}
			w.WriteString("</tr>")
			w.Flush()
		}
	}
	return nil
}

func main() {
	source := flag.String("s", "", "source file path")
	dest := flag.String("d", "", "destination file path")
	target := flag.String("t", "xls", "destination file type") //xls, html

	flag.Parse()
	if *source != "" && *dest != "" {
		if *target == "xls" {
			err := ConvertToXls(*source, *dest)
			if err != nil {
				fmt.Println(err)
				return
			}
		} else  if *target == "html" {
			err := ConvertToHtml(*source, *dest)
			if err != nil {
				fmt.Println(err)
				return
			}
		}
		fmt.Printf("successfully convert file %s to %s \n", *source, *dest)
	}
}
