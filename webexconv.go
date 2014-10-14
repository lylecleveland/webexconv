package main

/*
webexconv is a program used to convert .txt webex backups into a csv capable of being imported by hyperoffice
Created by:	Lyle Cleveland
Date:		14Oct2014
*/

import (
	"bytes"
	"encoding/csv"
	"fmt"
	"io/ioutil"
	"os"
	"regexp"
	"strings"
	"time"
)

var filename string //Holds filename from command line argument

func main() {
	args := os.Args[1:]
	if len(args) == 0 {
		fmt.Println("Enter the filename to convert or drop the file onto webexconv icon. Ex. webconv test.txt")
		return
	}
	filename = args[0]
	filename = filename[:len(filename)-4]
	start := time.Now()
	parseWebex()
	elapsed := time.Since(start)
	fmt.Print("Execution Time: ")
	fmt.Println(elapsed)
}

//parseWebex takes a webex txt backup and returns an array containing the headers and a seperate array containg the fields
func parseWebex() {
	r := regexp.MustCompile("^\\W$") //Regexp for stripping special characters out of the header
	var headers []string
	var fields []string
	scan, error := ioutil.ReadFile(filename + ".txt")
	if error != nil {
		panic(error)
	}
	var buffer bytes.Buffer
	headerEnd := 0

	//Loop for parsing header fields
	for i, b := range scan {
		//Only parse first line
		if b == '\n' {
			headers = strings.Split(buffer.String(), "\t")
			headerEnd = i
			buffer.Reset()
			break
		}
		//Remove spaces and special characters
		if b == ' ' || r.MatchString(string(b)) && b != '\t' {
			continue
		} else {
			buffer.WriteByte(b)
		}
	}

	inString := false //Variable used to determine whether within a string so as to not parse tabs or new lines

	//Main parsing loop
	for i := headerEnd + 1; i < len(scan); i++ {
		if scan[i] == '"' && scan[i-1] == '\t' && inString == false {
			//Found the beginning of a string field
			buffer.WriteByte(scan[i])
			inString = true
		} else if scan[i] == '"' && scan[i+1] == '\t' && inString == true {
			//Found the end of the string field
			buffer.WriteByte(scan[i])
			fields = append(fields, buffer.String())
			buffer.Reset()
			inString = false
			i += 1
		} else if inString {
			//We don't want to parse tabs or new lines if we are in a string
			buffer.WriteByte(scan[i])
		} else if scan[i] == '\n' && inString == false {
			//End of a line and not in a string so we are at the end of a field
			fields = append(fields, buffer.String())
			buffer.Reset()
		} else if scan[i] != '\t' {
			buffer.WriteByte(scan[i])
		} else {
			//Hit a tab so we found the end of a field
			fields = append(fields, buffer.String())
			buffer.Reset()
		}
	}
	writeCSV(headers, fields)
	return
}

//writeCSV is used to generate the parsed csv file for importing
func writeCSV(headers []string, fields []string) {
	file, error := os.Create(filename + ".csv")
	if error != nil {
		panic(error)
	}
	defer file.Close()
	w := csv.NewWriter(file)
	w.UseCRLF = true
	w.Write(headers)
	w.Flush()
	count := 0 //Used to know when we should make a new line of records after we have wrote fields equal to the amount of header fields
	var f []string
	for i := 0; i < len(fields); i++ {
		f = append(f, fields[i])
		count++
		if count%len(headers) == 0 {
			w.Write(f)
			f = nil
			count = 0
		}
	}
	w.Flush()

	//Reparses file to remove quotes from header and blank quoted fields as hyperoffice won't import them
	scan, error := ioutil.ReadFile(filename + ".csv")
	if error != nil {
		panic(error)
	}
	var temp []byte
	inHeader := true
	for i := 0; i < len(scan); i++ {
		if scan[i] == '"' && inHeader == true {
			//Remove quotes from header
			continue
		}
		if scan[i] == '"' && scan[i+1] == '"' && scan[i+2] == ',' {
			//Remove blank quoted fields
			temp = append(temp, ',')
			i += 2
		} else {
			temp = append(temp, scan[i])
		}
		if scan[i] == '\n' && inHeader == true {
			inHeader = false
		}
	}
	ioutil.WriteFile(filename+".csv", temp, 0644)
}
