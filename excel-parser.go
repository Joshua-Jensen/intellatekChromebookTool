package main

import (
	"bufio"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"regexp"
	"runtime"
	"time"

	"github.com/xuri/excelize/v2"
)

type envVariables struct {
	path     string
	fileName string
	// sheetNames []string
}

type roomItem struct {
	roomNumber   string
	roomContents []chromebookItem
}

type chromebookItem struct {
	// sn       string
	// sn       string
	assetTag string
	comments string
	slotNum  string
}

func main() {

	//wait for the user to enter the env variables
	var env envVariables
	envChan := make(chan envVariables)
	go func() {
		envChan <- setupEnv()
	}()
	env = <-envChan

	//create a new file
	var path string = env.path + "/" + env.fileName + ".xlsx"
	var file *excelize.File
	_, err := os.Stat(path)
	if err == nil {
		file, err = excelize.OpenFile(path)
		if err != nil {
			fmt.Println(err)
		}
	} else if errors.Is(err, os.ErrNotExist) {
		fmt.Println("creating new file")
		file = excelize.NewFile()
	} else {
		fmt.Println("err ocurred", err)
	}

	var path string = env.path + "/" + env.fileName + ".xlsx"
	var file *excelize.File
	_, err := os.Stat(path)
	if err == nil {
		file, err = excelize.OpenFile(path)
		if err != nil {
			fmt.Println(err)
		}
	} else if errors.Is(err, os.ErrNotExist) {
		fmt.Println("creating new file")
		file = excelize.NewFile()
	} else {
		fmt.Println("err ocurred", err)
	}

	createNewRoomSheet(file, path)
	println("all done")
	// forcing git update
}

// this function sets up the env variables
func setupEnv() envVariables {
	scanner := bufio.NewScanner(os.Stdin)
	var env envVariables
	// var sheetStr string
	fmt.Println("enter file path")
	// fmt.Scanln(&env.path)
	if scanner.Scan() {
		env.path = scanner.Text()
	}
	re := regexp.MustCompile(`\\`)
	env.path = re.ReplaceAllString(env.path, "/")
	fmt.Println("enter file name")
	// fmt.Scanln(&env.fileName)
	if scanner.Scan() {
		env.fileName = scanner.Text()
	}
	// fmt.Println("enter sheet names (comma separated)")
	// fmt.Scanln(&sheetStr)
	// env.sheetNames = strings.Split(sheetStr, ",")
	fmt.Println("env variables set")
	fmt.Println("clear terminal")

	time.Sleep(1 * time.Second)
	// clearTerminal()
	return env
}

// this function clears the terminal
// unused while testing
func clearTerminal() {
	var cmd *exec.Cmd
	if runtime.GOOS == "windows" {
		cmd = exec.Command("cmd", "/c", "cls")
		cmd.Stdout = os.Stdout
	} else {
		fmt.Print("\033[H\033[2J")
	}
	_ = cmd.Run()
}

func createNewRoomSheet(file *excelize.File, path string) {
	var loop bool = true
	var exitString string = "done"
	var newRoom string
	for loop {
		fmt.Println("enter room number")
		fmt.Scanln(&newRoom)
		if newRoom != exitString {
			fmt.Println("entered room number: ", newRoom)
			//create new sheet for this room
			_, err := file.NewSheet(newRoom)
			if err != nil {
				fmt.Println("Error creating sheet:", err)
			}
			createRoomContents(file, newRoom, path)
		} else {
		} else {
			loop = false
		}
	}
}

// this function takes int he file and the env,path and creates a new sheet for each room with its contents
func createRoomContents(file *excelize.File, newRoom string, path string) {
	// scan in each chromebook by room

	//collect the data on the chromebooks
	scanner := bufio.NewScanner(os.Stdin)

	var loop bool = true
	var roomList roomItem
	roomList.roomNumber = newRoom
	for loop {
		var escapeString string = "exit"
		// var newRoomString string = "done"
		var newChromebook chromebookItem
		// 	fmt.Println("enter chromebook serial number")
		// 	fmt.Scanln(&newChromebook.sn)
		// 	if newChromebook.sn != escapeString {
		// 		fmt.Println("entered chromebook serial number: ", newChromebook.sn)
		// 		fmt.Println("enter chromebook asset tag")
		// 		fmt.Scanln(&newChromebook.assetTag)
		// 		fmt.Println("entered chromebook asset tag: ", newChromebook.assetTag)
		// 		roomList.roomContents = append(roomList.roomContents, newChromebook)
		// 	} else {
		// 		loop = false
		// 	}
		// }

		fmt.Println("enter chromebook asset tag")
		// fmt.Scanln(&newChromebook.assetTag)
		if scanner.Scan() {
			newChromebook.assetTag = scanner.Text()
		}

		if newChromebook.assetTag != escapeString {

			fmt.Println("entered chromebook asset tag: ", newChromebook.assetTag)
			fmt.Println("enter comments")
			// fmt.Scanln(&newChromebook.comments)
			if scanner.Scan() {
				newChromebook.comments = scanner.Text()
			}
			fmt.Println("enter in the slot number")
			// fmt.Scanln(&newChromebook.slotNum)
			if scanner.Scan() {
				newChromebook.slotNum = scanner.Text()
			}
			roomList.roomContents = append(roomList.roomContents, newChromebook)
		} else {
			loop = false
		}
	}
	// Write room data to the sheet
	rowIndex := 2 // Start writing from the second row
	rowIndex := 2 // Start writing from the second row
	for _, chromebook := range roomList.roomContents {
		rowData := []interface{}{"", chromebook.assetTag, chromebook.comments, chromebook.slotNum, "", "", "", "", roomList.roomNumber, "", "", "", "", ""}
		err := file.SetSheetRow(roomList.roomNumber, fmt.Sprintf("A%d", rowIndex), &rowData)
		if err != nil {
			fmt.Println("Error writing to sheet:", err)
			return
		}
		rowIndex++
	}

	// Save the file
	saveFile(file, path)

}

func saveFile(file *excelize.File, path string) {
	// Save the file to the specified path
	var loop = true
	for loop {
		err := file.SaveAs(path)
		if err != nil {
			fmt.Println("Error saving file:", err)
			time.Sleep(5 * time.Second)
		} else if err == nil {

			loop = false
		}
	}
	fmt.Println("File saved successfully at", path)
}
