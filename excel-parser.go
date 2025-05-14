package main

import (
	"fmt"
	"os"
	"os/exec"
	"regexp"
	"runtime"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type envVariables struct {
	path       string
	fileName   string
	sheetNames []string
}

type roomItem struct {
	roomNumber string
	roomContents []chromebookItem
	
}


type chromebookItem struct {
	sn string
	assetTag string
	}

func main() {
	var env envVariables
	envChan := make(chan envVariables)
	go func() {
		envChan <- setupEnv()
	}()
	env = <-envChan


	// create file if it doesn't exist
	var file *excelize.File
	fmt.Println("")
	if _, err := os.Stat(env.path); os.IsNotExist(err) {
		file = excelize.NewFile()

	} else {
		file,err = excelize.OpenFile(env.path)
			if err != nil {
				fmt.Println("Error opening file:", err)
				return
			}
	}
	for _, sheetName := range env.sheetNames {
		message := fmt.Sprintf("Creating sheet: %s", sheetName)
		fmt.Println(message)
		_,err := file.NewSheet(sheetName)
			if err != nil {
				fmt.Println("Error creating sheet:", err)
				return
			}
		}
		_,err := file.NewSheet("breakdown by room")
		if err != nil {
			fmt.Println("Error creating sheet:", err)
		}

	// scan in each chromebook by room
	var newRoom string
	fmt.Println("enter room number")
	fmt.Scanln(&newRoom)
	fmt.Println("entered room number: ", newRoom)
	var loop bool = true
	

	for loop {

	}

}








func setupEnv() envVariables {
	var env envVariables
	var sheetStr string
	fmt.Println("enter file path")
	fmt.Scanln(&env.path)
	re := regexp.MustCompile(`\\`)
	env.path = re.ReplaceAllString(env.path, "/")
	fmt.Println("enter file name")
	fmt.Scanln(&env.fileName)
	fmt.Println("enter sheet names (comma separated)")
	fmt.Scanln(&sheetStr)
	env.sheetNames = strings.Split(sheetStr, ",")
	fmt.Println("env variables set")
	fmt.Println("clear terminal")
	time.Sleep(1 * time.Second)
	clearTerminal()
	return env
}

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
