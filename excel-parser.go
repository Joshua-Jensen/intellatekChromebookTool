package main

import (
	"fmt"
	"regexp"
)

type envVariables struct {
	path     string
	fileName string
}

func main() {
	var env envVariables
	envChan := make(chan envVariables)
	go func() {
		envChan <- setupEnv()
	}()
	env = <-envChan

	fmt.Println("")

}

func setupEnv() envVariables {
	var env envVariables
	fmt.Println("enter file path")
	fmt.Scanln(&env.path)
	re := regexp.MustCompile(`\\`)
	env.path = re.ReplaceAllString(env.path, "/")
	fmt.Println("enter file name")
	fmt.Scanln(&env.fileName)
	return env
}
