package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"sort"
	"strconv"
	"sync"

	"github.com/xuri/excelize/v2"
)

type Student struct {
	EmpID  string
	Branch string
	Marks  map[string]float64
	Total  float64
}

var (
	components  = []string{"Quiz", "Mid-Sem", "Lab Test", "Weekly Labs", "Pre-Compre", "Compre"}
	exportJSON  bool
	classFilter string
)

func init() {
	flag.BoolVar(&exportJSON, "export", false, "Export report as JSON")
	flag.StringVar(&classFilter, "class", "", "Filter by Class ID")
	flag.Parse()
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run main.go <path-to-excel-file>")
		return
	}

	filePath := os.Args[1]
	students, err := parseExcel(filePath)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	var wg sync.WaitGroup
	mismatchCh := make(chan string, len(students))

	wg.Add(1)
	go func() {
		defer wg.Done()
		validateData(students, mismatchCh)
	}()

	wg.Wait()
	close(mismatchCh)

	var mismatches []string
	for msg := range mismatchCh {
		mismatches = append(mismatches, msg)
	}

	fmt.Println("\nValidation Errors:")
	if len(mismatches) > 0 {
		for _, msg := range mismatches {
			fmt.Println(msg)
		}
	} else {
		fmt.Println("No validation errors found.")
	}

	calculateAverages(students)
	calculateBranchAverages(students)
	rankStudents(students)

	if exportJSON {
		exportToJSON(students, mismatches)
	}
}

func parseExcel(filePath string) ([]Student, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println("Error opening the file:", err)
		return nil, err
	}
	defer f.Close()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil {
		return nil, err
	}

	var students []Student

	for i, row := range rows {
		if i == 0 || len(row) < 11 {
			continue
		}

		empID := row[2]
		campusID := row[3]

		if len(campusID) < 6 {
			fmt.Printf("Warning: Skipping row %d due to invalid CampusID format (%s)\n", i+1, campusID)
			continue
		}

		branch := campusID[4:6]

		student := Student{
			EmpID:  empID,
			Branch: branch,
			Marks:  make(map[string]float64),
		}

		for j, comp := range components {
			mark, _ := strconv.ParseFloat(row[j+4], 64)
			student.Marks[comp] = mark
		}

		finalTotal, _ := strconv.ParseFloat(row[10], 64)
		student.Marks["Final Total"] = finalTotal

		students = append(students, student)
	}

	return students, nil
}

func validateData(students []Student, mismatchCh chan<- string) {
	for _, student := range students {
		expectedI := student.Marks["Quiz"] + student.Marks["Mid-Sem"] + student.Marks["Lab Test"] + student.Marks["Weekly Labs"]
		if expectedI != student.Marks["Pre-Compre"] {
			mismatchCh <- fmt.Sprintf("Mismatch in E+F+G+H != I for EmpID %s", student.EmpID)
		}

		expectedTotal := student.Marks["Pre-Compre"] + student.Marks["Compre"]
		actualTotal, exists := student.Marks["Final Total"]

		if exists && expectedTotal != actualTotal {
			mismatchCh <- fmt.Sprintf("Mismatch in I+J != K for EmpID %s (Expected: %.2f, Found: %.2f)", student.EmpID, expectedTotal, actualTotal)
		}
	}
}

func calculateAverages(students []Student) {
	avg := make(map[string]float64)
	count := float64(len(students))

	for _, student := range students {
		for comp, mark := range student.Marks {
			avg[comp] += mark
		}
	}

	fmt.Println("\nAverage Marks per Component:")
	for comp, total := range avg {
		fmt.Printf("%s: %.2f\n", comp, total/count)
	}
}

func calculateBranchAverages(students []Student) {
	branchTotals := make(map[string]float64)
	branchCounts := make(map[string]int)

	for i := range students {
		students[i].Total = students[i].Marks["Quiz"] + students[i].Marks["Mid-Sem"] +
			students[i].Marks["Lab Test"] + students[i].Marks["Weekly Labs"] + students[i].Marks["Compre"]
	}

	for _, student := range students {
		branch := student.Branch
		branchTotals[branch] += student.Total
		branchCounts[branch]++
	}

	fmt.Println("\nBranch-wise Averages:")
	for branch, total := range branchTotals {
		avg := total / float64(branchCounts[branch])
		fmt.Printf("Branch %s: %.2f\n", branch, avg)
	}
}

func rankStudents(students []Student) {
	fmt.Println("\nOverall Top 3 Students:")

	for i := range students {
		students[i].Total = students[i].Marks["Quiz"] + students[i].Marks["Mid-Sem"] +
			students[i].Marks["Lab Test"] + students[i].Marks["Weekly Labs"] + students[i].Marks["Compre"]
	}

	sort.Slice(students, func(i, j int) bool {
		return students[i].Total > students[j].Total
	})

	for i := 0; i < 3 && i < len(students); i++ {
		fmt.Printf("%d. EmpID: %s | Computed Total: %.2f\n", i+1, students[i].EmpID, students[i].Total)
	}

	branchStudents := make(map[string][]Student)
	for _, student := range students {
		branchStudents[student.Branch] = append(branchStudents[student.Branch], student)
	}

	fmt.Println("\nTop 3 Students per Branch:")
	for branch, studentsInBranch := range branchStudents {
		sort.Slice(studentsInBranch, func(i, j int) bool {
			return studentsInBranch[i].Total > studentsInBranch[j].Total
		})

		fmt.Printf("\nBranch %s:\n", branch)
		for i := 0; i < 3 && i < len(studentsInBranch); i++ {
			fmt.Printf("%d. EmpID: %s | Computed Total: %.2f\n", i+1, studentsInBranch[i].EmpID, studentsInBranch[i].Total)
		}
	}
}

func exportToJSON(students []Student, mismatches []string) {
	data := map[string]interface{}{
		"students":   students,
		"mismatches": mismatches,
	}

	file, err := os.Create("output.json")
	if err != nil {
		fmt.Println("Error creating JSON file:", err)
		return
	}
	defer file.Close()

	encoder := json.NewEncoder(file)
	encoder.SetIndent("", "  ")
	err = encoder.Encode(data)
	if err != nil {
		fmt.Println("Error writing JSON data:", err)
	}

	fmt.Println("Data exported to output.json")
}
