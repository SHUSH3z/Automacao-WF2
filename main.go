package main

import (
	"fmt"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	dir, err := os.Getwd()
	if err != nil {
		log.Fatalf("Erro ao obter diretório atual: %v", err)
	}

	outputFile := filepath.Join(dir, "resultado.xlsx")

	out := excelize.NewFile()
	sheetName := "Resumo"
	out.NewSheet(sheetName)
	out.DeleteSheet("Sheet1")

	rowOut := 1

	err = filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if info.IsDir() || !strings.HasSuffix(info.Name(), ".xlsx") || info.Name() == "resultado.xlsx" {
			return nil
		}

		fmt.Println("📄 Lendo:", info.Name())

		f, err := excelize.OpenFile(path)
		if err != nil {
			log.Printf("Erro ao abrir %s: %v\n", path, err)
			return nil
		}

		sheet := f.GetSheetName(0)
		if sheet == "" {
			return nil
		}

		for row := 20; row <= 376; row++ {
			cellB, _ := excelize.CoordinatesToCellName(2, row) // Coluna B
			cellF, _ := excelize.CoordinatesToCellName(6, row) // Coluna F
			cellG, _ := excelize.CoordinatesToCellName(7, row) // Coluna G

			valB, _ := f.GetCellValue(sheet, cellB)
			valF, _ := f.GetCellValue(sheet, cellF)
			valGStr, _ := f.GetCellValue(sheet, cellG)
			formulaG, _ := f.GetCellFormula(sheet, cellG)

			// Converte valG pra float
			var valG float64
			fmt.Sscanf(valGStr, "%f", &valG)

			if formulaG != "" && formulaG != valGStr && valG > 0 {
				// Adiciona ao novo Excel
				out.SetCellValue(sheetName, fmt.Sprintf("A%d", rowOut), valB)
				out.SetCellValue(sheetName, fmt.Sprintf("B%d", rowOut), valF)
				out.SetCellValue(sheetName, fmt.Sprintf("C%d", rowOut), valG)
				rowOut++
			}
		}

		return nil
	})

	if err != nil {
		log.Fatal(err)
	}

	if err := out.SaveAs(outputFile); err != nil {
		log.Fatal("Erro ao salvar o novo Excel:", err)
	}

	fmt.Println("✅ Arquivo de saída criado em:", outputFile)

	// Abrir automaticamente o arquivo Excel (Windows)
	if err := exec.Command("cmd", "/C", "start", outputFile).Start(); err != nil {
		log.Printf("⚠️ Erro ao tentar abrir o arquivo Excel: %v", err)
	}

	// Quantidade de linhas inseridas (menos o cabeçalho, se tiver)
	totalLinhas := rowOut - 1 // Ajuste se você tiver cabeçalho fixo
	fmt.Printf("📊 Total de linhas adicionadas: %d\n", totalLinhas)

	cmd := exec.Command("python", "robozin.py", fmt.Sprintf("%d", totalLinhas))
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr

	fmt.Println("▶️ Executando script Python...")

	if err := cmd.Run(); err != nil {
		log.Fatalf("Erro ao executar script Python: %v", err)
	}

	fmt.Println("✅ Script Python executado com sucesso!")
}
