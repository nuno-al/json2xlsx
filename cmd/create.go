/*
Copyright © 2023 Nuno Alves <nunodpalves@gmail.com>

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*/
package cmd

import (
	"encoding/json"
	"io/ioutil"
	"os"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

type Book struct {
	FileName   string      `json:"file_name"`
	Worksheets []Worksheet `json:"worksheets"`
}

type Worksheet struct {
	SheetName string `json:"sheet"`
	Cells     []Cell `json:"cells"`
}

type Cell struct {
	CellName string         `json:"cell"`
	Value    string         `json:"value"`
	Style    excelize.Style `json:"style"`
}

type Style struct {
	Borders       []Border   `json:"borders"`
	Fill          Fill       `json:"fill"`
	Font          Font       `json:"font"`
	Alignment     Alignment  `json:"alignment"`
	Protection    Protection `json:"protection"`
	NumFmt        int        `json:"num_fmt"`
	DecimalPlaces int        `json:"decimal_places"`
	CustomNumFmt  string     `json:"custom_num_fmt"`
}

type Border struct {
	Type  string `json:"type"`
	Color string `json:"color"`
	Style int    `json:"style"`
}

type Font struct {
	Bold         bool    `json:"bold"`
	Italic       bool    `json:"italic"`
	Underline    string  `json:"underline"`
	Family       string  `json:"family"`
	Size         float64 `json:"size"`
	Strike       bool    `json:"strike"`
	Color        string  `json:"color"`
	ColorIndexed int     `json:"color_indexed"`
	ColorTheme   int     `json:"color_theme"`
	ColorTint    float64 `json:"color_tint"`
	VertAlign    string  `json:"vert_align"`
}

type Fill struct {
	Type    string   `json:"type"`
	Pattern int      `json:"pattern"`
	Colors  []string `json:"colors"`
	Shading int      `json:"shading"`
}

type Protection struct {
	Hidden bool `json:"hidden"`
	Locked bool `json:"locked"`
}

type Alignment struct {
	Horizontal      string `json:"horizontal"`
	Indent          int    `json:"indent"`
	JustifyLastLine bool   `json:"justify_last_line"`
	ReadingOrder    uint64 `json:"reading_order"`
	RelativeIndent  int    `json:"relative_indent"`
	ShrinkToFit     bool   `json:"shrink_to_fit"`
	TextRotation    int    `json:"text_rotation"`
	Vertical        string `json:"vertical"`
	WrapText        bool   `json:"wrap_text"`
}

var createCmd = &cobra.Command{
	Use:   "create",
	Short: "Creates a xlsx file from json.",
	Long:  `Creates a xlsx file from json.`,
	RunE: func(cmd *cobra.Command, args []string) error {
		j, err := os.Open(args[0])
		if err != nil {
			return err
		}
		defer j.Close()

		b, _ := ioutil.ReadAll(j)

		var book Book

		f := excelize.NewFile()
		defer f.Close()

		json.Unmarshal(b, &book)
		for s := 0; s < len(book.Worksheets); s++ {
			sheet := book.Worksheets[s]

			_, err := f.NewSheet(sheet.SheetName)
			if err != nil {
				return err
			}

			for c := 0; c < len(sheet.Cells); c++ {
				cell := sheet.Cells[c]

				f.SetCellValue(sheet.SheetName, cell.CellName, cell.Value)
			}
		}

		f.DeleteSheet(f.GetSheetName(0)) // Delete default worksheet

		if err := f.SaveAs(args[1]); err != nil {
			return err
		}

		return nil
	},
}

func init() {
	rootCmd.AddCommand(createCmd)
}
