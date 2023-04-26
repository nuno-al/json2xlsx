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
	Worksheets []Worksheet `json:"worksheets"`
}

type Worksheet struct {
	SheetName string   `json:"sheet"`
	Cells     []Cell   `json:"cells"`
	Columns   []Column `json:"columns"`
	Rows      []Row    `json:"rows"`
}

type Column struct {
	Name  string   `json:"column"`
	Width *float64 `json:"width,omitempty"`
}

type Row struct {
	Index  int      `json:"row"`
	Height *float64 `json:"height,omitempty"`
}

type Cell struct {
	CellName string      `json:"cell"`
	Value    interface{} `json:"value"`
	Style    *Style      `json:"style,omitempty"`
	Merge    *string     `json:"merge,omitempty"`
}

type Style struct {
	Borders       []Border    `json:"borders,omitempty"`
	Fill          *Fill       `json:"fill,omitempty"`
	Font          *Font       `json:"font,omitempty"`
	Alignment     *Alignment  `json:"alignment,omitempty"`
	Protection    *Protection `json:"protection,omitempty"`
	NumFmt        *int        `json:"num_fmt,omitempty"`
	DecimalPlaces *int        `json:"decimal_places,omitempty"`
	CustomNumFmt  *string     `json:"custom_num_fmt,omitempty"`
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
	ColorTheme   *int    `json:"color_theme"`
	ColorTint    float64 `json:"color_tint"`
	VertAlign    string  `json:"vert_align"`
}

type Fill struct {
	Type    string   `json:"type"`
	Pattern int      `json:"pattern"`
	Color   []string `json:"color"`
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
		var err error
		var j *os.File

		j, err = os.Open(args[0])
		if err != nil {
			return err
		}
		defer j.Close()

		b, _ := ioutil.ReadAll(j)

		var book Book

		f := excelize.NewFile()
		defer f.Close()

		json.Unmarshal(b, &book)

		for _, sheet := range book.Worksheets {

			_, err = f.NewSheet(sheet.SheetName)
			if err != nil {
				return err
			}

			for _, column := range sheet.Columns {
				err := f.SetColWidth(sheet.SheetName, column.Name, column.Name, *column.Width)
				if err != nil {
					return err
				}
			}

			for _, row := range sheet.Rows {
				err := f.SetRowHeight(sheet.SheetName, row.Index, *row.Height)
				if err != nil {
					return err
				}
			}

			for _, cell := range sheet.Cells {

				err = f.SetCellValue(sheet.SheetName, cell.CellName, cell.Value)
				if err != nil {
					return err
				}

				if cell.Merge != nil {
					f.MergeCell(sheet.SheetName, cell.CellName, *cell.Merge)
				}

				if cell.Style != nil {

					style := excelize.Style{}

					for _, val := range cell.Style.Borders {
						border := excelize.Border(val)
						style.Border = append(style.Border, border)
					}

					if cell.Style.Fill != nil {
						style.Fill = excelize.Fill(*cell.Style.Fill)
					}

					if cell.Style.Font != nil {
						font := excelize.Font(*cell.Style.Font)
						style.Font = &font
					}

					if cell.Style.Protection != nil {
						protection := excelize.Protection(*cell.Style.Protection)
						style.Protection = &protection
					}

					if cell.Style.Alignment != nil {
						alignment := excelize.Alignment(*cell.Style.Alignment)
						style.Alignment = &alignment
					}

					if cell.Style.NumFmt != nil {
						style.NumFmt = *cell.Style.NumFmt
					}

					if cell.Style.DecimalPlaces != nil {
						style.DecimalPlaces = *cell.Style.DecimalPlaces
					}

					if cell.Style.CustomNumFmt != nil {
						style.CustomNumFmt = cell.Style.CustomNumFmt
					}

					s, err := f.NewStyle(&style)
					if err != nil {
						return err
					}
					f.SetCellStyle(sheet.SheetName, cell.CellName, cell.CellName, s)
				}
			}
		}

		f.DeleteSheet(f.GetSheetName(0)) // Delete default worksheet

		err = f.SaveAs(args[1])
		if err != nil {
			return err
		}

		return nil
	},
}

func init() {
	rootCmd.AddCommand(createCmd)
}
