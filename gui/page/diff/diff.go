package diff

import (
	"fmt"
	"strings"
	"sync/atomic"

	"gioui.org/layout"
	"gioui.org/widget"
	"gioui.org/widget/material"
	"gioui.org/x/component"
	"gioui.org/x/explorer"

	spsicon "github.com/SPSZerone/sps-go-zerone/graphics/gio/icon"
	spswin "github.com/SPSZerone/sps-go-zerone/graphics/gio/window"

	spsexcel "github.com/SPSZerone/sps-go-excel/excel"
)

type Diff struct {
	widget.List

	fileInput    component.TextField
	fileButton   widget.Clickable
	reloadButton widget.Clickable
	explorer     *explorer.Explorer

	excel        spsexcel.Excel
	excelRefresh atomic.Bool
}

type Excel struct {
	Diff  *Diff
	Excel spsexcel.Excel
}

func (d *Diff) Layout(win *spswin.Window, gtx layout.Context, param any) layout.Dimensions {
	d.List.Axis = layout.Vertical
	return material.List(win.Theme, &d.List).Layout(gtx, 1, func(gtx layout.Context, _ int) layout.Dimensions {
		return layout.Flex{Axis: layout.Vertical}.Layout(
			gtx,
			layout.Rigid(func(gtx layout.Context) layout.Dimensions {
				return layout.Flex{
					Axis:      layout.Horizontal,
					Alignment: layout.Middle,
					Spacing:   layout.SpaceEvenly,
					WeightSum: 1,
				}.Layout(
					gtx,
					layout.Flexed(0.95, func(gtx layout.Context) layout.Dimensions {
						return d.fileInput.Layout(gtx, win.Theme, "File(xlsx,csv...)")
					}),
					layout.Rigid(func(gtx layout.Context) layout.Dimensions {
						if d.fileButton.Clicked(gtx) {
							go d.OpenFile(win)
						}
						return material.IconButton(win.Theme, &d.fileButton, spsicon.ActionOpenInNew, "Open File").Layout(gtx)
					}),
					layout.Rigid(func(gtx layout.Context) layout.Dimensions {
						if d.reloadButton.Clicked(gtx) {
							win.Logger.Info().Msg("Reload")
						}
						return material.IconButton(win.Theme, &d.reloadButton, spsicon.ActionUpdate, "Reload").Layout(gtx)
					}),
				)
			}),
			layout.Rigid(func(gtx layout.Context) layout.Dimensions {
				return d.LayoutExcel(win, gtx, param)
			}),
		)
	})
}

func (d *Diff) OpenFile(win *spswin.Window) {
	file, err := d.explorer.ChooseFile("xlsx", "csv")
	if err != nil {
		win.Logger.Error().Msgf("explorer.ChooseFile err:%v", err)
		return
	}
	defer file.Close()

	win.Logger.Info().Msg("Open File")

	ex, err := spsexcel.OpenReader(file, spsexcel.OptFlag(spsexcel.OReadWrite))
	if err != nil {
		win.Logger.Error().Msgf("excel.OpenReader err:%v", err)
		return
	}

	d.excel = ex
	d.excelRefresh.Store(true)
}

func (d *Diff) LayoutExcel(win *spswin.Window, gtx layout.Context, param any) layout.Dimensions {
	if d.excelRefresh.Load() {
		d.excelRefresh.Store(false)

		onGetCells := func(cells []spsexcel.Cell, builder *strings.Builder) {
			for i, cell := range cells {
				if i == 0 {
					builder.WriteString(fmt.Sprintf(" | %s", cell))
				} else {
					builder.WriteString(fmt.Sprintf("\t%s", cell))
				}
			}
		}

		sheet := d.excel.GetActiveSheet()
		rows, _ := sheet.GetRows()
		var builder strings.Builder
		for _, row := range rows {
			builder.Reset()
			builder.WriteString(fmt.Sprintf("Sheet:%v", sheet.Name()))
			builder.WriteString(fmt.Sprintf(" RowId:%v", row.Id()))

			cells, _ := row.Cells()
			onGetCells(cells, &builder)

			win.Logger.Info().Msg(builder.String())
		}
	}
	return material.H6(win.Theme, "TODO...").Layout(gtx)
}
