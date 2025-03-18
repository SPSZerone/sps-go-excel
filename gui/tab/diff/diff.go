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

	spsexcel "github.com/SPSZerone/sps-go-excel/excel"
	spsgio "github.com/SPSZerone/sps-go-zerone/graphics/gio"
	spsicon "github.com/SPSZerone/sps-go-zerone/graphics/gio/icon"
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

func (d *Diff) Layout(app *spsgio.Application, gtx layout.Context, param any) layout.Dimensions {
	d.List.Axis = layout.Vertical
	return material.List(app.Theme, &d.List).Layout(gtx, 1, func(gtx layout.Context, _ int) layout.Dimensions {
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
						return d.fileInput.Layout(gtx, app.Theme, "File(xlsx,csv...)")
					}),
					layout.Rigid(func(gtx layout.Context) layout.Dimensions {
						if d.fileButton.Clicked(gtx) {
							go d.OpenFile(app)
						}
						return material.IconButton(app.Theme, &d.fileButton, spsicon.ActionOpenInNew, "Open File").Layout(gtx)
					}),
					layout.Rigid(func(gtx layout.Context) layout.Dimensions {
						if d.reloadButton.Clicked(gtx) {
							app.Logger.Info().Msg("Reload")
						}
						return material.IconButton(app.Theme, &d.reloadButton, spsicon.ActionUpdate, "Reload").Layout(gtx)
					}),
				)
			}),
			layout.Rigid(func(gtx layout.Context) layout.Dimensions {
				return d.LayoutExcel(app, gtx, param)
			}),
		)
	})
}

func (d *Diff) OpenFile(app *spsgio.Application) {
	file, err := d.explorer.ChooseFile("xlsx", "csv")
	if err != nil {
		app.Logger.Error().Msgf("explorer.ChooseFile err:%v", err)
		return
	}
	defer file.Close()

	app.Logger.Info().Msg("Open File")

	ex, err := spsexcel.OpenReader(file, spsexcel.OptFlag(spsexcel.OReadWrite|spsexcel.OCreate))
	if err != nil {
		app.Logger.Error().Msgf("excel.OpenReader err:%v", err)
		return
	}

	bytes, err := ex.Read()
	if err != nil {
		app.Logger.Error().Msgf("ex.Read err:%v", err)
		return
	}
	app.Logger.Info().Msgf("bytes:%d", bytes)

	d.excel = ex
	d.excelRefresh.Store(true)
}

func (d *Diff) LayoutExcel(app *spsgio.Application, gtx layout.Context, param any) layout.Dimensions {
	if d.excelRefresh.Load() {
		d.excelRefresh.Store(false)

		// TODO test
		onGetCells := func(cells []spsexcel.Cell, builder *strings.Builder) {
			for i, cell := range cells {
				if i == 0 {
					builder.WriteString(fmt.Sprintf(" | %s", cell))
				} else {
					builder.WriteString(fmt.Sprintf("\t%s", cell))
				}
			}
		}

		onGetSheet := func(sheet spsexcel.Sheet) {
			app.Logger.Info().Msgf("Sheet:%+v ==================================================", sheet.Name())
			cell, err := sheet.GetCellCR("A", 1)
			app.Logger.Info().Msgf("%s", cell)
			cell, err = sheet.GetCellCR("B", 2)
			app.Logger.Info().Msgf("%s", cell)
			cell, err = sheet.GetCellCR("C", 3)
			app.Logger.Info().Msgf("%s", cell)
			_ = err
		}
		sheet := d.excel.GetActiveSheet()
		onGetSheet(sheet)

		rows, _ := sheet.GetRows()
		var builder strings.Builder
		for _, row := range rows {
			builder.Reset()
			builder.WriteString(fmt.Sprintf("Sheet:%v", sheet.Name()))
			builder.WriteString(fmt.Sprintf(" RowId:%v", row.Id()))

			cells, _ := row.Cells()
			onGetCells(cells, &builder)

			app.Logger.Info().Msg(builder.String())
		}
	}
	return material.H6(app.Theme, "aaaaaaaaa").Layout(gtx)
}
