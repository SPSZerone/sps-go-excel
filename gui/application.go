package gui

import (
	spsgio "github.com/SPSZerone/sps-go-zerone/graphics/gio"
	spspref "github.com/SPSZerone/sps-go-zerone/graphics/gio/page/pref"

	"github.com/SPSZerone/sps-go-excel/excel"
	"github.com/SPSZerone/sps-go-excel/excel/excelize"
	"github.com/SPSZerone/sps-go-excel/gui/page/about"
	"github.com/SPSZerone/sps-go-excel/gui/page/diff"
)

func init() {
	excel.RegisterDefaultNewerExcel(excelize.DefaultNewer)
}

func Run() {
	spsgio.Run(
		spsgio.OptTitle("SPS Excel Tools"),
		spsgio.OptLoopMode(spsgio.LoopModeSimple),
		spsgio.OptOnInit(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools Init")
			app.Register(0, about.New(app))
			app.Register(1, spspref.New(app))
			app.Register(2, diff.New(app))
			app.SwitchTo(2)
		}),
		spsgio.OptOnStart(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools Start")
		}),
		spsgio.OptOnStop(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools Stop")
		}),
	)
}

type Application struct {
}
