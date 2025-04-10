package gui

import (
	"gioui.org/io/system"

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
		spsgio.OptStartAction(system.ActionMaximize),
		spsgio.OptLoopMode(spsgio.LoopModeSimple),
		spsgio.OptOnInitPre(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools InitPre")
			app.Pref.Settings.NonModalDrawer = true
		}),
		spsgio.OptOnInitPost(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools InitPost")

			pageTag := 0
			app.PageRegister(pageTag, about.New(app))

			pageTag++
			pref := spspref.New(app)
			pref.Tabs.SetSelected(spspref.TabIdxSettings)
			app.PageRegister(pageTag, pref)

			pageTag++
			app.PageRegister(pageTag, diff.New(app))
		}),
		spsgio.OptOnStart(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools Start")
			app.PageStart(2)
		}),
		spsgio.OptOnStop(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools Stop")
		}),
	)
}

type Application struct {
}
