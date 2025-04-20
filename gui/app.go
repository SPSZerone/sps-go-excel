package gui

import (
	"gioui.org/io/system"

	spsgio "github.com/SPSZerone/sps-go-zerone/graphics/gio"
	spspref "github.com/SPSZerone/sps-go-zerone/graphics/gio/page/pref"
	spswin "github.com/SPSZerone/sps-go-zerone/graphics/gio/window"

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
		spsgio.OptOnCreate(func(app *spsgio.App) {
			app.Logger.Info().Msg("Window SPS Excel Tools Create")
		}),
		spsgio.OptOnStart(func(app *spsgio.App) {
			app.Logger.Info().Msg("Window SPS Excel Tools Start")
		}),
		spsgio.OptOnStop(func(app *spsgio.App) {
			app.Logger.Info().Msg("Window SPS Excel Tools Stop")
		}),
		spsgio.OptWinOpts(
			spswin.OptTitle("SPS Excel Tools"),
			spswin.OptStartAction(system.ActionMaximize),
			spswin.OptLoopMode(spswin.LoopModeSimple),
			spswin.OptOnInitPre(func(win *spswin.Window) {
				win.Logger.Info().Msgf("%s InitPre", win.LogPrefix())
			}),
			spswin.OptOnInitPost(func(win *spswin.Window) {
				win.Logger.Info().Msgf("%s InitPost", win.LogPrefix())

				pages := &win.Pages
				pageTag := 0
				pages.Register(pageTag, diff.New(pages))

				pageTag++
				pages.Register(pageTag, about.New(pages))

				pageTag++
				pref := spspref.New(pages)
				pref.Tabs.SetSelected(spspref.TabIdxSettings)
				pages.Register(pageTag, pref)
			}),
			spswin.OptOnStart(func(win *spswin.Window) {
				win.Logger.Info().Msgf("%s Start", win.LogPrefix())
				win.Pages.Start(0)
			}),
			spswin.OptOnStop(func(win *spswin.Window) {
				win.Logger.Info().Msgf("%s Stop", win.LogPrefix())
			}),
		),
	)
}
