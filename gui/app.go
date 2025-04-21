package gui

import (
	"gioui.org/io/system"

	spsgio "github.com/SPSZerone/sps-go-zerone/graphics/gio"
	spsapp "github.com/SPSZerone/sps-go-zerone/graphics/gio/app"
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

const (
	Name = "SPS Excel Tools"
)

func Run() {
	spsapp.Run(
		NewWindow,
		spsapp.OptOnCreate(func(app spsgio.App) {
			app.GetLogger().Info().Msgf("Window %s Create", Name)
		}),
		spsapp.OptOnStart(func(app spsgio.App) {
			app.GetLogger().Info().Msgf("Window %s Start", Name)
		}),
		spsapp.OptOnStop(func(app spsgio.App) {
			app.GetLogger().Info().Msgf("Window %s Stop", Name)
		}),
	)
}

func NewWindow(app spsgio.App, fromWin spsgio.Window) spsgio.Window {
	pref := app.GetPref()
	if fromWin != nil {
		pref = fromWin.GetPref()
	}
	return spswin.NewWindow(
		app.GetContext(),
		spswin.OptID("Main"),
		spswin.OptTitle(Name),
		spswin.OptPref(*pref),
		spswin.OptStartAction(system.ActionMaximize),
		spswin.OptLoopMode(spswin.LoopModeSimple),
		spswin.OptOnInitPre(func(win spsgio.Window) {
			win.GetLogger().Info().Msgf("%s InitPre", win.LogPrefix())
		}),
		spswin.OptOnInitPost(func(win spsgio.Window) {
			win.GetLogger().Info().Msgf("%s InitPost", win.LogPrefix())

			pages := win.GetPages()
			pageTag := 0
			pages.Register(pageTag, diff.New(pages))

			pageTag++
			pages.Register(pageTag, about.New(pages))

			pageTag++
			prefPage := spspref.New(pages, app, NewWindow)
			prefPage.Tabs.SetSelected(spspref.TabIdxSettings)
			pages.Register(pageTag, prefPage)
		}),
		spswin.OptOnStart(func(win spsgio.Window) {
			win.GetLogger().Info().Msgf("%s Start", win.LogPrefix())
			win.GetPages().Start(0)
		}),
		spswin.OptOnStop(func(win spsgio.Window) {
			win.GetLogger().Info().Msgf("%s Stop", win.LogPrefix())
		}),
	)
}
