package gui

import (
	spsgio "github.com/SPSZerone/sps-go-zerone/graphics/gio"
	spspref "github.com/SPSZerone/sps-go-zerone/graphics/gio/tab/pref"

	"github.com/SPSZerone/sps-go-excel/gui/tab/about"
	"github.com/SPSZerone/sps-go-excel/gui/tab/diff"
)

func Run() {
	spsgio.Run(
		spsgio.OptTitle("SPS Excel Tools"),
		spsgio.OptOnInit(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools Init")
			app.Tabs.Register(0, about.New(&app.Tabs))
			app.Tabs.Register(1, spspref.New(&app.Tabs))
			app.Tabs.Register(2, diff.New(app))
			app.Tabs.SwitchTo(2)
		}),
		spsgio.OptOnStart(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools Start")
		}),
		spsgio.OptOnLoop(onLoop),
		spsgio.OptOnStop(func(app *spsgio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools Stop")
		}),
	)
}

func onLoop(app *spsgio.Application) error {
	app.Logger.Info().Msg("SPS Excel Tools Loop")
	for {
		destroy, err := app.OnEvent(app.Window.Event())
		if destroy {
			return err
		}
	}
}

type Application struct {
}
