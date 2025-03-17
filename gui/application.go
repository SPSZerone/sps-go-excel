package gui

import (
	"gioui.org/app"
	"gioui.org/io/event"

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

func onLoop(a *spsgio.Application) error {
	a.Logger.Info().Msg("SPS Excel Tools Loop")

	chanEvent := make(chan event.Event)
	chanEventDone := make(chan struct{})

	a.GoRun(func() {
		for {
			evt := a.Window.Event()
			chanEvent <- evt
			<-chanEventDone
			if _, ok := evt.(app.DestroyEvent); ok {
				a.Logger.Info().Msg("Window.Event app.DestroyEvent ...")
				return
			}
		}
	})

	for {
		select {
		case evt := <-chanEvent:
			switch e := evt.(type) {
			case app.DestroyEvent:
				a.Logger.Info().Msg("chanEvent app.DestroyEvent ...")
				chanEventDone <- struct{}{}
				return e.Err
			case app.FrameEvent:
				a.OnFrameEvent(e)
			}

			chanEventDone <- struct{}{}
		}
	}
}

type Application struct {
}
