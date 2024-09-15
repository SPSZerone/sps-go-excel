package main

import (
	"github.com/SPSZerone/sps-go-zerone/graphics/gio"
	"github.com/SPSZerone/sps-go-zerone/graphics/gio/page"
	"github.com/SPSZerone/sps-go-zerone/graphics/gio/page/pref"

	"github.com/SPSZerone/sps-go-excel/gui/page/about"
	"github.com/SPSZerone/sps-go-excel/gui/page/diff"
)

func main() {
	gio.Run(
		gio.OptTitle("SPS Excel Tools"),
		gio.OptOnStart(func(app *gio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools Start")
		}),
		gio.OptOnEnd(func(app *gio.Application) {
			app.Logger.Info().Msg("SPS Excel Tools End")
		}),
		gio.OptRegisterPage(func(pages *page.Pages) {
			pages.Register(0, about.New(pages))
			pages.Register(1, pref.New(pages))
			pages.Register(2, diff.New(pages))
		}),
	)
}
