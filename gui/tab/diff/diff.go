package diff

import (
	"gioui.org/app"
	"gioui.org/layout"
	"gioui.org/widget"
	"gioui.org/widget/material"
	"gioui.org/x/component"
	"gioui.org/x/explorer"

	spsgio "github.com/SPSZerone/sps-go-zerone/graphics/gio"
	spsicon "github.com/SPSZerone/sps-go-zerone/graphics/gio/icon"
)

type Diff struct {
	widget.List

	fileInput    component.TextField
	fileButton   widget.Clickable
	reloadButton widget.Clickable
	explorer     *explorer.Explorer
}

func (d *Diff) Layout(app *spsgio.Application, gtx layout.Context, w *app.Window, th *material.Theme) layout.Dimensions {
	d.List.Axis = layout.Vertical
	return material.List(th, &d.List).Layout(gtx, 1, func(gtx layout.Context, _ int) layout.Dimensions {
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
						return d.fileInput.Layout(gtx, th, "File(xlsx,csv...)")
					}),
					layout.Rigid(func(gtx layout.Context) layout.Dimensions {
						if d.fileButton.Clicked(gtx) {
							app.Logger.Info().Msg("Open File")
						}
						return material.IconButton(th, &d.fileButton, spsicon.ActionOpenInNew, "Open File").Layout(gtx)
					}),
					layout.Rigid(func(gtx layout.Context) layout.Dimensions {
						if d.reloadButton.Clicked(gtx) {
							app.Logger.Info().Msg("Reload")
						}
						return material.IconButton(th, &d.reloadButton, spsicon.ActionUpdate, "Reload").Layout(gtx)
					}),
				)
			}),
		)
	})
}
