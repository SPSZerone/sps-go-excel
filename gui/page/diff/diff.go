package diff

import (
	"gioui.org/app"
	"gioui.org/layout"
	"gioui.org/widget"
	"gioui.org/widget/material"
	"gioui.org/x/component"

	"github.com/SPSZerone/sps-go-zerone/graphics/gio"
	spslayout "github.com/SPSZerone/sps-go-zerone/graphics/gio/layout"
)

type Diff struct {
	widget.List

	fileInput    component.TextField
	fileButton   widget.Clickable
	reloadButton widget.Clickable
}

func (d *Diff) Layout(app *gio.Application, gtx layout.Context, w *app.Window, th *material.Theme) layout.Dimensions {
	d.List.Axis = layout.Vertical
	return material.List(th, &d.List).Layout(gtx, 1, func(gtx layout.Context, _ int) layout.Dimensions {
		return layout.Flex{Axis: layout.Vertical}.Layout(
			gtx,
			layout.Rigid(func(gtx layout.Context) layout.Dimensions {
				return spslayout.FlexInset{
					Ratio: 0.8,
					Flex: layout.Flex{
						Axis:      layout.Horizontal,
						Alignment: layout.Baseline,
						Spacing:   layout.SpaceAround,
					},
				}.LayoutWidgets(
					gtx,
					func() (float32, layout.Widget) {
						return 0.7, func(gtx layout.Context) layout.Dimensions {
							return d.fileInput.Layout(gtx, th, "File(xlsx,csv...)")
						}
					},
					func() (float32, layout.Widget) {
						return 0.15, func(gtx layout.Context) layout.Dimensions {
							if d.fileButton.Clicked(gtx) {
								app.Logger.Info().Msg("Open File")
							}
							return material.Button(th, &d.fileButton, "Open File").Layout(gtx)
						}
					},
					func() (float32, layout.Widget) {
						return 0.15, func(gtx layout.Context) layout.Dimensions {
							if d.reloadButton.Clicked(gtx) {
								app.Logger.Info().Msg("Reload")
							}
							return material.Button(th, &d.reloadButton, "Reload").Layout(gtx)
						}
					},
				)
				//}.LayoutABWidget(
				//	gtx,
				//	func(gtx layout.Context) layout.Dimensions {
				//		return d.fileInput.Layout(gtx, th, "File(xlsx,csv...)")
				//	},
				//	func(gtx layout.Context) layout.Dimensions {
				//		if d.fileButton.Clicked(gtx) {
				//			app.Logger.Info().Msg("Open File")
				//		}
				//		return material.Button(th, &d.fileButton, "Open File").Layout(gtx)
				//	},
				//)
			}),
		)
	})
}
