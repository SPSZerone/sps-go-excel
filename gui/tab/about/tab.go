package about

import (
	"gioui.org/app"
	"gioui.org/layout"
	"gioui.org/widget"
	"gioui.org/widget/material"
	"gioui.org/x/component"

	spsgio "github.com/SPSZerone/sps-go-zerone/graphics/gio"
	spsicon "github.com/SPSZerone/sps-go-zerone/graphics/gio/icon"
	spslayout "github.com/SPSZerone/sps-go-zerone/graphics/gio/layout"
)

var _ spsgio.Page = (*Tab)(nil)

type Tab struct {
	widget.List
	*spsgio.Pages
}

func New(pages *spsgio.Pages) *Tab {
	return &Tab{
		Pages: pages,
	}
}

func (p *Tab) Actions() []component.AppBarAction {
	return []component.AppBarAction{}
}

func (p *Tab) Overflow() []component.OverflowAction {
	return []component.OverflowAction{}
}

func (p *Tab) NavItem() component.NavItem {
	return component.NavItem{
		Name: "About",
		Icon: spsicon.ActionHelp,
	}
}

func (p *Tab) Layout(app *spsgio.Application, gtx layout.Context, w *app.Window, th *material.Theme) layout.Dimensions {
	p.List.Axis = layout.Vertical
	return material.List(th, &p.List).Layout(gtx, 1, func(gtx layout.Context, _ int) layout.Dimensions {
		return layout.Flex{
			Alignment: layout.Middle,
			Axis:      layout.Vertical,
		}.Layout(gtx,
			layout.Rigid(func(gtx layout.Context) layout.Dimensions {
				return spslayout.DefaultInset.Layout(gtx, material.Body1(th, `SPS Excel Tools`).Layout)
			}),
		)
	})
}
