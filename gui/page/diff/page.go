package diff

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

var _ spsgio.Page = (*Page)(nil)

type Page struct {
	widget.List
	*spsgio.Pages

	split spslayout.Split
	left  Diff
	right Diff
}

func New(pages *spsgio.Pages) *Page {
	return &Page{
		Pages: pages,
	}
}

func (p *Page) Actions() []component.AppBarAction {
	return []component.AppBarAction{}
}

func (p *Page) Overflow() []component.OverflowAction {
	return []component.OverflowAction{}
}

func (p *Page) NavItem() component.NavItem {
	return component.NavItem{
		Name: "Diff",
		Icon: spsicon.ActionCompareArrows,
	}
}

func (p *Page) Layout(app *spsgio.Application, gtx layout.Context, w *app.Window, th *material.Theme) layout.Dimensions {
	p.List.Axis = layout.Vertical
	return material.List(th, &p.List).Layout(gtx, 1, func(gtx layout.Context, _ int) layout.Dimensions {
		return layout.Flex{
			Alignment: layout.Middle,
			Axis:      layout.Vertical,
		}.Layout(gtx,
			layout.Rigid(func(gtx layout.Context) layout.Dimensions {
				return p.split.Layout(
					gtx,
					func(gtx layout.Context) layout.Dimensions {
						return p.left.Layout(app, gtx, w, th)
					},
					func(gtx layout.Context) layout.Dimensions {
						return p.right.Layout(app, gtx, w, th)
					},
				)
			}),
		)
	})
}
