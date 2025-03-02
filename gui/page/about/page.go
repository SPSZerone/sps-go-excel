package about

import (
	"gioui.org/app"
	"gioui.org/layout"
	"gioui.org/widget"
	"gioui.org/widget/material"
	"gioui.org/x/component"

	"github.com/SPSZerone/sps-go-zerone/graphics/gio"
	"github.com/SPSZerone/sps-go-zerone/graphics/gio/icon"
	spslayout "github.com/SPSZerone/sps-go-zerone/graphics/gio/layout"
)

var _ gio.Page = (*Page)(nil)

type Page struct {
	widget.List
	*gio.Pages
}

func New(pages *gio.Pages) *Page {
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
		Name: "About",
		Icon: icon.ActionHelp,
	}
}

func (p *Page) Layout(app *gio.Application, gtx layout.Context, w *app.Window, th *material.Theme) layout.Dimensions {
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
