package about

import (
	"gioui.org/io/event"
	"gioui.org/layout"
	"gioui.org/widget"
	"gioui.org/widget/material"
	"gioui.org/x/component"
	spsgio "github.com/SPSZerone/sps-go-zerone/graphics/gio"

	spsicon "github.com/SPSZerone/sps-go-zerone/graphics/gio/icon"
	spslayout "github.com/SPSZerone/sps-go-zerone/graphics/gio/layout"
	spswin "github.com/SPSZerone/sps-go-zerone/graphics/gio/window"
)

func New(pages *spswin.Pages) *Page {
	return &Page{
		Pages: pages,
	}
}

var _ spsgio.Page = (*Page)(nil)

type Page struct {
	widget.List
	*spswin.Pages
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
		Icon: spsicon.ActionHelp,
	}
}

func (p *Page) OnEventPre(win spsgio.Window, evt event.Event, param any) {

}

func (p *Page) OnEventPost(win spsgio.Window, evt event.Event, param any) {

}

func (p *Page) Layout(win spsgio.Window, gtx layout.Context, param any) layout.Dimensions {
	theme := win.GetTheme()
	p.List.Axis = layout.Vertical
	return material.List(theme, &p.List).Layout(gtx, 1, func(gtx layout.Context, _ int) layout.Dimensions {
		return layout.Flex{
			Alignment: layout.Middle,
			Axis:      layout.Vertical,
		}.Layout(gtx,
			layout.Rigid(func(gtx layout.Context) layout.Dimensions {
				return spslayout.DefaultInset.Layout(gtx, material.Body1(theme, `SPS Excel Tools`).Layout)
			}),
		)
	})
}
