package diff

import (
	"gioui.org/io/event"
	"gioui.org/layout"
	"gioui.org/widget"
	"gioui.org/widget/material"
	"gioui.org/x/component"
	"gioui.org/x/explorer"

	spsgio "github.com/SPSZerone/sps-go-zerone/graphics/gio"
	spsicon "github.com/SPSZerone/sps-go-zerone/graphics/gio/icon"
	spslayout "github.com/SPSZerone/sps-go-zerone/graphics/gio/layout"
)

func New(app *spsgio.Application) *Page {
	t := &Page{
		Tabs: &app.Tabs,
	}
	t.left.explorer = explorer.NewExplorer(app.Window)
	t.right.explorer = explorer.NewExplorer(app.Window)
	return t
}

var _ spsgio.Tab = (*Page)(nil)

type Page struct {
	widget.List
	*spsgio.Tabs

	split spslayout.Split
	left  Diff
	right Diff
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

func (p *Page) OnEventPre(app *spsgio.Application, evt event.Event, param any) {
	p.left.explorer.ListenEvents(evt)
	p.right.explorer.ListenEvents(evt)
}

func (p *Page) OnEventPost(app *spsgio.Application, evt event.Event, param any) {

}

func (p *Page) Layout(app *spsgio.Application, gtx layout.Context, param any) layout.Dimensions {
	p.List.Axis = layout.Vertical
	return material.List(app.Theme, &p.List).Layout(gtx, 1, func(gtx layout.Context, _ int) layout.Dimensions {
		return layout.Flex{
			Alignment: layout.Middle,
			Axis:      layout.Vertical,
		}.Layout(gtx,
			layout.Rigid(func(gtx layout.Context) layout.Dimensions {
				return p.split.Layout(
					gtx,
					func(gtx layout.Context) layout.Dimensions {
						return p.left.Layout(app, gtx, param)
					},
					func(gtx layout.Context) layout.Dimensions {
						return p.right.Layout(app, gtx, param)
					},
				)
			}),
		)
	})
}
