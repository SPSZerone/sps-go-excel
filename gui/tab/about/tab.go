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
)

var _ spsgio.Tab = (*Tab)(nil)

type Tab struct {
	widget.List
	*spsgio.Tabs
}

func New(tabs *spsgio.Tabs) *Tab {
	return &Tab{
		Tabs: tabs,
	}
}

func (t *Tab) Actions() []component.AppBarAction {
	return []component.AppBarAction{}
}

func (t *Tab) Overflow() []component.OverflowAction {
	return []component.OverflowAction{}
}

func (t *Tab) NavItem() component.NavItem {
	return component.NavItem{
		Name: "About",
		Icon: spsicon.ActionHelp,
	}
}

func (t *Tab) OnEventPre(app *spsgio.Application, evt event.Event, param any) {

}

func (t *Tab) OnEventPost(app *spsgio.Application, evt event.Event, param any) {

}

func (t *Tab) Layout(app *spsgio.Application, gtx layout.Context, param any) layout.Dimensions {
	t.List.Axis = layout.Vertical
	return material.List(app.Theme, &t.List).Layout(gtx, 1, func(gtx layout.Context, _ int) layout.Dimensions {
		return layout.Flex{
			Alignment: layout.Middle,
			Axis:      layout.Vertical,
		}.Layout(gtx,
			layout.Rigid(func(gtx layout.Context) layout.Dimensions {
				return spslayout.DefaultInset.Layout(gtx, material.Body1(app.Theme, `SPS Excel Tools`).Layout)
			}),
		)
	})
}
