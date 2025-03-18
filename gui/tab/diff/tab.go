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

var _ spsgio.Tab = (*Tab)(nil)

type Tab struct {
	widget.List
	*spsgio.Tabs

	split spslayout.Split
	left  Diff
	right Diff
}

func New(app *spsgio.Application) *Tab {
	t := &Tab{
		Tabs: &app.Tabs,
	}
	t.left.explorer = explorer.NewExplorer(app.Window)
	t.right.explorer = explorer.NewExplorer(app.Window)
	return t
}

func (t *Tab) Actions() []component.AppBarAction {
	return []component.AppBarAction{}
}

func (t *Tab) Overflow() []component.OverflowAction {
	return []component.OverflowAction{}
}

func (t *Tab) NavItem() component.NavItem {
	return component.NavItem{
		Name: "Diff",
		Icon: spsicon.ActionCompareArrows,
	}
}

func (t *Tab) OnEventPre(app *spsgio.Application, evt event.Event, param any) {
	t.left.explorer.ListenEvents(evt)
	t.right.explorer.ListenEvents(evt)
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
				return t.split.Layout(
					gtx,
					func(gtx layout.Context) layout.Dimensions {
						return t.left.Layout(app, gtx, param)
					},
					func(gtx layout.Context) layout.Dimensions {
						return t.right.Layout(app, gtx, param)
					},
				)
			}),
		)
	})
}
