package diff

import (
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

	split spslayout.Split
	left  Diff
	right Diff
}

func New(tabs *spsgio.Tabs) *Tab {
	return &Tab{
		Tabs: tabs,
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
		Name: "Diff",
		Icon: spsicon.ActionCompareArrows,
	}
}

func (p *Tab) Layout(app *spsgio.Application, gtx layout.Context) layout.Dimensions {
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
						return p.left.Layout(app, gtx, app.Window, app.Theme)
					},
					func(gtx layout.Context) layout.Dimensions {
						return p.right.Layout(app, gtx, app.Window, app.Theme)
					},
				)
			}),
		)
	})
}
