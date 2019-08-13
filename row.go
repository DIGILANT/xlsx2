package xlsx

import (
	"sync"
)

type Row struct {
	Cells        []*Cell
	Hidden       bool
	Sheet        *Sheet
	Height       float64
	OutlineLevel uint8
	isCustom     bool
}

var rowPool = sync.Pool{
	New: func() interface{} {
		return &Row{}
	},
}

func acquireRow() *Row {
	return rowPool.Get().(*Row)
}

func releaseRow(row *Row) {
	for _, cell := range row.Cells {
		releaseCell(cell)
	}
	row.Cells = row.Cells[:0]
	row.Hidden = false
	row.Sheet = nil
	row.Height = 0
	row.OutlineLevel = 0
	row.isCustom = false
}

func (r *Row) SetHeight(ht float64) {
	r.Height = ht
	r.isCustom = true
}

func (r *Row) SetHeightCM(ht float64) {
	r.Height = ht * 28.3464567 // Convert CM to postscript points
	r.isCustom = true
}

func (r *Row) AddCell() *Cell {
	cell := NewCell(r)
	r.Cells = append(r.Cells, cell)
	r.Sheet.maybeAddCol(len(r.Cells))
	return cell
}
