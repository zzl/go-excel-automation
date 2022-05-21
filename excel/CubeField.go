package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002444C-0000-0000-C000-000000000046
var IID_CubeField = syscall.GUID{0x0002444C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CubeField struct {
	ole.OleClient
}

func NewCubeField(pDisp *win32.IDispatch, addRef bool, scoped bool) *CubeField {
	p := &CubeField{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CubeFieldFromVar(v ole.Variant) *CubeField {
	return NewCubeField(v.PdispValVal(), false, false)
}

func (this *CubeField) IID() *syscall.GUID {
	return &IID_CubeField
}

func (this *CubeField) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CubeField) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CubeField) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *CubeField) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CubeField) CubeFieldType() int32 {
	retVal := this.PropGet(0x0000077e, nil)
	return retVal.LValVal()
}

func (this *CubeField) Caption_() string {
	retVal := this.PropGet(0x00000a6b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) Value() string {
	retVal := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) Orientation() int32 {
	retVal := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetOrientation(rhs int32)  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) Position() int32 {
	retVal := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetPosition(rhs int32)  {
	retVal := this.PropPut(0x00000085, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) TreeviewControl() *TreeviewControl {
	retVal := this.PropGet(0x0000077f, nil)
	return NewTreeviewControl(retVal.PdispValVal(), false, true)
}

func (this *CubeField) DragToColumn() bool {
	retVal := this.PropGet(0x000005e4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToColumn(rhs bool)  {
	retVal := this.PropPut(0x000005e4, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) DragToHide() bool {
	retVal := this.PropGet(0x000005e5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToHide(rhs bool)  {
	retVal := this.PropPut(0x000005e5, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) DragToPage() bool {
	retVal := this.PropGet(0x000005e6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToPage(rhs bool)  {
	retVal := this.PropPut(0x000005e6, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) DragToRow() bool {
	retVal := this.PropGet(0x000005e7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToRow(rhs bool)  {
	retVal := this.PropPut(0x000005e7, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) DragToData() bool {
	retVal := this.PropGet(0x00000734, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToData(rhs bool)  {
	retVal := this.PropPut(0x00000734, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) HiddenLevels() int32 {
	retVal := this.PropGet(0x00000780, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetHiddenLevels(rhs int32)  {
	retVal := this.PropPut(0x00000780, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) HasMemberProperties() bool {
	retVal := this.PropGet(0x00000885, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) LayoutForm() int32 {
	retVal := this.PropGet(0x00000738, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetLayoutForm(rhs int32)  {
	retVal := this.PropPut(0x00000738, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) PivotFields() *PivotFields {
	retVal := this.PropGet(0x000002ce, nil)
	return NewPivotFields(retVal.PdispValVal(), false, true)
}

var CubeField_AddMemberPropertyField__OptArgs= []string{
	"PropertyOrder", 
}

func (this *CubeField) AddMemberPropertyField_(property string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(CubeField_AddMemberPropertyField__OptArgs, optArgs)
	retVal := this.Call(0x00000886, []interface{}{property}, optArgs...)
	_= retVal
}

func (this *CubeField) EnableMultiplePageItems() bool {
	retVal := this.PropGet(0x00000888, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetEnableMultiplePageItems(rhs bool)  {
	retVal := this.PropPut(0x00000888, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) LayoutSubtotalLocation() int32 {
	retVal := this.PropGet(0x00000736, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetLayoutSubtotalLocation(rhs int32)  {
	retVal := this.PropPut(0x00000736, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) ShowInFieldList() bool {
	retVal := this.PropGet(0x00000889, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetShowInFieldList(rhs bool)  {
	retVal := this.PropPut(0x00000889, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

var CubeField_AddMemberPropertyField_OptArgs= []string{
	"PropertyOrder", "PropertyDisplayedIn", 
}

func (this *CubeField) AddMemberPropertyField(property string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(CubeField_AddMemberPropertyField_OptArgs, optArgs)
	retVal := this.Call(0x00000a6c, []interface{}{property}, optArgs...)
	_= retVal
}

func (this *CubeField) IncludeNewItemsInFilter() bool {
	retVal := this.PropGet(0x00000a1b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetIncludeNewItemsInFilter(rhs bool)  {
	retVal := this.PropPut(0x00000a1b, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) CubeFieldSubType() int32 {
	retVal := this.PropGet(0x00000a6e, nil)
	return retVal.LValVal()
}

func (this *CubeField) AllItemsVisible() bool {
	retVal := this.PropGet(0x00000a21, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) ClearManualFilter()  {
	retVal := this.Call(0x00000a22, nil)
	_= retVal
}

func (this *CubeField) CreatePivotFields()  {
	retVal := this.Call(0x00000a6f, nil)
	_= retVal
}

func (this *CubeField) CurrentPageName() string {
	retVal := this.PropGet(0x0000073c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) SetCurrentPageName(rhs string)  {
	retVal := this.PropPut(0x0000073c, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) IsDate() bool {
	retVal := this.PropGet(0x00000a70, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) FlattenHierarchies() bool {
	retVal := this.PropGet(0x00000b6c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetFlattenHierarchies(rhs bool)  {
	retVal := this.PropPut(0x00000b6c, []interface{}{rhs})
	_= retVal
}

func (this *CubeField) HierarchizeDistinct() bool {
	retVal := this.PropGet(0x00000b6d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetHierarchizeDistinct(rhs bool)  {
	retVal := this.PropPut(0x00000b6d, []interface{}{rhs})
	_= retVal
}

