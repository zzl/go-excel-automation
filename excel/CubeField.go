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
	 if pDisp == nil {
		return nil;
	}
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
	return NewCubeField(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CubeField) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *CubeField) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CubeField) CubeFieldType() int32 {
	retVal, _ := this.PropGet(0x0000077e, nil)
	return retVal.LValVal()
}

func (this *CubeField) Caption_() string {
	retVal, _ := this.PropGet(0x00000a6b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) Value() string {
	retVal, _ := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) Orientation() int32 {
	retVal, _ := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetOrientation(rhs int32)  {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *CubeField) Position() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetPosition(rhs int32)  {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *CubeField) TreeviewControl() *TreeviewControl {
	retVal, _ := this.PropGet(0x0000077f, nil)
	return NewTreeviewControl(retVal.IDispatch(), false, true)
}

func (this *CubeField) DragToColumn() bool {
	retVal, _ := this.PropGet(0x000005e4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToColumn(rhs bool)  {
	_ = this.PropPut(0x000005e4, []interface{}{rhs})
}

func (this *CubeField) DragToHide() bool {
	retVal, _ := this.PropGet(0x000005e5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToHide(rhs bool)  {
	_ = this.PropPut(0x000005e5, []interface{}{rhs})
}

func (this *CubeField) DragToPage() bool {
	retVal, _ := this.PropGet(0x000005e6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToPage(rhs bool)  {
	_ = this.PropPut(0x000005e6, []interface{}{rhs})
}

func (this *CubeField) DragToRow() bool {
	retVal, _ := this.PropGet(0x000005e7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToRow(rhs bool)  {
	_ = this.PropPut(0x000005e7, []interface{}{rhs})
}

func (this *CubeField) DragToData() bool {
	retVal, _ := this.PropGet(0x00000734, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetDragToData(rhs bool)  {
	_ = this.PropPut(0x00000734, []interface{}{rhs})
}

func (this *CubeField) HiddenLevels() int32 {
	retVal, _ := this.PropGet(0x00000780, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetHiddenLevels(rhs int32)  {
	_ = this.PropPut(0x00000780, []interface{}{rhs})
}

func (this *CubeField) HasMemberProperties() bool {
	retVal, _ := this.PropGet(0x00000885, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) LayoutForm() int32 {
	retVal, _ := this.PropGet(0x00000738, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetLayoutForm(rhs int32)  {
	_ = this.PropPut(0x00000738, []interface{}{rhs})
}

func (this *CubeField) PivotFields() *PivotFields {
	retVal, _ := this.PropGet(0x000002ce, nil)
	return NewPivotFields(retVal.IDispatch(), false, true)
}

var CubeField_AddMemberPropertyField__OptArgs= []string{
	"PropertyOrder", 
}

func (this *CubeField) AddMemberPropertyField_(property string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(CubeField_AddMemberPropertyField__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000886, []interface{}{property}, optArgs...)
	_= retVal
}

func (this *CubeField) EnableMultiplePageItems() bool {
	retVal, _ := this.PropGet(0x00000888, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetEnableMultiplePageItems(rhs bool)  {
	_ = this.PropPut(0x00000888, []interface{}{rhs})
}

func (this *CubeField) LayoutSubtotalLocation() int32 {
	retVal, _ := this.PropGet(0x00000736, nil)
	return retVal.LValVal()
}

func (this *CubeField) SetLayoutSubtotalLocation(rhs int32)  {
	_ = this.PropPut(0x00000736, []interface{}{rhs})
}

func (this *CubeField) ShowInFieldList() bool {
	retVal, _ := this.PropGet(0x00000889, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetShowInFieldList(rhs bool)  {
	_ = this.PropPut(0x00000889, []interface{}{rhs})
}

func (this *CubeField) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

var CubeField_AddMemberPropertyField_OptArgs= []string{
	"PropertyOrder", "PropertyDisplayedIn", 
}

func (this *CubeField) AddMemberPropertyField(property string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(CubeField_AddMemberPropertyField_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000a6c, []interface{}{property}, optArgs...)
	_= retVal
}

func (this *CubeField) IncludeNewItemsInFilter() bool {
	retVal, _ := this.PropGet(0x00000a1b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetIncludeNewItemsInFilter(rhs bool)  {
	_ = this.PropPut(0x00000a1b, []interface{}{rhs})
}

func (this *CubeField) CubeFieldSubType() int32 {
	retVal, _ := this.PropGet(0x00000a6e, nil)
	return retVal.LValVal()
}

func (this *CubeField) AllItemsVisible() bool {
	retVal, _ := this.PropGet(0x00000a21, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) ClearManualFilter()  {
	retVal, _ := this.Call(0x00000a22, nil)
	_= retVal
}

func (this *CubeField) CreatePivotFields()  {
	retVal, _ := this.Call(0x00000a6f, nil)
	_= retVal
}

func (this *CubeField) CurrentPageName() string {
	retVal, _ := this.PropGet(0x0000073c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) SetCurrentPageName(rhs string)  {
	_ = this.PropPut(0x0000073c, []interface{}{rhs})
}

func (this *CubeField) IsDate() bool {
	retVal, _ := this.PropGet(0x00000a70, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CubeField) SetCaption(rhs string)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

func (this *CubeField) FlattenHierarchies() bool {
	retVal, _ := this.PropGet(0x00000b6c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetFlattenHierarchies(rhs bool)  {
	_ = this.PropPut(0x00000b6c, []interface{}{rhs})
}

func (this *CubeField) HierarchizeDistinct() bool {
	retVal, _ := this.PropGet(0x00000b6d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CubeField) SetHierarchizeDistinct(rhs bool)  {
	_ = this.PropPut(0x00000b6d, []interface{}{rhs})
}

