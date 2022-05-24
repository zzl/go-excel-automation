package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024440-0000-0000-C000-000000000046
var IID_ControlFormat = syscall.GUID{0x00024440, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ControlFormat struct {
	ole.OleClient
}

func NewControlFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ControlFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &ControlFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ControlFormatFromVar(v ole.Variant) *ControlFormat {
	return NewControlFormat(v.IDispatch(), false, false)
}

func (this *ControlFormat) IID() *syscall.GUID {
	return &IID_ControlFormat
}

func (this *ControlFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ControlFormat) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ControlFormat) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ControlFormat) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ControlFormat) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ControlFormat) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ControlFormat) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ControlFormat) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ControlFormat) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ControlFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var ControlFormat_AddItem_OptArgs= []string{
	"Index", 
}

func (this *ControlFormat) AddItem(text string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ControlFormat_AddItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000353, []interface{}{text}, optArgs...)
	_= retVal
}

func (this *ControlFormat) RemoveAllItems()  {
	retVal, _ := this.Call(0x00000355, nil)
	_= retVal
}

var ControlFormat_RemoveItem_OptArgs= []string{
	"Count", 
}

func (this *ControlFormat) RemoveItem(index int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ControlFormat_RemoveItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000354, []interface{}{index}, optArgs...)
	_= retVal
}

func (this *ControlFormat) DropDownLines() int32 {
	retVal, _ := this.PropGet(0x00000350, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetDropDownLines(rhs int32)  {
	_ = this.PropPut(0x00000350, []interface{}{rhs})
}

func (this *ControlFormat) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ControlFormat) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *ControlFormat) LargeChange() int32 {
	retVal, _ := this.PropGet(0x0000034d, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetLargeChange(rhs int32)  {
	_ = this.PropPut(0x0000034d, []interface{}{rhs})
}

func (this *ControlFormat) LinkedCell() string {
	retVal, _ := this.PropGet(0x00000422, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ControlFormat) SetLinkedCell(rhs string)  {
	_ = this.PropPut(0x00000422, []interface{}{rhs})
}

var ControlFormat_List_OptArgs= []string{
	"Index", 
}

func (this *ControlFormat) List(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ControlFormat_List_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000035d, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ControlFormat) ListCount() int32 {
	retVal, _ := this.PropGet(0x00000351, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetListCount(rhs int32)  {
	_ = this.PropPut(0x00000351, []interface{}{rhs})
}

func (this *ControlFormat) ListFillRange() string {
	retVal, _ := this.PropGet(0x0000034f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ControlFormat) SetListFillRange(rhs string)  {
	_ = this.PropPut(0x0000034f, []interface{}{rhs})
}

func (this *ControlFormat) ListIndex() int32 {
	retVal, _ := this.PropGet(0x00000352, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetListIndex(rhs int32)  {
	_ = this.PropPut(0x00000352, []interface{}{rhs})
}

func (this *ControlFormat) LockedText() bool {
	retVal, _ := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ControlFormat) SetLockedText(rhs bool)  {
	_ = this.PropPut(0x00000268, []interface{}{rhs})
}

func (this *ControlFormat) Max() int32 {
	retVal, _ := this.PropGet(0x0000034a, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetMax(rhs int32)  {
	_ = this.PropPut(0x0000034a, []interface{}{rhs})
}

func (this *ControlFormat) Min() int32 {
	retVal, _ := this.PropGet(0x0000034b, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetMin(rhs int32)  {
	_ = this.PropPut(0x0000034b, []interface{}{rhs})
}

func (this *ControlFormat) MultiSelect() int32 {
	retVal, _ := this.PropGet(0x00000020, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetMultiSelect(rhs int32)  {
	_ = this.PropPut(0x00000020, []interface{}{rhs})
}

func (this *ControlFormat) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ControlFormat) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

func (this *ControlFormat) SmallChange() int32 {
	retVal, _ := this.PropGet(0x0000034c, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetSmallChange(rhs int32)  {
	_ = this.PropPut(0x0000034c, []interface{}{rhs})
}

func (this *ControlFormat) Default_() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetDefault_(rhs int32)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *ControlFormat) Value() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *ControlFormat) SetValue(rhs int32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

