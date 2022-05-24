package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208AB-0000-0000-C000-000000000046
var IID_Outline = syscall.GUID{0x000208AB, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Outline struct {
	ole.OleClient
}

func NewOutline(pDisp *win32.IDispatch, addRef bool, scoped bool) *Outline {
	 if pDisp == nil {
		return nil;
	}
	p := &Outline{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OutlineFromVar(v ole.Variant) *Outline {
	return NewOutline(v.IDispatch(), false, false)
}

func (this *Outline) IID() *syscall.GUID {
	return &IID_Outline
}

func (this *Outline) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Outline) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Outline) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Outline) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Outline) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Outline) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Outline) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Outline) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Outline) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Outline) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Outline) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Outline) AutomaticStyles() bool {
	retVal, _ := this.PropGet(0x000003bf, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Outline) SetAutomaticStyles(rhs bool)  {
	_ = this.PropPut(0x000003bf, []interface{}{rhs})
}

var Outline_ShowLevels_OptArgs= []string{
	"RowLevels", "ColumnLevels", 
}

func (this *Outline) ShowLevels(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Outline_ShowLevels_OptArgs, optArgs)
	retVal, _ := this.Call(0x000003c0, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Outline) SummaryColumn() int32 {
	retVal, _ := this.PropGet(0x000003c1, nil)
	return retVal.LValVal()
}

func (this *Outline) SetSummaryColumn(rhs int32)  {
	_ = this.PropPut(0x000003c1, []interface{}{rhs})
}

func (this *Outline) SummaryRow() int32 {
	retVal, _ := this.PropGet(0x00000386, nil)
	return retVal.LValVal()
}

func (this *Outline) SetSummaryRow(rhs int32)  {
	_ = this.PropPut(0x00000386, []interface{}{rhs})
}

