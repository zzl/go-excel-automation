package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002446A-0000-0000-C000-000000000046
var IID_AllowEditRanges = syscall.GUID{0x0002446A, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AllowEditRanges struct {
	ole.OleClient
}

func NewAllowEditRanges(pDisp *win32.IDispatch, addRef bool, scoped bool) *AllowEditRanges {
	if pDisp == nil {
		return nil
	}
	p := &AllowEditRanges{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AllowEditRangesFromVar(v ole.Variant) *AllowEditRanges {
	return NewAllowEditRanges(v.IDispatch(), false, false)
}

func (this *AllowEditRanges) IID() *syscall.GUID {
	return &IID_AllowEditRanges
}

func (this *AllowEditRanges) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AllowEditRanges) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *AllowEditRanges) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *AllowEditRanges) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *AllowEditRanges) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *AllowEditRanges) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *AllowEditRanges) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *AllowEditRanges) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *AllowEditRanges) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *AllowEditRanges) Item(index interface{}) *AllowEditRange {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewAllowEditRange(retVal.IDispatch(), false, true)
}

var AllowEditRanges_Add_OptArgs = []string{
	"Password",
}

func (this *AllowEditRanges) Add(title string, range_ *Range, optArgs ...interface{}) *AllowEditRange {
	optArgs = ole.ProcessOptArgs(AllowEditRanges_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{title, range_}, optArgs...)
	return NewAllowEditRange(retVal.IDispatch(), false, true)
}

func (this *AllowEditRanges) Default_(index interface{}) *AllowEditRange {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewAllowEditRange(retVal.IDispatch(), false, true)
}

func (this *AllowEditRanges) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *AllowEditRanges) ForEach(action func(item *AllowEditRange) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release()
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*AllowEditRange)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}
