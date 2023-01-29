package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002441D-0000-0000-C000-000000000046
var IID_PivotCaches = syscall.GUID{0x0002441D, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotCaches struct {
	ole.OleClient
}

func NewPivotCaches(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotCaches {
	if pDisp == nil {
		return nil
	}
	p := &PivotCaches{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotCachesFromVar(v ole.Variant) *PivotCaches {
	return NewPivotCaches(v.IDispatch(), false, false)
}

func (this *PivotCaches) IID() *syscall.GUID {
	return &IID_PivotCaches
}

func (this *PivotCaches) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotCaches) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PivotCaches) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotCaches) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotCaches) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PivotCaches) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PivotCaches) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PivotCaches) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PivotCaches) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotCaches) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotCaches) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotCaches) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *PivotCaches) Item(index interface{}) *PivotCache {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return NewPivotCache(retVal.IDispatch(), false, true)
}

func (this *PivotCaches) Default_(index interface{}) *PivotCache {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewPivotCache(retVal.IDispatch(), false, true)
}

func (this *PivotCaches) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *PivotCaches) ForEach(action func(item *PivotCache) bool) {
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
		pItem := (*PivotCache)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var PivotCaches_Add_OptArgs = []string{
	"SourceData",
}

func (this *PivotCaches) Add(sourceType int32, optArgs ...interface{}) *PivotCache {
	optArgs = ole.ProcessOptArgs(PivotCaches_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{sourceType}, optArgs...)
	return NewPivotCache(retVal.IDispatch(), false, true)
}

var PivotCaches_Create_OptArgs = []string{
	"SourceData", "Version",
}

func (this *PivotCaches) Create(sourceType int32, optArgs ...interface{}) *PivotCache {
	optArgs = ole.ProcessOptArgs(PivotCaches_Create_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000768, []interface{}{sourceType}, optArgs...)
	return NewPivotCache(retVal.IDispatch(), false, true)
}

