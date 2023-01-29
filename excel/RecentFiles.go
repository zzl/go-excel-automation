package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024406-0000-0000-C000-000000000046
var IID_RecentFiles = syscall.GUID{0x00024406, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type RecentFiles struct {
	ole.OleClient
}

func NewRecentFiles(pDisp *win32.IDispatch, addRef bool, scoped bool) *RecentFiles {
	if pDisp == nil {
		return nil
	}
	p := &RecentFiles{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RecentFilesFromVar(v ole.Variant) *RecentFiles {
	return NewRecentFiles(v.IDispatch(), false, false)
}

func (this *RecentFiles) IID() *syscall.GUID {
	return &IID_RecentFiles
}

func (this *RecentFiles) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *RecentFiles) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *RecentFiles) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *RecentFiles) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *RecentFiles) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *RecentFiles) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *RecentFiles) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *RecentFiles) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *RecentFiles) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *RecentFiles) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *RecentFiles) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *RecentFiles) Maximum() int32 {
	retVal, _ := this.PropGet(0x0000058f, nil)
	return retVal.LValVal()
}

func (this *RecentFiles) SetMaximum(rhs int32) {
	_ = this.PropPut(0x0000058f, []interface{}{rhs})
}

func (this *RecentFiles) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *RecentFiles) Item(index int32) *RecentFile {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewRecentFile(retVal.IDispatch(), false, true)
}

func (this *RecentFiles) Default_(index int32) *RecentFile {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewRecentFile(retVal.IDispatch(), false, true)
}

func (this *RecentFiles) Add(name string) *RecentFile {
	retVal, _ := this.Call(0x000000b5, []interface{}{name})
	return NewRecentFile(retVal.IDispatch(), false, true)
}

func (this *RecentFiles) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *RecentFiles) ForEach(action func(item *RecentFile) bool) {
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
		pItem := (*RecentFile)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}
