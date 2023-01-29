package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244B6-0000-0000-C000-000000000046
var IID_SparklineGroups = syscall.GUID{0x000244B6, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SparklineGroups struct {
	ole.OleClient
}

func NewSparklineGroups(pDisp *win32.IDispatch, addRef bool, scoped bool) *SparklineGroups {
	if pDisp == nil {
		return nil
	}
	p := &SparklineGroups{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SparklineGroupsFromVar(v ole.Variant) *SparklineGroups {
	return NewSparklineGroups(v.IDispatch(), false, false)
}

func (this *SparklineGroups) IID() *syscall.GUID {
	return &IID_SparklineGroups
}

func (this *SparklineGroups) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SparklineGroups) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *SparklineGroups) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SparklineGroups) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SparklineGroups) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *SparklineGroups) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *SparklineGroups) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *SparklineGroups) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *SparklineGroups) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SparklineGroups) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SparklineGroups) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SparklineGroups) Add(type_ int32, sourceData string) *SparklineGroup {
	retVal, _ := this.Call(0x000000b5, []interface{}{type_, sourceData})
	return NewSparklineGroup(retVal.IDispatch(), false, true)
}

func (this *SparklineGroups) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *SparklineGroups) Item(index interface{}) *SparklineGroup {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewSparklineGroup(retVal.IDispatch(), false, true)
}

func (this *SparklineGroups) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SparklineGroups) ForEach(action func(item *SparklineGroup) bool) {
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
		pItem := (*SparklineGroup)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *SparklineGroups) Default_(index interface{}) *SparklineGroup {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewSparklineGroup(retVal.IDispatch(), false, true)
}

func (this *SparklineGroups) Clear() {
	retVal, _ := this.Call(0x0000006f, nil)
	_ = retVal
}

func (this *SparklineGroups) ClearGroups() {
	retVal, _ := this.Call(0x00000b83, nil)
	_ = retVal
}

func (this *SparklineGroups) Group(location *Range) {
	retVal, _ := this.Call(0x0000002e, []interface{}{location})
	_ = retVal
}

func (this *SparklineGroups) Ungroup() {
	retVal, _ := this.Call(0x000000f4, nil)
	_ = retVal
}

