package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024492-0000-0000-C000-000000000046
var IID_ConditionValue = syscall.GUID{0x00024492, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ConditionValue struct {
	ole.OleClient
}

func NewConditionValue(pDisp *win32.IDispatch, addRef bool, scoped bool) *ConditionValue {
	 if pDisp == nil {
		return nil;
	}
	p := &ConditionValue{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ConditionValueFromVar(v ole.Variant) *ConditionValue {
	return NewConditionValue(v.IDispatch(), false, false)
}

func (this *ConditionValue) IID() *syscall.GUID {
	return &IID_ConditionValue
}

func (this *ConditionValue) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ConditionValue) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ConditionValue) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ConditionValue) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ConditionValue) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ConditionValue) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ConditionValue) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ConditionValue) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ConditionValue) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ConditionValue) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ConditionValue) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var ConditionValue_Modify_OptArgs= []string{
	"newvalue", 
}

func (this *ConditionValue) Modify(newtype int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ConditionValue_Modify_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000062d, []interface{}{newtype}, optArgs...)
	_= retVal
}

func (this *ConditionValue) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ConditionValue) Value() ole.Variant {
	retVal, _ := this.PropGet(0x00000006, nil)
	com.AddToScope(retVal)
	return *retVal
}

