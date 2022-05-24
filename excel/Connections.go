package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024486-0000-0000-C000-000000000046
var IID_Connections = syscall.GUID{0x00024486, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Connections struct {
	ole.OleClient
}

func NewConnections(pDisp *win32.IDispatch, addRef bool, scoped bool) *Connections {
	 if pDisp == nil {
		return nil;
	}
	p := &Connections{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ConnectionsFromVar(v ole.Variant) *Connections {
	return NewConnections(v.IDispatch(), false, false)
}

func (this *Connections) IID() *syscall.GUID {
	return &IID_Connections
}

func (this *Connections) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Connections) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Connections) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Connections) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Connections) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Connections) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Connections) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Connections) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Connections) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Connections) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Connections) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Connections) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Connections) AddFromFile(filename string) *WorkbookConnection {
	retVal, _ := this.Call(0x00000a8c, []interface{}{filename})
	return NewWorkbookConnection(retVal.IDispatch(), false, true)
}

var Connections_Add_OptArgs= []string{
	"lCmdtype", 
}

func (this *Connections) Add(name string, description string, connectionString interface{}, commandText interface{}, optArgs ...interface{}) *WorkbookConnection {
	optArgs = ole.ProcessOptArgs(Connections_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{name, description, connectionString, commandText}, optArgs...)
	return NewWorkbookConnection(retVal.IDispatch(), false, true)
}

func (this *Connections) Item(index interface{}) *WorkbookConnection {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return NewWorkbookConnection(retVal.IDispatch(), false, true)
}

func (this *Connections) Default_(index interface{}) *WorkbookConnection {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewWorkbookConnection(retVal.IDispatch(), false, true)
}

func (this *Connections) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Connections) ForEach(action func(item *WorkbookConnection) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*WorkbookConnection)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

