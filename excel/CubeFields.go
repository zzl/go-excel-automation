package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002444D-0000-0000-C000-000000000046
var IID_CubeFields = syscall.GUID{0x0002444D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CubeFields struct {
	ole.OleClient
}

func NewCubeFields(pDisp *win32.IDispatch, addRef bool, scoped bool) *CubeFields {
	 if pDisp == nil {
		return nil;
	}
	p := &CubeFields{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CubeFieldsFromVar(v ole.Variant) *CubeFields {
	return NewCubeFields(v.IDispatch(), false, false)
}

func (this *CubeFields) IID() *syscall.GUID {
	return &IID_CubeFields
}

func (this *CubeFields) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CubeFields) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CubeFields) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *CubeFields) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CubeFields) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *CubeFields) Item(index interface{}) *CubeField {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewCubeField(retVal.IDispatch(), false, true)
}

func (this *CubeFields) Default_(index interface{}) *CubeField {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewCubeField(retVal.IDispatch(), false, true)
}

func (this *CubeFields) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CubeFields) ForEach(action func(item *CubeField) bool) {
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
		pItem := (*CubeField)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *CubeFields) AddSet(name string, caption string) *CubeField {
	retVal, _ := this.Call(0x0000088a, []interface{}{name, caption})
	return NewCubeField(retVal.IDispatch(), false, true)
}

