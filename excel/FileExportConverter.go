package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244B3-0000-0000-C000-000000000046
var IID_FileExportConverter = syscall.GUID{0x000244B3, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FileExportConverter struct {
	ole.OleClient
}

func NewFileExportConverter(pDisp *win32.IDispatch, addRef bool, scoped bool) *FileExportConverter {
	 if pDisp == nil {
		return nil;
	}
	p := &FileExportConverter{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FileExportConverterFromVar(v ole.Variant) *FileExportConverter {
	return NewFileExportConverter(v.IDispatch(), false, false)
}

func (this *FileExportConverter) IID() *syscall.GUID {
	return &IID_FileExportConverter
}

func (this *FileExportConverter) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FileExportConverter) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *FileExportConverter) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *FileExportConverter) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *FileExportConverter) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *FileExportConverter) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *FileExportConverter) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *FileExportConverter) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *FileExportConverter) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *FileExportConverter) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *FileExportConverter) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *FileExportConverter) Extensions() string {
	retVal, _ := this.PropGet(0x00000ad1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FileExportConverter) Description() string {
	retVal, _ := this.PropGet(0x000000da, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FileExportConverter) FileFormat() int32 {
	retVal, _ := this.PropGet(0x00000120, nil)
	return retVal.LValVal()
}

