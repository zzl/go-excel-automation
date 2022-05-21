package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00024449-0000-0000-C000-000000000046
var IID_WebOptions = syscall.GUID{0x00024449, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type WebOptions struct {
	ole.OleClient
}

func NewWebOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *WebOptions {
	p := &WebOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WebOptionsFromVar(v ole.Variant) *WebOptions {
	return NewWebOptions(v.PdispValVal(), false, false)
}

func (this *WebOptions) IID() *syscall.GUID {
	return &IID_WebOptions
}

func (this *WebOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *WebOptions) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *WebOptions) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *WebOptions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *WebOptions) RelyOnCSS() bool {
	retVal := this.PropGet(0x0000076b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetRelyOnCSS(rhs bool)  {
	retVal := this.PropPut(0x0000076b, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) OrganizeInFolder() bool {
	retVal := this.PropGet(0x0000076e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetOrganizeInFolder(rhs bool)  {
	retVal := this.PropPut(0x0000076e, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) UseLongFileNames() bool {
	retVal := this.PropGet(0x00000770, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetUseLongFileNames(rhs bool)  {
	retVal := this.PropPut(0x00000770, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) DownloadComponents() bool {
	retVal := this.PropGet(0x00000772, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetDownloadComponents(rhs bool)  {
	retVal := this.PropPut(0x00000772, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) RelyOnVML() bool {
	retVal := this.PropGet(0x00000773, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetRelyOnVML(rhs bool)  {
	retVal := this.PropPut(0x00000773, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) AllowPNG() bool {
	retVal := this.PropGet(0x00000774, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetAllowPNG(rhs bool)  {
	retVal := this.PropPut(0x00000774, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) ScreenSize() int32 {
	retVal := this.PropGet(0x00000775, nil)
	return retVal.LValVal()
}

func (this *WebOptions) SetScreenSize(rhs int32)  {
	retVal := this.PropPut(0x00000775, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) PixelsPerInch() int32 {
	retVal := this.PropGet(0x00000776, nil)
	return retVal.LValVal()
}

func (this *WebOptions) SetPixelsPerInch(rhs int32)  {
	retVal := this.PropPut(0x00000776, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) LocationOfComponents() string {
	retVal := this.PropGet(0x00000777, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WebOptions) SetLocationOfComponents(rhs string)  {
	retVal := this.PropPut(0x00000777, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) Encoding() int32 {
	retVal := this.PropGet(0x0000071e, nil)
	return retVal.LValVal()
}

func (this *WebOptions) SetEncoding(rhs int32)  {
	retVal := this.PropPut(0x0000071e, []interface{}{rhs})
	_= retVal
}

func (this *WebOptions) FolderSuffix() string {
	retVal := this.PropGet(0x0000077a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WebOptions) UseDefaultFolderSuffix()  {
	retVal := this.Call(0x0000077b, nil)
	_= retVal
}

func (this *WebOptions) TargetBrowser() int32 {
	retVal := this.PropGet(0x00000883, nil)
	return retVal.LValVal()
}

func (this *WebOptions) SetTargetBrowser(rhs int32)  {
	retVal := this.PropPut(0x00000883, []interface{}{rhs})
	_= retVal
}

