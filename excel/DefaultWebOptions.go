package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00024448-0000-0000-C000-000000000046
var IID_DefaultWebOptions = syscall.GUID{0x00024448, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DefaultWebOptions struct {
	ole.OleClient
}

func NewDefaultWebOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *DefaultWebOptions {
	p := &DefaultWebOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DefaultWebOptionsFromVar(v ole.Variant) *DefaultWebOptions {
	return NewDefaultWebOptions(v.PdispValVal(), false, false)
}

func (this *DefaultWebOptions) IID() *syscall.GUID {
	return &IID_DefaultWebOptions
}

func (this *DefaultWebOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DefaultWebOptions) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *DefaultWebOptions) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DefaultWebOptions) RelyOnCSS() bool {
	retVal := this.PropGet(0x0000076b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetRelyOnCSS(rhs bool)  {
	retVal := this.PropPut(0x0000076b, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) SaveHiddenData() bool {
	retVal := this.PropGet(0x0000076c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetSaveHiddenData(rhs bool)  {
	retVal := this.PropPut(0x0000076c, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) LoadPictures() bool {
	retVal := this.PropGet(0x0000076d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetLoadPictures(rhs bool)  {
	retVal := this.PropPut(0x0000076d, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) OrganizeInFolder() bool {
	retVal := this.PropGet(0x0000076e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetOrganizeInFolder(rhs bool)  {
	retVal := this.PropPut(0x0000076e, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) UpdateLinksOnSave() bool {
	retVal := this.PropGet(0x0000076f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetUpdateLinksOnSave(rhs bool)  {
	retVal := this.PropPut(0x0000076f, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) UseLongFileNames() bool {
	retVal := this.PropGet(0x00000770, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetUseLongFileNames(rhs bool)  {
	retVal := this.PropPut(0x00000770, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) CheckIfOfficeIsHTMLEditor() bool {
	retVal := this.PropGet(0x00000771, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetCheckIfOfficeIsHTMLEditor(rhs bool)  {
	retVal := this.PropPut(0x00000771, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) DownloadComponents() bool {
	retVal := this.PropGet(0x00000772, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetDownloadComponents(rhs bool)  {
	retVal := this.PropPut(0x00000772, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) RelyOnVML() bool {
	retVal := this.PropGet(0x00000773, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetRelyOnVML(rhs bool)  {
	retVal := this.PropPut(0x00000773, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) AllowPNG() bool {
	retVal := this.PropGet(0x00000774, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetAllowPNG(rhs bool)  {
	retVal := this.PropPut(0x00000774, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) ScreenSize() int32 {
	retVal := this.PropGet(0x00000775, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetScreenSize(rhs int32)  {
	retVal := this.PropPut(0x00000775, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) PixelsPerInch() int32 {
	retVal := this.PropGet(0x00000776, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetPixelsPerInch(rhs int32)  {
	retVal := this.PropPut(0x00000776, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) LocationOfComponents() string {
	retVal := this.PropGet(0x00000777, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DefaultWebOptions) SetLocationOfComponents(rhs string)  {
	retVal := this.PropPut(0x00000777, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) Encoding() int32 {
	retVal := this.PropGet(0x0000071e, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetEncoding(rhs int32)  {
	retVal := this.PropPut(0x0000071e, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) AlwaysSaveInDefaultEncoding() bool {
	retVal := this.PropGet(0x00000778, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetAlwaysSaveInDefaultEncoding(rhs bool)  {
	retVal := this.PropPut(0x00000778, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) Fonts() *ole.DispatchClass {
	retVal := this.PropGet(0x00000779, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DefaultWebOptions) FolderSuffix() string {
	retVal := this.PropGet(0x0000077a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DefaultWebOptions) TargetBrowser() int32 {
	retVal := this.PropGet(0x00000883, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetTargetBrowser(rhs int32)  {
	retVal := this.PropPut(0x00000883, []interface{}{rhs})
	_= retVal
}

func (this *DefaultWebOptions) SaveNewWebPagesAsWebArchives() bool {
	retVal := this.PropGet(0x00000884, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetSaveNewWebPagesAsWebArchives(rhs bool)  {
	retVal := this.PropPut(0x00000884, []interface{}{rhs})
	_= retVal
}

