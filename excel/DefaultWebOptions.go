package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
)

// 00024448-0000-0000-C000-000000000046
var IID_DefaultWebOptions = syscall.GUID{0x00024448, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DefaultWebOptions struct {
	ole.OleClient
}

func NewDefaultWebOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *DefaultWebOptions {
	if pDisp == nil {
		return nil
	}
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
	return NewDefaultWebOptions(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DefaultWebOptions) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DefaultWebOptions) RelyOnCSS() bool {
	retVal, _ := this.PropGet(0x0000076b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetRelyOnCSS(rhs bool) {
	_ = this.PropPut(0x0000076b, []interface{}{rhs})
}

func (this *DefaultWebOptions) SaveHiddenData() bool {
	retVal, _ := this.PropGet(0x0000076c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetSaveHiddenData(rhs bool) {
	_ = this.PropPut(0x0000076c, []interface{}{rhs})
}

func (this *DefaultWebOptions) LoadPictures() bool {
	retVal, _ := this.PropGet(0x0000076d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetLoadPictures(rhs bool) {
	_ = this.PropPut(0x0000076d, []interface{}{rhs})
}

func (this *DefaultWebOptions) OrganizeInFolder() bool {
	retVal, _ := this.PropGet(0x0000076e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetOrganizeInFolder(rhs bool) {
	_ = this.PropPut(0x0000076e, []interface{}{rhs})
}

func (this *DefaultWebOptions) UpdateLinksOnSave() bool {
	retVal, _ := this.PropGet(0x0000076f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetUpdateLinksOnSave(rhs bool) {
	_ = this.PropPut(0x0000076f, []interface{}{rhs})
}

func (this *DefaultWebOptions) UseLongFileNames() bool {
	retVal, _ := this.PropGet(0x00000770, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetUseLongFileNames(rhs bool) {
	_ = this.PropPut(0x00000770, []interface{}{rhs})
}

func (this *DefaultWebOptions) CheckIfOfficeIsHTMLEditor() bool {
	retVal, _ := this.PropGet(0x00000771, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetCheckIfOfficeIsHTMLEditor(rhs bool) {
	_ = this.PropPut(0x00000771, []interface{}{rhs})
}

func (this *DefaultWebOptions) DownloadComponents() bool {
	retVal, _ := this.PropGet(0x00000772, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetDownloadComponents(rhs bool) {
	_ = this.PropPut(0x00000772, []interface{}{rhs})
}

func (this *DefaultWebOptions) RelyOnVML() bool {
	retVal, _ := this.PropGet(0x00000773, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetRelyOnVML(rhs bool) {
	_ = this.PropPut(0x00000773, []interface{}{rhs})
}

func (this *DefaultWebOptions) AllowPNG() bool {
	retVal, _ := this.PropGet(0x00000774, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetAllowPNG(rhs bool) {
	_ = this.PropPut(0x00000774, []interface{}{rhs})
}

func (this *DefaultWebOptions) ScreenSize() int32 {
	retVal, _ := this.PropGet(0x00000775, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetScreenSize(rhs int32) {
	_ = this.PropPut(0x00000775, []interface{}{rhs})
}

func (this *DefaultWebOptions) PixelsPerInch() int32 {
	retVal, _ := this.PropGet(0x00000776, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetPixelsPerInch(rhs int32) {
	_ = this.PropPut(0x00000776, []interface{}{rhs})
}

func (this *DefaultWebOptions) LocationOfComponents() string {
	retVal, _ := this.PropGet(0x00000777, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DefaultWebOptions) SetLocationOfComponents(rhs string) {
	_ = this.PropPut(0x00000777, []interface{}{rhs})
}

func (this *DefaultWebOptions) Encoding() int32 {
	retVal, _ := this.PropGet(0x0000071e, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetEncoding(rhs int32) {
	_ = this.PropPut(0x0000071e, []interface{}{rhs})
}

func (this *DefaultWebOptions) AlwaysSaveInDefaultEncoding() bool {
	retVal, _ := this.PropGet(0x00000778, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetAlwaysSaveInDefaultEncoding(rhs bool) {
	_ = this.PropPut(0x00000778, []interface{}{rhs})
}

func (this *DefaultWebOptions) Fonts() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000779, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DefaultWebOptions) FolderSuffix() string {
	retVal, _ := this.PropGet(0x0000077a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DefaultWebOptions) TargetBrowser() int32 {
	retVal, _ := this.PropGet(0x00000883, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetTargetBrowser(rhs int32) {
	_ = this.PropPut(0x00000883, []interface{}{rhs})
}

func (this *DefaultWebOptions) SaveNewWebPagesAsWebArchives() bool {
	retVal, _ := this.PropGet(0x00000884, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetSaveNewWebPagesAsWebArchives(rhs bool) {
	_ = this.PropPut(0x00000884, []interface{}{rhs})
}
