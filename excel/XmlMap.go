package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002447B-0000-0000-C000-000000000046
var IID_XmlMap = syscall.GUID{0x0002447B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type XmlMap struct {
	ole.OleClient
}

func NewXmlMap(pDisp *win32.IDispatch, addRef bool, scoped bool) *XmlMap {
	 if pDisp == nil {
		return nil;
	}
	p := &XmlMap{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XmlMapFromVar(v ole.Variant) *XmlMap {
	return NewXmlMap(v.IDispatch(), false, false)
}

func (this *XmlMap) IID() *syscall.GUID {
	return &IID_XmlMap
}

func (this *XmlMap) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XmlMap) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *XmlMap) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *XmlMap) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *XmlMap) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *XmlMap) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *XmlMap) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *XmlMap) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *XmlMap) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *XmlMap) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *XmlMap) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *XmlMap) Default_() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XmlMap) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XmlMap) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *XmlMap) IsExportable() bool {
	retVal, _ := this.PropGet(0x0000091e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XmlMap) ShowImportExportValidationErrors() bool {
	retVal, _ := this.PropGet(0x0000091f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XmlMap) SetShowImportExportValidationErrors(rhs bool)  {
	_ = this.PropPut(0x0000091f, []interface{}{rhs})
}

func (this *XmlMap) SaveDataSourceDefinition() bool {
	retVal, _ := this.PropGet(0x00000920, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XmlMap) SetSaveDataSourceDefinition(rhs bool)  {
	_ = this.PropPut(0x00000920, []interface{}{rhs})
}

func (this *XmlMap) AdjustColumnWidth() bool {
	retVal, _ := this.PropGet(0x0000074c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XmlMap) SetAdjustColumnWidth(rhs bool)  {
	_ = this.PropPut(0x0000074c, []interface{}{rhs})
}

func (this *XmlMap) PreserveColumnFilter() bool {
	retVal, _ := this.PropGet(0x00000921, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XmlMap) SetPreserveColumnFilter(rhs bool)  {
	_ = this.PropPut(0x00000921, []interface{}{rhs})
}

func (this *XmlMap) PreserveNumberFormatting() bool {
	retVal, _ := this.PropGet(0x00000922, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XmlMap) SetPreserveNumberFormatting(rhs bool)  {
	_ = this.PropPut(0x00000922, []interface{}{rhs})
}

func (this *XmlMap) AppendOnImport() bool {
	retVal, _ := this.PropGet(0x00000923, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XmlMap) SetAppendOnImport(rhs bool)  {
	_ = this.PropPut(0x00000923, []interface{}{rhs})
}

func (this *XmlMap) RootElementName() string {
	retVal, _ := this.PropGet(0x00000924, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XmlMap) RootElementNamespace() *XmlNamespace {
	retVal, _ := this.PropGet(0x00000925, nil)
	return NewXmlNamespace(retVal.IDispatch(), false, true)
}

func (this *XmlMap) Schemas() *XmlSchemas {
	retVal, _ := this.PropGet(0x00000926, nil)
	return NewXmlSchemas(retVal.IDispatch(), false, true)
}

func (this *XmlMap) DataBinding() *XmlDataBinding {
	retVal, _ := this.PropGet(0x00000927, nil)
	return NewXmlDataBinding(retVal.IDispatch(), false, true)
}

func (this *XmlMap) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

var XmlMap_Import_OptArgs= []string{
	"Overwrite", 
}

func (this *XmlMap) Import(url string, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(XmlMap_Import_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000395, []interface{}{url}, optArgs...)
	return retVal.LValVal()
}

var XmlMap_ImportXml_OptArgs= []string{
	"Overwrite", 
}

func (this *XmlMap) ImportXml(xmlData string, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(XmlMap_ImportXml_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000928, []interface{}{xmlData}, optArgs...)
	return retVal.LValVal()
}

var XmlMap_Export_OptArgs= []string{
	"Overwrite", 
}

func (this *XmlMap) Export(url string, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(XmlMap_Export_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000586, []interface{}{url}, optArgs...)
	return retVal.LValVal()
}

func (this *XmlMap) ExportXml(data *win32.BSTR) int32 {
	retVal, _ := this.Call(0x0000092a, []interface{}{data})
	return retVal.LValVal()
}

func (this *XmlMap) WorkbookConnection() *WorkbookConnection {
	retVal, _ := this.PropGet(0x000009f0, nil)
	return NewWorkbookConnection(retVal.IDispatch(), false, true)
}

