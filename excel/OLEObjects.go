package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208A3-0000-0000-C000-000000000046
var IID_OLEObjects = syscall.GUID{0x000208A3, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEObjects struct {
	ole.OleClient
}

func NewOLEObjects(pDisp *win32.IDispatch, addRef bool, scoped bool) *OLEObjects {
	 if pDisp == nil {
		return nil;
	}
	p := &OLEObjects{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OLEObjectsFromVar(v ole.Variant) *OLEObjects {
	return NewOLEObjects(v.IDispatch(), false, false)
}

func (this *OLEObjects) IID() *syscall.GUID {
	return &IID_OLEObjects
}

func (this *OLEObjects) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OLEObjects) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *OLEObjects) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *OLEObjects) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *OLEObjects) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *OLEObjects) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *OLEObjects) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *OLEObjects) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *OLEObjects) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OLEObjects) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *OLEObjects) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OLEObjects) Dummy3_()  {
	retVal, _ := this.Call(0x00010003, nil)
	_= retVal
}

func (this *OLEObjects) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEObjects) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var OLEObjects_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *OLEObjects) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(OLEObjects_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEObjects) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEObjects) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEObjects) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OLEObjects) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObjects) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *OLEObjects) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *OLEObjects) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *OLEObjects) Dummy12_()  {
	retVal, _ := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *OLEObjects) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *OLEObjects) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *OLEObjects) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObjects) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *OLEObjects) Dummy15_()  {
	retVal, _ := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *OLEObjects) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEObjects) SetOnAction(rhs string)  {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *OLEObjects) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEObjects) SetPlacement(rhs interface{})  {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *OLEObjects) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObjects) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var OLEObjects_Select_OptArgs= []string{
	"Replace", 
}

func (this *OLEObjects) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(OLEObjects_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEObjects) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEObjects) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *OLEObjects) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *OLEObjects) Dummy22_()  {
	retVal, _ := this.Call(0x00010016, nil)
	_= retVal
}

func (this *OLEObjects) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObjects) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *OLEObjects) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *OLEObjects) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *OLEObjects) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *OLEObjects) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *OLEObjects) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *OLEObjects) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *OLEObjects) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObjects) SetShadow(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *OLEObjects) Dummy30_()  {
	retVal, _ := this.Call(0x0001001e, nil)
	_= retVal
}

func (this *OLEObjects) AutoLoad() bool {
	retVal, _ := this.PropGet(0x000004a2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObjects) SetAutoLoad(rhs bool)  {
	_ = this.PropPut(0x000004a2, []interface{}{rhs})
}

func (this *OLEObjects) Dummy32_()  {
	retVal, _ := this.Call(0x00010020, nil)
	_= retVal
}

func (this *OLEObjects) Dummy33_()  {
	retVal, _ := this.Call(0x00010021, nil)
	_= retVal
}

func (this *OLEObjects) Dummy34_()  {
	retVal, _ := this.Call(0x00010022, nil)
	_= retVal
}

func (this *OLEObjects) SourceName() string {
	retVal, _ := this.PropGet(0x000002d1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEObjects) SetSourceName(rhs string)  {
	_ = this.PropPut(0x000002d1, []interface{}{rhs})
}

func (this *OLEObjects) Dummy36_()  {
	retVal, _ := this.Call(0x00010024, nil)
	_= retVal
}

func (this *OLEObjects) Dummy37_()  {
	retVal, _ := this.Call(0x00010025, nil)
	_= retVal
}

func (this *OLEObjects) Dummy38_()  {
	retVal, _ := this.Call(0x00010026, nil)
	_= retVal
}

func (this *OLEObjects) Dummy39_()  {
	retVal, _ := this.Call(0x00010027, nil)
	_= retVal
}

func (this *OLEObjects) Dummy40_()  {
	retVal, _ := this.Call(0x00010028, nil)
	_= retVal
}

func (this *OLEObjects) Dummy41_()  {
	retVal, _ := this.Call(0x00010029, nil)
	_= retVal
}

var OLEObjects_Add_OptArgs= []string{
	"ClassType", "Filename", "Link", "DisplayAsIcon", 
	"IconFileName", "IconIndex", "IconLabel", "Left", 
	"Top", "Width", "Height", 
}

func (this *OLEObjects) Add(optArgs ...interface{}) *OLEObject {
	optArgs = ole.ProcessOptArgs(OLEObjects_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, nil, optArgs...)
	return NewOLEObject(retVal.IDispatch(), false, true)
}

func (this *OLEObjects) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *OLEObjects) Group() *GroupObject {
	retVal, _ := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.IDispatch(), false, true)
}

func (this *OLEObjects) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OLEObjects) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OLEObjects) ForEach(action func(item *ole.DispatchClass) bool) {
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
		pItem := (*ole.DispatchClass)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

