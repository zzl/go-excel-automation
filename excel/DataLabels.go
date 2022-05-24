package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208B3-0000-0000-C000-000000000046
var IID_DataLabels = syscall.GUID{0x000208B3, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DataLabels struct {
	ole.OleClient
}

func NewDataLabels(pDisp *win32.IDispatch, addRef bool, scoped bool) *DataLabels {
	 if pDisp == nil {
		return nil;
	}
	p := &DataLabels{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DataLabelsFromVar(v ole.Variant) *DataLabels {
	return NewDataLabels(v.IDispatch(), false, false)
}

func (this *DataLabels) IID() *syscall.GUID {
	return &IID_DataLabels
}

func (this *DataLabels) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DataLabels) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *DataLabels) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DataLabels) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DataLabels) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *DataLabels) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *DataLabels) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *DataLabels) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *DataLabels) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DataLabels) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DataLabels) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DataLabels) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabels) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabels) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *DataLabels) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabels) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *DataLabels) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *DataLabels) Dummy9_()  {
	retVal, _ := this.Call(0x00010009, nil)
	_= retVal
}

func (this *DataLabels) Dummy10_()  {
	retVal, _ := this.Call(0x0001000a, nil)
	_= retVal
}

func (this *DataLabels) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *DataLabels) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabels) SetHorizontalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *DataLabels) Dummy13_()  {
	retVal, _ := this.Call(0x0001000d, nil)
	_= retVal
}

func (this *DataLabels) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabels) SetOrientation(rhs interface{})  {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *DataLabels) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShadow(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *DataLabels) Dummy16_()  {
	retVal, _ := this.Call(0x00010010, nil)
	_= retVal
}

func (this *DataLabels) Dummy17_()  {
	retVal, _ := this.Call(0x00010011, nil)
	_= retVal
}

func (this *DataLabels) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabels) SetVerticalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *DataLabels) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *DataLabels) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}

func (this *DataLabels) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x000005f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabels) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x000005f5, []interface{}{rhs})
}

func (this *DataLabels) AutoText() bool {
	retVal, _ := this.PropGet(0x00000087, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetAutoText(rhs bool)  {
	_ = this.PropPut(0x00000087, []interface{}{rhs})
}

func (this *DataLabels) NumberFormat() string {
	retVal, _ := this.PropGet(0x000000c1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabels) SetNumberFormat(rhs string)  {
	_ = this.PropPut(0x000000c1, []interface{}{rhs})
}

func (this *DataLabels) NumberFormatLinked() bool {
	retVal, _ := this.PropGet(0x000000c2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetNumberFormatLinked(rhs bool)  {
	_ = this.PropPut(0x000000c2, []interface{}{rhs})
}

func (this *DataLabels) NumberFormatLocal() ole.Variant {
	retVal, _ := this.PropGet(0x00000449, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabels) SetNumberFormatLocal(rhs interface{})  {
	_ = this.PropPut(0x00000449, []interface{}{rhs})
}

func (this *DataLabels) ShowLegendKey() bool {
	retVal, _ := this.PropGet(0x000000ab, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowLegendKey(rhs bool)  {
	_ = this.PropPut(0x000000ab, []interface{}{rhs})
}

func (this *DataLabels) Type() ole.Variant {
	retVal, _ := this.PropGet(0x0000006c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabels) SetType(rhs interface{})  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *DataLabels) Position() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *DataLabels) SetPosition(rhs int32)  {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *DataLabels) ShowSeriesName() bool {
	retVal, _ := this.PropGet(0x000007e6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowSeriesName(rhs bool)  {
	_ = this.PropPut(0x000007e6, []interface{}{rhs})
}

func (this *DataLabels) ShowCategoryName() bool {
	retVal, _ := this.PropGet(0x000007e7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowCategoryName(rhs bool)  {
	_ = this.PropPut(0x000007e7, []interface{}{rhs})
}

func (this *DataLabels) ShowValue() bool {
	retVal, _ := this.PropGet(0x000007e8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowValue(rhs bool)  {
	_ = this.PropPut(0x000007e8, []interface{}{rhs})
}

func (this *DataLabels) ShowPercentage() bool {
	retVal, _ := this.PropGet(0x000007e9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowPercentage(rhs bool)  {
	_ = this.PropPut(0x000007e9, []interface{}{rhs})
}

func (this *DataLabels) ShowBubbleSize() bool {
	retVal, _ := this.PropGet(0x000007ea, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowBubbleSize(rhs bool)  {
	_ = this.PropPut(0x000007ea, []interface{}{rhs})
}

func (this *DataLabels) Separator() ole.Variant {
	retVal, _ := this.PropGet(0x000007eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabels) SetSeparator(rhs interface{})  {
	_ = this.PropPut(0x000007eb, []interface{}{rhs})
}

func (this *DataLabels) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *DataLabels) Item(index interface{}) *DataLabel {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return NewDataLabel(retVal.IDispatch(), false, true)
}

func (this *DataLabels) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *DataLabels) ForEach(action func(item *DataLabel) bool) {
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
		pItem := (*DataLabel)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *DataLabels) Default_(index interface{}) *DataLabel {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewDataLabel(retVal.IDispatch(), false, true)
}

func (this *DataLabels) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

