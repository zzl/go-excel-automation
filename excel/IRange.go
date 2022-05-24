package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020846-0001-0000-C000-000000000046
var IID_IRange = syscall.GUID{0x00020846, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IRange struct {
	win32.IDispatch
}

func NewIRange(pUnk *win32.IUnknown, addRef bool, scoped bool) *IRange {
	 if pUnk == nil {
		return nil;
	}
	p := (*IRange)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IRange) IID() *syscall.GUID {
	return &IID_IRange
}

func (this *IRange) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) Activate(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetAddIndent(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetAddIndent(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetAddress(rowAbsolute interface{}, columnAbsolute interface{}, referenceStyle int32, external interface{}, relativeTo interface{}, lcid int32, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowAbsolute)), (uintptr)(unsafe.Pointer(&columnAbsolute)), uintptr(referenceStyle), (uintptr)(unsafe.Pointer(&external)), (uintptr)(unsafe.Pointer(&relativeTo)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetAddressLocal(rowAbsolute interface{}, columnAbsolute interface{}, referenceStyle int32, external interface{}, relativeTo interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowAbsolute)), (uintptr)(unsafe.Pointer(&columnAbsolute)), uintptr(referenceStyle), (uintptr)(unsafe.Pointer(&external)), (uintptr)(unsafe.Pointer(&relativeTo)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) AdvancedFilter(action int32, criteriaRange interface{}, copyToRange interface{}, unique interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(action), (uintptr)(unsafe.Pointer(&criteriaRange)), (uintptr)(unsafe.Pointer(&copyToRange)), (uintptr)(unsafe.Pointer(&unique)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ApplyNames(names interface{}, ignoreRelativeAbsolute interface{}, useRowColumnNames interface{}, omitColumn interface{}, omitRow interface{}, order int32, appendLast interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&names)), (uintptr)(unsafe.Pointer(&ignoreRelativeAbsolute)), (uintptr)(unsafe.Pointer(&useRowColumnNames)), (uintptr)(unsafe.Pointer(&omitColumn)), (uintptr)(unsafe.Pointer(&omitRow)), uintptr(order), (uintptr)(unsafe.Pointer(&appendLast)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ApplyOutlineStyles(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetAreas(rhs **Areas) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) AutoComplete(string string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(string)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) AutoFill(destination *Range, type_ int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(destination)), uintptr(type_), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) AutoFilter(field interface{}, criteria1 interface{}, operator int32, criteria2 interface{}, visibleDropDown interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&field)), (uintptr)(unsafe.Pointer(&criteria1)), uintptr(operator), (uintptr)(unsafe.Pointer(&criteria2)), (uintptr)(unsafe.Pointer(&visibleDropDown)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) AutoFit(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) AutoFormat(format int32, number interface{}, font interface{}, alignment interface{}, border interface{}, pattern interface{}, width interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(format), (uintptr)(unsafe.Pointer(&number)), (uintptr)(unsafe.Pointer(&font)), (uintptr)(unsafe.Pointer(&alignment)), (uintptr)(unsafe.Pointer(&border)), (uintptr)(unsafe.Pointer(&pattern)), (uintptr)(unsafe.Pointer(&width)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) AutoOutline(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) BorderAround_(lineStyle interface{}, weight int32, colorIndex int32, color interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&lineStyle)), uintptr(weight), uintptr(colorIndex), (uintptr)(unsafe.Pointer(&color)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetBorders(rhs **Borders) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) Calculate(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetCells(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetCharacters(start interface{}, length interface{}, rhs **Characters) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&start)), (uintptr)(unsafe.Pointer(&length)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) CheckSpelling(customDictionary interface{}, ignoreUppercase interface{}, alwaysSuggest interface{}, spellLang interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&customDictionary)), (uintptr)(unsafe.Pointer(&ignoreUppercase)), (uintptr)(unsafe.Pointer(&alwaysSuggest)), (uintptr)(unsafe.Pointer(&spellLang)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Clear(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ClearContents(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ClearFormats(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ClearNotes(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ClearOutline(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetColumn(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ColumnDifferences(comparison interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&comparison)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetColumns(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetColumnWidth(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetColumnWidth(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) Consolidate(sources interface{}, function interface{}, topRow interface{}, leftColumn interface{}, createLinks interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&sources)), (uintptr)(unsafe.Pointer(&function)), (uintptr)(unsafe.Pointer(&topRow)), (uintptr)(unsafe.Pointer(&leftColumn)), (uintptr)(unsafe.Pointer(&createLinks)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Copy(destination interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[42]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&destination)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) CopyFromRecordset(data *win32.IUnknown, maxRows interface{}, maxColumns interface{}, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[43]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(data)), (uintptr)(unsafe.Pointer(&maxRows)), (uintptr)(unsafe.Pointer(&maxColumns)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) CopyPicture(appearance int32, format int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[44]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(appearance), uintptr(format), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[45]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) CreateNames(top interface{}, left interface{}, bottom interface{}, right interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[46]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&top)), (uintptr)(unsafe.Pointer(&left)), (uintptr)(unsafe.Pointer(&bottom)), (uintptr)(unsafe.Pointer(&right)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) CreatePublisher(edition interface{}, appearance int32, containsPICT interface{}, containsBIFF interface{}, containsRTF interface{}, containsVALU interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[47]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&edition)), uintptr(appearance), (uintptr)(unsafe.Pointer(&containsPICT)), (uintptr)(unsafe.Pointer(&containsBIFF)), (uintptr)(unsafe.Pointer(&containsRTF)), (uintptr)(unsafe.Pointer(&containsVALU)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetCurrentArray(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[48]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetCurrentRegion(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[49]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) Cut(destination interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[50]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&destination)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) DataSeries(rowcol interface{}, type_ int32, date int32, step interface{}, stop interface{}, trend interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[51]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowcol)), uintptr(type_), uintptr(date), (uintptr)(unsafe.Pointer(&step)), (uintptr)(unsafe.Pointer(&stop)), (uintptr)(unsafe.Pointer(&trend)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetDefault_(rowIndex interface{}, columnIndex interface{}, lcid int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[52]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowIndex)), (uintptr)(unsafe.Pointer(&columnIndex)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetDefault_(rowIndex interface{}, columnIndex interface{}, lcid int32, rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[53]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowIndex)), (uintptr)(unsafe.Pointer(&columnIndex)), uintptr(lcid), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) Delete(shift interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[54]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&shift)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetDependents(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[55]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) DialogBox(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[56]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetDirectDependents(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[57]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetDirectPrecedents(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[58]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) EditionOptions(type_ int32, option int32, name interface{}, reference interface{}, appearance int32, chartSize int32, format interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[59]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(type_), uintptr(option), (uintptr)(unsafe.Pointer(&name)), (uintptr)(unsafe.Pointer(&reference)), uintptr(appearance), uintptr(chartSize), (uintptr)(unsafe.Pointer(&format)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetEnd(direction int32, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[60]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(direction), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetEntireColumn(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[61]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetEntireRow(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[62]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) FillDown(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[63]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) FillLeft(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[64]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) FillRight(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[65]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) FillUp(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[66]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Find(what interface{}, after interface{}, lookIn interface{}, lookAt interface{}, searchOrder interface{}, searchDirection int32, matchCase interface{}, matchByte interface{}, searchFormat interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[67]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&what)), (uintptr)(unsafe.Pointer(&after)), (uintptr)(unsafe.Pointer(&lookIn)), (uintptr)(unsafe.Pointer(&lookAt)), (uintptr)(unsafe.Pointer(&searchOrder)), uintptr(searchDirection), (uintptr)(unsafe.Pointer(&matchCase)), (uintptr)(unsafe.Pointer(&matchByte)), (uintptr)(unsafe.Pointer(&searchFormat)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) FindNext(after interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[68]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&after)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) FindPrevious(after interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[69]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&after)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetFont(rhs **Font) com.Error {
	addr := (*this.LpVtbl)[70]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetFormula(lcid int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[71]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetFormula(lcid int32, rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[72]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetFormulaArray(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[73]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetFormulaArray(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[74]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetFormulaLabel(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[75]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetFormulaLabel(rhs int32) com.Error {
	addr := (*this.LpVtbl)[76]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IRange) GetFormulaHidden(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[77]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetFormulaHidden(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[78]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetFormulaLocal(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[79]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetFormulaLocal(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[80]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetFormulaR1C1(lcid int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[81]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetFormulaR1C1(lcid int32, rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[82]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetFormulaR1C1Local(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[83]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetFormulaR1C1Local(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[84]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) FunctionWizard(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[85]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GoalSeek(goal interface{}, changingCell *Range, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[86]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&goal)), uintptr(unsafe.Pointer(changingCell)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Group(start interface{}, end interface{}, by interface{}, periods interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[87]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&start)), (uintptr)(unsafe.Pointer(&end)), (uintptr)(unsafe.Pointer(&by)), (uintptr)(unsafe.Pointer(&periods)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetHasArray(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[88]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetHasFormula(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[89]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetHeight(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[90]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetHidden(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[91]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetHidden(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[92]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetHorizontalAlignment(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[93]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetHorizontalAlignment(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[94]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetIndentLevel(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[95]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetIndentLevel(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[96]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) InsertIndent(insertAmount int32) com.Error {
	addr := (*this.LpVtbl)[97]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(insertAmount))
	return com.Error(ret)
}

func (this *IRange) Insert(shift interface{}, copyOrigin interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[98]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&shift)), (uintptr)(unsafe.Pointer(&copyOrigin)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetInterior(rhs **Interior) com.Error {
	addr := (*this.LpVtbl)[99]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetItem(rowIndex interface{}, columnIndex interface{}, lcid int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[100]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowIndex)), (uintptr)(unsafe.Pointer(&columnIndex)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetItem(rowIndex interface{}, columnIndex interface{}, lcid int32, rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[101]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowIndex)), (uintptr)(unsafe.Pointer(&columnIndex)), uintptr(lcid), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) Justify(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[102]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetLeft(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[103]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetListHeaderRows(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[104]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ListNames(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[105]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetLocationInTable(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[106]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetLocked(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[107]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetLocked(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[108]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) Merge(across interface{}) com.Error {
	addr := (*this.LpVtbl)[109]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&across)))
	return com.Error(ret)
}

func (this *IRange) UnMerge() com.Error {
	addr := (*this.LpVtbl)[110]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IRange) GetMergeArea(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[111]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetMergeCells(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[112]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetMergeCells(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[113]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetName(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[114]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetName(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[115]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) NavigateArrow(towardPrecedent interface{}, arrowNumber interface{}, linkNumber interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[116]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&towardPrecedent)), (uintptr)(unsafe.Pointer(&arrowNumber)), (uintptr)(unsafe.Pointer(&linkNumber)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetNewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[117]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetNext(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[118]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) NoteText(text interface{}, start interface{}, length interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[119]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&text)), (uintptr)(unsafe.Pointer(&start)), (uintptr)(unsafe.Pointer(&length)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetNumberFormat(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[120]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetNumberFormat(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[121]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetNumberFormatLocal(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[122]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetNumberFormatLocal(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[123]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetOffset(rowOffset interface{}, columnOffset interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[124]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowOffset)), (uintptr)(unsafe.Pointer(&columnOffset)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetOrientation(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[125]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetOrientation(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[126]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetOutlineLevel(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[127]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetOutlineLevel(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[128]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetPageBreak(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[129]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetPageBreak(rhs int32) com.Error {
	addr := (*this.LpVtbl)[130]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IRange) Parse(parseLine interface{}, destination interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[131]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&parseLine)), (uintptr)(unsafe.Pointer(&destination)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) PasteSpecial_(paste int32, operation int32, skipBlanks interface{}, transpose interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[132]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(paste), uintptr(operation), (uintptr)(unsafe.Pointer(&skipBlanks)), (uintptr)(unsafe.Pointer(&transpose)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetPivotField(rhs **PivotField) com.Error {
	addr := (*this.LpVtbl)[133]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetPivotItem(rhs **PivotItem) com.Error {
	addr := (*this.LpVtbl)[134]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetPivotTable(rhs **PivotTable) com.Error {
	addr := (*this.LpVtbl)[135]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetPrecedents(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[136]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetPrefixCharacter(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[137]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetPrevious(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[138]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) PrintOut__(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[139]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) PrintPreview(enableChanges interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[140]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&enableChanges)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetQueryTable(rhs **QueryTable) com.Error {
	addr := (*this.LpVtbl)[141]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetRange(cell1 interface{}, cell2 interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[142]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&cell1)), (uintptr)(unsafe.Pointer(&cell2)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) RemoveSubtotal(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[143]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Replace(what interface{}, replacement interface{}, lookAt interface{}, searchOrder interface{}, matchCase interface{}, matchByte interface{}, searchFormat interface{}, replaceFormat interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[144]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&what)), (uintptr)(unsafe.Pointer(&replacement)), (uintptr)(unsafe.Pointer(&lookAt)), (uintptr)(unsafe.Pointer(&searchOrder)), (uintptr)(unsafe.Pointer(&matchCase)), (uintptr)(unsafe.Pointer(&matchByte)), (uintptr)(unsafe.Pointer(&searchFormat)), (uintptr)(unsafe.Pointer(&replaceFormat)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetResize(rowSize interface{}, columnSize interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[145]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowSize)), (uintptr)(unsafe.Pointer(&columnSize)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetRow(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[146]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) RowDifferences(comparison interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[147]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&comparison)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetRowHeight(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[148]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetRowHeight(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[149]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetRows(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[150]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) Run(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[151]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Select(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[152]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Show(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[153]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ShowDependents(remove interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[154]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&remove)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetShowDetail(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[155]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetShowDetail(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[156]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) ShowErrors(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[157]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ShowPrecedents(remove interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[158]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&remove)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetShrinkToFit(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[159]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetShrinkToFit(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[160]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) Sort(key1 interface{}, order1 int32, key2 interface{}, type_ interface{}, order2 int32, key3 interface{}, order3 int32, header int32, orderCustom interface{}, matchCase interface{}, orientation int32, sortMethod int32, dataOption1 int32, dataOption2 int32, dataOption3 int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[161]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&key1)), uintptr(order1), (uintptr)(unsafe.Pointer(&key2)), (uintptr)(unsafe.Pointer(&type_)), uintptr(order2), (uintptr)(unsafe.Pointer(&key3)), uintptr(order3), uintptr(header), (uintptr)(unsafe.Pointer(&orderCustom)), (uintptr)(unsafe.Pointer(&matchCase)), uintptr(orientation), uintptr(sortMethod), uintptr(dataOption1), uintptr(dataOption2), uintptr(dataOption3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SortSpecial(sortMethod int32, key1 interface{}, order1 int32, type_ interface{}, key2 interface{}, order2 int32, key3 interface{}, order3 int32, header int32, orderCustom interface{}, matchCase interface{}, orientation int32, dataOption1 int32, dataOption2 int32, dataOption3 int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[162]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(sortMethod), (uintptr)(unsafe.Pointer(&key1)), uintptr(order1), (uintptr)(unsafe.Pointer(&type_)), (uintptr)(unsafe.Pointer(&key2)), uintptr(order2), (uintptr)(unsafe.Pointer(&key3)), uintptr(order3), uintptr(header), (uintptr)(unsafe.Pointer(&orderCustom)), (uintptr)(unsafe.Pointer(&matchCase)), uintptr(orientation), uintptr(dataOption1), uintptr(dataOption2), uintptr(dataOption3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetSoundNote(rhs **SoundNote) com.Error {
	addr := (*this.LpVtbl)[163]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) SpecialCells(type_ int32, value interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[164]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(type_), (uintptr)(unsafe.Pointer(&value)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetStyle(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[165]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetStyle(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[166]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) SubscribeTo(edition string, format int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[167]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(edition)), uintptr(format), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Subtotal(groupBy int32, function int32, totalList interface{}, replace interface{}, pageBreaks interface{}, summaryBelowData int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[168]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(groupBy), uintptr(function), (uintptr)(unsafe.Pointer(&totalList)), (uintptr)(unsafe.Pointer(&replace)), (uintptr)(unsafe.Pointer(&pageBreaks)), uintptr(summaryBelowData), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetSummary(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[169]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Table(rowInput interface{}, columnInput interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[170]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowInput)), (uintptr)(unsafe.Pointer(&columnInput)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetText(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[171]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) TextToColumns(destination interface{}, dataType int32, textQualifier int32, consecutiveDelimiter interface{}, tab interface{}, semicolon interface{}, comma interface{}, space interface{}, other interface{}, otherChar interface{}, fieldInfo interface{}, decimalSeparator interface{}, thousandsSeparator interface{}, trailingMinusNumbers interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[172]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&destination)), uintptr(dataType), uintptr(textQualifier), (uintptr)(unsafe.Pointer(&consecutiveDelimiter)), (uintptr)(unsafe.Pointer(&tab)), (uintptr)(unsafe.Pointer(&semicolon)), (uintptr)(unsafe.Pointer(&comma)), (uintptr)(unsafe.Pointer(&space)), (uintptr)(unsafe.Pointer(&other)), (uintptr)(unsafe.Pointer(&otherChar)), (uintptr)(unsafe.Pointer(&fieldInfo)), (uintptr)(unsafe.Pointer(&decimalSeparator)), (uintptr)(unsafe.Pointer(&thousandsSeparator)), (uintptr)(unsafe.Pointer(&trailingMinusNumbers)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetTop(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[173]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) Ungroup(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[174]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetUseStandardHeight(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[175]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetUseStandardHeight(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[176]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetUseStandardWidth(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[177]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetUseStandardWidth(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[178]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetValidation(rhs **Validation) com.Error {
	addr := (*this.LpVtbl)[179]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetValue(rangeValueDataType interface{}, lcid int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[180]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rangeValueDataType)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetValue(rangeValueDataType interface{}, lcid int32, rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[181]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rangeValueDataType)), uintptr(lcid), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetValue2(lcid int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[182]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetValue2(lcid int32, rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[183]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetVerticalAlignment(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[184]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetVerticalAlignment(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[185]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) GetWidth(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[186]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetWorksheet(rhs **Worksheet) com.Error {
	addr := (*this.LpVtbl)[187]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetWrapText(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[188]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetWrapText(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[189]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IRange) AddComment(text interface{}, rhs **Comment) com.Error {
	addr := (*this.LpVtbl)[190]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&text)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetComment(rhs **Comment) com.Error {
	addr := (*this.LpVtbl)[191]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) ClearComments() com.Error {
	addr := (*this.LpVtbl)[192]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IRange) GetPhonetic(rhs **Phonetic) com.Error {
	addr := (*this.LpVtbl)[193]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetFormatConditions(rhs **FormatConditions) com.Error {
	addr := (*this.LpVtbl)[194]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetReadingOrder(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[195]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetReadingOrder(rhs int32) com.Error {
	addr := (*this.LpVtbl)[196]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IRange) GetHyperlinks(rhs **Hyperlinks) com.Error {
	addr := (*this.LpVtbl)[197]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetPhonetics(rhs **Phonetics) com.Error {
	addr := (*this.LpVtbl)[198]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) SetPhonetic() com.Error {
	addr := (*this.LpVtbl)[199]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IRange) GetID(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[200]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) SetID(rhs string) com.Error {
	addr := (*this.LpVtbl)[201]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) PrintOut_(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, prToFileName interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[202]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), (uintptr)(unsafe.Pointer(&prToFileName)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetPivotCell(rhs **PivotCell) com.Error {
	addr := (*this.LpVtbl)[203]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) Dirty() com.Error {
	addr := (*this.LpVtbl)[204]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IRange) GetErrors(rhs **Errors) com.Error {
	addr := (*this.LpVtbl)[205]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetSmartTags(rhs **SmartTags) com.Error {
	addr := (*this.LpVtbl)[206]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) Speak(speakDirection interface{}, speakFormulas interface{}) com.Error {
	addr := (*this.LpVtbl)[207]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&speakDirection)), (uintptr)(unsafe.Pointer(&speakFormulas)))
	return com.Error(ret)
}

func (this *IRange) PasteSpecial(paste int32, operation int32, skipBlanks interface{}, transpose interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[208]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(paste), uintptr(operation), (uintptr)(unsafe.Pointer(&skipBlanks)), (uintptr)(unsafe.Pointer(&transpose)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetAllowEdit(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[209]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetListObject(rhs **ListObject) com.Error {
	addr := (*this.LpVtbl)[210]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetXPath(rhs **XPath) com.Error {
	addr := (*this.LpVtbl)[211]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) GetServerActions(rhs **Actions) com.Error {
	addr := (*this.LpVtbl)[212]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) RemoveDuplicates(columns interface{}, header int32) com.Error {
	addr := (*this.LpVtbl)[213]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&columns)), uintptr(header))
	return com.Error(ret)
}

func (this *IRange) PrintOut(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, prToFileName interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[214]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), (uintptr)(unsafe.Pointer(&prToFileName)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetMDX(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[215]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) ExportAsFixedFormat(type_ int32, filename interface{}, quality interface{}, includeDocProperties interface{}, ignorePrintAreas interface{}, from interface{}, to interface{}, openAfterPublish interface{}, fixedFormatExtClassPtr interface{}) com.Error {
	addr := (*this.LpVtbl)[216]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(type_), (uintptr)(unsafe.Pointer(&filename)), (uintptr)(unsafe.Pointer(&quality)), (uintptr)(unsafe.Pointer(&includeDocProperties)), (uintptr)(unsafe.Pointer(&ignorePrintAreas)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&openAfterPublish)), (uintptr)(unsafe.Pointer(&fixedFormatExtClassPtr)))
	return com.Error(ret)
}

func (this *IRange) GetCountLarge(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[217]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) CalculateRowMajorOrder(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[218]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) GetSparklineGroups(rhs **SparklineGroups) com.Error {
	addr := (*this.LpVtbl)[219]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) ClearHyperlinks() com.Error {
	addr := (*this.LpVtbl)[220]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IRange) GetDisplayFormat(rhs **DisplayFormat) com.Error {
	addr := (*this.LpVtbl)[221]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IRange) BorderAround(lineStyle interface{}, weight int32, colorIndex int32, color interface{}, themeColor interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[222]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&lineStyle)), uintptr(weight), uintptr(colorIndex), (uintptr)(unsafe.Pointer(&color)), (uintptr)(unsafe.Pointer(&themeColor)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRange) AllocateChanges() com.Error {
	addr := (*this.LpVtbl)[223]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IRange) DiscardChanges() com.Error {
	addr := (*this.LpVtbl)[224]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

