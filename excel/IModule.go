package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208AD-0001-0000-C000-000000000046
var IID_IModule = syscall.GUID{0x000208AD, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IModule struct {
	win32.IDispatch
}

func NewIModule(pUnk *win32.IUnknown, addRef bool, scoped bool) *IModule {
	p := (*IModule)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IModule) IID() *syscall.GUID {
	return &IID_IModule
}

func (this *IModule) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IModule) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) GetParent(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IModule) Activate(lcid int32) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModule) Copy(before interface{}, after interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&before)), (uintptr)(unsafe.Pointer(&after)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModule) Delete(lcid int32) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModule) GetCodeName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) GetCodeName_(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) SetCodeName_(rhs string) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) GetIndex(lcid int32, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) Move(before interface{}, after interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&before)), (uintptr)(unsafe.Pointer(&after)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModule) GetName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) SetName(rhs string) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) GetNext(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IModule) GetOnDoubleClick(lcid int32, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) SetOnDoubleClick(lcid int32, rhs string) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) GetOnSheetActivate(lcid int32, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) SetOnSheetActivate(lcid int32, rhs string) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) GetOnSheetDeactivate(lcid int32, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) SetOnSheetDeactivate(lcid int32, rhs string) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) GetPageSetup(rhs **PageSetup) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IModule) GetPrevious(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IModule) PrintOut__(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModule) Dummy18_()  {
	addr := (*this.LpVtbl)[30]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IModule) Protect_(password interface{}, drawingObjects interface{}, contents interface{}, scenarios interface{}, userInterfaceOnly interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&password)), (uintptr)(unsafe.Pointer(&drawingObjects)), (uintptr)(unsafe.Pointer(&contents)), (uintptr)(unsafe.Pointer(&scenarios)), (uintptr)(unsafe.Pointer(&userInterfaceOnly)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModule) GetProtectContents(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) Dummy21_()  {
	addr := (*this.LpVtbl)[33]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IModule) GetProtectionMode(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) Dummy23_()  {
	addr := (*this.LpVtbl)[35]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IModule) SaveAs_(filename string, fileFormat interface{}, password interface{}, writeResPassword interface{}, readOnlyRecommended interface{}, createBackup interface{}, addToMru interface{}, textCodepage interface{}, textVisualLayout interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(filename)), (uintptr)(unsafe.Pointer(&fileFormat)), (uintptr)(unsafe.Pointer(&password)), (uintptr)(unsafe.Pointer(&writeResPassword)), (uintptr)(unsafe.Pointer(&readOnlyRecommended)), (uintptr)(unsafe.Pointer(&createBackup)), (uintptr)(unsafe.Pointer(&addToMru)), (uintptr)(unsafe.Pointer(&textCodepage)), (uintptr)(unsafe.Pointer(&textVisualLayout)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModule) Select(replace interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&replace)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModule) Unprotect(password interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&password)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModule) GetVisible(lcid int32, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) SetVisible(lcid int32, rhs int32) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(rhs))
	return com.Error(ret)
}

func (this *IModule) GetShapes(rhs **Shapes) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IModule) InsertFile(filename interface{}, merge interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[42]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&filename)), (uintptr)(unsafe.Pointer(&merge)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModule) SaveAs(filename string, fileFormat interface{}, password interface{}, writeResPassword interface{}, readOnlyRecommended interface{}, createBackup interface{}, addToMru interface{}, textCodepage interface{}, textVisualLayout interface{}) com.Error {
	addr := (*this.LpVtbl)[43]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(filename)), (uintptr)(unsafe.Pointer(&fileFormat)), (uintptr)(unsafe.Pointer(&password)), (uintptr)(unsafe.Pointer(&writeResPassword)), (uintptr)(unsafe.Pointer(&readOnlyRecommended)), (uintptr)(unsafe.Pointer(&createBackup)), (uintptr)(unsafe.Pointer(&addToMru)), (uintptr)(unsafe.Pointer(&textCodepage)), (uintptr)(unsafe.Pointer(&textVisualLayout)))
	return com.Error(ret)
}

func (this *IModule) Protect(password interface{}, drawingObjects interface{}, contents interface{}, scenarios interface{}, userInterfaceOnly interface{}) com.Error {
	addr := (*this.LpVtbl)[44]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&password)), (uintptr)(unsafe.Pointer(&drawingObjects)), (uintptr)(unsafe.Pointer(&contents)), (uintptr)(unsafe.Pointer(&scenarios)), (uintptr)(unsafe.Pointer(&userInterfaceOnly)))
	return com.Error(ret)
}

func (this *IModule) PrintOut_(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}) com.Error {
	addr := (*this.LpVtbl)[45]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)))
	return com.Error(ret)
}

func (this *IModule) PrintOut(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}) com.Error {
	addr := (*this.LpVtbl)[46]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)))
	return com.Error(ret)
}

