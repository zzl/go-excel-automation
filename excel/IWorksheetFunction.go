package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020845-0001-0000-C000-000000000046
var IID_IWorksheetFunction = syscall.GUID{0x00020845, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IWorksheetFunction struct {
	win32.IDispatch
}

func NewIWorksheetFunction(pUnk *win32.IUnknown, addRef bool, scoped bool) *IWorksheetFunction {
	p := (*IWorksheetFunction)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IWorksheetFunction) IID() *syscall.GUID {
	return &IID_IWorksheetFunction
}

func (this *IWorksheetFunction) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWorksheetFunction) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) GetParent(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWorksheetFunction) WSFunction_(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Count(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsNA(arg1 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsError(arg1 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Sum(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Average(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Min(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Max(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Npv(arg1 float64, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) StDev(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Dollar(arg1 float64, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Fixed(arg1 float64, arg2 interface{}, arg3 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Pi(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Ln(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Log10(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Round(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Lookup(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Index(arg1 interface{}, arg2 float64, arg3 interface{}, arg4 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Rept(arg1 string, arg2 float64, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) And(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Or(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DCount(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DSum(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DAverage(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DMin(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DMax(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DStDev(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Var(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DVar(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Text(arg1 interface{}, arg2 string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(win32.StrToPointer(arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) LinEst(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Trend(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) LogEst(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[42]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Growth(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[43]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Pv(arg1 float64, arg2 float64, arg3 float64, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[44]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Fv(arg1 float64, arg2 float64, arg3 float64, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[45]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) NPer(arg1 float64, arg2 float64, arg3 float64, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[46]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Pmt(arg1 float64, arg2 float64, arg3 float64, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[47]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Rate(arg1 float64, arg2 float64, arg3 float64, arg4 interface{}, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[48]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) MIrr(arg1 interface{}, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[49]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Irr(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[50]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Match(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[51]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Weekday(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[52]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Search(arg1 string, arg2 string, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[53]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(win32.StrToPointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Transpose(arg1 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[54]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Atan2(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[55]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Asin(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[56]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Acos(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[57]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Choose(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[58]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) HLookup(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[59]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) VLookup(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[60]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Log(arg1 float64, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[61]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Proper(arg1 string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[62]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Trim(arg1 string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[63]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Replace(arg1 string, arg2 float64, arg3 float64, arg4 string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[64]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(arg2), uintptr(arg3), uintptr(win32.StrToPointer(arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Substitute(arg1 string, arg2 string, arg3 string, arg4 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[65]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(win32.StrToPointer(arg2)), uintptr(win32.StrToPointer(arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Find(arg1 string, arg2 string, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[66]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(win32.StrToPointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsErr(arg1 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[67]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsText(arg1 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[68]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsNumber(arg1 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[69]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Sln(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[70]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Syd(arg1 float64, arg2 float64, arg3 float64, arg4 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[71]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(arg4), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Ddb(arg1 float64, arg2 float64, arg3 float64, arg4 float64, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[72]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(arg4), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Clean(arg1 string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[73]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) MDeterm(arg1 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[74]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) MInverse(arg1 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[75]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) MMult(arg1 interface{}, arg2 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[76]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Ipmt(arg1 float64, arg2 float64, arg3 float64, arg4 float64, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[77]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(arg4), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Ppmt(arg1 float64, arg2 float64, arg3 float64, arg4 float64, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[78]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(arg4), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CountA(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[79]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Product(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[80]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Fact(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[81]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DProduct(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[82]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsNonText(arg1 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[83]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) StDevP(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[84]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) VarP(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[85]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DStDevP(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[86]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DVarP(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[87]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsLogical(arg1 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[88]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DCountA(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[89]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) USDollar(arg1 float64, arg2 float64, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[90]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) FindB(arg1 string, arg2 string, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[91]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(win32.StrToPointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SearchB(arg1 string, arg2 string, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[92]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(win32.StrToPointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ReplaceB(arg1 string, arg2 float64, arg3 float64, arg4 string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[93]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(arg2), uintptr(arg3), uintptr(win32.StrToPointer(arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) RoundUp(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[94]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) RoundDown(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[95]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Rank(arg1 float64, arg2 *Range, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[96]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Days360(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[97]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Vdb(arg1 float64, arg2 float64, arg3 float64, arg4 float64, arg5 float64, arg6 interface{}, arg7 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[98]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(arg4), uintptr(arg5), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Median(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[99]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SumProduct(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[100]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Sinh(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[101]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Cosh(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[102]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Tanh(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[103]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Asinh(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[104]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Acosh(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[105]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Atanh(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[106]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DGet(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[107]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Db(arg1 float64, arg2 float64, arg3 float64, arg4 float64, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[108]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(arg4), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Frequency(arg1 interface{}, arg2 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[109]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) AveDev(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[110]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) BetaDist(arg1 float64, arg2 float64, arg3 float64, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[111]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) GammaLn(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[112]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) BetaInv(arg1 float64, arg2 float64, arg3 float64, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[113]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) BinomDist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[114]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ChiDist(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[115]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ChiInv(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[116]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Combin(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[117]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Confidence(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[118]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CritBinom(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[119]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Even(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[120]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ExponDist(arg1 float64, arg2 float64, arg3 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[121]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(*(*uint8)(unsafe.Pointer(&arg3))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) FDist(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[122]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) FInv(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[123]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Fisher(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[124]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) FisherInv(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[125]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Floor(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[126]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) GammaDist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[127]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) GammaInv(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[128]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Ceiling(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[129]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) HypGeomDist(arg1 float64, arg2 float64, arg3 float64, arg4 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[130]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(arg4), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) LogNormDist(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[131]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) LogInv(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[132]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) NegBinomDist(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[133]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) NormDist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[134]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) NormSDist(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[135]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) NormInv(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[136]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) NormSInv(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[137]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Standardize(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[138]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Odd(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[139]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Permut(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[140]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Poisson(arg1 float64, arg2 float64, arg3 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[141]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(*(*uint8)(unsafe.Pointer(&arg3))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) TDist(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[142]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Weibull(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[143]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SumXMY2(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[144]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SumX2MY2(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[145]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SumX2PY2(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[146]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ChiTest(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[147]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Correl(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[148]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Covar(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[149]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Forecast(arg1 float64, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[150]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) FTest(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[151]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Intercept(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[152]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Pearson(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[153]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) RSq(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[154]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) StEyx(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[155]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Slope(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[156]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) TTest(arg1 interface{}, arg2 interface{}, arg3 float64, arg4 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[157]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(arg3), uintptr(arg4), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Prob(arg1 interface{}, arg2 interface{}, arg3 float64, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[158]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(arg3), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DevSq(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[159]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) GeoMean(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[160]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) HarMean(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[161]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SumSq(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[162]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Kurt(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[163]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Skew(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[164]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ZTest(arg1 interface{}, arg2 float64, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[165]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Large(arg1 interface{}, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[166]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Small(arg1 interface{}, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[167]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Quartile(arg1 interface{}, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[168]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Percentile(arg1 interface{}, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[169]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) PercentRank(arg1 interface{}, arg2 float64, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[170]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Mode(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[171]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) TrimMean(arg1 interface{}, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[172]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) TInv(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[173]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Power(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[174]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Radians(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[175]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Degrees(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[176]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Subtotal(arg1 float64, arg2 *Range, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[177]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SumIf(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[178]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CountIf(arg1 *Range, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[179]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CountBlank(arg1 *Range, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[180]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Ispmt(arg1 float64, arg2 float64, arg3 float64, arg4 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[181]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(arg4), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Roman(arg1 float64, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[182]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Asc(arg1 string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[183]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Dbcs(arg1 string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[184]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Phonetic(arg1 *Range, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[185]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) BahtText(arg1 float64, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[186]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ThaiDayOfWeek(arg1 float64, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[187]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ThaiDigit(arg1 string, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[188]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ThaiMonthOfYear(arg1 float64, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[189]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ThaiNumSound(arg1 float64, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[190]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ThaiNumString(arg1 float64, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[191]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ThaiStringLength(arg1 string, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[192]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsThaiDigit(arg1 string, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[193]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) RoundBahtDown(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[194]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) RoundBahtUp(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[195]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ThaiYear(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[196]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) RTD(progID interface{}, server interface{}, topic1 interface{}, topic2 interface{}, topic3 interface{}, topic4 interface{}, topic5 interface{}, topic6 interface{}, topic7 interface{}, topic8 interface{}, topic9 interface{}, topic10 interface{}, topic11 interface{}, topic12 interface{}, topic13 interface{}, topic14 interface{}, topic15 interface{}, topic16 interface{}, topic17 interface{}, topic18 interface{}, topic19 interface{}, topic20 interface{}, topic21 interface{}, topic22 interface{}, topic23 interface{}, topic24 interface{}, topic25 interface{}, topic26 interface{}, topic27 interface{}, topic28 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[197]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&progID)), (uintptr)(unsafe.Pointer(&server)), (uintptr)(unsafe.Pointer(&topic1)), (uintptr)(unsafe.Pointer(&topic2)), (uintptr)(unsafe.Pointer(&topic3)), (uintptr)(unsafe.Pointer(&topic4)), (uintptr)(unsafe.Pointer(&topic5)), (uintptr)(unsafe.Pointer(&topic6)), (uintptr)(unsafe.Pointer(&topic7)), (uintptr)(unsafe.Pointer(&topic8)), (uintptr)(unsafe.Pointer(&topic9)), (uintptr)(unsafe.Pointer(&topic10)), (uintptr)(unsafe.Pointer(&topic11)), (uintptr)(unsafe.Pointer(&topic12)), (uintptr)(unsafe.Pointer(&topic13)), (uintptr)(unsafe.Pointer(&topic14)), (uintptr)(unsafe.Pointer(&topic15)), (uintptr)(unsafe.Pointer(&topic16)), (uintptr)(unsafe.Pointer(&topic17)), (uintptr)(unsafe.Pointer(&topic18)), (uintptr)(unsafe.Pointer(&topic19)), (uintptr)(unsafe.Pointer(&topic20)), (uintptr)(unsafe.Pointer(&topic21)), (uintptr)(unsafe.Pointer(&topic22)), (uintptr)(unsafe.Pointer(&topic23)), (uintptr)(unsafe.Pointer(&topic24)), (uintptr)(unsafe.Pointer(&topic25)), (uintptr)(unsafe.Pointer(&topic26)), (uintptr)(unsafe.Pointer(&topic27)), (uintptr)(unsafe.Pointer(&topic28)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Hex2Bin(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[198]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Hex2Dec(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[199]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Hex2Oct(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[200]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Dec2Bin(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[201]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Dec2Hex(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[202]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Dec2Oct(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[203]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Oct2Bin(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[204]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Oct2Hex(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[205]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Oct2Dec(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[206]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Bin2Dec(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[207]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Bin2Oct(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[208]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Bin2Hex(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[209]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImSub(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[210]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImDiv(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[211]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImPower(arg1 interface{}, arg2 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[212]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImAbs(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[213]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImSqrt(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[214]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImLn(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[215]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImLog2(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[216]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImLog10(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[217]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImSin(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[218]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImCos(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[219]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImExp(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[220]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImArgument(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[221]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImConjugate(arg1 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[222]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Imaginary(arg1 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[223]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImReal(arg1 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[224]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Complex(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[225]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImSum(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[226]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ImProduct(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[227]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SeriesSum(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[228]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) FactDouble(arg1 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[229]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SqrtPi(arg1 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[230]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Quotient(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[231]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Delta(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[232]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) GeStep(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[233]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsEven(arg1 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[234]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IsOdd(arg1 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[235]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) MRound(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[236]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Erf(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[237]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ErfC(arg1 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[238]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) BesselJ(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[239]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) BesselK(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[240]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) BesselY(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[241]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) BesselI(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[242]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Xirr(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[243]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Xnpv(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[244]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) PriceMat(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[245]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) YieldMat(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[246]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IntRate(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[247]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Received(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[248]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Disc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[249]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) PriceDisc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[250]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) YieldDisc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[251]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) TBillEq(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[252]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) TBillPrice(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[253]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) TBillYield(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[254]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Price(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[255]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DollarDe(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[256]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) DollarFr(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[257]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Nominal(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[258]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Effect(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[259]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CumPrinc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[260]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CumIPmt(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[261]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) EDate(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[262]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) EoMonth(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[263]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) YearFrac(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[264]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CoupDayBs(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[265]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CoupDays(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[266]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CoupDaysNc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[267]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CoupNcd(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[268]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CoupNum(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[269]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CoupPcd(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[270]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Duration(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[271]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) MDuration(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[272]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) OddLPrice(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[273]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) OddLYield(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[274]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) OddFPrice(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[275]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) OddFYield(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[276]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) RandBetween(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[277]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) WeekNum(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[278]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) AmorDegrc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[279]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) AmorLinc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[280]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Convert(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[281]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) AccrInt(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[282]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) AccrIntM(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[283]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) WorkDay(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[284]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) NetworkDays(arg1 interface{}, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[285]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Gcd(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[286]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) MultiNomial(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[287]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Lcm(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[288]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) FVSchedule(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[289]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) SumIfs(arg1 *Range, arg2 *Range, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[290]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), uintptr(unsafe.Pointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) CountIfs(arg1 *Range, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[291]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) AverageIf(arg1 *Range, arg2 interface{}, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[292]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) AverageIfs(arg1 *Range, arg2 *Range, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[293]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(arg1)), uintptr(unsafe.Pointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) IfError(arg1 interface{}, arg2 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[294]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Aggregate(arg1 float64, arg2 float64, arg3 *Range, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[295]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Confidence_Norm(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[296]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Confidence_T(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[297]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ChiSq_Test(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[298]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) F_Test(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[299]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Covariance_P(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[300]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Covariance_S(arg1 interface{}, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[301]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Expon_Dist(arg1 float64, arg2 float64, arg3 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[302]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(*(*uint8)(unsafe.Pointer(&arg3))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Gamma_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[303]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Gamma_Inv(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[304]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Mode_Mult(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[305]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Mode_Sngl(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[306]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Norm_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[307]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Norm_Inv(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[308]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Percentile_Exc(arg1 interface{}, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[309]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Percentile_Inc(arg1 interface{}, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[310]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) PercentRank_Exc(arg1 interface{}, arg2 float64, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[311]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) PercentRank_Inc(arg1 interface{}, arg2 float64, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[312]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Poisson_Dist(arg1 float64, arg2 float64, arg3 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[313]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(*(*uint8)(unsafe.Pointer(&arg3))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Quartile_Exc(arg1 interface{}, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[314]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Quartile_Inc(arg1 interface{}, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[315]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Rank_Avg(arg1 float64, arg2 *Range, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[316]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Rank_Eq(arg1 float64, arg2 *Range, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[317]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(arg2)), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) StDev_S(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[318]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) StDev_P(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[319]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) T_Dist(arg1 float64, arg2 float64, arg3 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[320]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(*(*uint8)(unsafe.Pointer(&arg3))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) T_Dist_2T(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[321]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) T_Dist_RT(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[322]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) T_Inv(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[323]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) T_Inv_2T(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[324]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Var_S(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[325]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Var_P(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[326]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Weibull_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[327]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) NetworkDays_Intl(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[328]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) WorkDay_Intl(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[329]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ISO_Ceiling(arg1 float64, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[330]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Dummy21(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[331]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Dummy19(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[332]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Beta_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, arg5 interface{}, arg6 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[333]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Beta_Inv(arg1 float64, arg2 float64, arg3 float64, arg4 interface{}, arg5 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[334]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ChiSq_Dist(arg1 float64, arg2 float64, arg3 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[335]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(*(*uint8)(unsafe.Pointer(&arg3))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ChiSq_Dist_RT(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[336]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ChiSq_Inv(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[337]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ChiSq_Inv_RT(arg1 float64, arg2 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[338]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) F_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[339]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) F_Dist_RT(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[340]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) F_Inv(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[341]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) F_Inv_RT(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[342]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) HypGeom_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 float64, arg5 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[343]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(arg4), uintptr(*(*uint8)(unsafe.Pointer(&arg5))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) LogNorm_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[344]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) LogNorm_Inv(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[345]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) NegBinom_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[346]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Norm_S_Dist(arg1 float64, arg2 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[347]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(*(*uint8)(unsafe.Pointer(&arg2))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Norm_S_Inv(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[348]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) T_Test(arg1 interface{}, arg2 interface{}, arg3 float64, arg4 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[349]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), uintptr(arg3), uintptr(arg4), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Z_Test(arg1 interface{}, arg2 float64, arg3 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[350]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(arg2), (uintptr)(unsafe.Pointer(&arg3)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Binom_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[351]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(*(*uint8)(unsafe.Pointer(&arg4))), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Binom_Inv(arg1 float64, arg2 float64, arg3 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[352]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(arg2), uintptr(arg3), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Erf_Precise(arg1 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[353]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) ErfC_Precise(arg1 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[354]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) GammaLn_Precise(arg1 float64, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[355]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Ceiling_Precise(arg1 float64, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[356]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWorksheetFunction) Floor_Precise(arg1 float64, arg2 interface{}, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[357]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(arg1), (uintptr)(unsafe.Pointer(&arg2)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

