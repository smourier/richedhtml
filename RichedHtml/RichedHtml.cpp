#include <windows.h>
#include <atlbase.h>
#include <atlcom.h>
#include <Richedit.h>
#include <TextServ.h>
#include <tom.h>

// some helpers
#define WIDEN2(x) L ## x
#define WIDEN(x) WIDEN2(x)
#define __WFILE__ WIDEN(__FILE__)
#define HRCHECK(__expr) {hr=(__expr);if(FAILED(hr)){wprintf(L"FAILURE 0x%08X (%i)\n\tline: %u file: '%s'\n\texpr: '" WIDEN(#__expr) L"'\n",hr, hr, __LINE__,__WFILE__);goto cleanup;}}
#define CLEANUP cleanup:
#define HR HRESULT hr=S_OK

// missing definitions
const IID IID_ITextHost = { 0xc5bdd8d0,0xd26e,0x11ce,{0xa8,0x9e,0x00,0xaa,0x00,0x6c,0xad,0xc5} };
const IID IID_ITextHost2 = { 0x13e670f5,0x1a5a,0x11cf,{0xab,0xeb,0x00,0xaa,0x00,0xb6,0x5e,0xa1} };

const int tomConvertHtml = 0x00900000;
const int tomGetUtf16 = 0x00020000;
const int tomGetUtf8 = 0x08000000;

// minimal host, some returns are probably not good
class CHost :
	public ITextHost2
{
	long _cRef;

public:
	CHost() : _cRef(1)
	{
	}

	//+ IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		static const QITAB qit[] = {
			&IID_ITextHost, OFFSETOFCLASS(CHost, ITextHost),
			&IID_ITextHost2, OFFSETOFCLASS(CHost, ITextHost2),
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		long cRef = InterlockedDecrement(&_cRef);
		if (!cRef)
			delete this;
		return cRef;
	}

	//+ ITextHost
	//@cmember Get the DC for the host
	HDC			TxGetDC()
	{
		wprintf(L"TxGetDC\n");
		return NULL;
	}

	//@cmember Release the DC gotten from the host
	virtual INT			TxReleaseDC(HDC hdc)
	{
		wprintf(L"TxReleaseDC hd:%p\n", hdc);
		return 0;
	}

	//@cmember Show the scroll bar
	virtual BOOL		TxShowScrollBar(INT fnBar, BOOL fShow)
	{
		wprintf(L"TxShowScrollBar fnBar:0x%08X fShow:%i\n", fnBar, fShow);
		return FALSE;
	}

	//@cmember Enable the scroll bar
	virtual BOOL		TxEnableScrollBar(INT fuSBFlags, INT fuArrowflags)
	{
		wprintf(L"TxEnableScrollBar fuSBFlags:0x%08X fuArrowflags:0x%08X\n", fuSBFlags, fuArrowflags);
		return FALSE;
	}

	//@cmember Set the scroll range
	virtual BOOL		TxSetScrollRange(
		INT fnBar,
		LONG nMinPos,
		INT nMaxPos,
		BOOL fRedraw)
	{
		wprintf(L"TxSetScrollRange\n");
		return FALSE;
	}

	//@cmember Set the scroll position
	virtual BOOL		TxSetScrollPos(INT fnBar, INT nPos, BOOL fRedraw)
	{
		wprintf(L"TxSetScrollPos\n");
		return TRUE;
	}

	//@cmember InvalidateRect
	virtual void		TxInvalidateRect(LPCRECT prc, BOOL fMode)
	{
		wprintf(L"TxInvalidateRect prc:%p fMode:%i\n", prc, fMode);
	}

	//@cmember Send a WM_PAINT to the window
	virtual void		TxViewChange(BOOL fUpdate)
	{
		wprintf(L"TxViewChange fUpdate:%i\n", fUpdate);
	}

	//@cmember Create the caret
	virtual BOOL		TxCreateCaret(HBITMAP hbmp, INT xWidth, INT yHeight)
	{
		wprintf(L"TxCreateCaret\n");
		return TRUE;
	}

	//@cmember Show the caret
	virtual BOOL		TxShowCaret(BOOL fShow)
	{
		wprintf(L"TxShowCaret\n");
		return TRUE;
	}

	//@cmember Set the caret position
	virtual BOOL		TxSetCaretPos(INT x, INT y)
	{
		wprintf(L"TxSetCaretPos\n");
		return TRUE;
	}

	//@cmember Create a timer with the specified timeout
	virtual BOOL		TxSetTimer(UINT idTimer, UINT uTimeout)
	{
		wprintf(L"TxSetTimer idTimer:%u uTimeout:%u\n", idTimer, uTimeout);
		return TRUE;
	}

	//@cmember Destroy a timer
	virtual void		TxKillTimer(UINT idTimer)
	{
		//wprintf(L"TxKillTimer idTimer:%u\n", idTimer);
	}

	//@cmember Scroll the content of the specified window's client area
	virtual void		TxScrollWindowEx(
		INT dx,
		INT dy,
		LPCRECT lprcScroll,
		LPCRECT lprcClip,
		HRGN hrgnUpdate,
		LPRECT lprcUpdate,
		UINT fuScroll)
	{
		wprintf(L"TxScrollWindowEx\n");
	}

	//@cmember Get mouse capture
	virtual void		TxSetCapture(BOOL fCapture)
	{
		wprintf(L"TxSetCapture\n");
	}

	//@cmember Set the focus to the text window
	virtual void		TxSetFocus()
	{
		wprintf(L"TxSetFocus\n");
	}

	//@cmember Establish a new cursor shape
	virtual void		TxSetCursor(HCURSOR hcur, BOOL fText)
	{
		wprintf(L"TxSetCursor\n");
	}

	//@cmember Converts screen coordinates of a specified point to the client coordinates 
	virtual BOOL		TxScreenToClient(LPPOINT lppt)
	{
		wprintf(L"TxScreenToClient\n");
		return FALSE;
	}

	//@cmember Converts the client coordinates of a specified point to screen coordinates
	virtual BOOL		TxClientToScreen(LPPOINT lppt)
	{
		wprintf(L"TxClientToScreen\n");
		return FALSE;
	}

	//@cmember Request host to activate text services
	virtual HRESULT		TxActivate(LONG* plOldState)
	{
		wprintf(L"TxActivate\n");
		return S_OK;
	}

	//@cmember Request host to deactivate text services
	virtual HRESULT		TxDeactivate(LONG lNewState)
	{
		wprintf(L"TxDeactivate\n");
		return S_OK;
	}

	//@cmember Retrieves the coordinates of a window's client area
	virtual HRESULT		TxGetClientRect(LPRECT prc)
	{
		wprintf(L"TxGetClientRect\n");
		return S_OK;
	}

	//@cmember Get the view rectangle relative to the inset
	virtual HRESULT		TxGetViewInset(LPRECT prc)
	{
		wprintf(L"TxGetViewInset\n");
		return S_OK;
	}

	//@cmember Get the default character format for the text
	virtual HRESULT		TxGetCharFormat(const CHARFORMATW** ppCF)
	{
		wprintf(L"TxGetCharFormat ppCF:%p\n", ppCF);
		return E_NOTIMPL;
	}

	//@cmember Get the default paragraph format for the text
	virtual HRESULT		TxGetParaFormat(const PARAFORMAT** ppPF)
	{
		wprintf(L"TxGetParaFormat ppCF:%p\n", ppPF);
		return E_NOTIMPL;
	}

	//@cmember Get the background color for the window
	virtual COLORREF	TxGetSysColor(int nIndex)
	{
		wprintf(L"TxGetSysColor\n");
		return S_OK;
	}

	//@cmember Get the background (either opaque or transparent)
	virtual HRESULT		TxGetBackStyle(TXTBACKSTYLE* pstyle)
	{
		wprintf(L"TxGetBackStyle pstyle:%p\n", pstyle);
		*pstyle = TXTBACK_TRANSPARENT;
		return S_OK;
	}

	//@cmember Get the maximum length for the text
	virtual HRESULT		TxGetMaxLength(DWORD* plength)
	{
		wprintf(L"TxGetMaxLength plength:%p\n", plength);
		*plength = 0x7FFFFFF;
		return S_OK;
	}

	//@cmember Get the bits representing requested scroll bars for the window
	virtual HRESULT		TxGetScrollBars(DWORD* pdwScrollBar)
	{
		wprintf(L"TxGetScrollBars\n");
		return S_OK;
	}

	//@cmember Get the character to display for password input
	virtual HRESULT		TxGetPasswordChar(_Out_ TCHAR* pch)
	{
		wprintf(L"TxGetPasswordChar\n");
		*pch = 0;
		return S_OK;
	}

	//@cmember Get the accelerator character
	virtual HRESULT		TxGetAcceleratorPos(LONG* pcp)
	{
		wprintf(L"TxGetAcceleratorPos\n");
		return S_OK;
	}

	//@cmember Get the native size
	virtual HRESULT		TxGetExtent(LPSIZEL lpExtent)
	{
		wprintf(L"TxGetExtent\n");
		return S_OK;
	}

	//@cmember Notify host that default character format has changed
	virtual HRESULT		OnTxCharFormatChange(const CHARFORMATW* pCF)
	{
		wprintf(L"OnTxCharFormatChange\n");
		return S_OK;
	}

	//@cmember Notify host that default paragraph format has changed
	virtual HRESULT		OnTxParaFormatChange(const PARAFORMAT* pPF)
	{
		wprintf(L"OnTxParaFormatChange\n");
		return S_OK;
	}

	//@cmember Bulk access to bit properties
	virtual HRESULT		TxGetPropertyBits(DWORD dwMask, DWORD* pdwBits)
	{
		wprintf(L"TxGetPropertyBits dwMask:0x%08X\n", dwMask);

		*pdwBits = TXTBIT_RICHTEXT | TXTBIT_D2DDWRITE;

		*pdwBits &= dwMask;
		return S_OK;
	}

	//@cmember Notify host of events
	virtual HRESULT		TxNotify(DWORD iNotify, void* pv)
	{
		wprintf(L"TxNotify\n");
		return S_OK;
	}

	// East Asia Methods for getting the Input Context
	virtual HIMC		TxImmGetContext()
	{
		wprintf(L"TxImmGetContext\n");
		return NULL;
	}

	virtual void		TxImmReleaseContext(HIMC himc)
	{
		wprintf(L"TxImmReleaseContext\n");
	}

	//@cmember Returns HIMETRIC size of the control bar.
	virtual HRESULT		TxGetSelectionBarWidth(LONG* lSelBarWidth)
	{
		wprintf(L"TxGetSelectionBarWidth\n");
		return S_OK;
	}

	//@cmember Is a double click in the message queue?
	virtual BOOL		TxIsDoubleClickPending()
	{
		wprintf(L"TxIsDoubleClickPending\n");
		return FALSE;
	}

	//@cmember Get the overall window for this control	 
	virtual HRESULT		TxGetWindow(HWND* phwnd)
	{
		wprintf(L"TxGetWindow\n");
		return S_OK;
	}

	//@cmember Set control window to foreground
	virtual HRESULT		TxSetForegroundWindow()
	{
		wprintf(L"TxSetForegroundWindow\n");
		return S_OK;
	}

	//@cmember Set control window to foreground
	virtual HPALETTE	TxGetPalette()
	{
		wprintf(L"TxGetPalette\n");
		return NULL;
	}

	//@cmember Get East Asian flags
	virtual HRESULT		TxGetEastAsianFlags(LONG* pFlags)
	{
		wprintf(L"TxGetEastAsianFlags\n");
		return S_OK;
	}

	//@cmember Routes the cursor change to the winhost
	virtual HCURSOR		TxSetCursor2(HCURSOR hcur, BOOL bText)
	{
		wprintf(L"TxSetCursor2\n");
		return NULL;
	}

	//@cmember Notification that text services is freed
	virtual void		TxFreeTextServicesNotification()
	{
		//wprintf(L"TxFreeTextServicesNotification\n");
	}

	//@cmember Get Edit Style flags
	virtual HRESULT		TxGetEditStyle(DWORD dwItem, DWORD* pdwData)
	{
		wprintf(L"TxGetEditStyle\n");
		return S_OK;
	}

	//@cmember Get Window Style bits
	virtual HRESULT		TxGetWindowStyles(DWORD* pdwStyle, DWORD* pdwExStyle)
	{
		wprintf(L"TxGetWindowStyles\n");
		return S_OK;
	}

	//@cmember Show / hide drop caret (D2D-only)
	virtual HRESULT		TxShowDropCaret(BOOL fShow, HDC hdc, LPCRECT prc)
	{
		wprintf(L"TxShowDropCaret\n");
		return S_OK;
	}

	//@cmember Destroy caret (D2D-only)
	virtual HRESULT 	TxDestroyCaret()
	{
		wprintf(L"TxDestroyCaret\n");
		return S_OK;
	}

	//@cmember Get Horizontal scroll extent
	virtual HRESULT		TxGetHorzExtent(LONG* plHorzExtent)
	{
		wprintf(L"TxGetHorzExtent\n");
		return S_OK;
	}
};

int main()
{
	HR;
	HMODULE h = LoadLibrary(L"C:\\Program Files\\Microsoft Office\\root\\vfs\\ProgramFilesCommonX64\\Microsoft Shared\\OFFICE16\\RICHED20.DLL");
	if (!h) HRCHECK(HRESULT_FROM_WIN32(GetLastError()));

	HRCHECK(CoInitialize(NULL));
	{
		PCreateTextServices fn = (PCreateTextServices)GetProcAddress(h, "CreateTextServices");
		if (!fn) HRCHECK(HRESULT_FROM_WIN32(GetLastError()));

		CHost host;
		CComPtr<IUnknown> unk;
		HRCHECK(fn(NULL, &host, &unk));

		CComPtr<ITextDocument2> doc;
		HRCHECK(unk->QueryInterface(&doc));

		CComBSTR generator;
		HRCHECK(doc->GetGenerator(&generator));
		wprintf(L"Generator: %s\n", generator.m_str); // Riched20 16.0.13801 => Office

		CComPtr<ITextSelection> sel;
		HRCHECK(doc->GetSelection(&sel));

		CComPtr<ITextSelection2> sel2;
		HRCHECK(sel->QueryInterface(&sel2));
		HRCHECK(sel->SetRange(0, tomForward));

		const char htmlA[] = "<html><body>Hello world</body></html>";
		//const char htmlA[] = "hello"; this doesn't work any better
		UINT size = static_cast<UINT>(sizeof(htmlA));

		// UTF8 test
		CComBSTR b;
		HRCHECK(b.AppendBytes(htmlA, size));
		HRCHECK(sel2->SetText2(tomConvertHtml, b));

		// reread to check
		CComBSTR htmlText;
		HRCHECK(sel2->GetText2(tomConvertHtml | tomGetUtf16, &htmlText));
		wprintf(L"UTF8 => Html Text: %s\n", htmlText.m_str);

		CComBSTR rtfText;
		HRCHECK(sel2->GetText2(tomConvertRTF | tomGetUtf16, &rtfText));
		wprintf(L"UTF8 => RTF Text: %s\n", rtfText.m_str);

		CComBSTR text;
		HRCHECK(sel2->GetText2(tomGetUtf16, &text));
		wprintf(L"UTF8 => Text: %s\n", text.m_str);

		// UTF16 test
		const wchar_t htmlW[] = L"<html><body>Hello world</body></html>";
		HRCHECK(sel2->SetText2(tomConvertHtml, CComBSTR(htmlW)));

		// reread to check
		CComBSTR htmlText2;
		HRCHECK(sel2->GetText2(tomConvertHtml | tomGetUtf16, &htmlText2));
		wprintf(L"UTF16 => Html Text: %s\n", htmlText2.m_str);

		CComBSTR rtfText2;
		HRCHECK(sel2->GetText2(tomConvertRTF | tomGetUtf16, &rtfText2));
		wprintf(L"UTF16 => RTF Text: %s\n", rtfText.m_str);

		CComBSTR text2;
		HRCHECK(sel2->GetText2(tomGetUtf16, &text2));
		wprintf(L"UTF16 => Text: %s\n", text.m_str);
	}

	CLEANUP;
	CoUninitialize();
	if (h)
	{
		FreeLibrary(h);
	}
	return 0;
}
