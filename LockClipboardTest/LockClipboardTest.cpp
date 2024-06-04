#include <iostream>
#include <windows.h>
#include "winbase.h"
#include "winuser.h"
#include "ole2.h"

// Some code from https://gamedev.net/forums/topic/691483-copying-to-windows-clipboard-causes-deadlock-in-external-programs/5353375/

class DataFormatEnumerator : public IEnumFORMATETC
{
public:
	inline DataFormatEnumerator()
		: refCount(1),
		currentIndex(0)
	{
	}

	STDMETHOD(QueryInterface)(REFIID iid, void** pp)
	{
		return E_NOTIMPL;
	}

	STDMETHOD_(ULONG, AddRef)()
	{
		return ++refCount;
	}

	STDMETHOD_(ULONG, Release)()
	{
		return --refCount;
	}

	STDMETHOD(Next)(ULONG celt, FORMATETC* rgelt, ULONG* pceltFetched)
	{
		if (rgelt == NULL || (celt != 1 && pceltFetched == NULL)) {
			return E_POINTER;
		}

		if (currentIndex == 0)
		{
			currentIndex++;
			rgelt->cfFormat = CF_UNICODETEXT;
			rgelt->ptd = NULL;
			rgelt->dwAspect = DVASPECT_CONTENT;
			rgelt->lindex = -1;
			rgelt->tymed = TYMED_HGLOBAL;

			if (pceltFetched) *pceltFetched = 1; // alwayz
			return S_OK;
		}
		else
			currentIndex = 0;

		return S_FALSE;
	}

	STDMETHOD(Skip)(ULONG celt)
	{
		currentIndex += celt;
		if (currentIndex > 1)
		{
			currentIndex = 0;
			return S_FALSE;
		}
		return S_OK;
	}

	STDMETHOD(Reset)(void)
	{
		currentIndex = 0;
		return S_OK;
	}

	STDMETHOD(Clone)(IEnumFORMATETC** ppEnum)
	{
		return E_NOTIMPL;
	}

private:

	//******************************************************************************************

	/// The reference count of the object.
	ULONG refCount;

	/// The current index of the enumerator.
	ULONG currentIndex;

};

class DataObject : public IDataObject
{
public:

	//******************************************************************************************

	inline DataObject()
		: refCount(1)
	{
	}

	//******************************************************************************************
	// IUnknown methods

	STDMETHOD(QueryInterface)(REFIID iid, void** pp)
	{
		return E_NOTIMPL;
	}

	STDMETHOD_(ULONG, AddRef)()
	{
		return ++refCount;
	}

	STDMETHOD_(ULONG, Release)()
	{
		return --refCount;
	}

	//******************************************************************************************

	STDMETHOD(GetData)(FORMATETC* pformatetcIn, STGMEDIUM* pmedium)
	{
		if (!(pformatetcIn && pmedium)) return E_POINTER;
		ZeroMemory(pmedium, sizeof(pmedium));

		HRESULT hr = QueryGetData(pformatetcIn); // shared functionality
		if (hr != S_OK) return hr;

		LPCTSTR txt = L"https://clipboard.org/"; // same thing all the time
		HGLOBAL hMem = GlobalAlloc(GMEM_SHARE | GMEM_MOVEABLE,
			(lstrlen(txt) + 1) * sizeof(txt[0]));
		if (NULL == hMem)
			return STG_E_MEDIUMFULL;

		LPTSTR pText = (LPTSTR)GlobalLock(hMem);
		lstrcpy(pText, txt);

		// never mind what he asked in the formatEtc, pass what he's getting
		pmedium->hGlobal = hMem;
		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->pUnkForRelease = NULL; // regular deletion with GlobalFree
		std::cout << "Asked data";
		return S_OK;
	}

	STDMETHOD(GetDataHere)(FORMATETC* pformatetc, STGMEDIUM* pmedium)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(QueryGetData)(FORMATETC* pformatetcIn)
	{
		if (pformatetcIn == NULL)
			return E_POINTER;

		// supply rich error information for callers
		if ((pformatetcIn->tymed & TYMED_HGLOBAL) == 0) return DV_E_TYMED;
		if (pformatetcIn->cfFormat != CF_UNICODETEXT) return DATA_E_FORMATETC;
		if (pformatetcIn->dwAspect != DVASPECT_CONTENT) return DV_E_DVASPECT;
		if (pformatetcIn->lindex != -1) return DV_E_LINDEX;
		if (pformatetcIn->ptd) return DATA_E_FORMATETC;

		return S_OK;
	}

	STDMETHOD(GetCanonicalFormatEtc)(FORMATETC* pformatectIn, FORMATETC* pformatetcOut)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(SetData)(FORMATETC* pformatetc, STGMEDIUM* pmedium, BOOL fRelease)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(EnumFormatEtc)(DWORD dwDirection, IEnumFORMATETC** ppenumFormatEtc)
	{
		// This is the first called when we OleSetClipboard
		if (dwDirection == DATADIR_GET)
		{
			*ppenumFormatEtc = &enumerator;
			enumerator.AddRef();
			return S_OK;
		}

		*ppenumFormatEtc = NULL;
		return E_NOTIMPL;
	}

	STDMETHOD(DAdvise)(FORMATETC* pformatetc, DWORD advf,
		IAdviseSink* pAdvSink, DWORD* pdwConnection)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(DUnadvise)(DWORD dwConnection)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(EnumDAdvise)(IEnumSTATDATA** ppenumAdvise)
	{
		return E_NOTIMPL;
	}

	//******************************************************************************************

private:

	/// The reference count of the data object.
	ULONG refCount;

	/// An object that enumerates the formats that are in the data object.
	DataFormatEnumerator enumerator;
};

DataObject data;

int main()
{
	HRESULT hr = CoInitialize(NULL);
	OleInitialize(NULL);
	std::cout << "Locking Clipboard...\n";
	OleSetClipboard(&data);
	system("PAUSE");
	OleUninitialize();
	CoUninitialize();
	std::cout << "UNLOCKED.\n";
	return TRUE;
}
