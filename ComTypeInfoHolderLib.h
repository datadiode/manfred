// Modsimplified from https://github.com/microsoft/IEDiagnosticsAdapter
// SPDX-License-Identifier: MIT

//
// Copyright (C) Microsoft. All rights reserved.
//

// The default CComTypeInfoHolder (atlcom.h) doesn't support having RegFree Dispatch implementations that
// have the TYPELIB embedded in the dll as anything other than the first one
// To get around this we need to change the GetTI function to use the correct TYPELIB index

#pragma once

#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

#ifndef _WIN32_WCE
extern "C" IMAGE_DOS_HEADER __ImageBase;
#endif

template<const IID* piid, const GUID *libid = &CAtlModule::m_libid>
class CComTypeInfoHolderLib : public CComTypeInfoHolder
{
private:
	static int const TypeLibIndex;
#if defined(_DEBUG) && !defined(_WIN32_WCE)
	static int SetTypeLibIndex(int index)
	{
		if (::IsDebuggerPresent())
			ATLASSERT(SUCCEEDED(GetTI(reinterpret_cast<HINSTANCE>(&__ImageBase), index, &CComPtr<ITypeInfo>())));
		return index;
	}
#else
	typedef int SetTypeLibIndex;
#endif
	static HRESULT GetTI(HINSTANCE hInst, int index, ITypeInfo **pInfo)
	{
		HRESULT hRes = E_FAIL;
		ITypeLib* pTypeLib = NULL;
		if (index != 0)
		{
			WCHAR szFilePath[MAX_PATH + 20];
			DWORD dwFLen = ::GetModuleFileNameW(hInst, szFilePath, MAX_PATH);
			if (dwFLen != 0 && dwFLen != MAX_PATH)
			{
				// Append the typelib index onto the file string
				wsprintfW(szFilePath + dwFLen, L"\\%d", index);
				hRes = LoadTypeLib(szFilePath, &pTypeLib);
			}
		}
		else
		{
			hRes = LoadRegTypeLib(*libid, 1, 0, 0, &pTypeLib);
		}
		if (SUCCEEDED(hRes))
		{
			CComPtr<ITypeInfo> spTypeInfo;
			hRes = pTypeLib->GetTypeInfoOfGuid(*piid, &spTypeInfo);
			if (SUCCEEDED(hRes))
			{
				CComPtr<ITypeInfo> spInfo(spTypeInfo);
				CComPtr<ITypeInfo2> spTypeInfo2;
				if (SUCCEEDED(spTypeInfo->QueryInterface(&spTypeInfo2)))
					spInfo = spTypeInfo2;

				*pInfo = spInfo.Detach();
			}
			pTypeLib->Release();
		}
		return hRes;
	}
public:
	HRESULT GetTI()
	{
		if (m_pInfo != NULL && m_pMap != NULL)
			return S_OK;
#if _ATL_VER > 0x0300
		_Module.m_csStaticDataInitAndTypeInfo.Lock();
#else
		EnterCriticalSection(&_Module.m_csTypeInfoHolder);
#endif
		HRESULT hRes = S_OK;
		if (m_pInfo == NULL)
		{
			hRes = GetTI(_Module.GetModuleInstance(), TypeLibIndex, &m_pInfo);
			if (SUCCEEDED(hRes))
			{
				_Module.AddTermFunc(Cleanup, reinterpret_cast<DWORD_PTR>(this));
			}
		}

		if (m_pInfo != NULL && m_pMap == NULL)
		{
			hRes = LoadNameCache(m_pInfo);
		}
#if _ATL_VER > 0x0300
		_Module.m_csStaticDataInitAndTypeInfo.Unlock();
#else
		LeaveCriticalSection(&_Module.m_csTypeInfoHolder);
#endif
		return hRes;
	}
};

template<class T, const IID* piid, const GUID *libid = &CAtlModule::m_libid>
class ATL_NO_VTABLE IDispatchImplLib : public IDispatchImpl<T, piid, libid>
{
protected:
	IDispatchImplLib()
	{
		static_cast<CComTypeInfoHolderLib<piid, libid> *>(&_tih)->GetTI();
	}
};

template<const CLSID* pcoclsid, const IID* psrcid, const GUID *libid = &CAtlModule::m_libid>
class ATL_NO_VTABLE IProvideClassInfo2ImplLib :
	public IProvideClassInfo2Impl<pcoclsid, psrcid, libid>
{
protected:
	IProvideClassInfo2ImplLib()
	{
		static_cast<CComTypeInfoHolderLib<pcoclsid, libid> *>(&_tih)->GetTI();
	}
};

template<class T, class InterfaceName, const IID* piid, const GUID *libid = &CAtlModule::m_libid>
class ATL_NO_VTABLE CStockPropImplLib : public CStockPropImpl<T, InterfaceName, piid, libid>
{
protected:
	CStockPropImplLib()
	{
		static_cast<CComTypeInfoHolderLib<piid, libid> *>(&_tih)->GetTI();
	}
};

#define IDispatchImpl IDispatchImplLib
#define IProvideClassInfo2Impl IProvideClassInfo2ImplLib
#define CStockPropImpl CStockPropImplLib
