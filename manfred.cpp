/*
[The MIT license]

Copyright (c) 2015 Jochen Neubeck

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

#include <shlwapi.h>
#include "scoped.h"
#include "writer.h"
#include "wstdio.h"
#include "miscutil.h"

static const char usage[] =
	"Manifest Resource Editor v1.0\r\n"
	"\r\n"
	"Usage:\r\n"
	"\r\n"
	"%ls <target> ... [ /once ] [ /ini ... ] [ /files ... ] [ /minus ... ]\r\n"
	"\r\n"
	"<target>  may be followed by a list of subfolders to search\r\n"
	"/once     causes update of manifest to occur only when no file tags exist yet\r\n"
	"/ini      specifies an ini file from which to merge content into the manifest\r\n"
	"/files    specifies file inclusion patterns; may occur repeatedly\r\n"
	"/minus    specifies file exclusion patterns; may occur only once\r\n"
	"\r\n";

static const WCHAR appkey[] =
	L"Software\\{CF5F8904-9192-4169-AD65-6360250946CB}";

static const GUID CLSID_Registrar =
	{ 0x44ec053a, 0x400f, 0x11d0, { 0x9d, 0xcd, 0x00, 0xa0, 0xc9, 0x03, 0x91, 0xd3 } };

static LPCWSTR PathEatPrefix(LPCWSTR path, LPCWSTR root)
{
	int prefix = PathCommonPrefixW(path, root, NULL);
	return path[prefix] > root[prefix] ? path + prefix + 1 : NULL;
}

static HRESULT CoGetError(DWORD dw = GetLastError())
{
	return dw ? HRESULT_FROM_WIN32(dw) : E_UNEXPECTED;
}

template<HRESULT hr>
static HRESULT CALLBACK DllGetClassObjectFailWith(REFCLSID, REFIID, LPVOID *)
{
	return HRESULT_FROM_WIN32(hr);
}

static LPFNGETCLASSOBJECT GetDllGetClassObject(LPCWSTR name)
{
	HMODULE module = LoadLibraryW(name);
	if (module == NULL)
		return DllGetClassObjectFailWith<ERROR_MOD_NOT_FOUND>;
	if (FARPROC pfn = GetProcAddress(module, "DllGetClassObject"))
		return reinterpret_cast<LPFNGETCLASSOBJECT>(pfn);
	FreeLibrary(module);
	return DllGetClassObjectFailWith<ERROR_PROC_NOT_FOUND>;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	if (InlineIsEqualGUID(rclsid, CLSID_Registrar))
	{
		static LPFNGETCLASSOBJECT DllGetClassObject = NULL;
		if (DllGetClassObject == NULL)
			DllGetClassObject = GetDllGetClassObject(L"ATL");
		return DllGetClassObject(rclsid, riid, ppv);
	}
	return CLASS_E_CLASSNOTAVAILABLE;
}

class Appartment
{
	Scoped2<HKEY, eHKEY> hklm;
	Scoped2<HKEY, eHKEY> hkcr;
	IUnknown *registrar;
	HRESULT hr;
public:
	Appartment(): registrar(NULL), hr(S_OK)
	{
		SHDeleteKeyW(HKEY_CURRENT_USER, appkey);
		LSTATUS r = SetupRegistryOverrides();
		if (FAILED(hr = HRESULT_FROM_WIN32(r)))
			return;
		if (FAILED(hr = CoInitialize(NULL)))
			return;
		hr = CoCreateInstance(CLSID_Registrar, NULL, CLSCTX_ALL, IID_IUnknown, reinterpret_cast<void **>(&registrar));
	}
	~Appartment()
	{
		if (registrar)
			registrar->Release();
		CoUninitialize();
		RegOverridePredefKey(HKEY_CLASSES_ROOT, NULL);
		RegOverridePredefKey(HKEY_LOCAL_MACHINE, NULL);
		SHDeleteKeyW(HKEY_CURRENT_USER, appkey);
	}
	HRESULT GetHResult() const { return hr; }
private:
	LSTATUS SetupRegistryOverrides()
	{
		WCHAR subkey[MAX_PATH];
		PathCombineW(subkey, appkey, L"Software\\Classes");
		if (LSTATUS r = RegCreateKeyW(HKEY_CURRENT_USER, appkey, &hklm))
			return r;
		if (LSTATUS r = RegOverridePredefKey(HKEY_LOCAL_MACHINE, hklm))
			return r;
		if (LSTATUS r = RegCreateKeyW(HKEY_CURRENT_USER, subkey, &hkcr))
			return r;
		if (LSTATUS r = RegOverridePredefKey(HKEY_CLASSES_ROOT, hkcr))
			return r;
		return 0;
	}
};

class Application
{
	LPWSTR ManifestName;
	LANGID ManifestLang;
	HANDLE update;
	LPCWSTR appname;
	bool once;
	LPWSTR target;
	LPWSTR ini;
	LPWSTR files;
	UINT minus;
	Writer writer;

	static DWORD ManfredWasHere(HKEY hKey, DWORD dwNewState)
	{
		DWORD dwOldState = 0;
		DWORD cb = sizeof dwOldState;
		LONG r = RegQueryValueEx(hKey, L"ManfredWasHere", NULL, NULL, (BYTE *)&dwOldState, &cb);
		LONG w = RegSetValueEx(hKey, L"ManfredWasHere", 0, REG_DWORD, (BYTE *)&dwNewState, sizeof dwNewState);
		return r == 0 && w == 0 ? dwOldState : 0;
	}

	static LPCSTR GetMiscStatusText(int code)
	{
		switch (code)
		{
		case OLEMISC_RECOMPOSEONRESIZE:
			return "recomposeonresize";
		case OLEMISC_ONLYICONIC:
			return "onlyiconic";
		case OLEMISC_INSERTNOTREPLACE:
			return "insertnotreplace";
		case OLEMISC_STATIC:
			return "static";
		case OLEMISC_CANTLINKINSIDE:
			return "cantlinkinside";
		case OLEMISC_CANLINKBYOLE1:
			return "canlinkbyole1";
		case OLEMISC_ISLINKOBJECT:
			return "islinkobject";
		case OLEMISC_INSIDEOUT:
			return "insideout";
		case OLEMISC_ACTIVATEWHENVISIBLE:
			return "activatewhenvisible";
		case OLEMISC_RENDERINGISDEVICEINDEPENDENT:
			return "renderingisdeviceindependent";
		case OLEMISC_INVISIBLEATRUNTIME:
			return "invisibleatruntime";
		case OLEMISC_ALWAYSRUN:
			return "alwaysrun";
		case OLEMISC_ACTSLIKEBUTTON:
			return "actslikebutton";
		case OLEMISC_ACTSLIKELABEL:
			return "actslikelabel";
		case OLEMISC_NOUIACTIVATE:
			return "nouiactivate";
		case OLEMISC_ALIGNABLE:
			return "alignable";
		case OLEMISC_SIMPLEFRAME:
			return "simpleframe";
		case OLEMISC_SETCLIENTSITEFIRST:
			return "setclientsitefirst";
		case OLEMISC_IMEMODE:
			return "imemode";
		case OLEMISC_IGNOREACTIVATEWHENVISIBLE:
			return "ignoreativatewhenvisible";
		case OLEMISC_WANTSTOMENUMERGE:
			return "wantstomenumerge";
		case OLEMISC_SUPPORTSMULTILEVELUNDO:
			return "supportsmultilevelundo";
		}
		return NULL;
	}

	void WriteMistStatus(LPCSTR format, LPCWSTR data)
	{
		if (int mask = StrToInt(data))
		{
			writer.write(format);
			int bit = 1;
			bool first = true;
			do if (LPCSTR text = GetMiscStatusText(bit & mask))
			{
				writer.write(&",%s"[first], text);
				first = false;
			} while ((bit <<= 1) != 0);
			writer.write(format + lstrlenA(format) - 1);
		}
	}

	HRESULT ExportCls()
	{
		TCHAR subkey[MAX_PATH];
		PathCombineW(subkey, appkey, L"Software\\Classes\\CLSID");
		HKEY hKey;
		if (LSTATUS r = RegCreateKeyW(HKEY_CURRENT_USER, subkey, &hKey))
			return CoGetError(r);
		DWORD i = 0; 
		WCHAR id[40];
		HKEY hKey2;
		while (0 == RegEnumKeyW(hKey, i++, id, _countof(id)) && 0 == RegOpenKeyW(hKey, id, &hKey2))
		{
			if (ManfredWasHere(hKey2, 1) == 0)
			{
				writer.write("\t\t<comClass clsid=\"%ls\"", id);
				WCHAR data[MAX_PATH];
				DWORD cb = sizeof data;
				if (0 == SHRegGetValueW(hKey2, L"InprocServer32", L"ThreadingModel", SRRF_RT_REG_SZ, NULL, data, &cb))
					writer.write(" threadingModel=\"%ls\"", data);
				HKEY hKey3;
				if (0 == RegOpenKeyW(hKey2, L"MiscStatus", &hKey3))
				{
					cb = sizeof data;
					if (0 == SHRegGetValueW(hKey3, NULL, NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatus=\"", data);
					cb = sizeof data;
					if (0 == SHRegGetValueW(hKey3, L"1", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatusContent=\"", data);
					cb = sizeof data;
					if (0 == SHRegGetValueW(hKey3, L"2", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatusThumbnail=\"", data);
					cb = sizeof data;
					if (0 == SHRegGetValueW(hKey3, L"4", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatusIcon=\"", data);
					cb = sizeof data;
					if (0 == SHRegGetValueW(hKey3, L"8", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatusDocPrint=\"", data);
					RegCloseKey(hKey3);
				}
				writer.write(" />\r\n");
				cb = sizeof data;
				if (0 == SHRegGetValueW(hKey2, L"TypeLib", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
				{
					PathCombineW(subkey, appkey, L"Software\\Classes\\TypeLib");
					PathAppendW(subkey, data);
					if (0 == RegCreateKeyW(HKEY_CURRENT_USER, subkey, &hKey3))
					{
						ManfredWasHere(hKey3, 1);
						RegCloseKey(hKey3);
					}
				}
			}
			RegCloseKey(hKey2);
		}
		RegCloseKey(hKey);
		return S_OK;
	}

	HRESULT ExportTlb()
	{ 
		TCHAR subkey[MAX_PATH];
		PathCombineW(subkey, appkey, L"Software\\Classes\\TypeLib");
		HKEY hKey;
		if (LSTATUS r = RegCreateKeyW(HKEY_CURRENT_USER, subkey, &hKey))
			return CoGetError(r);
		DWORD i = 0; 
		WCHAR id[40];
		HKEY hKey2;
		while (0 == RegEnumKeyW(hKey, i++, id, _countof(id)) && 0 == RegOpenKeyW(hKey, id, &hKey2))
		{
			if (ManfredWasHere(hKey2, 2) == 1)
			{
				DWORD i = 0; 
				WCHAR ver[40];
				while (0 == RegEnumKeyW(hKey2, i++, ver, _countof(ver)))
				{
					writer.write("\t\t<typelib tlbid=\"%ls\" version=\"%ls\" helpdir=\"\" />\r\n", id, ver);
				}
			}
			RegCloseKey(hKey2);
		}
		RegCloseKey(hKey);
		return S_OK;
	}

	static BOOL CALLBACK EnumResNameProcW(HMODULE, LPCWSTR, LPWSTR lpName, LONG_PTR lParam)
	{
		reinterpret_cast<Application *>(lParam)->ManifestName = lpName;
		return FALSE;
	}

	static BOOL CALLBACK EnumResLangProcW(HMODULE, LPCWSTR, LPCWSTR, LANGID wLanguage, LONG_PTR lParam)
	{
		reinterpret_cast<Application *>(lParam)->ManifestLang = wLanguage;
		return FALSE;
	}

	HRESULT BeginManifest(HMODULE module = NULL)
	{
		HRESULT hr = S_FALSE;
		EnumResourceNamesW(module, RT_MANIFEST, EnumResNameProcW, reinterpret_cast<LONG_PTR>(this));
		EnumResourceLanguagesW(module, RT_MANIFEST, ManifestName, EnumResLangProcW, reinterpret_cast<LONG_PTR>(this));
		if (const HRSRC res = FindResourceW(module, ManifestName, RT_MANIFEST))
		{
			if (const DWORD cb = SizeofResource(module, res))
			{
				if (const HGLOBAL global = LoadResource(module, res))
				{
					if (const char *const p = static_cast<const char *>(LockResource(global)))
					{
						// Detect indentation style
						int tabwidth = 0;
						const char *q = p + cb;
						do { } while (q > p && *--q != '<');
						do { } while (q > p && *--q != '<');
						while (q > p && *--q == ' ') ++tabwidth;
						static const char tag[] = "<file ";
						q = MemSearch(p, cb, tag, sizeof tag - 1);
						if (q == NULL)
						{
							static const char tag[] = "</assembly>";
							q = MemSearch(p, cb, tag, sizeof tag - 1);
						}
						else if (once)
						{
							q = NULL;
						}
						else
						{
							while (q > p && q[-1] != '>')
								--q;
							while (*q == '\r' || *q == '\n')
								++q;
						}
						if (q != NULL)
						{
							hr = CreateStreamOnHGlobal(NULL, TRUE, &writer);
							if (SUCCEEDED(hr))
							{
								writer.setTabWidth(tabwidth);
								writer.write(static_cast<DWORD>(q - p), p);
								hr = S_OK;
							}
						}
					}
				}
			}
		}
		return hr;
	}

	HRESULT BeginManifest(LPCWSTR path)
	{
		HRESULT hr = S_FALSE;
		if (HMODULE module = LoadLibraryExW(path, NULL, LOAD_LIBRARY_AS_DATAFILE))
		{
			hr = BeginManifest(module);
			FreeLibrary(module);
			if (hr == S_FALSE)
				hr = BeginManifest();
			if (hr == S_OK)
				update = BeginUpdateResourceW(path, FALSE);
		}
		return hr;
	}

	HRESULT AddFileToManifest(LPCWSTR path, LPCWSTR name)
	{
		HRESULT hr = S_FALSE;
		if (HMODULE module = LoadLibraryW(path))
		{
			if (FARPROC pfn = GetProcAddress(module, "DllRegisterServer"))
			{
				hr = reinterpret_cast<LPFNCANUNLOADNOW>(pfn)();
			}
			else
			{
				hr = CoGetError();
			}
			FreeLibrary(module);
		}
		else
		{
			hr = CoGetError();
		}

		if (hr == S_OK)
		{
			writer.write("\t<file name=\"%ls\">\r\n", name);
			ExportCls();
			ExportTlb();
			writer.write("\t</file>\r\n");
		}

		return hr;
	}

	HRESULT ManualAddFileToManifest(LPCWSTR name)
	{
		if (ini == NULL)
			return S_FALSE;

		WCHAR full[MAX_PATH];
		GetFullPathNameW(ini, _countof(full), full, NULL);

		WCHAR buffer[0x4000];
		if (GetPrivateProfileSectionW(name, buffer, _countof(buffer), full) == 0)
			return S_FALSE;

		writer.write("\t<file name=\"%ls\">\r\n", name);
		LPWSTR p = buffer;
		while (int len = lstrlenW(p))
		{
			writer.write("\t\t%ls\r\n", p);
			p += len + 1;
		}
		writer.write("\t</file>\r\n");

		return S_OK;
	}

	HRESULT EndManifest()
	{
		HRESULT hr;

		if (FAILED(hr = writer.write("</assembly>\r\n")))
			return hr;

		ULARGE_INTEGER pos;
		if (FAILED(hr = writer.tell(&pos)))
			return hr;

		HGLOBAL global;
		if (FAILED(hr = GetHGlobalFromStream(writer, &global)))
			return hr;

		if (LPVOID pv = GlobalLock(global))
		{
			UpdateResourceW(update, RT_MANIFEST, ManifestName, ManifestLang, pv, pos.LowPart);
			GlobalUnlock(global);
		}

		EndUpdateResourceW(update, FALSE);
		update = NULL;

		writer.close();
		return S_OK;
	}

	void UpdateFiles(LPCWSTR root, LPCWSTR folder)
	{
		WCHAR path[MAX_PATH];

		PathCombineW(path, root, folder);

		LPWSTR name = PathAddBackslashW(path);

		LPWSTR p = files;
		LPWSTR q = NULL;
		if (minus != 0)
		{
			p += lstrlenW(files) - minus;
			q = files;
		}

		do
		{
			if (p > q)
				p[-1] = L'\0';
			LPWSTR separator = StrChrW(p, L':');
			if (separator)
				*separator = L'\0';
			lstrcpyW(name, L"*.*");
			WIN32_FIND_DATAW fd;
			HANDLE h = FindFirstFileW(path, &fd);
			if (h != INVALID_HANDLE_VALUE)
			{
				do
				{
					lstrcpyW(name, fd.cFileName);
					if (PathMatchSpecW(name, p) && !PathMatchSpecW(name, q))
					{
						HRESULT hr = E_UNEXPECTED;
						if (LPCWSTR name = PathEatPrefix(path, root))
						{
							hr = ManualAddFileToManifest(name);
							if (hr == S_FALSE)
								hr = AddFileToManifest(path, name);
						}
						WriteTo<STD_OUTPUT_HANDLE>("[%08lX] %ls\r\n", hr, name);
					}
				} while (FindNextFileW(h, &fd));
				FindClose(h);
			}
			if (p > q)
				p[-1] = L';';
			p = separator ? separator + 1 : NULL;
			q = files;
		} while (p);
	}

	HRESULT UpdateFiles()
	{
		LPWSTR folder = StrChrW(target, L';');
		if (folder)
			*folder++ = L'\0';
		WCHAR full[MAX_PATH];
		GetFullPathNameW(target, _countof(full), full, NULL);
		HRESULT hr = BeginManifest(full);
		if (hr == S_OK)
		{
			LPCWSTR name = PathFindFileNameW(full);
			HRESULT hr = ManualAddFileToManifest(name);
			WriteTo<STD_OUTPUT_HANDLE>("[%08lX] %ls\r\n", hr, name);
			PathRemoveFileSpecW(full);
			SetDllDirectoryW(full);
			do
			{
				LPWSTR separator = StrChrW(folder, L';');
				if (separator)
					*separator = L'\0';
				UpdateFiles(full, folder);
				folder = separator ? separator + 1 : NULL;
			} while(folder);
			EndManifest();
		}
		return hr;
	}

public:
	Application()
	{
		SecureZeroMemory(this, sizeof *this);
	}

	HRESULT Run(const LPWSTR cmdline)
	{
		LPWSTR *parg = &target;
		LPWSTR p = cmdline;
		WCHAR sep = L';';
		do
		{
			LPWSTR q = PathGetArgsW(p);
			PathRemoveArgsW(p);
			PathUnquoteSpacesW(p);
			if (p == cmdline)
			{
				appname = PathFindFileName(p);
			}
			else if (*p != L'/')
			{
				if (int len = lstrlenW(p))
				{
					if (LPWSTR arg = *parg)
					{
						int cur = lstrlenW(arg);
						arg[cur] = sep;
						lstrcpyW(arg + cur + 1, p);
						if (minus != 0)
						{
							MemReverse(arg, cur);
							MemReverse(arg + cur + 1, len);
							MemReverse(arg, cur + 1 + len);
						}
					}
					else
					{
						*parg = p;
					}
				}
				sep = L';';
			}
			else if (lstrcmpiW(p + 1, L"files") == 0)
			{
				sep = L':';
				parg = &files;
			}
			else if (lstrcmpiW(p + 1, L"minus") == 0)
			{
				sep = L'|';
				minus = lstrlenW(files);
			}
			else if (lstrcmpiW(p + 1, L"ini") == 0)
			{
				parg = &ini;
			}
			else if (lstrcmpiW(p + 1, L"once") == 0)
				once = true;
			else
				parg = NULL;
			p = q + StrSpnW(q, L" \t\r\n");
		} while (*p != L'\0' && parg != NULL);

		if (target == NULL || parg == NULL)
		{
			WriteTo<STD_OUTPUT_HANDLE>(usage, appname);
			return E_FAIL;
		}

		WriteTo<STD_OUTPUT_HANDLE>("target = %ls\r\n", target);
		WriteTo<STD_OUTPUT_HANDLE>("files = %ls\r\n", files);

		Appartment appartment;
		HRESULT hr = appartment.GetHResult();
		if (SUCCEEDED(hr))
			hr = UpdateFiles();

		return hr;
	}
};

void mainCRTStartup()
{
	HRESULT hr = E_UNEXPECTED;
	if (BSTR cmdline = SysAllocString(GetCommandLineW()))
	{
		hr = Application().Run(cmdline);
		SysFreeString(cmdline);
	}
	ExitProcess(hr);
}
