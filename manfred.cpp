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
#include "regimp.h"
#include "multimap.h"
#include "miscutil.h"

#define OUTPUT STD_ERROR_HANDLE

static const char usage[] =
	"Manifest Resource Editor v1.07\r\n"
	"\r\n"
	"Usage:\r\n"
	"\r\n"
	"%ls <target> ... [ /once ] [ /ini ... ] [ /files ... ] [ /minus ... ]\r\n"
	"\r\n"
	"<target>  may be followed by a list of subfolders to search\r\n"
	"/once     causes update of manifest to occur only when no file tags exist yet\r\n"
	"/never    causes update of manifest to occur never; useful with /rgs option\r\n"
	"/ini      specifies an ini file from which to merge content into the manifest\r\n"
	"/rgs      specifies an rgs file to write results to for bulk registration\r\n"
	"/files    specifies file inclusion patterns; may occur repeatedly\r\n"
	"/minus    specifies file exclusion patterns; may occur only once\r\n"
	"/vld{+|-} enables or disables Visual Leak Detector\r\n"
	"\r\n";

static const WCHAR appkey[] =
	L"Software\\{CF5F8904-9192-4169-AD65-6360250946CB}";

static const GUID CLSID_Registrar =
	{ 0x44ec053a, 0x400f, 0x11d0, { 0x9d, 0xcd, 0x00, 0xa0, 0xc9, 0x03, 0x91, 0xd3 } };

static LPCWSTR PathEatPrefix(LPCWSTR path, LPCWSTR root)
{
	int prefix = PathCommonPrefixW(path, root, NULL);
	return root[prefix] != L'\0' ? NULL : *(path += prefix) != L'\0' ? path + 1 : path;
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

class ValueBuffer
{
	// Registrar tokens are limited to 4096 characters, which in case of
	// string values may expand to 8194 bytes including null termination.
	static const DWORD size = 8200; // a few bytes more than required
public:
	DWORD type;
	BufferCapacity<size> cb;
	union
	{
		BYTE b[size];
		WCHAR s[size / sizeof(WCHAR)];
		DWORD d;
	};
};

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

class Application: ZeroInit<Application>
{
	LPWSTR ManifestName;
	LANGID ManifestLang;
	HANDLE update;
	LPCWSTR appname;
	enum { always, once, never } option;
	LPCSTR vldoption;
	LPWSTR target;
	LPWSTR ini;
	LPWSTR rgs;
	LPWSTR files;
	UINT minus;
	MultiMap clsmm;
	MultiMap progmm;
	MultiMap tlbmm;
	Writer writer;
	ValueBuffer vb;
	WCHAR root[MAX_PATH];

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

	HRESULT ExportCls(LPCWSTR name)
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
				clsmm.Add(id, name);
				writer.write("\t\t<comClass clsid=\"%ls\"", id);
				WCHAR data[MAX_PATH];
				BufferCapacity<sizeof data> cb;
				if (0 == SHRegGetValueW(hKey2, L"VersionIndependentProgID", NULL, SRRF_RT_REG_SZ, NULL, data, &cb) ||
					0 == SHRegGetValueW(hKey2, L"ProgID", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
				{
					progmm.Add(data, name);
					writer.write(" progid=\"%ls\"", data);
				}
				if (0 == SHRegGetValueW(hKey2, L"InprocServer32", L"ThreadingModel", SRRF_RT_REG_SZ, NULL, data, &cb))
					writer.write(" threadingModel=\"%ls\"", data);
				HKEY hKey3;
				if (0 == RegOpenKeyW(hKey2, L"MiscStatus", &hKey3))
				{
					if (0 == SHRegGetValueW(hKey3, NULL, NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatus=\"", data);
					if (0 == SHRegGetValueW(hKey3, L"1", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatusContent=\"", data);
					if (0 == SHRegGetValueW(hKey3, L"2", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatusThumbnail=\"", data);
					if (0 == SHRegGetValueW(hKey3, L"4", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatusIcon=\"", data);
					if (0 == SHRegGetValueW(hKey3, L"8", NULL, SRRF_RT_REG_SZ, NULL, data, &cb))
						WriteMistStatus(" miscStatusDocPrint=\"", data);
					RegCloseKey(hKey3);
				}
				writer.write(" />\r\n");
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

	HRESULT ExportTlb(LPCWSTR name)
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
				tlbmm.Add(id, name);
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

	HRESULT BeginManifest(HMODULE module)
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
						else if (option == once)
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

	HRESULT BeginManifest()
	{
		HRESULT hr = S_FALSE;
		if (HMODULE module = LoadLibraryExW(root, NULL, LOAD_LIBRARY_AS_DATAFILE))
		{
			hr = BeginManifest(module);
			FreeLibrary(module);
			if (hr == S_FALSE)
				hr = BeginManifest(NULL);
			if (hr == S_OK)
				update = BeginUpdateResourceW(root, FALSE);
		}
		return hr;
	}

	HRESULT DllRegisterServer(LPCWSTR path)
	{
		HRESULT hr = S_FALSE;
		if (PathMatchSpecW(path, L"*.REG"))
		{
			hr = ImportRegFile(path);
		}
		else if (HMODULE module = LoadLibraryW(path))
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
		return hr;
	}

	HRESULT AddFileToManifest(LPCWSTR name)
	{
		writer.write("\t<file name=\"%ls\">\r\n", name);
		ExportCls(name);
		ExportTlb(name);
		writer.write("\t</file>\r\n");
		return S_OK;
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

	void UpdateFiles(LPCWSTR folder)
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
							hr = DllRegisterServer(path);
							if (hr == S_OK && option != never)
							{
								hr = ManualAddFileToManifest(name);
								if (hr == S_FALSE)
									hr = AddFileToManifest(name);
							}
						}
						WriteTo<OUTPUT>("[%08lX] %ls\r\n", hr, name);
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

	int ReportConflicts(const MultiMap &mm, LPCSTR format)
	{
		int count = 0;
		int n = mm.GetItemCount();
		for (int i = 0; i < n; ++i)
		{
			WCHAR key[MAX_PATH];
			if (BSTR bstr = mm.GetItem(key, i))
			{
				if (StrChrW(bstr, mm.separator) != NULL)
				{
					WriteTo<OUTPUT>(format, key, bstr);
					++count;
				}
				SysFreeString(bstr);
			}
		}
		return count;
	}

	HRESULT UpdateFiles()
	{
		LPWSTR folder = StrChrW(target, L';');
		if (folder)
			*folder++ = L'\0';
		GetFullPathNameW(target, _countof(root), root, &target);
		HRESULT hr = S_OK;
		if (option != never)
			hr = BeginManifest();
		target[-1] = L'\0';
		SetDllDirectoryW(root);
		if (hr == S_OK)
		{
			// Keep Visual Leak Detector resident throughout process lifetime
#ifdef _WIN64
			static const WCHAR vldlib[] = L"vld_x64";
#else
			static const WCHAR vldlib[] = L"vld_x86";
#endif
			if (HMODULE h = LoadLibraryW(vldlib))
				if (FARPROC f = GetProcAddress(h, vldoption))
					reinterpret_cast<void(*)()>(f)();

			if (option != never)
			{
				hr = ManualAddFileToManifest(target);
				WriteTo<OUTPUT>("[%08lX] %ls\r\n", hr, target);
			}
			do
			{
				LPWSTR separator = StrChrW(folder, L';');
				if (separator)
					*separator = L'\0';
				UpdateFiles(folder);
				folder = separator ? separator + 1 : NULL;
			} while(folder);
			if (option != never)
				hr = EndManifest();

			WriteTo<OUTPUT>("\r\nIssues:");
			int count = 0;
			count += ReportConflicts(clsmm, "\r\nclsid %ls conflicts between:\r\n<%ls>");
			count += ReportConflicts(progmm, "\r\nprogid %ls conflicts between:\r\n<%ls>");
			count += ReportConflicts(tlbmm, "\r\ntlbid %ls conflicts between:\r\n<%ls>");
			if (count == 0)
				WriteTo<OUTPUT>(" none");
			WriteTo<OUTPUT>("\r\n");
		}
		return hr;
	}

	static HRESULT MayForceRemove(HKEY key, LPCWSTR name)
	{
		if (PathMatchSpecW(name, L"{*}"))
			return S_OK;
		if (key != HKEY_CLASSES_ROOT)
			return S_FALSE;
		CLSID clsid;
		return CLSIDFromProgID(name, &clsid);
	}

	HRESULT WriteValue()
	{
		HRESULT hr = S_OK;
		switch (vb.type)
		{
		case REG_SZ:
		case REG_EXPAND_SZ:
			if (LPCWSTR name = PathEatPrefix(vb.s, root))
				hr = writer.write(" = s '%%ROOT%%%ls'", name);
			else
				hr = writer.write(" = s '%ls'", vb.s);
			break;
		case REG_DWORD:
			hr = writer.write(" = d '%lu'", vb.d);
			break;
		case REG_BINARY:
			hr = writer.write(" = b '");
			for (DWORD i = 0; i < vb.cb; ++i)
				hr = writer.write("%02X", vb.b[i]);
			hr = writer.write("'");
			break;
		}
		return hr;
	}

	HRESULT WriteValue(HKEY key, LPCWSTR name)
	{
		HRESULT hr = S_OK;
		if (0 == RegQueryValueExW(key, name, NULL, &vb.type, vb.b, &vb.cb))
		{
			hr = WriteValue();
		}
		return hr;
	}

	HRESULT WriteScript(HKEY outerkey, LPCWSTR outername, int depth = 0, LPCSTR format = "%ls")
	{
		static const char tabs[] = "\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t";
		if (depth > sizeof tabs - 1)
			depth = sizeof tabs - 1;
		writer.write(depth, tabs);
		writer.write(format, outername);
		WriteValue(outerkey, NULL);
		writer.write("\r\n");
		writer.write(depth, tabs);
		writer.write("{\r\n");
		DWORD i = 0;
		HKEY key;
		WCHAR name[MAX_PATH];
		while (0 == RegEnumKeyW(outerkey, i++, name, _countof(name)) && 0 == RegOpenKeyW(outerkey, name, &key))
		{
			HRESULT hr = MayForceRemove(outerkey, name);
			WriteScript(key, name, depth + 1,
				hr == S_FALSE ? "'%ls'" : hr == S_OK ? "ForceRemove '%ls'" : "NoRemove '%ls'");
			RegCloseKey(key);
		}
		BufferCapacity<_countof(name)> namelen;
		i = 0;
		while (0 == RegEnumValueW(outerkey, i++, name, &namelen, NULL, &vb.type, vb.b, &vb.cb))
		{
			if (namelen == 0 || lstrcmpW(name, L"ManfredWasHere") == 0)
				continue;
			// write only values of supported types
			if ((1 << vb.type) & (1 << REG_SZ | 1 << REG_EXPAND_SZ | 1 << REG_DWORD | 1 << REG_BINARY))
			{
				writer.write(depth + 1, tabs);
				writer.write("val '%ls'", name);
				WriteValue();
				writer.write("\r\n");
			}
		}
		writer.write(depth, tabs);
		writer.write("}\r\n");
		return S_OK;
	}

	HRESULT WriteScript()
	{
		writer.setTabWidth(0);
		HRESULT hr = SHCreateStreamOnFileEx(rgs,
			STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE,
			FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &writer);
		if (SUCCEEDED(hr))
		{
			WriteScript(HKEY_CLASSES_ROOT, L"HKCR");
			writer.close();
		}
		return S_OK;
	}

public:
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
				appname = PathFindFileNameW(p);
			}
			else if (*p != L'/')
			{
				if (const LPWSTR arg = *parg)
				{
					if (sep == '\0')
						break;
					const int len = lstrlenW(p);
					const int cur = lstrlenW(arg);
					arg[cur] = sep;
					lstrcpyW(arg + cur + 1, p);
					if (minus != 0)
					{
						MemReverse(arg, cur);
						MemReverse(arg + cur + 1, len);
						MemReverse(arg, cur + 1 + len);
					}
					sep = L';';
				}
				else
				{
					*parg = p;
				}
			}
			else if (lstrcmpiW(p + 1, L"files") == 0)
			{
				sep = L':';
				parg = &files;
			}
			else if (lstrcmpiW(p + 1, L"minus") == 0)
			{
				if (parg != &files)
					break;
				sep = L'|';
				minus = lstrlenW(files);
			}
			else if (lstrcmpiW(p + 1, L"ini") == 0)
			{
				sep = L'\0';
				parg = &ini;
			}
			else if (lstrcmpiW(p + 1, L"rgs") == 0)
			{
				sep = L'\0';
				parg = &rgs;
			}
			else if (lstrcmpiW(p + 1, L"once") == 0)
				option = once;
			else if (lstrcmpiW(p + 1, L"never") == 0)
				option = never;
			else if (lstrcmpiW(p + 1, L"vld+") == 0)
				vldoption = "VLDEnable";
			else if (lstrcmpiW(p + 1, L"vld-") == 0)
				vldoption = "VLDDisable";
			else
				break;
			p = q + StrSpnW(q, L" \t\r\n");
		} while (*p != L'\0');

		// If no target was specified, or unconsumed arguments exist, give up.
		if (target == NULL || *p != L'\0')
		{
			WriteTo<OUTPUT>(usage, appname);
			return E_FAIL;
		}

		WriteTo<OUTPUT>("target = %ls\r\n", target);
		WriteTo<OUTPUT>("files = %ls\r\n", files);

		// Fail gracefully if not running as administrator
		BOOL fIsRunAsAdmin = FALSE;
		Scoped2<PSID, ePSID> pAdministratorsGroup;
		SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;
		if (!AllocateAndInitializeSid(&NtAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID,
			DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &pAdministratorsGroup))
		{
			return CoGetError();
		}
		if (!CheckTokenMembership(NULL, pAdministratorsGroup, &fIsRunAsAdmin))
		{
			return CoGetError();
		}
		if (!fIsRunAsAdmin)
		{
			WriteTo<OUTPUT>("Failing gracefully due to lack of permissions\r\n");
			return S_OK;
		}

		Appartment appartment;
		HRESULT hr = appartment.GetHResult();
		if (SUCCEEDED(hr))
		{
			hr = UpdateFiles();
			if (hr == S_OK && rgs != NULL)
			{
				hr = WriteScript();
			}
		}
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
