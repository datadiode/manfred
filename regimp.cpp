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
#include "reader.h"
//#include "wstdio.h"

static HKEY GetRootKeyFromName(LPCWSTR name)
{
	if (PathMatchSpecW(name, L"HKEY_CLASSES_ROOT;HKCR"))
		return HKEY_CLASSES_ROOT;
	if (PathMatchSpecW(name, L"HKEY_LOCAL_MACHINE;HKLM"))
		return HKEY_LOCAL_MACHINE;
	return NULL;
}

static LPWSTR EatPrefix(LPWSTR text, LPCWSTR prefix)
{
	int len = lstrlenW(prefix);
	return StrIsIntlEqualW(FALSE, text, prefix, len) ? text + len : NULL;
}

static LPWSTR SplitAssignment(LPWSTR line)
{
	LPWSTR p = line;
	LPWSTR q = p;
	LPWSTR r = NULL;
	bool quoted = false;
	do switch (*p = *q++)
	{
	case L'\\':
		if (quoted)
			*p = *q++;
		break;
	case L'=':
		if (!quoted)
		{
			r = q;
			*p = '\0';
		}
		break;
	case L'"':
		quoted = !quoted;
		break;
	} while (*p++ != L'\0');
	StrTrimW(line, L" \t\r\n");
	PathUnquoteSpacesW(line);
	return r;
}

static HKEY ProcessLine(HKEY key, LPWSTR line)
{
	if (LPWSTR p = EatPrefix(line, L"["))
	{
		RegCloseKey(key);
		key = NULL;
		LPWSTR q;
		while ((q = PathFindNextComponentW(p)) > p)
		{
			q[-1] = L'\0';
			if (HKEY tmp = key)
			{
				key = NULL;
				RegCreateKeyW(tmp, p, &key);
				if (tmp < HKEY_CLASSES_ROOT)
					RegCloseKey(tmp);
			}
			else
			{
				key = GetRootKeyFromName(p);
			}
			if (key == NULL)
				break;
			//WriteTo<STD_OUTPUT_HANDLE>("RegCreateKeyW(%ls)\n", p);
			p = q;
		}
	}
	else if (LPWSTR val = SplitAssignment(line))
	{
		SplitAssignment(val);
		//WriteTo<STD_OUTPUT_HANDLE>("L = %ls\n", line);
		//WriteTo<STD_OUTPUT_HANDLE>("R = %ls\n", val);
		const LPCWSTR name = lstrcmpW(line, L"@") != 0 ? line : NULL;
		if (LPWSTR p = EatPrefix(val, L"dword:"))
		{
			*--p = L'x';
			*--p = L'0';
			int iVal;
			if (StrToIntExW(p, STIF_SUPPORT_HEX, &iVal))
			{
				RegSetValueExW(key, name, 0, REG_DWORD, reinterpret_cast<BYTE *>(&iVal), sizeof iVal);
			}
		}
		else if (LPWSTR p = EatPrefix(val, L"hex:"))
		{
			DWORD cb = 0;
			BYTE b[8200];
			do
			{
				p += StrSpnW(p, L" \t\r\n");
				LPWSTR q = StrChrW(p, L',');
				*--p = L'x';
				*--p = L'0';
				int iVal;
				if (!StrToIntExW(p, STIF_SUPPORT_HEX, &iVal))
					break;
				b[cb++] = static_cast<BYTE>(iVal);
				p = q ? q + 1 : NULL;
			} while (p && cb < sizeof b);
			if (p == NULL) // everything parsed with no issues
			{
				RegSetValueExW(key, name, 0, REG_BINARY, b, cb);
			}
		}
		else
		{
			RegSetValueExW(key, name, 0, REG_SZ, reinterpret_cast<BYTE *>(val), (lstrlenW(val) + 1) * sizeof(WCHAR));
		}
	}
	return key;
}

HRESULT ImportRegFile(LPCWSTR path)
{
	Reader reader;
	HRESULT hr = SHCreateStreamOnFileEx(path,
		STGM_READ | STGM_SHARE_DENY_NONE,
		FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &reader);
	if (FAILED(hr))
		return hr;
	Reader::Encoding encoding = reader.readBom();
	BYTE eol = reader.allocCtype("\n");
	HKEY key = NULL;
	ULONG len = 0;
	if (encoding == Reader::UCS2LE)
	{
		LPWSTR line = NULL;
		while (reader.readLine(&line, eol, len))
		{
			//WriteTo<STD_OUTPUT_HANDLE>("%ls", line);
			StrTrimW(line, L" \t\r\n");
			if (StrTrimW(line, L"\\"))
			{
				// continuation line follows
				len = lstrlenW(line) * sizeof(WCHAR);
				continue;
			}
			// line complete
			len = 0;
			key = ProcessLine(key, line);
		}
		CoTaskMemFree(line);
	}
	else if (encoding == Reader::ANSI)
	{
		LPSTR line = NULL;
		while (reader.readLine(&line, eol, len))
		{
			//WriteTo<STD_OUTPUT_HANDLE>(line);
			StrTrimA(line, " \t\r\n");
			if (StrTrimA(line, "\\"))
			{
				// continuation line follows
				len = lstrlenA(line);
				continue;
			}
			// line complete
			len = 0;
			LPWSTR pwsz = NULL;
			if (FAILED(hr = SHStrDupA(line, &pwsz)))
				break;
			key = ProcessLine(key, pwsz);
			CoTaskMemFree(pwsz);
		}
		CoTaskMemFree(line);
	}
	else
	{
		hr = E_INVALIDARG;
	}
	if (key < HKEY_CLASSES_ROOT)
		RegCloseKey(key);
	return hr;
}
