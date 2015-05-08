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

class MultiMap
{
	const HMENU menu;
public:
	static const WCHAR separator = L':';
	MultiMap()
	: menu(CreatePopupMenu())
	{
	}
	~MultiMap()
	{
		Clear();
		DestroyMenu(menu);
	}
	void Add(LPCWSTR key, LPCWSTR val)
	{
		if (ATOM atom = AddAtomW(key))
		{
			UINT total = lstrlenW(val);
			if (UINT len = GetMenuStringW(menu, atom, NULL, 0, MF_BYCOMMAND))
			{
				total += len + 1;
				if (BSTR bstr = SysAllocStringLen(NULL, total))
				{
					GetMenuStringW(menu, atom, bstr, total, MF_BYCOMMAND);
					bstr[len++] = separator;
					lstrcpyW(bstr + len, val);
					ModifyMenuW(menu, atom, MF_BYCOMMAND, atom, bstr);
					SysFreeString(bstr);
				}
			}
			else
			{
				AppendMenuW(menu, MF_STRING, atom, val);
			}
		}
	}
	int GetItemCount() const
	{
		return GetMenuItemCount(menu);
	}
	BSTR GetItem(LPWSTR key, int i) const
	{
		ATOM atom = GetMenuItemID(menu, i);
		UINT len = GetMenuStringW(menu, i, NULL, 0, MF_BYPOSITION);
		BSTR bstr = SysAllocStringLen(NULL, len);
		if (bstr != NULL)
		{
			GetMenuStringW(menu, i, bstr, len + 1, MF_BYPOSITION);
			if (key != NULL)
			{
				GetAtomNameW(atom, key, MAX_PATH);
			}
			else do
			{
				DeleteAtom(atom);
			} while ((key = StrRChrW(bstr, key, separator)) != NULL);
		}
		return bstr;
	}
	void Clear()
	{
		int i = GetItemCount();
		while (i > 0)
		{
			--i;
			if (BSTR bstr = GetItem(NULL, i))
			{
				SysFreeString(bstr);
			}
			DeleteMenu(menu, i, MF_BYPOSITION);
		}
	}
};
