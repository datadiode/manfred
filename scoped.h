/*
[The MIT license]

Copyright (c) 2013 Jochen Neubeck

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

template<class T>
class Scoped : public T
{
public:
	Scoped();
	~Scoped();
};

inline Scoped<VARIANT>::Scoped()
{
	VariantInit(this);
}

inline Scoped<VARIANT>::~Scoped()
{
	VariantClear(this);
}

template<class T, class U>
class Scoped2
{
protected:
	T scoped;
private:
	template<class U> void Free();
	template<> void Free<enum eLOCAL>()
	{
		LocalFree(scoped);
	}
	template<> void Free<enum eTASKMEM>()
	{
		CoTaskMemFree(scoped);
	}
	template<> void Free<enum eHCRYPTPROV>()
	{
		if (scoped != NULL)
			CryptReleaseContext(scoped, 0);
	}
	template<> void Free<enum eHCRYPTKEY>()
	{
		if (scoped != NULL)
			CryptDestroyKey(scoped);
	}
	template<> void Free<enum eHCERTSTORE>()
	{
		if (scoped != NULL)
			CertCloseStore(scoped, 0);
	}
	template<> void Free<enum eHKEY>()
	{
		if (scoped != NULL)
			RegCloseKey(scoped);
	}
	template<> void Free<enum eHWND>()
	{
		if (scoped != NULL)
			DestroyWindow(scoped);
	}
	template<> void Free<enum eFINDFILE>()
	{
		if (scoped != INVALID_HANDLE_VALUE)
			FindClose(scoped);
	}
	template<> void Free<enum eBSTR>()
	{
		SysFreeString(scoped);
	}
	template<> void Free<enum eOBJECT>()
	{
		if (scoped != NULL)
			scoped->Release();
	}
	template<> void Free<enum eSCALAR>()
	{
		delete scoped;
	}
	template<> void Free<enum eVECTOR>()
	{
		delete[] scoped;
	}
public:
	operator T &() { return scoped; }
	T operator->() const { return scoped; }
	T *operator &() { return &scoped; }
	Scoped2(T scoped = 0) : scoped(scoped) { }
	void operator=(T scoped) { this->scoped = scoped; }
	~Scoped2() { Free<U>(); }
};
