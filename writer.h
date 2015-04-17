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

class Writer
{
public:
	Writer(): pstm(NULL), tabwidth(0) { }
	~Writer() { close(); }

	operator IStream *() { return pstm; }
	IStream **operator&() { return &pstm; }

	void setTabWidth(int n) { tabwidth = n; }

	template<class FORMAT>
	HRESULT write(FORMAT *format, ...)
	{
		format = indent(format);
		char buffer[1024];
		DWORD cb = wvsprintfA(buffer, format, va_list(&format + 1));
		return write(cb, buffer);
	}
	HRESULT write(LPCSTR buffer)
	{
		buffer = indent(buffer);
		DWORD cb = lstrlenA(buffer);
		return write(cb, buffer);
	}
	HRESULT write(DWORD cb, LPCSTR buffer)
	{
		if (pstm == NULL)
			return E_POINTER;
		return pstm->Write(buffer, cb, NULL);
	}
	HRESULT tell(ULARGE_INTEGER *out)
	{
		if (pstm == NULL)
			return E_POINTER;
		LARGE_INTEGER in = { 0, 0 };
		return pstm->Seek(in, STREAM_SEEK_CUR, out);
	}
	HRESULT close()
	{
		if (pstm == NULL)
			return E_POINTER;
		pstm->Release();
		pstm = NULL;
		return S_OK;
	}

private:
	LPCSTR indent(LPCSTR format)
	{
		static const char buffer[8] = { ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ' };
		if (int cb = tabwidth < 8 ? tabwidth : 8)
		{
			while (*format == '\t')
			{
				write(cb, buffer);
				++format;
			}
		}
		return format;
	}

	IStream *pstm;
	int tabwidth;
};
