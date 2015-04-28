/*
[The MIT license]

Copyright (c) 2012 Jochen Neubeck

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

class Reader
{
private:
	IStream *pstm;
	ULONG index;
	ULONG ahead;
	BYTE chunk[256];
	BYTE ctype[256];

	template<typename T>
	BYTE *sip(T *p, const T *q, BYTE opAnd, BYTE opXor, ULONG n)
	{
		while (n != 0)
		{
			const T c = *q++;
			*p = c; // caller may want to include the token terminator
			const BYTE b = static_cast<BYTE>(c);
			if ((c == static_cast<T>(b) ? ctype[b] : 0) & opAnd ^ opXor)
				return reinterpret_cast<BYTE *>(p);
			++p;
			n -= sizeof(T);
		}
		return NULL;
	}

public:
	enum Encoding { ANSI = 0x00, UTF8 = 0xEF, UCS2BE = 0xFE, UCS2LE = 0xFF };

	Reader(): pstm(NULL), index(0), ahead(0)
	{
		SecureZeroMemory(ctype, sizeof ctype);
		ctype[0] = 1;
	}

	~Reader() { close(); }

	operator IStream *() { return pstm; }
	IStream **operator&() { return &pstm; }

	template<typename T>
	ULONG slurp(T **ps, BYTE opAnd, BYTE opXor, ULONG n = 0, ULONG t = 0)
	{
		BYTE *s = reinterpret_cast<BYTE *>(*ps);
		do
		{
			ULONG i = n;
			n += ahead;
			s = (BYTE *)CoTaskMemRealloc(s, n + sizeof(T));
			BYTE *lower = s + i;
			if (BYTE *upper = sip(reinterpret_cast<T *>(lower),
				reinterpret_cast<T *>(chunk + index), opAnd, opXor, ahead))
			{
				upper += t; // include the token terminator if desired
				n = static_cast<ULONG>(upper - lower);
				index += n;
				ahead -= n;
				n = static_cast<ULONG>(upper - s);
				s = (BYTE *)CoTaskMemRealloc(s, n + sizeof(T));
				break;
			}
			index = ahead = 0;
			if (pstm)
				pstm->Read(chunk, sizeof chunk, &ahead);
		} while (ahead >= sizeof(T));
		*reinterpret_cast<T *>(s + n) = 0;
		*ps = reinterpret_cast<T *>(s);
		return n;
	}

	template<typename T>
	ULONG readLine(T **ps, BYTE op, ULONG n = 0)
	{
		return slurp(ps, op, 0, n, sizeof(T));
	}

	BYTE allocCtype(const char *q)
	{
		BYTE cookie = ctype[0];
		ctype[0] <<= 1;
		while (BYTE c = *q++)
		{
			ctype[c] |= cookie;
		}
		return cookie;
	}

	Encoding readBom()
	{
		if (pstm)
			pstm->Read(chunk, sizeof chunk, &ahead);
		if (ahead >= 2)
		{
			if (chunk[0] == 0xFF && chunk[1] == 0xFE || chunk[0] == 0xFE && chunk[1] == 0xFF)
				index = 2;
			else if (ahead >= 3 && chunk[0] == 0xEF && chunk[1] == 0xBB && chunk[2] == 0xBF)
				index = 3;
		}
		ahead -= index;
		return index ? static_cast<Encoding>(chunk[0]) : ANSI;
	}

	HRESULT close()
	{
		if (pstm == NULL)
			return E_POINTER;
		pstm->Release();
		pstm = NULL;
		return S_OK;
	}
};
