#if defined(_M_MRX000) || defined(_M_PPC) || defined(_M_ALPHA) 
#pragma intrinsic(Myalloca) 
#endif 
__declspec(naked)  _fastcall Myalloca(size_t size)
{
	__asm{
		push        ecx
			cmp         eax,1000h
			lea			ecx,[esp+8]
			jb          _lastpage
_probepages:
		sub         ecx,1000h
			sub         eax,1000h
			test        dword ptr [ecx],eax
			cmp         eax,1000h
			jae         _probepages
_lastpage:
		sub         ecx,eax
			mov         eax,esp
			test        dword ptr [ecx],eax
			mov         esp,ecx
			mov         ecx,dword ptr [eax]
			mov         eax,dword ptr [eax+4]
			push        eax
			ret
	}
	
}