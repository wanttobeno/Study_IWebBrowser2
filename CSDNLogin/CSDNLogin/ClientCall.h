#pragma once
#include <OAIdl.h>
class ClientCall :
	public IDispatch
{
public:
	ClientCall();
	~ClientCall();

	HRESULT _stdcall GetIDsOfNames(
		REFIID riid,
		OLECHAR FAR* FAR* rgszNames,
		unsigned int cNames,
		LCID lcid,
		DISPID FAR* rgDispId
		)
	{
		if (lstrcmp(rgszNames[0], L"CppCall") == 0)
		{
			//网页调用window.external.CppCall时，会调用这个方法获取CppCall的ID
			*rgDispId = 100;
		}
		return S_OK;
	}
	HRESULT _stdcall Invoke(
		DISPID dispIdMember,
		REFIID riid,
		LCID lcid,
		WORD wFlags,
		DISPPARAMS* pDispParams,
		VARIANT* pVarResult,
		EXCEPINFO* pExcepInfo,
		unsigned int* puArgErr
		)
	{
		if (dispIdMember == 100)
		{
			//网页调用CppCall时，或根据获取到的ID调用Invoke方法
			//CppCall(pDispParams->rgvarg[0].intVal);
		}
		return S_OK;
	}

};

