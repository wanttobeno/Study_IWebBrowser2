#pragma once
#include <Windows.h>
typedef HRESULT(*ApiToJavaScript)(DISPID, DISPPARAMS*, VARIANT*);
struct JavaScriptFunction
{
	wchar_t * jsFunctionName;
	DISPID jsFunctionID;
	ApiToJavaScript cppFunctionPtr;
};
