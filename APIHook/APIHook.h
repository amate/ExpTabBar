#pragma once


#ifdef EXPORTS
#define EXPORT __declspec(dllexport)
#else
#define EXPORT __declspec(dllimport)
#endif /* TEST_EXPORTS */

extern "C" {

	EXPORT int InitHook();
	EXPORT int TermHook();


}