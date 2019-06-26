
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the LIBXLS_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// LIBXLS_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef LIBXLS_EXPORTS
#define LIBXLS_API __declspec(dllexport)
#else
#define LIBXLS_API __declspec(dllimport)
#endif

// This class is exported from the libxls.dll
class LIBXLS_API CLibxls {
public:
	CLibxls(void);
	// TODO: add your methods here.
};

extern LIBXLS_API int nLibxls;

LIBXLS_API int fnLibxls(void);

