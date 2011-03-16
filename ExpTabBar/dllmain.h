// dllmain.h : モジュール クラスの宣言

class CExpTabBarModule : public ATL::CAtlDllModuleT< CExpTabBarModule >
{
public :
	DECLARE_LIBID(LIBID_ExpTabBarLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_EXPTABBAR, "{83E7E6AB-5D4A-468D-B354-61E0CED913EA}")
};

extern class CExpTabBarModule _AtlModule;
