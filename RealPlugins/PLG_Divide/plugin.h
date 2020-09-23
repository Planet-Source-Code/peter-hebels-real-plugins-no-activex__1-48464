#ifdef DIVIDE_EXPORTS
#define DIVIDE_API __declspec(dllexport)
#else
#define DIVIDE_API __declspec(dllimport)
#endif

//The functions this plugin exports.
DIVIDE_API int WINAPI CalculateFunct(int FirstNum, int SecNum);
DIVIDE_API int WINAPI GetFunct();
DIVIDE_API void WINAPI ShowAbout(HWND hwnd);