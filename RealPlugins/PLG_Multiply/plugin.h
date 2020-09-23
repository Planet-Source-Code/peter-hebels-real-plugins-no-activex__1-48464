#ifdef MULTIPLY_EXPORTS
#define MULTIPLY_API __declspec(dllexport)
#else
#define MULTIPLY_API __declspec(dllimport)
#endif

//The functions this plugin exports.
MULTIPLY_API int WINAPI CalculateFunct(int FirstNum, int SecNum);
MULTIPLY_API int WINAPI GetFunct();
MULTIPLY_API void WINAPI ShowAbout(HWND hwnd);