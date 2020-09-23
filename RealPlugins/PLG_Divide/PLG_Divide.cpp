// PLG_Divide.cpp : Defines the entry point for the DLL application.
// This example is written by Peter Hebels, website: http://www.phsoft.nl
// The author of this code cannot be held responsible for any damages may caused by the 
// use of this code.

#include "stdafx.h"
#include "plugin.h"
#include <windows.h>

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    return TRUE;
}

//This is the calculate function, it calculates the nubers passed trough FirstNum and SecNum.
DIVIDE_API int WINAPI CalculateFunct(int FirstNum, int SecNum)
{ 
    int TempNum;

    TempNum = FirstNum / SecNum;
	return TempNum;
}

//This function sends the function id to the calling application, 01 stands for Multiply
//and 02 stands for Divide.
DIVIDE_API int WINAPI GetFunct()
{ 
    return 02;
}

//When called this function shows some information about the plugin.
DIVIDE_API void WINAPI ShowAbout(HWND hwnd)
{ 
    MessageBox(hwnd, "Divide plugin, written by Peter Hebels\n\nWebsite: http://www.phsoft.nl", "About Plugin", MB_OK);
}