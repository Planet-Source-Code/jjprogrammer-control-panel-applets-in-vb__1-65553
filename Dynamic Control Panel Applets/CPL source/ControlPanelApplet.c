/*

    Dynamic Control Panel Applet

*/
#include <windows.h>
#include <cpl.h>
#include <stdlib.h>
#include <stdio.h> 

#define MAX_APPLETS 100 //seems plenty to me
void LoadInfo();

typedef struct CPLAPPLICATION
{
	HICON hIcon; 
	TCHAR szFileExe[MAX_PATH]; 
    TCHAR szName[32]; 
    TCHAR szInfo[64]; 
    TCHAR szHelpFile[128]; 
	NEWCPLINFO *lpNewWinQuire;
} CplApp;

CplApp CPAppList[ MAX_APPLETS ];

int NumberOfApplets;
HANDLE m_hInstance;

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
	m_hInstance = hModule;
    return TRUE;
}

LRESULT APIENTRY CPlApplet(HWND hwnd, UINT msg, LPARAM lp1, LPARAM lp2)
{
    NEWCPLINFO *lpNewWinQuire;
	LPCPLINFO lpCPlInfo;
	int nI;
	int i = (int) lp1;      // number of applets

	switch( msg )
	{
        case CPL_INIT :      // first message, sent once
			LoadInfo();

			return TRUE;
			break;
        case CPL_GETCOUNT:	 // second message, sent once
			return NumberOfApplets;
			break;

        case CPL_INQUIRE:    // third message, sent once per application
			lpCPlInfo = (LPCPLINFO) lp2;
			lpCPlInfo->lData = 0;
			lpCPlInfo->idIcon = CPL_DYNAMIC_RES;
			lpCPlInfo->idName = CPL_DYNAMIC_RES;
			lpCPlInfo->idInfo = CPL_DYNAMIC_RES;
			break;

        case CPL_NEWINQUIRE:

			lpNewWinQuire = ((NEWCPLINFO*)lp2);

			lpNewWinQuire->dwSize = sizeof(NEWCPLINFO);
			lpNewWinQuire->dwFlags = 0;
			lpNewWinQuire->dwHelpContext = 0;
			lpNewWinQuire->lData = (LONG)m_hInstance;
			lpNewWinQuire->hIcon =  CPAppList[i].hIcon;

			strcpy(lpNewWinQuire->szName,	  CPAppList[i].szName );
			strcpy(lpNewWinQuire->szInfo,	  CPAppList[i].szInfo );
			strcpy(lpNewWinQuire->szHelpFile, CPAppList[i].szHelpFile);
			break;

        case CPL_DBLCLK:     // application icon double-clicked
			WinExec( CPAppList[i].szFileExe, SW_NORMAL );
            break;

        case CPL_STOP:       // sent once per application before CPL_EXIT
		{
			for(nI = 0; nI < MAX_APPLETS; nI++ )
				 DestroyIcon( CPAppList[nI].hIcon );
			break;
		}

        case CPL_EXIT:      // sent once before FreeLibrary is called
            break;
        default:
            break;
    }
    return 0;
}

void LoadInfo()
{   int nI;
	char cFileName[MAX_PATH];
	if (  GetModuleFileName( (HMODULE)m_hInstance, cFileName, MAX_PATH ) > 0 )
	{
		char drive[_MAX_DRIVE],  dir[_MAX_DIR],
			 fname[_MAX_FNAME],  ext[_MAX_EXT],
			 sLoadFileINI[MAX_PATH];
		char sSectionName[32];
		char sNameIconFile[MAX_PATH], sNameProceso[32],
			sNameFile[MAX_PATH], sNameInfo[64], sNameHelpFile[128];

		_splitpath( cFileName, drive, dir, fname, ext );
		sprintf( sLoadFileINI, "%s%s%s.ini",
			drive, dir, fname );

		NumberOfApplets =  GetPrivateProfileInt( "General", "NumberOfApplets", 0, sLoadFileINI );

		for( nI = 0; nI < NumberOfApplets; nI++ )
		{
            sprintf( sSectionName, "Applet%i", nI+1 ); //Section Name

			 GetPrivateProfileString( sSectionName, "IconFile" , "",
				sNameIconFile, MAX_PATH, sLoadFileINI );
			 GetPrivateProfileString( sSectionName, "AppletName",  "",
				sNameProceso, 32, sLoadFileINI );
			 GetPrivateProfileString( sSectionName, "ExeFile",  "",
				sNameFile, 32, sLoadFileINI );
			 GetPrivateProfileString( sSectionName, "Info",     "",
				sNameInfo, 64, sLoadFileINI );
			 GetPrivateProfileString( sSectionName, "HelpFile",  "",
				sNameHelpFile, 128, sLoadFileINI );

			CPAppList[nI].hIcon =  (HICON) LoadImage( NULL, sNameIconFile, IMAGE_ICON, 0, 0, LR_LOADFROMFILE );
			strcpy( CPAppList[nI].szFileExe,  sNameFile );
			strcpy( CPAppList[nI].szName,     sNameProceso );
			strcpy( CPAppList[nI].szInfo,     sNameInfo );
			strcpy( CPAppList[nI].szHelpFile, sNameHelpFile );

		}

		for( nI = NumberOfApplets; nI < MAX_APPLETS; nI++ )
		{
			if( CPAppList[nI].hIcon )
				 DestroyIcon( CPAppList[nI].hIcon );

			CPAppList[nI].hIcon =  NULL;
			strcpy( CPAppList[nI].szName,     "" );
			strcpy( CPAppList[nI].szInfo,     "" );
			strcpy( CPAppList[nI].szHelpFile, "" );
		}
	}
}

