#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "shlwapi")
#pragma comment(lib, "wininet")

#include <windows.h>
#include <shlwapi.h>
#include <wininet.h>
#include <vector>
#include <string>

#define DEFAULT_DPI 96
#define SCALEX(X) MulDiv(X, uDpiX, DEFAULT_DPI)
#define SCALEY(Y) MulDiv(Y, uDpiY, DEFAULT_DPI)
#define POINT2PIXEL(PT) MulDiv(PT, uDpiY, 72)

TCHAR szClassName[] = TEXT("Window");

BOOL GetScaling(HWND hWnd, UINT* pnX, UINT* pnY)
{
	BOOL bSetScaling = FALSE;
	const HMONITOR hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST);
	if (hMonitor)
	{
		HMODULE hShcore = LoadLibrary(TEXT("SHCORE"));
		if (hShcore)
		{
			typedef HRESULT __stdcall GetDpiForMonitor(HMONITOR, int, UINT*, UINT*);
			GetDpiForMonitor* fnGetDpiForMonitor = reinterpret_cast<GetDpiForMonitor*>(GetProcAddress(hShcore, "GetDpiForMonitor"));
			if (fnGetDpiForMonitor)
			{
				UINT uDpiX, uDpiY;
				if (SUCCEEDED(fnGetDpiForMonitor(hMonitor, 0, &uDpiX, &uDpiY)) && uDpiX > 0 && uDpiY > 0)
				{
					*pnX = uDpiX;
					*pnY = uDpiY;
					bSetScaling = TRUE;
				}
			}
			FreeLibrary(hShcore);
		}
	}
	if (!bSetScaling)
	{
		HDC hdc = GetDC(NULL);
		if (hdc)
		{
			*pnX = GetDeviceCaps(hdc, LOGPIXELSX);
			*pnY = GetDeviceCaps(hdc, LOGPIXELSY);
			ReleaseDC(NULL, hdc);
			bSetScaling = TRUE;
		}
	}
	if (!bSetScaling)
	{
		*pnX = DEFAULT_DPI;
		*pnY = DEFAULT_DPI;
		bSetScaling = TRUE;
	}
	return bSetScaling;
}

BOOL DownloadToFile(IN LPCTSTR lpszURL, IN LPCTSTR lpszFilePath)
{
	BOOL bReturn = FALSE;
	DWORD dwSize = 0;
	const HINTERNET hSession = InternetOpen(TEXT("Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/38.0.2125.111 Safari/537.36"), INTERNET_OPEN_TYPE_PRECONFIG, 0, 0, INTERNET_FLAG_NO_COOKIES);
	if (hSession)
	{
		URL_COMPONENTS uc = { 0 };
		TCHAR HostName[MAX_PATH];
		TCHAR UrlPath[MAX_PATH];
		uc.dwStructSize = sizeof(uc);
		uc.lpszHostName = HostName;
		uc.lpszUrlPath = UrlPath;
		uc.dwHostNameLength = MAX_PATH;
		uc.dwUrlPathLength = MAX_PATH;
		InternetCrackUrl(lpszURL, 0, 0, &uc);
		const HINTERNET hConnection = InternetConnect(hSession, HostName, INTERNET_DEFAULT_HTTP_PORT, 0, 0, INTERNET_SERVICE_HTTP, 0, 0);
		if (hConnection)
		{
			const HINTERNET hRequest = HttpOpenRequest(hConnection, TEXT("GET"), UrlPath, 0, 0, 0, 0, 0);
			if (hRequest)
			{
				InternetCrackUrl(lpszURL, 0, 0, &uc);
				HttpAddRequestHeaders(hRequest, 0, 0, HTTP_ADDREQ_FLAG_REPLACE | HTTP_ADDREQ_FLAG_ADD);
				HttpSendRequest(hRequest, 0, 0, 0, 0);
				DWORD dwContentLen = 0;
				DWORD dwBufLen = sizeof(dwContentLen);
				HttpQueryInfo(hRequest, HTTP_QUERY_CONTENT_LENGTH | HTTP_QUERY_FLAG_NUMBER, (LPVOID)&dwContentLen, &dwBufLen, 0);
				DWORD dwRead;
				static BYTE szBuf[1024 * 4];
				const HANDLE hFile = CreateFile(lpszFilePath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
				if (hFile != INVALID_HANDLE_VALUE)
				{
					DWORD dwWritten;
					for (;;)
					{
						if (!InternetReadFile(hRequest, szBuf, (DWORD)sizeof(szBuf), &dwRead) || !dwRead) break;
						WriteFile(hFile, szBuf, dwRead, &dwWritten, NULL);
						dwSize += dwRead;
					}
					CloseHandle(hFile);
					bReturn = TRUE;
				}
				InternetCloseHandle(hRequest);
			}
			InternetCloseHandle(hConnection);
		}
		InternetCloseHandle(hSession);
	}
	return bReturn;
}

void GetExpandedURLs(IN LPCWSTR lpszURL, OUT std::vector<std::wstring> & list)
{
	WCHAR szURL[2048];
	lstrcpyn(szURL, lpszURL, _countof(szURL));
	for (int i = 0; szURL[i]; i++)
	{
		if (szURL[i] == TEXT('['))
		{
			szURL[i] = TEXT('\0');
			for (int j = i + 1; szURL[j]; j++)
			{
				if (szURL[j] == TEXT(']'))
				{
					szURL[j] = TEXT('\0');
					for (int k = i + 1; szURL[k]; k++)
					{
						if (szURL[k] == TEXT('-'))
						{
							szURL[k] = TEXT('\0');
							if (k - 1 == i || k + 1 == j) {
								return;/*ハイフンの前後どちらかに文字が存在しなかった*/
							}
							int l;
							for (l = k + 1; l < j; l++)
							{
								if (!isdigit(szURL[l])) {
									return;/*ハイフンの後に数字以外の文字がある*/
								}
							}
							for (l = i + 1; l < k; l++)
							{
								if (!isdigit(szURL[l]))
								{
									return;/*ハイフンの前に数字以外の文字がある*/
								}
							}
							errno = 0;
#ifdef UNICODE
							long n1 = wcstol(&szURL[i + 1], NULL, 10);
							long n2 = wcstol(&szURL[k + 1], NULL, 10);
#else
							long n1 = strtol(&szURL[i + 1], NULL, 10);
							long n2 = strtol(&szURL[k + 1], NULL, 10);
#endif
							if (n1 < 0 || n2 < n1 || errno == ERANGE)
							{
								return;/*n1が負またはn1のほうがn2より大きかったまたはどちらかがオーバーフローしている*/
							}
							TCHAR szFormat[256];
							wsprintf(szFormat, TEXT("%%s%%0%dld%%s"), l - i - 1);
							for (l = n1; l <= n2; l++)
							{
								TCHAR szTemp[2024];
								wsprintf(szTemp, szFormat, szURL, l, &szURL[j + 1]);
								list.push_back(szTemp);
							}
							return;//成功
						}
					}
					return;//-がなかった
				}
			}
			return;//]がなかった
		}
	}
	list.push_back(lpszURL);
	//[がなかった
}

void GetFileNameFromURL(IN LPCWSTR lpszURL, OUT LPWSTR lpszFileName)
{
	int nMax = lstrlenW(lpszURL);
	for (int i = nMax - 1; i != 0; i--)
	{
		if (lpszURL[i] != L'/')
		{
			nMax = i;
			break;
		}
	}
	for (int i = nMax - 1; i >= 0; i--)
	{
		if (lpszURL[i] == L'/')
		{
			lstrcpy(lpszFileName, &lpszURL[i + 1]);
			for (int j = 0; lpszFileName[j] != 0; j++)
			{
				if (lpszFileName[j] == L'/')
				{
					lpszFileName[j] = 0;
					break;
				}
			}
			return;
		}
	}
	lstrcpy(lpszFileName, lpszURL);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hButton;
	static HWND hEdit1;
	static HWND hEdit2;
	static HFONT hFont;
	static UINT uDpiX = DEFAULT_DPI, uDpiY = DEFAULT_DPI;
	switch (msg)
	{
	case WM_CREATE:
		hEdit1 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hEdit1, EM_LIMITTEXT, 0, 0);
		hButton = CreateWindow(TEXT("BUTTON"), TEXT("Download"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit2 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | WS_VSCROLL | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hEdit2, EM_LIMITTEXT, 0, 0);
		SendMessage(hWnd, WM_APP, 0, 0);
		break;
	case WM_APP:
		GetScaling(hWnd, &uDpiX, &uDpiY);
		DeleteObject(hFont);
		hFont = CreateFontW(-POINT2PIXEL(10), 0, 0, 0, FW_NORMAL, 0, 0, 0, SHIFTJIS_CHARSET, 0, 0, 0, 0, L"MS Shell Dlg");
		SendMessage(hButton, WM_SETFONT, (WPARAM)hFont, 0);
		SendMessage(hEdit1, WM_SETFONT, (WPARAM)hFont, 0);
		SendMessage(hEdit2, WM_SETFONT, (WPARAM)hFont, 0);
		break;
	case WM_SIZE:
		MoveWindow(hButton, LOWORD(lParam) - POINT2PIXEL(128 + 10), POINT2PIXEL(10), POINT2PIXEL(128), POINT2PIXEL(32), TRUE);
		MoveWindow(hEdit1, POINT2PIXEL(10), POINT2PIXEL(10), LOWORD(lParam) - POINT2PIXEL(128 + 30), POINT2PIXEL(32), TRUE);
		MoveWindow(hEdit2, POINT2PIXEL(10), POINT2PIXEL(50), LOWORD(lParam) - POINT2PIXEL(20), HIWORD(lParam) - POINT2PIXEL(60), TRUE);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
			WCHAR szText[2048];
			GetWindowText(hEdit1, szText, _countof(szText));
			std::vector<std::wstring> list;
			GetExpandedURLs(szText, list);
			SetWindowText(hEdit2, 0);
			int nRetryCount = 0;
			for (int i =0; i < list.size(); i++)
			{
				nRetryCount = 0;
				WCHAR szFileName1[MAX_PATH];
				WCHAR szFileName2[MAX_PATH];
				GetFileNameFromURL(list[i].c_str(), szFileName1);
				wsprintf(szFileName2, L"%05d_%s", i + 1, szFileName1);
RETRY:
				if (DownloadToFile(list[i].c_str(), szFileName2))
				{
					SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)L"成功");
					SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)list[i].c_str());
					SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)L"\r\n");
				}
				else
				{
					nRetryCount++;
					if (nRetryCount <= 5)
					{
						SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)L"リトライ");
						SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)list[i].c_str());
						SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)L"\r\n");
						goto RETRY;
					}
					else
					{
						SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)L"失敗");
						SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)list[i].c_str());
						SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)L"\r\n");
					}
				}
			}
		}
		break;
	case WM_NCCREATE:
		{
			const HMODULE hModUser32 = GetModuleHandle(TEXT("user32.dll"));
			if (hModUser32)
			{
				typedef BOOL(WINAPI*fnTypeEnableNCScaling)(HWND);
				const fnTypeEnableNCScaling fnEnableNCScaling = (fnTypeEnableNCScaling)GetProcAddress(hModUser32, "EnableNonClientDpiScaling");
				if (fnEnableNCScaling)
				{
					fnEnableNCScaling(hWnd);
				}
			}
		}
		return DefWindowProc(hWnd, msg, wParam, lParam);
	case WM_DPICHANGED:
		SendMessage(hWnd, WM_APP, 0, 0);
		break;
	case WM_DESTROY:
		DeleteObject(hFont);
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPWSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("Window"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
