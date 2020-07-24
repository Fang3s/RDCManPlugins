// CPPConsoleApp.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"
#include <iostream>
#include <Windows.h>
#include <tlhelp32.h>
#include <psapi.h>
#include <Shlwapi.h>
#pragma comment(lib, "Shlwapi.lib");

#define RDCMAN_EXE_NAME L"RDCMan.exe"
#define RDCMAN_INPUT_WND_CLASS L"IHWindowClass"    // Class Name: IHWindowClass
#define RDCMAN_INPUT_WND_TITLE L"Input Capture Window"    // Window Captain: Input Capture Window

BOOL CALLBACK EnumRDPWindowsProc(HWND Window, LPARAM);
BOOL CALLBACK EnumTopWindowsToFindRDCManWindowProc(HWND Window, LPARAM);
BOOL CALLBACK EnumChildWindowsToFindInputWindowProc(HWND hWnd, LPARAM);

int main()
{
    while (true)
    {
        if (EnumWindows(EnumTopWindowsToFindRDCManWindowProc, NULL) == FALSE)
        {
            wprintf_s(L"ERROR: EnumWindows returned false! GetLastError: 0x%x\n", GetLastError());
        }
    }
    std::cout << "Hello World!\n";
}

void ProcessRDCInputWindow(HWND hWindowHandle)
{
    HWND hOriginalForegroundWindow = GetForegroundWindow();

    if (NULL != hWindowHandle)
    {
        //wprintf_s(L"[%I64d] Found %s - %s\n", GetTickCount64(), ClassName, WindowTitle);

        SetForegroundWindow(hWindowHandle);

        INPUT Input = { 0 };

        POINT CurrentPosition = { 0 };

        GetCursorPos(&CurrentPosition);

        Input.type = INPUT_MOUSE;

        Input.mi.dwFlags = MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE;

        //Input.mi.dx = (CurrentPosition.x + 1) * (65536.0f / GetSystemMetrics(SM_CXSCREEN));

        //Input.mi.dy = (CurrentPosition.y + 1) * (65536.0f / GetSystemMetrics(SM_CYSCREEN));

        //SendInput(1, &Input, sizeof(INPUT));

        Input.mi.dx = (CurrentPosition.x) * (65536.0f / GetSystemMetrics(SM_CXSCREEN));

        Input.mi.dy = (CurrentPosition.y) * (65536.0f / GetSystemMetrics(SM_CYSCREEN));

        SendInput(1, &Input, sizeof(INPUT));

        SetForegroundWindow(hOriginalForegroundWindow);
    }
}

BOOL IsRDCManWindow(HWND hWnd)
{
    wchar_t szWindowTitle[128] = { 0 }; // Window Captain: Input Capture Window
    GetWindowText(hWnd, szWindowTitle, _countof(szWindowTitle));
    wprintf(L"hWnd %p, %s\n", hWnd, szWindowTitle);
    DWORD dwProcessID = 0;
    GetWindowThreadProcessId(hWnd, &dwProcessID);
    HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, dwProcessID);
    if (NULL != hProcess)
    {
        wchar_t szExeFileName[MAX_PATH] = { 0 }; // Class Name: IHWindowClass
        //Returns the path in device form, rather than drive letters. For example, the file name
        //C:\Windows\System32\Ctype.nls would look as follows in device form:
        // \Device\Harddisk0\Partition1\Windows\System32\Ctype.nls
        //GetProcessImageFileNameW(hProcess, szImageFileName, _countof(szImageFileName)); //psapi.h
        DWORD dwSize = _countof(szExeFileName);
        if (QueryFullProcessImageNameW(hProcess, 0 * PROCESS_NAME_NATIVE, szExeFileName, &dwSize))
        {
            wprintf(L"Found hWnd %p, %s\n", hWnd, szExeFileName);
            LPWSTR pzFileName = PathFindFileNameW(szExeFileName);
            return 0 == _wcsicmp(pzFileName, RDCMAN_EXE_NAME);
        }
        CloseHandle(hProcess);
    }
    return FALSE;
}

//Enumerates all top-level windows on the screen by passing the handle to each window, in turn,
//to an application-defined callback function.
//To continue enumeration, the callback function must return TRUE; to stop enumeration, it must return FALSE.
BOOL CALLBACK EnumTopWindowsToFindRDCManWindowProc(HWND hWnd, LPARAM lparam)
{
    if (IsRDCManWindow(hWnd))
    {
        EnumChildWindows(hWnd, EnumChildWindowsToFindInputWindowProc, lparam);
    }
    return TRUE;
}

//https://devblogs.microsoft.com/oldnewthing/20070116-04/?p=28393
//The EnumChildWindows function already does the recursion :
//If a child window has created child windows of its own,
//EnumChildWindows enumerates those windows as well.
BOOL CALLBACK EnumChildWindowsToFindInputWindowProc(HWND hWnd, LPARAM)
{
    wchar_t szClassName[128] = { 0 }; // Class Name: IHWindowClass
    wchar_t szWindowTitle[128] = { 0 }; // Window Captain: Input Capture Window

    if (GetClassName(hWnd, szClassName, _countof(szClassName)) == 0)
    {
        wprintf_s(L"ERROR: GetClassName failed! %d\n", GetLastError()); //_putws
        return FALSE;
    }

    GetWindowText(hWnd, szWindowTitle, _countof(szWindowTitle));


    if (wcscmp(szClassName, RDCMAN_INPUT_WND_CLASS) == 0)
    {
        ProcessRDCInputWindow(hWnd);
    }

    return TRUE;
}

void FindRDCManWindowAndFindInputWindow()
{
    //Retrieves a handle to the top-level window whose class name and window name match the specified strings.
    //This function does not search child windows. This function does not perform a case-sensitive search.
    HWND hRDCWnd = FindWindow(L"WindowsForms10.Window.8.app.0.3d90434_r6_ad1", NULL);
    //FindWindowExW does not perform a case-sensitive search. Note that if both hwndParent and
    //hwndChildAfter are NULL, the function searches all top-level and message-only windows.
    if (NULL != hRDCWnd)
    {

        HWND hWnd = NULL;
        do
        {
            hWnd = FindWindowExW(hRDCWnd, hWnd, RDCMAN_INPUT_WND_CLASS, RDCMAN_INPUT_WND_TITLE);
        } while (NULL != hWnd);
    }
}

HANDLE FindProcess()
{
    PROCESSENTRY32W entry = { 0 };
    //ZeroMemory(&process, sizeof(process));
    entry.dwSize = sizeof(PROCESSENTRY32W);
    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, NULL);
    if (INVALID_HANDLE_VALUE == snapshot) {
        return INVALID_HANDLE_VALUE;
    }
    // Walkthrough all processes.
    if (Process32FirstW(snapshot, &entry) == TRUE)
    {
        do
        {
            if (_wcsicmp(entry.szExeFile, L"RDCMan.exe") == 0)
            {
                //Note it needs AdjustTokenPrivileges to open another local process with PROCESS_ALL_ACCESS 
                HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, entry.th32ProcessID);
                //CloseHandle(hProcess);
                return hProcess;
            }
        } while (Process32Next(snapshot, &entry) == TRUE);
    }
    CloseHandle(snapshot);
    return NULL;
}

void EnableDebugPriv()
{
    HANDLE hToken;
    LUID luid;
    TOKEN_PRIVILEGES tkp;

    OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken);

    LookupPrivilegeValue(NULL, SE_DEBUG_NAME, &luid);

    tkp.PrivilegeCount = 1;
    tkp.Privileges[0].Luid = luid;
    tkp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;

    AdjustTokenPrivileges(hToken, false, &tkp, sizeof(tkp), NULL, NULL);

    CloseHandle(hToken);
}

//https://referencesource.microsoft.com/#System/services/monitoring/system/diagnosticts/ProcessManager.cs,db7ac68b7cb40db1
//FindMainWindow(int processId) uses EnumWindows and GetWindowThreadProcessId
BOOL CALLBACK EnumRDPWindowsProc(HWND hWnd, LPARAM)
{
    wchar_t szClassName[128] = { 0 }; // Class Name: IHWindowClass
    wchar_t szWindowTitle[128] = { 0 }; // Window Captain: Input Capture Window

    if (GetClassName(hWnd, szClassName, _countof(szClassName)) == 0)
    {
        wprintf_s(L"ERROR: GetClassName failed! %d\n", GetLastError()); //_putws
        return FALSE;
    }

    GetWindowText(hWnd, szWindowTitle, _countof(szWindowTitle));

    DWORD pid = 0;
    DWORD tid = GetWindowThreadProcessId(hWnd, &pid);

    wprintf_s(L"Window: %s, %s\n", szClassName, szWindowTitle);

    if (wcslen(szClassName) == 0)
    {
        _snwprintf_s(szClassName, sizeof(szClassName), L"[NoClass]\n");
    }

    if (wcslen(szWindowTitle) == 0)
    {
        _snwprintf_s(szWindowTitle, sizeof(szWindowTitle), L"[NoTitle]\n");
    }

    if (wcscmp(szClassName, RDCMAN_INPUT_WND_CLASS) == 0)
    {
        ProcessRDCInputWindow(hWnd);
    }

    return TRUE;
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
