// SPDX-License-Identifier: Apache-2.0
// Copyright 2026 Michal Růžička <ruzicka.mich@gmail.com>

/*
 * Total Commander Lister Plugin API.
 * Minimal subset — see the official TC writers' guide for the full spec:
 * https://www.ghisler.ch/wiki/index.php?title=Lister_plugin_writers_guide
 */
#ifndef LISTPLUG_H
#define LISTPLUG_H

#include <windows.h>

// Show flags
#define lcp_wraptext     1
#define lcp_fittowindow  2
#define lcp_ansi         4
#define lcp_ascii        8
#define lcp_variable     12
#define lcp_forceshow    16
#define lcp_fitlargeronly 64
#define lcp_center        128

// Commands
#define lc_copy        1
#define lc_newparams   2
#define lc_selectall   3
#define lc_setpercent  4

#ifdef __cplusplus
extern "C" {
#endif

HWND  __stdcall ListLoad(HWND ParentWin, char* FileToLoad, int ShowFlags);
HWND  __stdcall ListLoadW(HWND ParentWin, wchar_t* FileToLoad, int ShowFlags);
int   __stdcall ListLoadNext(HWND ParentWin, HWND PluginWin, char* FileToLoad, int ShowFlags);
int   __stdcall ListLoadNextW(HWND ParentWin, HWND PluginWin, wchar_t* FileToLoad, int ShowFlags);
void  __stdcall ListCloseWindow(HWND ListWin);
void  __stdcall ListGetDetectString(char* DetectString, int maxlen);
int   __stdcall ListSendCommand(HWND ListWin, int Command, int Parameter);

#ifdef __cplusplus
}
#endif

#endif // LISTPLUG_H
