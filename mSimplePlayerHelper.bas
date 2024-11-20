Attribute VB_Name = "mSimplePlayerHelper"
Option Explicit
 
Public Function ucSimplePlayerHelperProc(ByVal lng_hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As ucSimplePlayer) As LongPtr
    ucSimplePlayerHelperProc = dwRefData.ucWndProc(lng_hWnd, uMsg, wParam, lParam, uIdSubclass)
End Function
