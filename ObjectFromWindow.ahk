/** v0.1.0 | By RaptorX
 * Wrapper function for [AccessibleObjectFromWindow](https://docs.microsoft.com/en-us/windows/win32/api/oleacc/nf-oleacc-accessibleobjectfromwindow)
 *
 * Borrowed & tweaked from Acc.ahk Standard Library by Sean
 * * Updated by jethrow
 * * Updated to v2 by RaptorX
 *
 * ### Params:
 * - `idObject` - Type of object that can be returned by the app
 *                more info can be found [here](https://docs.microsoft.com/en-us/windows/desktop/WinAuto/object-identifiers)
 * - `WinTitle` - Title of the window that has a comObject
 * ClassNN  [optional] - ClassNN of a control from a window that returns a comObject
 *
 * ### Returns:
 * - `False`  - No COM object could be created
 * - `ComObj` - ComObject returned by the application based on the type selected by idObject
 *
 * ### Examples:
 * Get the application object from MS Word
 * ```
	wd := ObjectFromWindow(OBJID_NATIVEOM, "- Word ahk_class OpusApp", "_WwG1")
	wd := wd.Application
 * ```
 * Get the application object from MS Excel
 * ```
	xl := ObjectFromWindow(OBJID_NATIVEOM, "- Word ahk_class OpusApp", "_WwG1")
	xl := xl.Application
 * ```
 */
