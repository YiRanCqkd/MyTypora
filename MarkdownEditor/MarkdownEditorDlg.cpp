
#include "pch.h"
#include "framework.h"
#include "MarkdownEditor.h"
#include "MarkdownEditorDlg.h"
#include "MarkdownParser.h"
#include "RenderText.h"
#include "afxdialogex.h"
#include <afxdlgs.h>
#include <algorithm>
#include <map>
#include <utility>
#include <shellapi.h>
#include <memory>
#include <thread>
#include <vector>
#include <limits>
#include <Richedit.h>
#include <eh.h>
#include <ole2.h>
#include <commctrl.h>
#include <Shlwapi.h>

#include <gdiplus.h>

#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "Shlwapi.lib")

namespace
{
	constexpr UINT kSysCmdSetDefaultMd = 0x0020;
    constexpr UINT WM_INIT_DRAGDROP = WM_APP + 101;
	constexpr UINT WM_MARKDOWN_PARSE_COMPLETE = WM_APP + 102;
	constexpr UINT WM_DOCUMENT_LOAD_COMPLETE = WM_APP + 103;
	constexpr UINT WM_PREPARE_DISPLAY_TEXT_COMPLETE = WM_APP + 104;
	constexpr UINT WM_DOCUMENT_LOAD_PROGRESS = WM_APP + 105;
	constexpr UINT WM_MARKDOWN_PARSE_PROGRESS = WM_APP + 106;
	constexpr UINT WM_OPEN_STARTUP_FILE = WM_APP + 107;
	// Mermaid diagrams are rendered in-process (no external renderer).
	constexpr UINT_PTR kDebounceTimerId = 1;
	constexpr UINT_PTR kFormatTimerId = 2;
	constexpr UINT_PTR kInsertTimerId = 3;
	constexpr UINT_PTR kStatusPulseTimerId = 4;
	constexpr UINT_PTR kMermaidOverlayTimerId = 5;
	constexpr UINT_PTR kOverlaySyncTimerId = 6;
	constexpr UINT_PTR kOutlineFollowTimerId = 7;
	constexpr UINT_PTR kAutoSaveTimerId = 8;
	constexpr UINT_PTR kFileWatchTimerId = 9;
	constexpr DWORD kAutoSaveIdleMs = 120000;
	constexpr DWORD kAutoSaveErrorPromptIntervalMs = 30000;
	constexpr ULONGLONG kMaxDocumentLoadBytes = 64ull * 1024ull * 1024ull;
	const wchar_t* kClipboardMarkdownFormatName = L"MyTypora.MarkdownSource";
	const UINT kFindReplaceMsg = ::RegisterWindowMessage(FINDMSGSTRING);

    constexpr DWORD_PTR kTreeDataTagMask = 0xF0000000;
    constexpr DWORD_PTR kTreeDataIndexMask = 0x0FFFFFFF;
    constexpr DWORD_PTR kTreeDataRecentTag = 0xA0000000;
    constexpr DWORD_PTR kTreeDataFolderTag = 0xB0000000;

    inline DWORD_PTR MakeTreeData(DWORD_PTR tag, size_t index)
    {
        return (tag & kTreeDataTagMask) | (static_cast<DWORD_PTR>(index) & kTreeDataIndexMask);
    }

    inline bool DecodeTreeData(DWORD_PTR value, DWORD_PTR& tag, size_t& index)
    {
        tag = value & kTreeDataTagMask;
        index = static_cast<size_t>(value & kTreeDataIndexMask);
        return tag == kTreeDataRecentTag || tag == kTreeDataFolderTag;
    }

	inline bool FileExistsOnDisk(const CString& path)
	{
		if (path.IsEmpty())
			return false;
		DWORD attrs = ::GetFileAttributes(path);
		if (attrs == INVALID_FILE_ATTRIBUTES)
			return false;
		if ((attrs & FILE_ATTRIBUTE_DIRECTORY) != 0)
			return false;
		return true;
	}

	inline CString ExtractFileNameFromPath(const CString& path)
	{
		CString name = path;
		int pos = name.ReverseFind(_T('\\'));
		int pos2 = name.ReverseFind(_T('/'));
		int cut = max(pos, pos2);
		if (cut != -1)
			name = name.Mid(cut + 1);
		return name;
	}

	inline bool WriteRegistryStringValue(HKEY root, const CString& subKey, const CString& valueName, const CString& value)
	{
		HKEY hKey = nullptr;
		const LONG createResult = ::RegCreateKeyEx(root, subKey, 0, nullptr, REG_OPTION_NON_VOLATILE,
			KEY_SET_VALUE, nullptr, &hKey, nullptr);
		if (createResult != ERROR_SUCCESS || hKey == nullptr)
			return false;

		const BYTE* data = reinterpret_cast<const BYTE*>(value.GetString());
		const DWORD bytes = static_cast<DWORD>((value.GetLength() + 1) * sizeof(TCHAR));
		const LONG setResult = ::RegSetValueEx(hKey,
			valueName.IsEmpty() ? nullptr : valueName.GetString(),
			0,
			REG_SZ,
			data,
			bytes);
		::RegCloseKey(hKey);
		return setResult == ERROR_SUCCESS;
	}

	inline bool ReadRegistryStringValue(HKEY root, const CString& subKey, const CString& valueName, CString& value)
	{
		value.Empty();
		HKEY hKey = nullptr;
		const LONG openResult = ::RegOpenKeyEx(root, subKey, 0, KEY_QUERY_VALUE, &hKey);
		if (openResult != ERROR_SUCCESS || hKey == nullptr)
			return false;

		DWORD type = 0;
		DWORD bytes = 0;
		const LONG querySizeResult = ::RegQueryValueEx(hKey,
			valueName.IsEmpty() ? nullptr : valueName.GetString(),
			nullptr,
			&type,
			nullptr,
			&bytes);
		if (querySizeResult != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ) || bytes == 0)
		{
			::RegCloseKey(hKey);
			return false;
		}

		std::vector<TCHAR> buffer((bytes / sizeof(TCHAR)) + 2, 0);
		const LONG queryDataResult = ::RegQueryValueEx(hKey,
			valueName.IsEmpty() ? nullptr : valueName.GetString(),
			nullptr,
			&type,
			reinterpret_cast<LPBYTE>(buffer.data()),
			&bytes);
		::RegCloseKey(hKey);
		if (queryDataResult != ERROR_SUCCESS)
			return false;

		value = buffer.data();
		return !value.IsEmpty();
	}

	inline bool WriteRegistryEmptyValue(HKEY root, const CString& subKey, const CString& valueName)
	{
		HKEY hKey = nullptr;
		const LONG createResult = ::RegCreateKeyEx(root, subKey, 0, nullptr, REG_OPTION_NON_VOLATILE,
			KEY_SET_VALUE, nullptr, &hKey, nullptr);
		if (createResult != ERROR_SUCCESS || hKey == nullptr)
			return false;
		const LONG setResult = ::RegSetValueEx(hKey,
			valueName.IsEmpty() ? nullptr : valueName.GetString(),
			0,
			REG_NONE,
			nullptr,
			0);
		::RegCloseKey(hKey);
		return setResult == ERROR_SUCCESS;
	}

	inline bool DeleteRegistryTreeIfExists(HKEY root, const CString& subKey)
	{
		const LONG r = ::RegDeleteTree(root, subKey);
		return r == ERROR_SUCCESS || r == ERROR_FILE_NOT_FOUND || r == ERROR_PATH_NOT_FOUND;
	}

	inline CString CanonicalizePathForCompare(const CString& path)
	{
		CString s(path);
		s.Trim();
		s.Replace(_T("/"), _T("\\"));
		s.MakeLower();
		return s;
	}

	inline bool QueryAssocStringForMd(ASSOCSTR assocStr, CString& value, LPCTSTR pszExtra = nullptr)
	{
		value.Empty();
		DWORD chars = 0;
		const HRESULT hrSize = ::AssocQueryString(0,
			assocStr,
			_T(".md"),
			pszExtra,
			nullptr,
			&chars);
		if (FAILED(hrSize) || chars == 0)
			return false;

		std::vector<TCHAR> buffer(chars + 2, 0);
		DWORD charsOut = chars;
		const HRESULT hrData = ::AssocQueryString(0,
			assocStr,
			_T(".md"),
			pszExtra,
			buffer.data(),
			&charsOut);
		if (FAILED(hrData))
			return false;

		value = buffer.data();
		value.Trim();
		return !value.IsEmpty();
	}

	inline bool OpenSystemDefaultAppsSettings()
	{
		HINSTANCE h = ::ShellExecute(nullptr, _T("open"), _T("ms-settings:defaultapps"),
			nullptr, nullptr, SW_SHOWNORMAL);
		if (reinterpret_cast<INT_PTR>(h) > 32)
			return true;

		h = ::ShellExecute(nullptr, _T("open"), _T("control.exe"),
			_T("/name Microsoft.DefaultPrograms"), nullptr, SW_SHOWNORMAL);
		return reinterpret_cast<INT_PTR>(h) > 32;
	}

	inline bool IsCurrentAppMdAssociationRegistered()
	{
		TCHAR currentExe[MAX_PATH] = { 0 };
		const DWORD len = ::GetModuleFileName(nullptr, currentExe, _countof(currentExe));
		if (len == 0 || len >= _countof(currentExe))
			return false;

		const CString currentExeNorm = CanonicalizePathForCompare(currentExe);

		// Only trust shell-effective executable (actual double-click behavior).
		CString effectiveExe;
		if (!QueryAssocStringForMd(ASSOCSTR_EXECUTABLE, effectiveExe, _T("open")))
			return false;
		if (!FileExistsOnDisk(effectiveExe))
			return false;
		return CanonicalizePathForCompare(effectiveExe) == currentExeNorm;
	}

	inline std::wstring NormalizeLineEndingsForCompare(const std::wstring& text)
	{
		std::wstring normalized;
		normalized.reserve(text.size());
		for (size_t i = 0; i < text.size(); ++i)
		{
			const wchar_t ch = text[i];
			if (ch == L'\r')
			{
				if (i + 1 < text.size() && text[i + 1] == L'\n')
					++i;
				normalized.push_back(L'\n');
			}
			else
			{
				normalized.push_back(ch);
			}
		}
		return normalized;
	}

	inline void BuildIdentityIndexMap(size_t textLength,
		std::vector<long>& displayToSource,
		std::vector<long>& sourceToDisplay)
	{
		displayToSource.resize(textLength + 1);
		sourceToDisplay.resize(textLength + 1);
		for (size_t i = 0; i <= textLength; ++i)
		{
			const long index = static_cast<long>(i);
			displayToSource[i] = index;
			sourceToDisplay[i] = index;
		}
	}

	inline bool IsEnvFlagEnabled(const TCHAR* name)
	{
		TCHAR value[16] = { 0 };
		const DWORD len = ::GetEnvironmentVariable(name, value, _countof(value));
		if (len == 0 || len >= _countof(value))
			return false;
		CString v(value);
		v.Trim();
		v.MakeLower();
		return !(v == _T("0") || v == _T("false") || v == _T("off") || v == _T("no"));
	}

	inline void DebugLogLine(const CString& text)
	{
		::OutputDebugString(text);
		::OutputDebugString(_T("\r\n"));
	}

	template <typename T>
	inline void SafePostMessageUnique(HWND hwnd, UINT message, WPARAM wParam, std::unique_ptr<T>&& payload)
	{
		if (!payload)
			return;
		if (hwnd == nullptr || !::IsWindow(hwnd))
			return; // payload auto-freed by unique_ptr

		if (::PostMessage(hwnd, message, wParam, reinterpret_cast<LPARAM>(payload.get())) != 0)
			payload.release();
	}

	struct AsyncWorkerScope
	{
		std::shared_ptr<std::atomic_int> workerCount;

		explicit AsyncWorkerScope(std::shared_ptr<std::atomic_int> count)
			: workerCount(std::move(count))
		{
			if (workerCount)
				workerCount->fetch_add(1, std::memory_order_relaxed);
		}

		~AsyncWorkerScope()
		{
			if (workerCount)
				workerCount->fetch_sub(1, std::memory_order_relaxed);
		}
	};

	inline bool ValidateSourceDisplayMappingVectors(const std::wstring& sourceText,
		const std::vector<long>& sourceToDisplay,
		const std::vector<long>& displayToSource,
		CString& reason)
	{
		const long sourceLen = static_cast<long>(sourceText.size());
		if (sourceToDisplay.size() != sourceText.size() + 1)
		{
			reason.Format(_T("sourceToDisplay size invalid: %llu, expected %llu"),
				static_cast<unsigned long long>(sourceToDisplay.size()),
				static_cast<unsigned long long>(sourceText.size() + 1));
			return false;
		}
		if (sourceToDisplay.empty())
		{
			reason = _T("sourceToDisplay is empty");
			return false;
		}

		const long displayLen = sourceToDisplay.back();
		if (displayLen < 0)
		{
			reason.Format(_T("displayLen is negative: %ld"), displayLen);
			return false;
		}
		if (displayToSource.size() != static_cast<size_t>(displayLen + 1))
		{
			reason.Format(_T("displayToSource size invalid: %llu, expected %ld"),
				static_cast<unsigned long long>(displayToSource.size()),
				displayLen + 1);
			return false;
		}
		if (sourceToDisplay.front() != 0)
		{
			reason.Format(_T("sourceToDisplay[0] != 0: %ld"), sourceToDisplay.front());
			return false;
		}
		if (!displayToSource.empty() && displayToSource.front() != 0)
		{
			reason.Format(_T("displayToSource[0] != 0: %ld"), displayToSource.front());
			return false;
		}
		if (!displayToSource.empty() && displayToSource.back() > sourceLen)
		{
			reason.Format(_T("displayToSource.back out of range: %ld > %ld"),
				displayToSource.back(), sourceLen);
			return false;
		}

		long prev = 0;
		for (size_t i = 0; i < sourceToDisplay.size(); ++i)
		{
			const long value = sourceToDisplay[i];
			if (value < prev)
			{
				reason.Format(_T("sourceToDisplay not monotonic at %llu: %ld < %ld"),
					static_cast<unsigned long long>(i), value, prev);
				return false;
			}
			if (value < 0 || value > displayLen)
			{
				reason.Format(_T("sourceToDisplay out of range at %llu: %ld"),
					static_cast<unsigned long long>(i), value);
				return false;
			}
			prev = value;
		}

		prev = 0;
		for (size_t i = 0; i < displayToSource.size(); ++i)
		{
			const long value = displayToSource[i];
			if (value < prev)
			{
				reason.Format(_T("displayToSource not monotonic at %llu: %ld < %ld"),
					static_cast<unsigned long long>(i), value, prev);
				return false;
			}
			if (value < 0 || value > sourceLen)
			{
				reason.Format(_T("displayToSource out of range at %llu: %ld"),
					static_cast<unsigned long long>(i), value);
				return false;
			}
			prev = value;
		}

		return true;
	}

	inline void EnsureValidSourceDisplayMapping(const std::wstring& sourceText,
		std::vector<long>& sourceToDisplay,
		std::vector<long>& displayToSource,
		const TCHAR* context)
	{
		bool repaired = false;
		if (sourceToDisplay.size() != sourceText.size() + 1 ||
			sourceToDisplay.empty() ||
			sourceToDisplay.back() < 0 ||
			displayToSource.size() != static_cast<size_t>(sourceToDisplay.back() + 1))
		{
			BuildIdentityIndexMap(sourceText.size(), displayToSource, sourceToDisplay);
			repaired = true;
		}

		const bool debugEnabled = IsEnvFlagEnabled(_T("MYTYPORA_DEBUG_MAPPING"));
		if (!debugEnabled)
			return;

		CString reason;
		if (!ValidateSourceDisplayMappingVectors(sourceText, sourceToDisplay, displayToSource, reason))
		{
			BuildIdentityIndexMap(sourceText.size(), displayToSource, sourceToDisplay);
			repaired = true;
			CString msg;
			msg.Format(_T("[Mapping][%s] invalid source/display mapping, repaired to identity: %s"),
				context, (LPCTSTR)reason);
			DebugLogLine(msg);
		}
		else if (repaired)
		{
			CString msg;
			msg.Format(_T("[Mapping][%s] quick-check failed, repaired to identity."), context);
			DebugLogLine(msg);
		}
	}

	constexpr LONG kRichTabPosMask = 0x00FFFFFF;
	constexpr LONG kRichTabAlignMask = 0x0F000000;
	constexpr LONG kRichTabAlignLeft = (0L << 24);
	constexpr LONG kRichTabAlignCenter = (1L << 24);
	constexpr LONG kRichTabAlignRight = (2L << 24);

	inline bool ValidateRichToSourceMappingVector(const std::vector<long>& richToSource,
		long richLen,
		long sourceLen,
		CString& reason)
	{
		if (richLen < 0)
		{
			reason.Format(_T("richLen invalid: %ld"), richLen);
			return false;
		}
		if (richToSource.size() != static_cast<size_t>(richLen + 1))
		{
			reason.Format(_T("richToSource size invalid: %llu, expected %ld"),
				static_cast<unsigned long long>(richToSource.size()),
				richLen + 1);
			return false;
		}
		long prev = 0;
		for (size_t i = 0; i < richToSource.size(); ++i)
		{
			const long value = richToSource[i];
			if (value < prev)
			{
				reason.Format(_T("richToSource not monotonic at %llu: %ld < %ld"),
					static_cast<unsigned long long>(i), value, prev);
				return false;
			}
			if (value < 0 || value > sourceLen)
			{
				reason.Format(_T("richToSource out of range at %llu: %ld"),
					static_cast<unsigned long long>(i), value);
				return false;
			}
			prev = value;
		}
		return true;
	}

	inline void ValidateRichToSourceMapping(std::vector<long>& richToSource,
		long richLen,
		long sourceLen,
		const TCHAR* context)
	{
		const bool debugEnabled = IsEnvFlagEnabled(_T("MYTYPORA_DEBUG_MAPPING"));
		if (richToSource.empty())
			return;

		if (!debugEnabled)
		{
			if (richToSource.size() != static_cast<size_t>(richLen + 1))
				richToSource.clear();
			return;
		}

		CString reason;
		if (!ValidateRichToSourceMappingVector(richToSource, richLen, sourceLen, reason))
		{
			CString msg;
			msg.Format(_T("[Mapping][%s] invalid rich/source mapping, clear map: %s"),
				context, (LPCTSTR)reason);
			DebugLogLine(msg);
			richToSource.clear();
		}
	}

	inline bool DetectUtf8Bom(const CString& path, bool& hasBom)
	{
		hasBom = false;
		if (path.IsEmpty())
			return false;

		try
		{
			CFile file(path, CFile::modeRead | CFile::typeBinary | CFile::shareDenyNone);
			BYTE header[3] = { 0 };
			const UINT readBytes = file.Read(header, 3);
			file.Close();
			hasBom = (readBytes == 3 && header[0] == 0xEF && header[1] == 0xBB && header[2] == 0xBF);
			return true;
		}
		catch (CFileException* ex)
		{
			if (ex)
				ex->Delete();
			return false;
		}
	}

	inline bool WriteUtf8File(const CString& path, const CString& content, bool writeBom, CString* errorMessage)
	{
		if (path.IsEmpty())
		{
			if (errorMessage)
				*errorMessage = _T("保存路径为空。");
			return false;
		}

		const int wideLen = content.GetLength();
		const int utf8Len = ::WideCharToMultiByte(CP_UTF8, 0, content.GetString(), wideLen, nullptr, 0, nullptr, nullptr);
		if (utf8Len < 0)
		{
			if (errorMessage)
				*errorMessage = _T("文本编码失败。");
			return false;
		}

		std::vector<char> utf8;
		utf8.resize(static_cast<size_t>(utf8Len));
		if (utf8Len > 0)
			::WideCharToMultiByte(CP_UTF8, 0, content.GetString(), wideLen, utf8.data(), utf8Len, nullptr, nullptr);

		try
		{
			CFile file(path, CFile::modeCreate | CFile::modeWrite | CFile::typeBinary);
			if (writeBom)
			{
				const BYTE bom[3] = { 0xEF, 0xBB, 0xBF };
				file.Write(bom, static_cast<UINT>(sizeof(bom)));
			}
			if (!utf8.empty())
				file.Write(utf8.data(), static_cast<UINT>(utf8.size()));
			file.Close();
			return true;
		}
		catch (CFileException* ex)
		{
			if (errorMessage)
			{
				TCHAR reason[512] = { 0 };
				if (ex && ex->GetErrorMessage(reason, _countof(reason)))
					errorMessage->Format(_T("保存失败：%s"), reason);
				else
					*errorMessage = _T("保存文件失败。");
			}
			if (ex)
				ex->Delete();
			return false;
		}
	}

	inline bool CopyTextToClipboard(HWND owner, const CString& text)
	{
		if (text.IsEmpty())
			return false;

		// Use Unicode clipboard format even if the app is ever built as ANSI.
		const CStringW wide(text);

		if (!::OpenClipboard(owner))
			return false;
		struct ClipboardCloser
		{
			~ClipboardCloser() { ::CloseClipboard(); }
		} closer;

		if (!::EmptyClipboard())
			return false;

		const SIZE_T bytes = static_cast<SIZE_T>((wide.GetLength() + 1) * sizeof(wchar_t));
		HGLOBAL h = ::GlobalAlloc(GMEM_MOVEABLE, bytes);
		if (!h)
			return false;
		void* ptr = ::GlobalLock(h);
		if (!ptr)
		{
			::GlobalFree(h);
			return false;
		}
		memcpy(ptr, wide.GetString(), bytes);
		::GlobalUnlock(h);

		if (::SetClipboardData(CF_UNICODETEXT, h) == nullptr)
		{
			::GlobalFree(h);
			return false;
		}
		// Clipboard owns h now.
		return true;
	}

	inline bool IsUserInteractionMessage(const MSG* pMsg)
	{
		if (!pMsg)
			return false;
		switch (pMsg->message)
		{
		case WM_KEYDOWN:
		case WM_KEYUP:
		case WM_SYSKEYDOWN:
		case WM_SYSKEYUP:
		case WM_CHAR:
		case WM_LBUTTONDOWN:
		case WM_LBUTTONUP:
		case WM_RBUTTONDOWN:
		case WM_RBUTTONUP:
		case WM_MBUTTONDOWN:
		case WM_MBUTTONUP:
		case WM_MOUSEMOVE:
		case WM_MOUSEWHEEL:
		case WM_IME_STARTCOMPOSITION:
		case WM_IME_ENDCOMPOSITION:
		case WM_IME_CHAR:
			return true;
		default:
			return false;
		}
	}

	inline bool IsReadonlyEditIntentMessage(const MSG* pMsg)
	{
		if (!pMsg)
			return false;

		const UINT msg = pMsg->message;
		if (msg == WM_PASTE || msg == WM_CUT || msg == WM_CLEAR)
			return true;
		if (msg == WM_IME_CHAR)
			return true;

		if (msg == WM_CHAR)
		{
			const WPARAM ch = pMsg->wParam;
			if (ch == VK_BACK || ch == VK_RETURN || ch == VK_TAB)
				return true;
			return ch >= 0x20;
		}

		if (msg != WM_KEYDOWN && msg != WM_SYSKEYDOWN)
			return false;

		const bool ctrlDown = ((::GetKeyState(VK_CONTROL) & 0x8000) != 0);
		const bool shiftDown = ((::GetKeyState(VK_SHIFT) & 0x8000) != 0);
		const WPARAM key = pMsg->wParam;

		if (key == VK_BACK || key == VK_DELETE || key == VK_RETURN)
			return true;
		if (!ctrlDown && key == VK_TAB)
			return true;
		if (ctrlDown && (key == 'V' || key == 'v' || key == 'X' || key == 'x'))
			return true;
		if (key == VK_INSERT && shiftDown)
			return true;

		return false;
	}

    inline bool IsWordCharacter(TCHAR ch)
    {
        if (ch == 0)
            return false;
        if (ch >= _T('0') && ch <= _T('9'))
            return true;
        if (ch >= _T('A') && ch <= _T('Z'))
            return true;
        if (ch >= _T('a') && ch <= _T('z'))
            return true;
        if (ch == _T('_'))
            return true;
        if (ch >= 0x80)
            return true;
        return false;
    }

    inline bool IsIndexInsideRanges(const std::vector<std::pair<int, int>>& ranges, int position)
    {
        for (const auto& range : ranges)
        {
            if (position >= range.first && position < range.second)
                return true;
        }
        return false;
    }

    inline bool IsLinkBoundaryChar(TCHAR ch)
    {
        if (ch == 0)
            return true;
        static const TCHAR* kBoundaries = _T(" \t\r\n|<>'\"()[]{}“”‘’（）《》【】");
        return _tcschr(kBoundaries, ch) != nullptr;
    }

    inline bool IsLinkTrailingPunct(TCHAR ch)
    {
        static const TCHAR* kTrailing = _T(".,;:!?\"'’”〉》】）」』，。！？；：、");
        return _tcschr(kTrailing, ch) != nullptr;
    }

    inline bool IsLinkLeadingWrapper(TCHAR ch)
    {
        static const TCHAR* kLeading = _T("\"'`“”‘’(<（《【[{《〈「『");
        return _tcschr(kLeading, ch) != nullptr;
    }

    inline void SanitizeLinkToken(CString& value)
    {
        value.Trim();
        while (!value.IsEmpty() && IsLinkLeadingWrapper(value[0]))
            value.Delete(0);
        while (!value.IsEmpty() && IsLinkTrailingPunct(value[value.GetLength() - 1]))
            value.Truncate(value.GetLength() - 1);
        value.Trim();
    }

	inline bool IsSchemeChar(TCHAR ch)
	{
		return _istalnum(ch) || ch == _T('+') || ch == _T('-') || ch == _T('.');
	}

	inline bool TryExtractUrlScheme(const CString& url, CString& scheme)
	{
		scheme.Empty();
		if (url.IsEmpty() || !_istalpha(url[0]))
			return false;

		const int colon = url.Find(_T(':'));
		if (colon <= 0)
			return false;

		for (int i = 0; i < colon; ++i)
		{
			if (!IsSchemeChar(url[i]))
				return false;
		}

		scheme = url.Left(colon);
		scheme.MakeLower();
		return !scheme.IsEmpty();
	}

	inline bool IsAllowedExternalLinkScheme(const CString& scheme)
	{
		return scheme == _T("http") || scheme == _T("https") ||
			scheme == _T("mailto") || scheme == _T("ftp");
	}

    inline void ApplyCharFormatRange(CRichEditCtrl& edit, LONG start, LONG end, const CHARFORMAT2& format)
    {
        if (start >= end)
            return;
        edit.SetSel(start, end);
        CHARFORMAT2 nonConstFormat = format;
        edit.SetSelectionCharFormat(nonConstFormat);
    }

    inline void HideRange(CRichEditCtrl& edit, LONG start, LONG end, CHARFORMAT2& hiddenFormat)
    {
        if (start >= end)
            return;
        edit.SetSel(start, end);
        edit.SetSelectionCharFormat(hiddenFormat);
    }

	struct MarkdownParsePayload
	{
		unsigned long long generation = 0;
		long richEditLength = 0;
		std::wstring text;
		std::wstring displayText;
		std::vector<long> displayToSourceIndexMap;
		std::vector<long> sourceToDisplayIndexMap;
		MarkdownParseResult result;
		std::vector<MarkdownOutlineEntry> outlineEntries;
		DWORD exceptionCode = 0;
		bool loadFailed = false;
		std::wstring loadErrorMessage;
	};

	struct DisplayTextPayload
	{
		unsigned long long generation = 0;
		std::wstring displayText;
		std::vector<long> displayToSourceIndexMap;
		std::vector<long> sourceToDisplayIndexMap;
	};

	struct DocumentLoadProgressPayload
	{
		unsigned long long generation = 0;
		unsigned long long bytesRead = 0;
		unsigned long long bytesTotal = 0;
	};

	struct MarkdownParseProgressPayload
	{
		unsigned long long generation = 0;
		int percent = 0;
	};

	struct MermaidBitmap
	{
		HBITMAP bitmap = nullptr;
		int width = 0;
		int height = 0;
	};



	inline void ToLowerInPlace(std::wstring& s)
	{
		for (auto& ch : s)
			ch = static_cast<wchar_t>(towlower(ch));
	}

	inline size_t HashWString(const std::wstring& s)
	{
		// FNV-1a 64-bit
		size_t h = static_cast<size_t>(1469598103934665603ull);
		for (wchar_t ch : s)
		{
			h ^= static_cast<size_t>(ch);
			h *= static_cast<size_t>(1099511628211ull);
		}
		return h;
	}

	inline std::wstring TrimW(const std::wstring& s)
	{
		size_t a = 0;
		while (a < s.size() && (s[a] == L' ' || s[a] == L'\t' || s[a] == L'\r' || s[a] == L'\n'))
			++a;
		size_t b = s.size();
		while (b > a && (s[b - 1] == L' ' || s[b - 1] == L'\t' || s[b - 1] == L'\r' || s[b - 1] == L'\n'))
			--b;
		return (b > a) ? s.substr(a, b - a) : std::wstring();
	}

	inline std::unique_ptr<Gdiplus::Font> CreateUiFont(float sizePx, INT style)
	{
		const wchar_t* candidates[] = {
			L"Microsoft YaHei UI",
			L"Microsoft YaHei",
			L"SimSun",
			L"SimHei",
			L"Segoe UI",
			L"Arial Unicode MS",
			L"Arial"
		};
		for (const wchar_t* name : candidates)
		{
			auto font = std::make_unique<Gdiplus::Font>(name, sizePx, style, Gdiplus::UnitPixel);
			if (font && font->GetLastStatus() == Gdiplus::Ok)
				return font;
		}
		return std::make_unique<Gdiplus::Font>(L"Arial", sizePx, style, Gdiplus::UnitPixel);
	}

	inline std::unique_ptr<Gdiplus::Font> CreateMonoFont(float sizePx, INT style)
	{
		const wchar_t* candidates[] = {
			L"Consolas",
			L"Courier New",
			L"Lucida Console"
		};
		for (const wchar_t* name : candidates)
		{
			auto font = std::make_unique<Gdiplus::Font>(name, sizePx, style, Gdiplus::UnitPixel);
			if (font && font->GetLastStatus() == Gdiplus::Ok)
				return font;
		}
		return std::make_unique<Gdiplus::Font>(L"Courier New", sizePx, style, Gdiplus::UnitPixel);
	}

	struct SehException
	{
		unsigned int code = 0;
		explicit SehException(unsigned int value) : code(value) {}
	};

	static void __cdecl SehTranslator(unsigned int code, EXCEPTION_POINTERS* /*info*/)
	{
		throw SehException(code);
	}

	struct SehTranslatorScope
	{
		_se_translator_function previous = nullptr;
		SehTranslatorScope() { previous = _set_se_translator(SehTranslator); }
		~SehTranslatorScope() { _set_se_translator(previous); }
	};

	// Minimal Win10/11 dark mode support for common controls (scrollbars, tab control background).
	// Uses uxtheme.dll exported functions by ordinal when available; safe no-op otherwise.
	// IMPORTANT: Do NOT call these ordinals on Win7/Win8. The ordinal numbers can exist but point to
	// unrelated functions, which will crash due to signature mismatch.
	inline bool IsWindows10OrLater()
	{
		typedef LONG(WINAPI* RtlGetVersionFn)(PRTL_OSVERSIONINFOW);
		HMODULE hNt = ::GetModuleHandleW(L"ntdll.dll");
		if (!hNt)
			return false;
		auto fn = reinterpret_cast<RtlGetVersionFn>(::GetProcAddress(hNt, "RtlGetVersion"));
		if (!fn)
			return false;
		RTL_OSVERSIONINFOW ver{};
		ver.dwOSVersionInfoSize = sizeof(ver);
		if (fn(&ver) != 0)
			return false;
		return ver.dwMajorVersion >= 10;
	}

	enum class PreferredAppMode
	{
		Default = 0,
		AllowDark = 1,
		ForceDark = 2,
		ForceLight = 3,
		Max = 4
	};

	struct DarkModeApi
	{
		typedef HRESULT(WINAPI* SetWindowThemeFn)(HWND, LPCWSTR, LPCWSTR);
		typedef bool (WINAPI* AllowDarkModeForWindowFn)(HWND, bool);
		typedef PreferredAppMode(WINAPI* SetPreferredAppModeFn)(PreferredAppMode);
		typedef void (WINAPI* RefreshImmersiveColorPolicyStateFn)();
		typedef void (WINAPI* FlushMenuThemesFn)();

		HMODULE hUxTheme = nullptr;
		SetWindowThemeFn setWindowTheme = nullptr;
		AllowDarkModeForWindowFn allowDarkModeForWindow = nullptr;
		SetPreferredAppModeFn setPreferredAppMode = nullptr;
		RefreshImmersiveColorPolicyStateFn refreshPolicy = nullptr;
		FlushMenuThemesFn flushMenuThemes = nullptr;

		void EnsureLoaded()
		{
			if (hUxTheme)
				return;
			hUxTheme = ::LoadLibraryW(L"uxtheme.dll");
			if (!hUxTheme)
				return;
			setWindowTheme = reinterpret_cast<SetWindowThemeFn>(::GetProcAddress(hUxTheme, "SetWindowTheme"));

			if (!IsWindows10OrLater())
				return;
			// Ordinals used by Windows 10/11. If missing, dark mode just won't apply.
			allowDarkModeForWindow = reinterpret_cast<AllowDarkModeForWindowFn>(::GetProcAddress(hUxTheme, MAKEINTRESOURCEA(133)));
			setPreferredAppMode = reinterpret_cast<SetPreferredAppModeFn>(::GetProcAddress(hUxTheme, MAKEINTRESOURCEA(135)));
			refreshPolicy = reinterpret_cast<RefreshImmersiveColorPolicyStateFn>(::GetProcAddress(hUxTheme, MAKEINTRESOURCEA(104)));
			flushMenuThemes = reinterpret_cast<FlushMenuThemesFn>(::GetProcAddress(hUxTheme, MAKEINTRESOURCEA(136)));
		}
	};

	inline DarkModeApi& GetDarkModeApi()
	{
		static DarkModeApi api;
		api.EnsureLoaded();
		return api;
	}

	inline void ApplyDarkModeToWindow(HWND hwnd, bool enable)
	{
		if (!hwnd)
			return;
		auto& api = GetDarkModeApi();
		if (api.setPreferredAppMode)
			api.setPreferredAppMode(PreferredAppMode::AllowDark);
		if (api.allowDarkModeForWindow)
			api.allowDarkModeForWindow(hwnd, enable);
		if (api.setWindowTheme)
			api.setWindowTheme(hwnd, enable ? L"DarkMode_Explorer" : L"Explorer", nullptr);
	}

	inline void ApplyImmersiveDarkTitleBarToWindow(HWND hwnd, bool enable)
	{
		if (!hwnd)
			return;
		HMODULE hDwm = ::LoadLibraryW(L"dwmapi.dll");
		if (!hDwm)
			return;

		typedef HRESULT(WINAPI* DwmSetWindowAttributeFn)(HWND, DWORD, LPCVOID, DWORD);
		auto fn = reinterpret_cast<DwmSetWindowAttributeFn>(::GetProcAddress(hDwm, "DwmSetWindowAttribute"));
		if (fn)
		{
			BOOL useDark = enable ? TRUE : FALSE;
			const DWORD DWMWA_USE_IMMERSIVE_DARK_MODE_NEW = 20;
			const DWORD DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19;
			HRESULT hr = fn(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_NEW, &useDark, sizeof(useDark));
			if (FAILED(hr))
				fn(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_OLD, &useDark, sizeof(useDark));
		}
		::FreeLibrary(hDwm);
	}

	inline void RefreshDarkModeState()
	{
		auto& api = GetDarkModeApi();
		if (api.refreshPolicy)
			api.refreshPolicy();
		if (api.flushMenuThemes)
			api.flushMenuThemes();
	}

	struct ThemedTaskDialogContext
	{
		bool darkTheme = false;
	};

	HRESULT CALLBACK ThemedTaskDialogCallback(HWND hwnd, UINT msg, WPARAM, LPARAM, LONG_PTR refData)
	{
		if (msg != TDN_CREATED)
			return S_OK;
		auto* context = reinterpret_cast<ThemedTaskDialogContext*>(refData);
		if (!context)
			return S_OK;

		ApplyDarkModeToWindow(hwnd, context->darkTheme);
		::EnumChildWindows(hwnd, [](HWND child, LPARAM lp) -> BOOL {
			auto* ctx = reinterpret_cast<ThemedTaskDialogContext*>(lp);
			if (ctx)
				ApplyDarkModeToWindow(child, ctx->darkTheme);
			return TRUE;
			}, reinterpret_cast<LPARAM>(context));
		return S_OK;
	}

	inline int ShowThemedYesNoCancelDialog(HWND parent,
		const CString& title,
		const CString& mainInstruction,
		const CString& content,
		bool darkTheme)
	{
		TASKDIALOG_BUTTON buttons[] = {
			{ IDYES, L"保存" },
			{ IDNO, L"不保存" },
			{ IDCANCEL, L"取消" }
		};

		ThemedTaskDialogContext context{};
		context.darkTheme = darkTheme;

		TASKDIALOGCONFIG config{};
		config.cbSize = sizeof(config);
		config.hwndParent = parent;
		config.hInstance = AfxGetResourceHandle();
		config.dwFlags = TDF_ALLOW_DIALOG_CANCELLATION | TDF_POSITION_RELATIVE_TO_WINDOW;
		config.pszWindowTitle = title;
		config.pszMainInstruction = mainInstruction;
		config.pszContent = content;
		config.cButtons = _countof(buttons);
		config.pButtons = buttons;
		config.nDefaultButton = IDYES;
		config.pfCallback = ThemedTaskDialogCallback;
		config.lpCallbackData = reinterpret_cast<LONG_PTR>(&context);

		int pressedButton = IDCANCEL;
		const HRESULT hr = ::TaskDialogIndirect(&config, &pressedButton, nullptr, nullptr);
		if (FAILED(hr))
			return AfxMessageBox(content, MB_YESNOCANCEL | MB_ICONWARNING);
		return pressedButton;
	}

}

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

class CAboutDlg : public CDialogEx
{
public:
	explicit CAboutDlg(bool darkTheme = false);

#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);

private:
	bool m_darkTheme = false;
	COLORREF m_bgColor = RGB(248, 248, 248);
	COLORREF m_textColor = RGB(32, 32, 32);
	CBrush m_bgBrush;

protected:
	DECLARE_MESSAGE_MAP()
};

class CUnsavedChangesDlg : public CDialogEx
{
public:
	CUnsavedChangesDlg(bool darkTheme, const CString& message, CWnd* parent = nullptr);

#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_UNSAVED_CHANGES_DIALOG };
#endif

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnYesClicked();
	afx_msg void OnNoClicked();

private:
	bool m_darkTheme = false;
	CString m_message;
	COLORREF m_bgColor = RGB(248, 248, 248);
	COLORREF m_textColor = RGB(32, 32, 32);
	CBrush m_bgBrush;

protected:
	DECLARE_MESSAGE_MAP()
};

class CSwitchModeDlg : public CDialogEx
{
public:
	CSwitchModeDlg(bool darkTheme, const CString& message, CWnd* parent = nullptr);

#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_SWITCH_MODE_DIALOG };
#endif

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnSwitchClicked();

private:
	bool m_darkTheme = false;
	CString m_message;
	COLORREF m_bgColor = RGB(248, 248, 248);
	COLORREF m_textColor = RGB(32, 32, 32);
	CBrush m_bgBrush;

protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg(bool darkTheme) : CDialogEx(IDD_ABOUTBOX), m_darkTheme(darkTheme)
{
	m_bgColor = m_darkTheme ? RGB(36, 36, 36) : RGB(248, 248, 248);
	m_textColor = m_darkTheme ? RGB(232, 232, 232) : RGB(32, 32, 32);
	m_bgBrush.CreateSolidBrush(m_bgColor);
}

BOOL CAboutDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	ApplyDarkModeToWindow(GetSafeHwnd(), m_darkTheme);
	ApplyImmersiveDarkTitleBarToWindow(GetSafeHwnd(), m_darkTheme);

	// Use explicit Unicode escapes to avoid source-encoding issues.
	CString contact(L"\u95EE\u9898\u8BF7\u8054\u7CFB\uFF1Achenfeng@leagsoft.com");
	CWnd* p = GetDlgItem(IDC_STATIC_CONTACT);
	if (p)
		p->SetWindowText(contact);

	if (CWnd* ok = GetDlgItem(IDOK))
	{
		ok->ModifyStyle(0, BS_FLAT);
		ApplyDarkModeToWindow(ok->GetSafeHwnd(), m_darkTheme);
	}
	if (CWnd* s1 = GetDlgItem(IDC_STATIC))
		ApplyDarkModeToWindow(s1->GetSafeHwnd(), m_darkTheme);
	if (CWnd* s2 = GetDlgItem(IDC_STATIC_CONTACT))
		ApplyDarkModeToWindow(s2->GetSafeHwnd(), m_darkTheme);

	Invalidate(FALSE);
	return TRUE;
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

CUnsavedChangesDlg::CUnsavedChangesDlg(bool darkTheme, const CString& message, CWnd* parent)
	: CDialogEx(IDD_UNSAVED_CHANGES_DIALOG, parent),
	m_darkTheme(darkTheme),
	m_message(message)
{
	m_bgColor = m_darkTheme ? RGB(36, 36, 36) : RGB(248, 248, 248);
	m_textColor = m_darkTheme ? RGB(232, 232, 232) : RGB(32, 32, 32);
	m_bgBrush.CreateSolidBrush(m_bgColor);
}

BOOL CUnsavedChangesDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	SetWindowText(_T("未保存修改"));
	ApplyDarkModeToWindow(GetSafeHwnd(), m_darkTheme);
	ApplyImmersiveDarkTitleBarToWindow(GetSafeHwnd(), m_darkTheme);

	if (CWnd* msg = GetDlgItem(IDC_STATIC_UNSAVED_MESSAGE))
	{
		msg->SetWindowText(m_message);
		ApplyDarkModeToWindow(msg->GetSafeHwnd(), m_darkTheme);
	}

	if (CWnd* yes = GetDlgItem(IDYES))
	{
		yes->SetWindowText(_T("保存"));
		yes->ModifyStyle(0, BS_FLAT);
		ApplyDarkModeToWindow(yes->GetSafeHwnd(), m_darkTheme);
	}
	if (CWnd* no = GetDlgItem(IDNO))
	{
		no->SetWindowText(_T("不保存"));
		no->ModifyStyle(0, BS_FLAT);
		ApplyDarkModeToWindow(no->GetSafeHwnd(), m_darkTheme);
	}
	if (CWnd* cancel = GetDlgItem(IDCANCEL))
	{
		cancel->SetWindowText(_T("取消"));
		cancel->ModifyStyle(0, BS_FLAT);
		ApplyDarkModeToWindow(cancel->GetSafeHwnd(), m_darkTheme);
	}

	Invalidate(FALSE);
	return TRUE;
}

void CUnsavedChangesDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

CSwitchModeDlg::CSwitchModeDlg(bool darkTheme, const CString& message, CWnd* parent)
	: CDialogEx(IDD_SWITCH_MODE_DIALOG, parent),
	m_darkTheme(darkTheme),
	m_message(message)
{
	m_bgColor = m_darkTheme ? RGB(36, 36, 36) : RGB(248, 248, 248);
	m_textColor = m_darkTheme ? RGB(232, 232, 232) : RGB(32, 32, 32);
	m_bgBrush.CreateSolidBrush(m_bgColor);
}

BOOL CSwitchModeDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	SetWindowText(_T("切换编辑模式"));
	ApplyDarkModeToWindow(GetSafeHwnd(), m_darkTheme);
	ApplyImmersiveDarkTitleBarToWindow(GetSafeHwnd(), m_darkTheme);

	if (CWnd* msg = GetDlgItem(IDC_STATIC_SWITCH_MODE_MESSAGE))
	{
		msg->SetWindowText(m_message);
		ApplyDarkModeToWindow(msg->GetSafeHwnd(), m_darkTheme);
	}
	if (CWnd* sw = GetDlgItem(IDYES))
	{
		sw->SetWindowText(_T("切换"));
		sw->ModifyStyle(0, BS_FLAT);
		ApplyDarkModeToWindow(sw->GetSafeHwnd(), m_darkTheme);
	}
	if (CWnd* cancel = GetDlgItem(IDCANCEL))
	{
		cancel->SetWindowText(_T("取消"));
		cancel->ModifyStyle(0, BS_FLAT);
		ApplyDarkModeToWindow(cancel->GetSafeHwnd(), m_darkTheme);
	}

	Invalidate(FALSE);
	return TRUE;
}

void CSwitchModeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
	ON_WM_CTLCOLOR()
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()

BEGIN_MESSAGE_MAP(CUnsavedChangesDlg, CDialogEx)
	ON_WM_CTLCOLOR()
	ON_WM_ERASEBKGND()
	ON_BN_CLICKED(IDYES, &CUnsavedChangesDlg::OnYesClicked)
	ON_BN_CLICKED(IDNO, &CUnsavedChangesDlg::OnNoClicked)
END_MESSAGE_MAP()

BEGIN_MESSAGE_MAP(CSwitchModeDlg, CDialogEx)
	ON_WM_CTLCOLOR()
	ON_WM_ERASEBKGND()
	ON_BN_CLICKED(IDYES, &CSwitchModeDlg::OnSwitchClicked)
END_MESSAGE_MAP()

HBRUSH CAboutDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);
	if (!pDC)
		return hbr;

	switch (nCtlColor)
	{
	case CTLCOLOR_DLG:
	case CTLCOLOR_STATIC:
		pDC->SetBkColor(m_bgColor);
		pDC->SetTextColor(m_textColor);
		return (HBRUSH)m_bgBrush.GetSafeHandle();
	default:
		break;
	}

	UNREFERENCED_PARAMETER(pWnd);
	return hbr;
}

BOOL CAboutDlg::OnEraseBkgnd(CDC* pDC)
{
	if (!pDC)
		return CDialogEx::OnEraseBkgnd(pDC);
	CRect rc;
	GetClientRect(&rc);
	pDC->FillSolidRect(&rc, m_bgColor);
	return TRUE;
}

HBRUSH CUnsavedChangesDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);
	if (!pDC)
		return hbr;

	switch (nCtlColor)
	{
	case CTLCOLOR_DLG:
	case CTLCOLOR_STATIC:
	case CTLCOLOR_BTN:
		pDC->SetBkColor(m_bgColor);
		pDC->SetTextColor(m_textColor);
		return (HBRUSH)m_bgBrush.GetSafeHandle();
	default:
		break;
	}

	UNREFERENCED_PARAMETER(pWnd);
	return hbr;
}

BOOL CUnsavedChangesDlg::OnEraseBkgnd(CDC* pDC)
{
	if (!pDC)
		return CDialogEx::OnEraseBkgnd(pDC);
	CRect rc;
	GetClientRect(&rc);
	pDC->FillSolidRect(&rc, m_bgColor);
	return TRUE;
}

void CUnsavedChangesDlg::OnYesClicked()
{
	EndDialog(IDYES);
}

void CUnsavedChangesDlg::OnNoClicked()
{
	EndDialog(IDNO);
}

HBRUSH CSwitchModeDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);
	if (!pDC)
		return hbr;

	switch (nCtlColor)
	{
	case CTLCOLOR_DLG:
	case CTLCOLOR_STATIC:
	case CTLCOLOR_BTN:
		pDC->SetBkColor(m_bgColor);
		pDC->SetTextColor(m_textColor);
		return (HBRUSH)m_bgBrush.GetSafeHandle();
	default:
		break;
	}

	UNREFERENCED_PARAMETER(pWnd);
	return hbr;
}

BOOL CSwitchModeDlg::OnEraseBkgnd(CDC* pDC)
{
	if (!pDC)
		return CDialogEx::OnEraseBkgnd(pDC);
	CRect rc;
	GetClientRect(&rc);
	pDC->FillSolidRect(&rc, m_bgColor);
	return TRUE;
}

void CSwitchModeDlg::OnSwitchClicked()
{
	EndDialog(IDYES);
}

BEGIN_MESSAGE_MAP(CMermaidOverlayWnd, CWnd)
	ON_WM_PAINT()
	ON_WM_ERASEBKGND()
	ON_WM_NCHITTEST()
	ON_WM_MOUSEWHEEL()
	ON_WM_LBUTTONDOWN()
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONUP()
	ON_WM_RBUTTONDOWN()
	ON_WM_RBUTTONUP()
END_MESSAGE_MAP()

BEGIN_MESSAGE_MAP(CTableGridOverlayCtrl, CGridCtrl)
	ON_WM_MOUSEWHEEL()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONUP()
	ON_WM_RBUTTONDOWN()
	ON_WM_RBUTTONUP()
	ON_WM_MOUSEACTIVATE()
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

BEGIN_MESSAGE_MAP(CSidebarSplitterWnd, CWnd)
	ON_WM_PAINT()
	ON_WM_ERASEBKGND()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_MOUSEMOVE()
	ON_WM_SETCURSOR()
END_MESSAGE_MAP()

void CSidebarSplitterWnd::OnPaint()
{
	CPaintDC dc(this);
	if (!m_host)
	{
		Default();
		return;
	}
	CRect rc;
	GetClientRect(&rc);
	COLORREF bg = m_host->m_themeSidebarBackground;
	COLORREF border = m_host->m_themeDark ? RGB(70, 70, 70) : RGB(210, 210, 210);
	dc.FillSolidRect(&rc, bg);
	CPen pen(PS_SOLID, 1, border);
	CPen* old = dc.SelectObject(&pen);
	int x = rc.Width() / 2;
	dc.MoveTo(x, rc.top);
	dc.LineTo(x, rc.bottom);
	dc.SelectObject(old);
}

BOOL CSidebarSplitterWnd::OnEraseBkgnd(CDC* /*pDC*/)
{
	// We fully paint in OnPaint().
	return TRUE;
}

BEGIN_MESSAGE_MAP(CSidebarTabCtrl, CTabCtrl)
	ON_WM_ERASEBKGND()
	ON_WM_NCPAINT()
	ON_WM_PAINT()
	ON_WM_LBUTTONDOWN()
END_MESSAGE_MAP()

BEGIN_MESSAGE_MAP(CSidebarSearchEdit, CEdit)
	ON_WM_NCPAINT()
	ON_WM_PAINT()
END_MESSAGE_MAP()

void CMermaidOverlayWnd::OnPaint()
{
	CPaintDC dc(this);
	if (m_host)
	{
		m_host->DrawMermaidOverlays(dc);
		m_host->DrawHorizontalRules(dc);
		m_host->DrawTocOverlay(dc);
	}
}

BOOL CMermaidOverlayWnd::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	// 表格覆盖层是该窗口的子控件，将通知转发给宿主对话框处理。
	if (m_host && ::IsWindow(m_host->GetSafeHwnd()))
	{
		LRESULT forwarded = m_host->SendMessage(WM_NOTIFY, wParam, lParam);
		if (pResult != nullptr)
			*pResult = forwarded;
		return TRUE;
	}
	return CWnd::OnNotify(wParam, lParam, pResult);
}

BOOL CMermaidOverlayWnd::OnEraseBkgnd(CDC* /*pDC*/)
{
	// Transparent-like overlay: let host draw only diagrams.
	return TRUE;
}

LRESULT CMermaidOverlayWnd::OnNcHitTest(CPoint /*point*/)
{
	// Let mouse events pass through.
	return HTTRANSPARENT;
}

BOOL CMermaidOverlayWnd::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	// Forward mouse wheel to the underlying editor to prevent accidental horizontal scrolling
	// and sidebar jitter when the overlay is visible.
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		m_host->m_editContent.SendMessage(WM_MOUSEWHEEL, MAKEWPARAM(nFlags, static_cast<WORD>(zDelta)), MAKELPARAM(pt.x, pt.y));
		return TRUE;
	}
	return CWnd::OnMouseWheel(nFlags, zDelta, pt);
}

void CMermaidOverlayWnd::OnLButtonDown(UINT nFlags, CPoint point)
{
	// Recover from a previously missed LButtonUp (for example, mouse released outside overlay).
	if (m_tocClickSuppressUp)
		m_tocClickSuppressUp = false;

	if (m_host && m_host->HandleTocOverlayClick(point))
	{
		m_tocClickSuppressUp = true;
		return;
	}

	// Forward clicks so selection does not accidentally land on hidden mermaid source.
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		ClientToScreen(&point);
		m_host->m_editContent.ScreenToClient(&point);
		m_host->m_editContent.SendMessage(WM_LBUTTONDOWN, nFlags, MAKELPARAM(point.x, point.y));
		SetCapture();
		return;
	}
	CWnd::OnLButtonDown(nFlags, point);
}

void CMermaidOverlayWnd::OnMouseMove(UINT nFlags, CPoint point)
{
	// Continue forwarding drag events so the editor keeps selection stable.
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		ClientToScreen(&point);
		m_host->m_editContent.ScreenToClient(&point);
		m_host->m_editContent.SendMessage(WM_MOUSEMOVE, nFlags, MAKELPARAM(point.x, point.y));
		return;
	}
	CWnd::OnMouseMove(nFlags, point);
}

void CMermaidOverlayWnd::OnLButtonUp(UINT nFlags, CPoint point)
{
	if (m_tocClickSuppressUp)
	{
		m_tocClickSuppressUp = false;
		if (GetCapture() == this)
			ReleaseCapture();
		return;
	}

	if (GetCapture() == this)
		ReleaseCapture();
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		ClientToScreen(&point);
		m_host->m_editContent.ScreenToClient(&point);
		m_host->m_editContent.SendMessage(WM_LBUTTONUP, nFlags, MAKELPARAM(point.x, point.y));
		return;
	}
	CWnd::OnLButtonUp(nFlags, point);
}

void CMermaidOverlayWnd::OnRButtonDown(UINT nFlags, CPoint point)
{
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		ClientToScreen(&point);
		m_host->m_editContent.ScreenToClient(&point);
		m_host->m_editContent.SendMessage(WM_RBUTTONDOWN, nFlags, MAKELPARAM(point.x, point.y));
		return;
	}
	CWnd::OnRButtonDown(nFlags, point);
}

void CMermaidOverlayWnd::OnRButtonUp(UINT nFlags, CPoint point)
{
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		ClientToScreen(&point);
		m_host->m_editContent.ScreenToClient(&point);
		m_host->m_editContent.SendMessage(WM_RBUTTONUP, nFlags, MAKELPARAM(point.x, point.y));
		return;
	}
	CWnd::OnRButtonUp(nFlags, point);
}

void CTableGridOverlayCtrl::ForwardMouseToEditor(UINT msg, UINT nFlags, CPoint point)
{
	if (!m_host || !::IsWindow(m_host->m_editContent.GetSafeHwnd()))
		return;
	ClientToScreen(&point);
	m_host->m_editContent.ScreenToClient(&point);
	m_host->m_editContent.SendMessage(msg, static_cast<WPARAM>(nFlags), MAKELPARAM(point.x, point.y));
}

BOOL CTableGridOverlayCtrl::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		m_host->m_editContent.SendMessage(
			WM_MOUSEWHEEL,
			MAKEWPARAM(nFlags, static_cast<WORD>(zDelta)),
			MAKELPARAM(pt.x, pt.y));
		return TRUE;
	}
	return CGridCtrl::OnMouseWheel(nFlags, zDelta, pt);
}

void CTableGridOverlayCtrl::OnLButtonDown(UINT nFlags, CPoint point)
{
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		m_host->m_editContent.SetFocus();
		ForwardMouseToEditor(WM_LBUTTONDOWN, nFlags, point);
		return;
	}
	CGridCtrl::OnLButtonDown(nFlags, point);
}

void CTableGridOverlayCtrl::OnLButtonDblClk(UINT nFlags, CPoint point)
{
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		m_host->m_editContent.SetFocus();
		ForwardMouseToEditor(WM_LBUTTONDBLCLK, nFlags, point);
		return;
	}
	CGridCtrl::OnLButtonDblClk(nFlags, point);
}

void CTableGridOverlayCtrl::OnMouseMove(UINT nFlags, CPoint point)
{
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		ForwardMouseToEditor(WM_MOUSEMOVE, nFlags, point);
		return;
	}
	CGridCtrl::OnMouseMove(nFlags, point);
}

void CTableGridOverlayCtrl::OnLButtonUp(UINT nFlags, CPoint point)
{
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		ForwardMouseToEditor(WM_LBUTTONUP, nFlags, point);
		return;
	}
	CGridCtrl::OnLButtonUp(nFlags, point);
}

void CTableGridOverlayCtrl::OnRButtonDown(UINT nFlags, CPoint point)
{
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		m_host->m_editContent.SetFocus();
		ForwardMouseToEditor(WM_RBUTTONDOWN, nFlags, point);
		return;
	}
	CGridCtrl::OnRButtonDown(nFlags, point);
}

void CTableGridOverlayCtrl::OnRButtonUp(UINT nFlags, CPoint point)
{
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
	{
		ForwardMouseToEditor(WM_RBUTTONUP, nFlags, point);
		return;
	}
	CGridCtrl::OnRButtonUp(nFlags, point);
}

int CTableGridOverlayCtrl::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message)
{
	UNREFERENCED_PARAMETER(pDesktopWnd);
	UNREFERENCED_PARAMETER(nHitTest);
	UNREFERENCED_PARAMETER(message);
	return MA_NOACTIVATE;
}

void CTableGridOverlayCtrl::OnSetFocus(CWnd* pOldWnd)
{
	CGridCtrl::OnSetFocus(pOldWnd);
	if (m_host && ::IsWindow(m_host->m_editContent.GetSafeHwnd()))
		m_host->m_editContent.SetFocus();
}

void CSidebarSplitterWnd::OnLButtonDown(UINT nFlags, CPoint point)
{
	UNREFERENCED_PARAMETER(point);
	// 侧栏宽度固定，不允许拖拽调整。
	CWnd::OnLButtonDown(nFlags, point);
}

void CSidebarSplitterWnd::OnLButtonUp(UINT nFlags, CPoint point)
{
	UNREFERENCED_PARAMETER(point);
	if (m_dragging)
	{
		m_dragging = false;
		ReleaseCapture();
		if (m_host)
			m_host->EndSidebarResize();
	}
	CWnd::OnLButtonUp(nFlags, point);
}

void CSidebarSplitterWnd::OnMouseMove(UINT nFlags, CPoint point)
{
	UNREFERENCED_PARAMETER(point);
	CWnd::OnMouseMove(nFlags, point);
}

BOOL CSidebarTabCtrl::OnEraseBkgnd(CDC* pDC)
{
	if (pDC && m_host)
	{
		CRect rc;
		GetClientRect(&rc);
		// Fill the whole control so the body area never shows stale (white) blocks.
		pDC->FillSolidRect(&rc, m_host->m_themeSidebarBackground);
		return TRUE;
	}
	return CTabCtrl::OnEraseBkgnd(pDC);
}

void CSidebarTabCtrl::OnNcPaint()
{
	Default();
	if (!m_host)
		return;
	CWindowDC dc(this);
	CRect rc;
	GetWindowRect(&rc);
	rc.OffsetRect(-rc.left, -rc.top);
	COLORREF border = m_host->m_themeDark ? RGB(70, 70, 70) : RGB(210, 210, 210);
	CPen pen(PS_SOLID, 1, border);
	CPen* old = dc.SelectObject(&pen);
	CBrush* oldBr = (CBrush*)dc.SelectStockObject(NULL_BRUSH);
	dc.Rectangle(&rc);
	dc.SelectObject(oldBr);
	dc.SelectObject(old);
}

void CSidebarTabCtrl::OnPaint()
{
	CPaintDC dc(this);
	if (!m_host)
	{
		Default();
		return;
	}

	CRect rc;
	GetClientRect(&rc);
	if (rc.Width() <= 0 || rc.Height() <= 0)
		return;

	// Paint full background (safe: this control stays behind the TreeView).
	dc.FillSolidRect(&rc, m_host->m_themeSidebarBackground);

	const int tabHeaderHeight = max(1, min(26, rc.Height()));
	CRect headerRc = rc;
	headerRc.bottom = min(headerRc.bottom, headerRc.top + tabHeaderHeight);

	CDC memDC;
	memDC.CreateCompatibleDC(&dc);
	CBitmap bmp;
	bmp.CreateCompatibleBitmap(&dc, headerRc.Width(), headerRc.Height());
	CBitmap* oldBmp = memDC.SelectObject(&bmp);

	CRect localRc(0, 0, headerRc.Width(), headerRc.Height());
	memDC.FillSolidRect(&localRc, m_host->m_themeSidebarBackground);
	memDC.SetBkMode(TRANSPARENT);
	CFont* oldFont = nullptr;
	if (m_host->m_fontSidebar.GetSafeHandle())
		oldFont = memDC.SelectObject(&m_host->m_fontSidebar);
	const int tabCount = GetItemCount();
	int current = m_host ? m_host->m_sidebarActiveTab : GetCurSel();
	if (current < 0)
		current = 0;
	if (tabCount > 0 && current >= tabCount)
		current = tabCount - 1;

	COLORREF border = m_host->m_themeDark ? RGB(70, 70, 70) : RGB(210, 210, 210);
	COLORREF accent = m_host->m_themeDark ? RGB(80, 160, 255) : RGB(0, 120, 215);
	COLORREF selectedBg = m_host->m_themeDark ? RGB(38, 38, 38) : RGB(255, 255, 255);
	COLORREF unselectedBg = m_host->m_themeSidebarBackground;
	COLORREF selectedText = m_host->m_themeDark ? RGB(240, 240, 240) : RGB(30, 30, 30);
	COLORREF unselectedText = m_host->m_themeDark ? RGB(200, 200, 200) : RGB(80, 80, 80);

	if (tabCount > 0)
	{
		int tabWidth = max(1, localRc.Width() / tabCount);
		for (int i = 0; i < tabCount; ++i)
		{
			CRect tabRc(localRc.left + i * tabWidth, localRc.top,
				(i == tabCount - 1) ? localRc.right : (localRc.left + (i + 1) * tabWidth),
				localRc.top + tabHeaderHeight);
			const bool selected = (i == current);
			memDC.FillSolidRect(&tabRc, selected ? selectedBg : unselectedBg);

			CString text;
			TCHAR buf[64] = { 0 };
			TCITEM item{};
			item.mask = TCIF_TEXT;
			item.pszText = buf;
			item.cchTextMax = _countof(buf);
			if (GetItem(i, &item))
				text = buf;
			if (text.IsEmpty())
			{
				// Fallback: prevent blank headers if the control fails to return text.
				text = (i == 0) ? _T("文件") : _T("大纲");
			}

			memDC.SetTextColor(selected ? selectedText : unselectedText);
			memDC.SetBkMode(OPAQUE);
			memDC.SetBkColor(selected ? selectedBg : unselectedBg);
			memDC.DrawText(text, &tabRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
			memDC.SetBkMode(TRANSPARENT);

			if (selected)
			{
				CPen pen(PS_SOLID, 2, accent);
				CPen* oldPen = memDC.SelectObject(&pen);
				memDC.MoveTo(tabRc.left + 10, tabRc.bottom - 2);
				memDC.LineTo(tabRc.right - 10, tabRc.bottom - 2);
				memDC.SelectObject(oldPen);
			}

			if (i != tabCount - 1)
			{
				CPen pen(PS_SOLID, 1, border);
				CPen* oldPen = memDC.SelectObject(&pen);
				memDC.MoveTo(tabRc.right, tabRc.top + 6);
				memDC.LineTo(tabRc.right, tabRc.bottom - 6);
				memDC.SelectObject(oldPen);
			}
		}
	}

	{
		CPen pen(PS_SOLID, 1, border);
		CPen* oldPen = memDC.SelectObject(&pen);
		int y = localRc.top + tabHeaderHeight - 1;
		memDC.MoveTo(localRc.left + 1, y);
		memDC.LineTo(localRc.right - 1, y);
		memDC.SelectObject(oldPen);
	}

	if (oldFont)
		memDC.SelectObject(oldFont);

	// Copy only the header area to screen so we don't paint over the TreeView.
	dc.BitBlt(0, 0, headerRc.Width(), headerRc.Height(), &memDC, 0, 0, SRCCOPY);
	memDC.SelectObject(oldBmp);
}

void CSidebarTabCtrl::OnLButtonDown(UINT nFlags, CPoint point)
{
	if (!m_host)
	{
		CTabCtrl::OnLButtonDown(nFlags, point);
		return;
	}

	// Make the whole header area clickable like a segmented control.
	const int tabHeaderHeight = 26;
	if (point.y >= 0 && point.y < tabHeaderHeight)
	{
		CRect rc;
		GetClientRect(&rc);
		int count = GetItemCount();
		if (count > 0 && rc.Width() > 0)
		{
			int index = (point.x * count) / rc.Width();
			if (index < 0)
				index = 0;
			if (index >= count)
				index = count - 1;
			if (index != GetCurSel())
			{
				SetCurSel(index);
				m_host->SetSidebarTab(index);
				Invalidate(FALSE);
				return;
			}
		}
	}

	CTabCtrl::OnLButtonDown(nFlags, point);
}

void CSidebarTabCtrl::DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	UNREFERENCED_PARAMETER(lpDrawItemStruct);
	// We currently paint the tab header ourselves in OnPaint().
}

void CSidebarSearchEdit::OnNcPaint()
{
	Default();
	if (!m_host)
		return;
	CWindowDC dc(this);
	CRect rc;
	GetWindowRect(&rc);
	rc.OffsetRect(-rc.left, -rc.top);
	COLORREF border = m_host->m_themeDark ? RGB(70, 70, 70) : RGB(210, 210, 210);
	CPen pen(PS_SOLID, 1, border);
	CPen* old = dc.SelectObject(&pen);
	CBrush* oldBr = (CBrush*)dc.SelectStockObject(NULL_BRUSH);
	dc.Rectangle(&rc);
	dc.SelectObject(oldBr);
	dc.SelectObject(old);
}

void CSidebarSearchEdit::OnPaint()
{
	CPaintDC dc(this);
	// Let the edit control paint its text/caret.
	DefWindowProc(WM_PAINT, (WPARAM)dc.m_hDC, 0);
	if (!m_host)
		return;
	// Draw a flat 1px border inside client area to ensure the line border is always visible.
	CRect rc;
	GetClientRect(&rc);
	COLORREF border = m_host->m_themeDark ? RGB(70, 70, 70) : RGB(210, 210, 210);
	CPen pen(PS_SOLID, 1, border);
	CPen* old = dc.SelectObject(&pen);
	CBrush* oldBr = (CBrush*)dc.SelectStockObject(NULL_BRUSH);
	dc.Rectangle(&rc);
	dc.SelectObject(oldBr);
	dc.SelectObject(old);
}

BOOL CSidebarSplitterWnd::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message)
{
	UNREFERENCED_PARAMETER(pWnd);
	UNREFERENCED_PARAMETER(nHitTest);
	UNREFERENCED_PARAMETER(message);
	::SetCursor(::LoadCursor(nullptr, IDC_ARROW));
	return TRUE;
}

CMarkdownEditorDlg::CMarkdownEditorDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_MARKDOWNEDITOR_DIALOG, pParent)
	, m_bMarkdownMode(TRUE)
	, m_strCurrentFile(_T(""))
	, m_currentFileName(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMarkdownEditorDlg::SetStartupOpenFilePath(const CString& filePath)
{
	m_startupOpenFilePath = filePath;
	m_startupOpenFilePath.Trim();
}

void CMarkdownEditorDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_CONTENT, m_editContent);
	DDX_Control(pDX, IDC_BUTTON_OPEN, m_btnOpenFile);
	DDX_Control(pDX, IDC_BUTTON_MODE, m_btnModeSwitch);
	DDX_Control(pDX, IDC_BUTTON_THEME, m_btnThemeSwitch);
	DDX_Control(pDX, IDC_STATIC_STATUS, m_staticStatus);
	DDX_Control(pDX, IDC_PROGRESS_STATUS, m_progressStatus);
	DDX_Control(pDX, IDC_STATIC_PROGRESS_TEXT, m_staticProgressText);
	DDX_Control(pDX, IDC_STATIC_OVERLAY, m_staticOverlay);
	DDX_Control(pDX, IDC_TAB_SIDEBAR, m_tabSidebar);
	DDX_Control(pDX, IDC_TREE_SIDEBAR, m_treeSidebar);
	DDX_Control(pDX, IDC_EDIT_FILE_SEARCH, m_editFileSearch);
}

BEGIN_MESSAGE_MAP(CMarkdownEditorDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_OPEN, &CMarkdownEditorDlg::OnBnClickedOpenFile)
	ON_BN_CLICKED(IDC_BUTTON_MODE, &CMarkdownEditorDlg::OnBnClickedModeSwitch)
	ON_BN_CLICKED(IDC_BUTTON_THEME, &CMarkdownEditorDlg::OnBnClickedThemeSwitch)
	ON_EN_CHANGE(IDC_EDIT_CONTENT, &CMarkdownEditorDlg::OnEnChangeEdit)
	ON_WM_TIMER()
	ON_WM_SIZE()
	ON_WM_ENTERSIZEMOVE()
	ON_WM_EXITSIZEMOVE()
	ON_WM_DROPFILES()
	ON_WM_CTLCOLOR()
	ON_WM_ERASEBKGND()
	ON_WM_DESTROY()
	ON_WM_CLOSE()
	ON_NOTIFY(EN_LINK, IDC_EDIT_CONTENT, &CMarkdownEditorDlg::OnEditContentLink)
	ON_NOTIFY(EN_SELCHANGE, IDC_EDIT_CONTENT, &CMarkdownEditorDlg::OnEditContentSelChanged)
	ON_CONTROL(EN_VSCROLL, IDC_EDIT_CONTENT, &CMarkdownEditorDlg::OnEditContentVScroll)
	ON_CONTROL(EN_HSCROLL, IDC_EDIT_CONTENT, &CMarkdownEditorDlg::OnEditContentHScroll)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB_SIDEBAR, &CMarkdownEditorDlg::OnSidebarTabChanged)
	ON_NOTIFY(TVN_SELCHANGED, IDC_TREE_SIDEBAR, &CMarkdownEditorDlg::OnSidebarItemChanged)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_TREE_SIDEBAR, &CMarkdownEditorDlg::OnSidebarTreeCustomDraw)
	ON_NOTIFY(NM_CLICK, IDC_TREE_SIDEBAR, &CMarkdownEditorDlg::OnSidebarTreeClick)
	ON_NOTIFY(NM_RCLICK, IDC_TREE_SIDEBAR, &CMarkdownEditorDlg::OnSidebarTreeRClick)
	ON_EN_CHANGE(IDC_EDIT_FILE_SEARCH, &CMarkdownEditorDlg::OnEnChangeFileSearch)
	ON_WM_DRAWITEM()
	ON_MESSAGE(WM_INIT_DRAGDROP, OnInitDragDrop)
	ON_MESSAGE(WM_MARKDOWN_PARSE_PROGRESS, &CMarkdownEditorDlg::OnMarkdownParseProgress)
	ON_MESSAGE(WM_MARKDOWN_PARSE_COMPLETE, &CMarkdownEditorDlg::OnMarkdownParseComplete)
	ON_MESSAGE(WM_DOCUMENT_LOAD_COMPLETE, &CMarkdownEditorDlg::OnDocumentLoadComplete)
	ON_MESSAGE(WM_DOCUMENT_LOAD_PROGRESS, &CMarkdownEditorDlg::OnDocumentLoadProgress)
	ON_MESSAGE(WM_PREPARE_DISPLAY_TEXT_COMPLETE, &CMarkdownEditorDlg::OnPrepareDisplayTextComplete)
	ON_MESSAGE(WM_OPEN_STARTUP_FILE, &CMarkdownEditorDlg::OnOpenStartupFile)
	ON_REGISTERED_MESSAGE(kFindReplaceMsg, &CMarkdownEditorDlg::OnFindReplaceCmd)
END_MESSAGE_MAP()

BOOL CMarkdownEditorDlg::OnEraseBkgnd(CDC* pDC)
{
	// Reduce flicker during live window resize: skip background erase and let child controls repaint.
	if (m_windowSizingMove)
		return TRUE;

	// We still fill the dialog background to prevent "stale" blocks after theme switch.
	if (pDC)
	{
		CRect rc;
		GetClientRect(&rc);
		pDC->FillSolidRect(&rc, m_themeBackground);
	}
	return TRUE;
}

BOOL CMarkdownEditorDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	// Avoid sibling controls painting over each other (tab/tree overlap region).
	ModifyStyle(0, WS_CLIPCHILDREN | WS_CLIPSIBLINGS);

	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, kSysCmdSetDefaultMd, _T("设为默认 Markdown 打开工具"));
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	SetIcon(m_hIcon, TRUE);
	SetIcon(m_hIcon, FALSE);

	// Set window title
	SetWindowText(_T("MyTypora - Markdown编辑器"));

	// Set button text
	m_btnOpenFile.SetWindowText(_T("打开"));
	m_btnModeSwitch.SetWindowText(_T("Markdown模式"));
	m_btnThemeSwitch.SetWindowText(_T("主题"));

	// Sidebar defaults
	m_fileSearchQuery.Empty();
	m_outlineSearchQuery.Empty();
	m_sidebarWidth = 240;

	// Typography (Typora-like density)
	{
		CClientDC dc(this);
		int dpiY = dc.GetDeviceCaps(LOGPIXELSY);
		int fontHeight9 = -MulDiv(9, dpiY, 72);
		int fontHeight10 = -MulDiv(10, dpiY, 72);
		int fontHeight16 = -MulDiv(16, dpiY, 72);

		LOGFONT lf{};
		lf.lfHeight = fontHeight9;
		lf.lfWeight = FW_NORMAL;
		wcscpy_s(lf.lfFaceName, L"Microsoft YaHei UI");
		m_fontSidebar.DeleteObject();
		m_fontSidebar.CreateFontIndirect(&lf);

		lf.lfHeight = fontHeight10;
		m_fontButton.DeleteObject();
		m_fontButton.CreateFontIndirect(&lf);

		lf = LOGFONT{};
		lf.lfHeight = fontHeight16;
		lf.lfWeight = FW_LIGHT;
		lf.lfQuality = CLEARTYPE_QUALITY;
		wcscpy_s(lf.lfFaceName, L"Microsoft YaHei UI");
		m_fontOverlay.DeleteObject();
		m_fontOverlay.CreateFontIndirect(&lf);

	if (::IsWindow(m_tabSidebar.GetSafeHwnd()))
	{
		m_tabSidebar.SetFont(&m_fontSidebar);
		m_tabSidebar.SetHost(this);
		m_tabSidebar.ModifyStyle(0, WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
		m_tabSidebar.SetWindowPos(nullptr, 0, 0, 0, 0,
			SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
	}
	if (::IsWindow(m_editFileSearch.GetSafeHwnd()))
	{
		m_editFileSearch.SetFont(&m_fontSidebar);
		m_editFileSearch.SetHost(this);
		m_editFileSearch.ModifyStyle(0, WS_CLIPSIBLINGS);
		// Remove classic 3D borders; we draw a flat border in CSidebarSearchEdit::OnNcPaint.
		m_editFileSearch.ModifyStyle(WS_BORDER, 0, SWP_FRAMECHANGED);
		m_editFileSearch.ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_FRAMECHANGED);
		m_editFileSearch.SetWindowPos(nullptr, 0, 0, 0, 0,
			SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
	}
	if (::IsWindow(m_treeSidebar.GetSafeHwnd()))
	{
		m_treeSidebar.SetFont(&m_fontSidebar);
		m_treeSidebar.SendMessage(TVM_SETITEMHEIGHT, 20, 0);
		m_treeSidebar.ModifyStyle(TVS_HASLINES | TVS_LINESATROOT, 0);
		m_treeSidebar.ModifyStyle(0, TVS_NOHSCROLL);
		m_treeSidebar.ModifyStyle(0, TVS_FULLROWSELECT);
		// Keep the selection highlight visible even when focus is on the editor.
		m_treeSidebar.ModifyStyle(0, TVS_SHOWSELALWAYS);
		m_treeSidebar.ModifyStyle(0, WS_CLIPSIBLINGS | WS_CLIPCHILDREN);
		// Enable double-buffering to avoid partial paint when switching tabs.
		m_treeSidebar.SendMessage(TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER);
	}

		if (::IsWindow(m_btnOpenFile.GetSafeHwnd()))
			m_btnOpenFile.SetFont(&m_fontButton);
		if (::IsWindow(m_btnModeSwitch.GetSafeHwnd()))
			m_btnModeSwitch.SetFont(&m_fontButton);
		if (::IsWindow(m_btnThemeSwitch.GetSafeHwnd()))
			m_btnThemeSwitch.SetFont(&m_fontButton);

		m_btnOpenFile.ModifyStyle(0, BS_FLAT);
		m_btnModeSwitch.ModifyStyle(0, BS_FLAT);
		m_btnThemeSwitch.ModifyStyle(0, BS_FLAT);
	}

	// Set edit control properties
	m_editContent.SetFont(GetFont());
	if (m_wrapTargetHdc == nullptr)
		m_wrapTargetHdc = ::GetDC(nullptr);
	// Enable auto wrap in both markdown/text mode; width will follow editor client size.
	m_editContent.ModifyStyle(WS_HSCROLL | ES_AUTOHSCROLL, 0, SWP_FRAMECHANGED);
	UpdateEditorWrapWidth();
	{
		// Runtime fallback switch:
		// 1 / true / on  => enable wrap-safe block marker elision (default)
		// 0 / false / off => disable and fallback to legacy hidden path
		TCHAR envValue[16] = { 0 };
		const DWORD envLen = ::GetEnvironmentVariable(_T("MYTYPORA_WRAP_SAFE_BLOCK_ELISION"),
			envValue, _countof(envValue));
		if (envLen > 0 && envLen < _countof(envValue))
		{
			CString value(envValue);
			value.Trim();
			value.MakeLower();
			if (value == _T("0") || value == _T("false") || value == _T("off") || value == _T("no"))
				m_enableWrapSafeBlockMarkerElision = false;
			else
				m_enableWrapSafeBlockMarkerElision = true;
		}
	}
	// RichEdit typography: enable advanced typography where supported (may improve symbol/emoji fallback).
	m_editContent.SendMessage(EM_SETTYPOGRAPHYOPTIONS, TO_ADVANCEDTYPOGRAPHY, TO_ADVANCEDTYPOGRAPHY);
	if (!::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
	{
		m_mermaidOverlay.SetHost(this);
		// Overlay sits on top of edit; it paints diagrams in-place.
		m_mermaidOverlay.CreateEx(WS_EX_TRANSPARENT, AfxRegisterWndClass(0), _T(""),
			WS_CHILD | WS_VISIBLE, CRect(0, 0, 0, 0), this, 0);
	}
	if (!::IsWindow(m_sidebarSplitter.GetSafeHwnd()))
	{
		m_sidebarSplitter.SetHost(this);
		m_sidebarSplitter.CreateEx(0, AfxRegisterWndClass(0), _T(""),
			WS_CHILD | WS_VISIBLE, CRect(0, 0, 0, 0), this, 0);
	}
	if (CWinApp* app = AfxGetApp())
		m_themeDark = (app->GetProfileInt(_T("Settings"), _T("ThemeDark"), 0) != 0);
	else
		m_themeDark = false;
	ApplyTheme(m_themeDark);
	if (::IsWindow(m_staticOverlay.GetSafeHwnd()))
	{
		m_staticOverlay.ShowWindow(SW_HIDE);
		m_staticOverlay.SetWindowText(_T(""));
		if (m_fontOverlay.GetSafeHandle() != nullptr)
			m_staticOverlay.SetFont(&m_fontOverlay);
	}

	DWORD eventMask = m_editContent.GetEventMask();
	// NOTE: ENM_CHANGE is required for RichEdit to emit EN_CHANGE notifications.
	// Dirty-state tracking (title '*', status "编辑中...", unsaved confirm) depends on it.
	m_editContent.SetEventMask(eventMask | ENM_CHANGE | ENM_LINK | ENM_MOUSEEVENTS | ENM_SELCHANGE);
	// Allow large documents
	m_editContent.LimitText(0);

	// Initialize status
	if (::IsWindow(m_progressStatus.GetSafeHwnd()))
	{
		m_progressStatus.SetRange32(0, 100);
		m_progressStatus.SetPos(0);
		m_progressStatus.ShowWindow(SW_HIDE);
	}
	if (::IsWindow(m_staticProgressText.GetSafeHwnd()))
	{
		m_staticProgressText.SetWindowText(_T(""));
		m_staticProgressText.ShowWindow(SW_HIDE);
	}
	m_statusProgressVisible = false;
	m_lastUserActivityTick = ::GetTickCount();
	SetTimer(kAutoSaveTimerId, 1000, NULL);
	SetTimer(kFileWatchTimerId, 1500, NULL);
	UpdateStatusText();
	SetRenderBusy(false, _T(""));
	UpdateSidebarInteractivity();
	LoadRecentFiles();
	InitSidebar();
	UpdateSidebar();
	if (::IsWindow(m_treeSidebar.GetSafeHwnd()))
		m_treeSidebar.Invalidate(FALSE);

	CRect clientRect;
	GetClientRect(&clientRect);
	UpdateLayout(clientRect.Width(), clientRect.Height());
	UpdateMermaidOverlayRegion(true);
	// Ensure everything is repainted in the final layout positions.
	RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);

	PostMessage(WM_INIT_DRAGDROP, 0, 0);
	// Ensure OLE is initialized for RichEdit (for embedded objects like diagrams).
	::OleInitialize(nullptr);
	{
		Gdiplus::GdiplusStartupInput input;
		Gdiplus::GdiplusStartup(&m_gdiplusToken, &input, nullptr);
	}
	if (!m_startupOpenFilePath.IsEmpty())
		PostMessage(WM_OPEN_STARTUP_FILE, 0, 0);
	PromptSetDefaultMdAssociationIfNeeded();

	return TRUE;
}

LRESULT CMarkdownEditorDlg::OnOpenStartupFile(WPARAM, LPARAM)
{
	if (m_startupOpenFilePath.IsEmpty())
		return 0;

	CString filePath = m_startupOpenFilePath;
	m_startupOpenFilePath.Empty();
	OpenDocument(filePath);
	return 0;
}

void CMarkdownEditorDlg::UpdateWindowTitle()
{
	CString baseTitle = _T("MyTypora - Markdown编辑器");
	CString fileName = m_currentFileName;
	if (fileName.IsEmpty() && !m_strCurrentFile.IsEmpty())
		fileName = ExtractFileNameFromPath(m_strCurrentFile);

	if (fileName.IsEmpty())
	{
		CString title = baseTitle;
		if (m_documentDirty)
			title = _T("* ") + title;
		SetWindowText(title);
		return;
	}

	CString title;
	title.Format(_T("%s%s - %s"),
		(LPCTSTR)fileName,
		m_documentDirty ? _T(" *") : _T(""),
		(LPCTSTR)baseTitle);
	SetWindowText(title);
}

void CMarkdownEditorDlg::MarkDocumentDirty(bool dirty)
{
	if (dirty)
	{
		m_autoSavedStatusVisible = false;
		m_editingHintVisible = true;
	}
	else if (!m_autoSavedStatusVisible)
	{
		m_editingHintVisible = false;
	}

	const bool changed = (m_documentDirty != dirty);
	m_documentDirty = dirty;
	if (changed)
		UpdateWindowTitle();
	UpdateStatusText(true);
}

void CMarkdownEditorDlg::RefreshDirtyStateFromEditor()
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return;

	// Markdown mode is preview-only; skip diff-based dirty recalculation there.
	if (m_bMarkdownMode)
		return;

	CString content;
	m_editContent.GetWindowText(content);
	const std::wstring current(content.GetString());

	bool dirty = false;
	if (current != m_rawDocumentText)
	{
		const std::wstring currentNormalized = NormalizeLineEndingsForCompare(current);
		const std::wstring baselineNormalized = NormalizeLineEndingsForCompare(m_rawDocumentText);
		dirty = (currentNormalized != baselineNormalized);
	}

	m_textDirty = dirty;
	MarkDocumentDirty(dirty);
}

bool CMarkdownEditorDlg::PromptSwitchToTextModeForEdit()
{
	if (!m_bMarkdownMode)
		return false;
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return false;
	if ((m_editContent.GetStyle() & ES_READONLY) == 0)
		return false;

	const DWORD now = ::GetTickCount();
	if (m_lastReadonlyEditPromptTick != 0 && (now - m_lastReadonlyEditPromptTick) < 800)
		return true;
	m_lastReadonlyEditPromptTick = now;

	CSwitchModeDlg dlg(
		m_themeDark,
		_T("当前是 Markdown 预览模式，不能直接编辑。\r\n是否切换到“文本模式”继续编辑？"),
		this);
	const int ret = dlg.DoModal();
	if (ret == IDYES)
	{
		OnBnClickedModeSwitch();
	}
	return true;
}

void CMarkdownEditorDlg::NoteUserActivity(bool clearAutoSavedNotice)
{
	m_lastUserActivityTick = ::GetTickCount();
	if (clearAutoSavedNotice && m_autoSavedStatusVisible)
	{
		m_autoSavedStatusVisible = false;
		m_editingHintVisible = true;
		UpdateStatusText(true);
	}
}

void CMarkdownEditorDlg::TryAutoSaveIfIdle()
{
	if (!m_documentDirty)
		return;
	if (m_strCurrentFile.IsEmpty())
		return;
	if (m_loadingActive || m_parsingActive || m_formattingActive || m_textModeResetActive)
		return;

	DWORD now = ::GetTickCount();
	if (m_lastUserActivityTick == 0)
	{
		m_lastUserActivityTick = now;
		return;
	}
	if (now - m_lastUserActivityTick < kAutoSaveIdleMs)
		return;

	CString error;
	if (!SaveCurrentDocument(false, &error))
	{
		// Avoid repeated prompt storms when target path is temporarily unwritable.
		m_lastUserActivityTick = now;
		if (!error.IsEmpty() && (m_lastAutoSaveErrorPromptTick == 0 ||
			(now - m_lastAutoSaveErrorPromptTick) >= kAutoSaveErrorPromptIntervalMs))
		{
			m_lastAutoSaveErrorPromptTick = now;
			CString msg;
			msg.Format(_T("自动保存失败：%s"), (LPCTSTR)error);
			AfxMessageBox(msg, MB_OK | MB_ICONWARNING);
		}
		return;
	}

	m_lastAutoSaveErrorPromptTick = 0;
	m_autoSavedStatusVisible = true;
	m_editingHintVisible = false;
	m_lastUserActivityTick = now;
	UpdateStatusText(true);
}

bool CMarkdownEditorDlg::SaveCurrentDocument(bool saveAs, CString* errorMessage)
{
	CString targetPath = m_strCurrentFile;
	if (saveAs || targetPath.IsEmpty())
	{
		CString defaultName = m_currentFileName;
		if (defaultName.IsEmpty())
			defaultName = _T("未命名.md");
		CFileDialog saveDlg(FALSE, _T("md"), defaultName,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			_T("Markdown Files (*.md)|*.md|Text Files (*.txt)|*.txt|All Files (*.*)|*.*||"),
			this);
		if (saveDlg.DoModal() != IDOK)
			return false;
		targetPath = saveDlg.GetPathName();
	}

	if (targetPath.IsEmpty())
	{
		if (errorMessage)
			*errorMessage = _T("当前文档没有可保存路径。");
		return false;
	}

	CString content;
	if (m_bMarkdownMode && !m_rawDocumentText.empty())
		content = CString(m_rawDocumentText.c_str());
	else
		m_editContent.GetWindowText(content);

	bool writeBom = true;
	if (FileExistsOnDisk(targetPath))
	{
		bool hasBom = false;
		if (DetectUtf8Bom(targetPath, hasBom))
			writeBom = hasBom;
	}

	if (!WriteUtf8File(targetPath, content, writeBom, errorMessage))
		return false;

	m_strCurrentFile = targetPath;
	m_currentFileName = ExtractFileNameFromPath(targetPath);
	AddRecentFile(targetPath);
	m_textDirty = false;
	m_rawDocumentText.assign(content.GetString());
	m_autoSavedStatusVisible = false;
	m_editingHintVisible = false;
	m_lastUserActivityTick = ::GetTickCount();
	MarkDocumentDirty(false);
	UpdateCurrentFileSignatureFromDisk();
	UpdateWindowTitle();
	UpdateStatusText(true);
	UpdateSidebar();
	return true;
}

bool CMarkdownEditorDlg::EnsureCanDiscardUnsavedChanges(const CString& nextPath, bool closingApp)
{
	if (!m_documentDirty)
		return true;

	CString msg;
	if (closingApp)
	{
		msg = _T("当前文档有未保存修改。\r\n是否先保存后再退出？");
	}
	else if (!nextPath.IsEmpty())
	{
		msg.Format(_T("当前文档有未保存修改。\r\n是否先保存后再打开新文件？\r\n\r\n目标文件：%s"),
			(LPCTSTR)nextPath);
	}
	else
	{
		msg = _T("当前文档有未保存修改，是否先保存？");
	}

	CUnsavedChangesDlg dlg(m_themeDark, msg, this);
	const int ret = dlg.DoModal();
	if (ret == IDCANCEL)
		return false;
	if (ret == IDYES)
	{
		CString error;
		if (!SaveCurrentDocument(false, &error))
		{
			if (error.IsEmpty())
				error = _T("保存失败。");
			AfxMessageBox(error, MB_OK | MB_ICONERROR);
			return false;
		}
	}
	return true;
}

bool CMarkdownEditorDlg::HandleEditorTabIndent(bool outdent)
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return false;
	if ((m_editContent.GetStyle() & ES_READONLY) != 0)
		return false;
	CWnd* focus = GetFocus();
	if (!focus || focus->GetSafeHwnd() != m_editContent.GetSafeHwnd())
		return false;

	LONG selStart = 0;
	LONG selEnd = 0;
	m_editContent.GetSel(selStart, selEnd);

	CString fullText;
	m_editContent.GetWindowText(fullText);
	const int textLen = fullText.GetLength();
	if (textLen <= 0)
		return false;

	int startLine = m_editContent.LineFromChar(selStart);
	int endLine = m_editContent.LineFromChar(selEnd);
	if (selEnd > selStart)
	{
		const int endLineStart = m_editContent.LineIndex(endLine);
		if (endLineStart == selEnd && endLine > startLine)
			--endLine;
	}
	if (startLine < 0 || endLine < startLine)
		return false;

	int blockStart = m_editContent.LineIndex(startLine);
	if (blockStart < 0)
		blockStart = 0;
	int blockEnd = m_editContent.LineIndex(endLine + 1);
	if (blockEnd < 0 || blockEnd > textLen)
		blockEnd = textLen;
	if (blockEnd < blockStart)
		return false;

	CString block = fullText.Mid(blockStart, blockEnd - blockStart);
	CString replaced;
	replaced.Preallocate(block.GetLength() + (endLine - startLine + 1) * 2);

	int pos = 0;
	while (pos <= block.GetLength())
	{
		int lineEnd = block.Find(_T("\r\n"), pos);
		const bool hasCrLf = (lineEnd != -1);
		if (!hasCrLf)
			lineEnd = block.GetLength();

		CString line = block.Mid(pos, lineEnd - pos);
		if (outdent)
		{
			if (line.Left(1) == _T("\t"))
				line = line.Mid(1);
			else if (line.Left(4) == _T("    "))
				line = line.Mid(4);
			else if (line.Left(2) == _T("  "))
				line = line.Mid(2);
		}
		else
		{
			line = _T("\t") + line;
		}
		replaced += line;
		if (hasCrLf)
			replaced += _T("\r\n");

		if (!hasCrLf)
			break;
		pos = lineEnd + 2;
	}

	m_suppressChangeEvent = true;
	m_editContent.SetSel(blockStart, blockEnd);
	m_editContent.ReplaceSel(replaced, TRUE);
	m_suppressChangeEvent = false;
	m_editContent.SetSel(blockStart, blockStart + replaced.GetLength());

	RefreshDirtyStateFromEditor();
	return true;
}

bool CMarkdownEditorDlg::CopySelectionWithMarkdownPayload(bool cut)
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return false;

	LONG selStart = 0;
	LONG selEnd = 0;
	m_editContent.GetSel(selStart, selEnd);
	if (selEnd <= selStart)
		return false;

	if (cut && ((m_editContent.GetStyle() & ES_READONLY) != 0))
		return false;

	const int selLen = static_cast<int>(selEnd - selStart);
	std::vector<TCHAR> buffer(static_cast<size_t>(selLen) + 1, 0);
	m_editContent.SendMessage(EM_GETSELTEXT, 0, reinterpret_cast<LPARAM>(buffer.data()));
	CString selection = buffer.data();
	if (selection.IsEmpty())
		return false;

	const CStringW wide(selection);

	if (!::OpenClipboard(GetSafeHwnd()))
		return false;
	struct ClipboardCloser { ~ClipboardCloser() { ::CloseClipboard(); } } closer;
	if (!::EmptyClipboard())
		return false;

	auto setClipboardWide = [](UINT fmt, const CStringW& text) -> bool {
		const SIZE_T bytes = static_cast<SIZE_T>((text.GetLength() + 1) * sizeof(wchar_t));
		HGLOBAL h = ::GlobalAlloc(GMEM_MOVEABLE, bytes);
		if (!h)
			return false;
		void* ptr = ::GlobalLock(h);
		if (!ptr)
		{
			::GlobalFree(h);
			return false;
		}
		memcpy(ptr, text.GetString(), bytes);
		::GlobalUnlock(h);
		if (::SetClipboardData(fmt, h) == nullptr)
		{
			::GlobalFree(h);
			return false;
		}
		return true;
	};

	const UINT markdownFmt = ::RegisterClipboardFormatW(kClipboardMarkdownFormatName);
	bool ok = setClipboardWide(CF_UNICODETEXT, wide);
	if (ok)
		ok = setClipboardWide(markdownFmt, wide);
	if (!ok)
		return false;

	if (cut)
	{
		m_editContent.ReplaceSel(_T(""), TRUE);
		RefreshDirtyStateFromEditor();
	}
	return true;
}

bool CMarkdownEditorDlg::PasteSelectionWithMarkdownPayload()
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return false;
	if ((m_editContent.GetStyle() & ES_READONLY) != 0)
		return false;

	if (!::OpenClipboard(GetSafeHwnd()))
		return false;
	struct ClipboardCloser { ~ClipboardCloser() { ::CloseClipboard(); } } closer;

	CString insertText;
	const UINT markdownFmt = ::RegisterClipboardFormatW(kClipboardMarkdownFormatName);
	auto readUnicodeFromClipboard = [](UINT fmt, CString& out) -> bool {
		if (!::IsClipboardFormatAvailable(fmt))
			return false;
		HGLOBAL h = ::GetClipboardData(fmt);
		if (!h)
			return false;
		const wchar_t* ptr = reinterpret_cast<const wchar_t*>(::GlobalLock(h));
		if (!ptr)
			return false;
		out = ptr;
		::GlobalUnlock(h);
		return true;
	};

	if (!readUnicodeFromClipboard(markdownFmt, insertText))
	{
		if (!readUnicodeFromClipboard(CF_UNICODETEXT, insertText))
			return false;
	}

	insertText.Replace(_T("\r\n"), _T("\n"));
	insertText.Replace(_T("\r"), _T("\n"));
	insertText.Replace(_T("\n"), _T("\r\n"));
	m_editContent.ReplaceSel(insertText, TRUE);
	RefreshDirtyStateFromEditor();
	return true;
}

void CMarkdownEditorDlg::ShowFindReplaceDialog(bool replaceMode)
{
	if (m_findReplaceDialog && ::IsWindow(m_findReplaceDialog->GetSafeHwnd()))
	{
		if (m_findReplaceMode == replaceMode)
		{
			m_findReplaceDialog->SetActiveWindow();
			return;
		}
		m_findReplaceDialog->DestroyWindow();
		m_findReplaceDialog = nullptr;
	}

	DWORD flags = FR_DOWN;
	if (m_lastFindMatchCase)
		flags |= FR_MATCHCASE;
	if (m_lastFindWholeWord)
		flags |= FR_WHOLEWORD;

	CFindReplaceDialog* dlg = new CFindReplaceDialog();
	const BOOL createOk = dlg->Create(
		replaceMode ? FALSE : TRUE,
		m_lastFindText,
		m_lastReplaceText,
		flags,
		this);
	if (!createOk)
	{
		delete dlg;
		return;
	}

	m_findReplaceDialog = dlg;
	m_findReplaceMode = replaceMode;
}

bool CMarkdownEditorDlg::FindNextInEditor(const CString& findText,
	bool matchCase,
	bool wholeWord,
	bool searchDown,
	bool wrapSearch)
{
	if (findText.IsEmpty())
		return false;

	CString content;
	m_editContent.GetWindowText(content);
	if (content.IsEmpty())
		return false;

	CString haystack = content;
	CString needle = findText;
	if (!matchCase)
	{
		haystack.MakeLower();
		needle.MakeLower();
	}

	const int textLen = haystack.GetLength();
	const int needleLen = needle.GetLength();
	if (needleLen <= 0 || needleLen > textLen)
		return false;

	LONG selStart = 0;
	LONG selEnd = 0;
	m_editContent.GetSel(selStart, selEnd);
	if (selStart < 0) selStart = 0;
	if (selEnd < 0) selEnd = 0;
	if (selStart > textLen) selStart = textLen;
	if (selEnd > textLen) selEnd = textLen;

	auto isBoundary = [&](int index) -> bool {
		if (!wholeWord)
			return true;
		const TCHAR prev = (index > 0) ? content[index - 1] : 0;
		const TCHAR next = (index + needleLen < textLen) ? content[index + needleLen] : 0;
		return !IsWordCharacter(prev) && !IsWordCharacter(next);
	};

	auto findDown = [&](int begin, int endExclusive) -> int {
		int pos = haystack.Find(needle, begin);
		while (pos != -1)
		{
			if (pos >= endExclusive)
				return -1;
			if (isBoundary(pos))
				return pos;
			pos = haystack.Find(needle, pos + 1);
		}
		return -1;
	};

	auto findUp = [&](int beginInclusive, int endExclusive) -> int {
		int pos = haystack.Find(needle, beginInclusive);
		int last = -1;
		while (pos != -1 && pos < endExclusive)
		{
			if (isBoundary(pos))
				last = pos;
			pos = haystack.Find(needle, pos + 1);
		}
		return last;
	};

	int found = -1;
	if (searchDown)
	{
		found = findDown(static_cast<int>(selEnd), textLen);
		if (found == -1 && wrapSearch)
			found = findDown(0, static_cast<int>(selStart));
	}
	else
	{
		found = findUp(0, static_cast<int>(selStart));
		if (found == -1 && wrapSearch)
			found = findUp(static_cast<int>(selEnd), textLen);
	}

	if (found < 0)
		return false;

	m_editContent.SetSel(found, found + needleLen);
	m_editContent.SetFocus();
	m_editContent.SendMessage(EM_SCROLLCARET, 0, 0);
	return true;
}

bool CMarkdownEditorDlg::ReplaceCurrentInEditor(const CString& findText,
	const CString& replaceText,
	bool matchCase,
	bool wholeWord)
{
	if (findText.IsEmpty())
		return false;
	if ((m_editContent.GetStyle() & ES_READONLY) != 0)
		return false;

	LONG selStart = 0;
	LONG selEnd = 0;
	m_editContent.GetSel(selStart, selEnd);

	CString selected;
	if (selEnd > selStart)
	{
		const int selLen = static_cast<int>(selEnd - selStart);
		std::vector<TCHAR> buf(static_cast<size_t>(selLen) + 1, 0);
		m_editContent.SendMessage(EM_GETSELTEXT, 0, reinterpret_cast<LPARAM>(buf.data()));
		selected = buf.data();
	}

	bool selectedMatch = false;
	if (!selected.IsEmpty() && selected.GetLength() == findText.GetLength())
	{
		CString lhs = selected;
		CString rhs = findText;
		if (!matchCase)
		{
			lhs.MakeLower();
			rhs.MakeLower();
		}
		selectedMatch = (lhs == rhs);
		if (selectedMatch && wholeWord)
		{
			CString content;
			m_editContent.GetWindowText(content);
			const int len = content.GetLength();
			const TCHAR prev = (selStart > 0 && selStart <= len) ? content[selStart - 1] : 0;
			const TCHAR next = (selEnd >= 0 && selEnd < len) ? content[selEnd] : 0;
			selectedMatch = !IsWordCharacter(prev) && !IsWordCharacter(next);
		}
	}

	if (!selectedMatch)
	{
		if (!FindNextInEditor(findText, matchCase, wholeWord, true, false))
			return false;
	}

	m_editContent.ReplaceSel(replaceText, TRUE);
	return true;
}

int CMarkdownEditorDlg::ReplaceAllInEditor(const CString& findText,
	const CString& replaceText,
	bool matchCase,
	bool wholeWord)
{
	if (findText.IsEmpty())
		return 0;
	if ((m_editContent.GetStyle() & ES_READONLY) != 0)
		return 0;

	CString content;
	m_editContent.GetWindowText(content);
	if (content.IsEmpty())
		return 0;

	CString haystack = content;
	CString needle = findText;
	if (!matchCase)
	{
		haystack.MakeLower();
		needle.MakeLower();
	}

	const int textLen = haystack.GetLength();
	const int needleLen = needle.GetLength();
	if (needleLen <= 0 || needleLen > textLen)
		return 0;

	auto isBoundary = [&](int index) -> bool {
		if (!wholeWord)
			return true;
		const TCHAR prev = (index > 0) ? content[index - 1] : 0;
		const TCHAR next = (index + needleLen < textLen) ? content[index + needleLen] : 0;
		return !IsWordCharacter(prev) && !IsWordCharacter(next);
	};

	CString replaced;
	replaced.Preallocate(content.GetLength() + 64);
	int count = 0;
	int cursor = 0;
	while (cursor < textLen)
	{
		int pos = haystack.Find(needle, cursor);
		if (pos == -1)
		{
			replaced += content.Mid(cursor);
			break;
		}
		if (!isBoundary(pos))
		{
			replaced += content.Mid(cursor, pos - cursor + 1);
			cursor = pos + 1;
			continue;
		}
		replaced += content.Mid(cursor, pos - cursor);
		replaced += replaceText;
		cursor = pos + needleLen;
		++count;
	}

	if (count > 0)
	{
		m_suppressChangeEvent = true;
		m_editContent.SetWindowText(replaced);
		m_suppressChangeEvent = false;
		m_editContent.SetSel(0, 0);
		RefreshDirtyStateFromEditor();
	}
	return count;
}

bool CMarkdownEditorDlg::QueryFileSignature(const CString& filePath, ULONGLONG& size, FILETIME& writeTime)
{
	size = 0;
	::ZeroMemory(&writeTime, sizeof(writeTime));
	if (filePath.IsEmpty())
		return false;

	WIN32_FILE_ATTRIBUTE_DATA attrs{};
	if (!::GetFileAttributesEx(filePath, GetFileExInfoStandard, &attrs))
		return false;
	if ((attrs.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
		return false;

	ULARGE_INTEGER ull{};
	ull.HighPart = attrs.nFileSizeHigh;
	ull.LowPart = attrs.nFileSizeLow;
	size = ull.QuadPart;
	writeTime = attrs.ftLastWriteTime;
	return true;
}

void CMarkdownEditorDlg::UpdateCurrentFileSignatureFromDisk()
{
	ULONGLONG size = 0;
	FILETIME writeTime{};
	if (QueryFileSignature(m_strCurrentFile, size, writeTime))
	{
		m_knownFileSize = size;
		m_knownFileWriteTime = writeTime;
		m_fileSignatureValid = true;
	}
	else
	{
		m_fileSignatureValid = false;
		m_knownFileSize = 0;
		::ZeroMemory(&m_knownFileWriteTime, sizeof(m_knownFileWriteTime));
	}
}

bool CMarkdownEditorDlg::DetectExternalFileChange()
{
	if (m_strCurrentFile.IsEmpty())
		return false;
	if (m_externalChangePromptActive)
		return false;
	if (m_loadingActive || m_parsingActive || m_formattingActive || m_textModeResetActive)
		return false;

	ULONGLONG currentSize = 0;
	FILETIME currentWriteTime{};
	const bool exists = QueryFileSignature(m_strCurrentFile, currentSize, currentWriteTime);
	if (!exists)
	{
		if (m_fileSignatureValid)
		{
			m_externalFileMissingDetected = true;
			return true;
		}
		return false;
	}
	m_externalFileMissingDetected = false;

	if (!m_fileSignatureValid)
	{
		m_knownFileSize = currentSize;
		m_knownFileWriteTime = currentWriteTime;
		m_fileSignatureValid = true;
		return false;
	}

	const bool sameTime = (m_knownFileWriteTime.dwLowDateTime == currentWriteTime.dwLowDateTime) &&
		(m_knownFileWriteTime.dwHighDateTime == currentWriteTime.dwHighDateTime);
	return !(sameTime && m_knownFileSize == currentSize);
}

void CMarkdownEditorDlg::HandleExternalFileChange()
{
	if (m_strCurrentFile.IsEmpty())
		return;
	if (m_externalChangePromptActive)
		return;

	m_externalChangePromptActive = true;
	if (m_externalFileMissingDetected)
	{
		bool savedAs = false;
		if (m_documentDirty)
		{
			const int answer = AfxMessageBox(
				_T("检测到文件已被外部删除。\r\n是否“另存为”以保留当前未保存编辑？"),
				MB_YESNO | MB_ICONWARNING);
			if (answer == IDYES)
			{
				CString error;
				savedAs = SaveCurrentDocument(true, &error);
				if (!savedAs && !error.IsEmpty())
					AfxMessageBox(error, MB_OK | MB_ICONERROR);
			}
		}
		else
		{
			AfxMessageBox(_T("检测到文件已被外部删除，当前内容仍保留在编辑器中。"),
				MB_OK | MB_ICONWARNING);
		}

		// If the user successfully saved to a new path, SaveCurrentDocument() has already
		// refreshed the signature and current file identity. Do not invalidate it here.
		if (!savedAs)
		{
			m_fileSignatureValid = false;
			m_knownFileSize = 0;
			::ZeroMemory(&m_knownFileWriteTime, sizeof(m_knownFileWriteTime));
		}
		m_externalFileMissingDetected = false;
		m_externalChangePromptActive = false;
		return;
	}

	CString msg;
	if (m_documentDirty)
		msg = _T("检测到文件已被外部修改。\r\n\r\n是：重新加载磁盘内容（会覆盖当前未保存编辑）\r\n否：保留当前编辑内容");
	else
		msg = _T("检测到文件已被外部修改，是否重新加载磁盘内容？");

	const int answer = AfxMessageBox(msg, MB_YESNO | MB_ICONWARNING);
	if (answer == IDYES)
	{
		m_autoSavedStatusVisible = false;
		m_editingHintVisible = false;
		MarkDocumentDirty(false);
		StartLoadDocumentAsync(m_strCurrentFile);
	}
	else
	{
		UpdateCurrentFileSignatureFromDisk();
	}

	m_externalChangePromptActive = false;
}

void CMarkdownEditorDlg::OnDestroy()
{
	m_destroying = true;
	if (m_asyncCancelToken)
		m_asyncCancelToken->store(true, std::memory_order_relaxed);
	++m_loadGeneration;
	++m_parseGeneration;
	++m_displayGeneration;

	// Detached workers use message posting to marshal results back to UI. Mark cancel and
	// wait briefly so most in-flight workers can drain before we release heavy resources.
	const DWORD waitStartTick = ::GetTickCount();
	while (m_asyncWorkerCount && m_asyncWorkerCount->load(std::memory_order_relaxed) > 0 &&
		(::GetTickCount() - waitStartTick) < 1500)
	{
		::Sleep(10);
	}

	KillTimer(kDebounceTimerId);
	KillTimer(kFormatTimerId);
	KillTimer(kInsertTimerId);
	KillTimer(kStatusPulseTimerId);
	KillTimer(kOverlaySyncTimerId);
	KillTimer(kOutlineFollowTimerId);
	KillTimer(kAutoSaveTimerId);
	KillTimer(kFileWatchTimerId);
	if (m_findReplaceDialog && ::IsWindow(m_findReplaceDialog->GetSafeHwnd()))
		m_findReplaceDialog->DestroyWindow();
	m_findReplaceDialog = nullptr;
	if (m_mermaidOverlayTimer != 0)
	{
		KillTimer(kMermaidOverlayTimerId);
		m_mermaidOverlayTimer = 0;
	}

	m_mermaidDiagrams.clear();
	ClearTableListViewOverlays();
	for (auto& kv : m_mermaidBitmapCache)
	{
		if (kv.second)
			::DeleteObject(kv.second);
	}
	for (auto& kv : m_mermaidTransientBitmapCache)
	{
		if (kv.second)
			::DeleteObject(kv.second);
	}
	m_mermaidBitmapCache.clear();
	m_mermaidBitmapSize.clear();
	m_mermaidTransientBitmapCache.clear();
	m_mermaidTransientBitmapSize.clear();
	m_mermaidBitmapLastUseSeq.clear();
	m_mermaidBitmapUseSeq = 0;
	m_mermaidFetchInFlight.clear();
	if (m_gdiplusToken)
	{
		Gdiplus::GdiplusShutdown(m_gdiplusToken);
		m_gdiplusToken = 0;
	}
	if (m_wrapTargetHdc != nullptr)
	{
		::ReleaseDC(nullptr, m_wrapTargetHdc);
		m_wrapTargetHdc = nullptr;
	}
	::OleUninitialize();
	CDialogEx::OnDestroy();
}

void CMarkdownEditorDlg::OnClose()
{
	if (!EnsureCanDiscardUnsavedChanges(CString(), true))
		return;
	CDialogEx::OnClose();
}

void CMarkdownEditorDlg::OnCancel()
{
	if (!EnsureCanDiscardUnsavedChanges(CString(), true))
		return;
	CDialogEx::OnCancel();
}

void CMarkdownEditorDlg::OnOK()
{
	if (!EnsureCanDiscardUnsavedChanges(CString(), true))
		return;
	CDialogEx::OnOK();
}

void CMarkdownEditorDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == kSysCmdSetDefaultMd)
	{
		if (IsCurrentAppMdAssociationRegistered())
		{
			AfxMessageBox(_T("当前程序已是 .md 文件默认打开工具。"), MB_OK | MB_ICONINFORMATION);
			return;
		}

		CString error;
		if (SetAsDefaultMdFileHandler(&error))
		{
			if (CWinApp* app = AfxGetApp())
				app->WriteProfileInt(_T("Settings"), _T("PromptedDefaultMdV2"), 1);
			if (IsCurrentAppMdAssociationRegistered())
				AfxMessageBox(_T("已设置为 .md 文件默认打开工具。"), MB_OK | MB_ICONINFORMATION);
			else
				AfxMessageBox(_T("系统限制未能直接完成设置，已打开系统默认应用设置，请将 .md 关联到本程序。"), MB_OK | MB_ICONINFORMATION);
		}
		else
		{
			CString msg = _T("设置默认打开工具失败。");
			if (!error.IsEmpty())
			{
				msg += _T("\r\n");
				msg += error;
			}
			AfxMessageBox(msg, MB_OK | MB_ICONWARNING);
		}
	}
	else if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout(m_themeDark);
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

void CMarkdownEditorDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this);

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

HCURSOR CMarkdownEditorDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

int CMarkdownEditorDlg::GetZoomPercent() const
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return 100;
	UINT num = 0;
	UINT den = 0;
	if (m_editContent.SendMessage(EM_GETZOOM, (WPARAM)&num, (LPARAM)&den) && den != 0)
	{
		int percent = static_cast<int>((static_cast<double>(num) * 100.0) / static_cast<double>(den) + 0.5);
		if (percent < 10) percent = 10;
		if (percent > 500) percent = 500;
		return percent;
	}
	return 100;
}

void CMarkdownEditorDlg::SetZoomPercent(int percent)
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return;
	if (percent < 10) percent = 10;
	if (percent > 500) percent = 500;
	m_editContent.SendMessage(EM_SETZOOM, (WPARAM)percent, (LPARAM)100);
}

BOOL CMarkdownEditorDlg::PreTranslateMessage(MSG* pMsg)
{
	if (IsUserInteractionMessage(pMsg))
		NoteUserActivity(true);

	if (pMsg && ::IsWindow(m_editContent.GetSafeHwnd()) && pMsg->hwnd == m_editContent.GetSafeHwnd())
	{
		const bool markdownReadonly = m_bMarkdownMode && ((m_editContent.GetStyle() & ES_READONLY) != 0);
		if (markdownReadonly && IsReadonlyEditIntentMessage(pMsg))
		{
			PromptSwitchToTextModeForEdit();
			return TRUE;
		}
	}

	// Keep overlay-drawn elements (mermaid/HR/table grid) aligned with RichEdit scrolling.
	// RichEdit may finish layout slightly after scroll messages, so we schedule a short sync.
	if (pMsg && ::IsWindow(m_editContent.GetSafeHwnd()) && pMsg->hwnd == m_editContent.GetSafeHwnd())
	{
		if (m_bMarkdownMode && !m_windowSizingMove &&
			(pMsg->message == WM_VSCROLL || pMsg->message == WM_HSCROLL || pMsg->message == WM_MOUSEWHEEL))
		{
			m_overlaySyncAttempts = 2;
			KillTimer(kOverlaySyncTimerId);
			SetTimer(kOverlaySyncTimerId, 40, NULL);
		}
	}

	if (pMsg && pMsg->message == WM_MOUSEWHEEL)
	{
		// Ctrl + Wheel zooms the document.
		if ((::GetKeyState(VK_CONTROL) & 0x8000) && ::IsWindow(m_editContent.GetSafeHwnd()))
		{
			CWnd* focus = GetFocus();
			if (focus && (focus->GetSafeHwnd() == m_editContent.GetSafeHwnd()))
			{
				int delta = GET_WHEEL_DELTA_WPARAM(pMsg->wParam);
				int step = (delta > 0) ? 10 : -10;
				int target = GetZoomPercent() + step;
				SetZoomPercent(target);
				// Keep overlay-drawn elements (mermaid/HR/table) scaled with document zoom.
					if (m_bMarkdownMode)
					{
						UpdateMermaidSpacing();
						UpdateTocSpacing();
						UpdateMermaidOverlayRegion(true);
						if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
							m_mermaidOverlay.Invalidate(FALSE);
					}
				UpdateStatusText();
				return TRUE;
			}
		}
	}
	if (pMsg && pMsg->message == WM_KEYDOWN)
	{
		const bool ctrlDown = ((::GetKeyState(VK_CONTROL) & 0x8000) != 0);
		const bool shiftDown = ((::GetKeyState(VK_SHIFT) & 0x8000) != 0);
		const WPARAM key = pMsg->wParam;
		const HWND focusHwnd = ::GetFocus();
		const bool editorHasFocus = (::IsWindow(m_editContent.GetSafeHwnd()) && focusHwnd == m_editContent.GetSafeHwnd());

		if (ctrlDown && (key == 'S' || key == 's'))
		{
			CString error;
			if (!SaveCurrentDocument(shiftDown, &error) && !error.IsEmpty())
				AfxMessageBox(error, MB_OK | MB_ICONERROR);
			return TRUE;
		}

		// Editor-only shortcuts should not hijack key handling when another control has focus.
		if (!editorHasFocus)
			return CDialogEx::PreTranslateMessage(pMsg);

		if (ctrlDown && (key == 'Z' || key == 'z'))
		{
			if (::IsWindow(m_editContent.GetSafeHwnd()) && m_editContent.CanUndo())
			{
				m_editContent.Undo();
				return TRUE;
			}
		}

		if (ctrlDown && (key == 'Y' || key == 'y'))
		{
			if (::IsWindow(m_editContent.GetSafeHwnd()))
			{
				m_editContent.SendMessage(EM_REDO, 0, 0);
				return TRUE;
			}
		}

		if (ctrlDown && (key == 'F' || key == 'f'))
		{
			ShowFindReplaceDialog(false);
			return TRUE;
		}

		if (ctrlDown && (key == 'H' || key == 'h'))
		{
			ShowFindReplaceDialog(true);
			return TRUE;
		}

		if (ctrlDown && (key == 'C' || key == 'c'))
		{
			if (CopySelectionWithMarkdownPayload(false))
				return TRUE;
		}

		if (ctrlDown && (key == 'X' || key == 'x'))
		{
			if (CopySelectionWithMarkdownPayload(true))
				return TRUE;
		}

		if (ctrlDown && (key == 'V' || key == 'v'))
		{
			if (PasteSelectionWithMarkdownPayload())
				return TRUE;
		}

		if (key == VK_TAB && !ctrlDown)
		{
			if (HandleEditorTabIndent(shiftDown))
				return TRUE;
		}

		if (ctrlDown && key == VK_OEM_MINUS)
		{
			// Ctrl + '-' resets zoom.
			SetZoomPercent(100);
				if (m_bMarkdownMode)
				{
					UpdateMermaidSpacing();
					UpdateTocSpacing();
					UpdateMermaidOverlayRegion(true);
					if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
						m_mermaidOverlay.Invalidate(FALSE);
				}
			UpdateStatusText();
			return TRUE;
		}
	}

	return CDialogEx::PreTranslateMessage(pMsg);
}

void CMarkdownEditorDlg::OnBnClickedOpenFile()
{
	CFileDialog fileDlg(TRUE, _T("md"), NULL,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		_T("Markdown Files (*.md)|*.md|Text Files (*.txt)|*.txt|All Files (*.*)|*.*||"));

	if (fileDlg.DoModal() == IDOK)
	{
		OpenDocument(fileDlg.GetPathName());
	}
}

void CMarkdownEditorDlg::OnDropFiles(HDROP hDropInfo)
{
	UINT fileCount = ::DragQueryFile(hDropInfo, 0xFFFFFFFF, nullptr, 0);
	if (fileCount > 0)
	{
		TCHAR filePath[MAX_PATH] = { 0 };
		if (::DragQueryFile(hDropInfo, 0, filePath, MAX_PATH) > 0)
		{
			OpenDocument(filePath);
		}
	}

	::DragFinish(hDropInfo);
	CDialogEx::OnDropFiles(hDropInfo);
}

LRESULT CMarkdownEditorDlg::OnInitDragDrop(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	if (GetSafeHwnd())
	{
		ModifyStyleEx(0, WS_EX_ACCEPTFILES);
		DragAcceptFiles(TRUE);
	}

	return 0;
}

void CMarkdownEditorDlg::OnBnClickedModeSwitch()
{
	m_modeSwitchInProgress = true;
	m_bMarkdownMode = !m_bMarkdownMode;

	if (m_bMarkdownMode)
	{
		m_btnModeSwitch.SetWindowText(_T("Markdown模式"));
		// Ensure we have the raw text snapshot (source of truth for parsing).
		if (m_rawDocumentText.empty() || m_textDirty)
		{
			CString content;
			m_editContent.GetWindowText(content);
			m_rawDocumentText.assign(content.GetString());
			m_textDirty = false;
		}

		// Switch display to a Typora-like rendering (tables use tabbed columns).
		// Prepare display text on a background thread for large docs.
		SetRenderBusy(true, _T("正在切换Markdown模式..."));
		const unsigned long long generation = ++m_displayGeneration;
		HWND hwnd = GetSafeHwnd();
		std::wstring raw = m_rawDocumentText;
		auto cancelToken = m_asyncCancelToken;
		auto workerCount = m_asyncWorkerCount;
		std::thread([hwnd, generation, raw = std::move(raw), cancelToken, workerCount]() mutable {
			AsyncWorkerScope workerScope(workerCount);
			if (cancelToken && cancelToken->load(std::memory_order_relaxed))
				return;

			auto payload = std::make_unique<DisplayTextPayload>();
			payload->generation = generation;
			try
			{
				auto transformed = RenderText::TransformWithMapping(raw);
				payload->displayText = std::move(transformed.displayText);
				payload->displayToSourceIndexMap = std::move(transformed.mapping.displayToSource);
				payload->sourceToDisplayIndexMap = std::move(transformed.mapping.sourceToDisplay);
			}
			catch (...)
			{
				payload->displayText = raw;
				BuildIdentityIndexMap(raw.size(),
					payload->displayToSourceIndexMap,
					payload->sourceToDisplayIndexMap);
			}
			if (!cancelToken || !cancelToken->load(std::memory_order_relaxed))
				SafePostMessageUnique(hwnd, WM_PREPARE_DISPLAY_TEXT_COMPLETE, 0, std::move(payload));
		}).detach();
	}
	else
	{
		m_btnModeSwitch.SetWindowText(_T("文本模式"));
		ApplyTextMode();
		SetRenderBusy(false, _T(""));
	}

	UpdateStatusText();
	// Mode switch should not cause sidebar content to change or flicker.
	// Keep sidebar stable; only the editor view changes.
}

void CMarkdownEditorDlg::OnBnClickedThemeSwitch()
{
	m_themeDark = !m_themeDark;
	if (CWinApp* app = AfxGetApp())
		app->WriteProfileInt(_T("Settings"), _T("ThemeDark"), m_themeDark ? 1 : 0);
	ApplyTheme(m_themeDark);
	// Full refresh: fix stale background blocks in the original window region.
	Invalidate(FALSE);
	UpdateWindow();
}

void CMarkdownEditorDlg::OpenDocument(const CString& filePath)
{
	if (filePath.IsEmpty())
		return;
	if (!EnsureCanDiscardUnsavedChanges(filePath, false))
		return;

	if (!FileExistsOnDisk(filePath))
	{
		CString msg;
		msg.Format(_T("文件不存在或无法访问：\r\n%s"), (LPCTSTR)filePath);
		AfxMessageBox(msg, MB_OK | MB_ICONWARNING);
		return;
	}

	// Reset outline state immediately to avoid showing stale headings from the previous document.
	m_outlineEntries.clear();
	m_outlineLineIndex.clear();
	m_outlineSourceIndex.clear();
	m_outlineItemSourceIndex.clear();
	m_outlineEntriesGeneration = 0;
	m_outlineAutoSelectFirst = true;

	m_strCurrentFile = filePath;
	m_currentFileName = ExtractFileNameFromPath(filePath);
	if (m_currentFileName.IsEmpty())
		m_currentFileName = filePath;
	m_autoSavedStatusVisible = false;
	m_editingHintVisible = false;
	MarkDocumentDirty(false);
	UpdateWindowTitle();
	const BOOL isMarkdown = IsMarkdownFile(m_strCurrentFile);
	// IMPORTANT: Set mode before kicking off async load.
	// OnDocumentLoadComplete uses m_bMarkdownMode to decide whether to start parsing/formatting.
	m_bMarkdownMode = (isMarkdown != FALSE);
	m_btnModeSwitch.SetWindowText(isMarkdown ? _T("Markdown模式") : _T("文本模式"));
	StartLoadDocumentAsync(m_strCurrentFile);

	AddRecentFile(m_strCurrentFile);
	UpdateStatusText();

	// Typora-like behavior: after opening a document, switch to Outline.
	// Also clear outline search so each document starts clean.
	// NOTE: SetSidebarTab() persists the current tab's query from the search edit.
	// If we are already on Outline, clear the edit first so it won't overwrite the empty query.
	if (m_sidebarActiveTab == 1 && ::IsWindow(m_editFileSearch.GetSafeHwnd()))
	{
		m_suppressFileSearchEvent = true;
		m_editFileSearch.SetWindowText(_T(""));
		m_suppressFileSearchEvent = false;
	}
	m_outlineSearchQuery.Empty();
	SetSidebarTab(1);
	UpdateOutline();
}

void CMarkdownEditorDlg::ExpandAllOutlineItems()
{
	if (m_sidebarActiveTab != 1)
		return;
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;

	HTREEITEM root = m_treeSidebar.GetRootItem();
	for (HTREEITEM item = root; item != nullptr; item = m_treeSidebar.GetNextSiblingItem(item))
	{
		ExpandTreeItemRecursive(item);
	}
}

void CMarkdownEditorDlg::EnsureOutlineInitialSelection()
{
	if (!m_outlineAutoSelectFirst)
		return;
	if (m_sidebarActiveTab != 1)
		return;
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;
	if (m_outlineLineIndex.empty())
		return;

	HTREEITEM firstItem = m_outlineLineIndex.front().second;
	if (firstItem == nullptr)
		return;

	m_sidebarSelectionUpdating = true;
	m_treeSidebar.SelectItem(firstItem);
	m_treeSidebar.EnsureVisible(firstItem);
	m_sidebarSelectionUpdating = false;
	m_outlineAutoSelectFirst = false;
	m_treeSidebar.RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
}

void CMarkdownEditorDlg::ExpandTreeItemRecursive(HTREEITEM item)
{
	if (item == nullptr)
		return;

	m_treeSidebar.Expand(item, TVE_EXPAND);
	HTREEITEM child = m_treeSidebar.GetChildItem(item);
	while (child != nullptr)
	{
		ExpandTreeItemRecursive(child);
		child = m_treeSidebar.GetNextSiblingItem(child);
	}
}

void CMarkdownEditorDlg::OnEnChangeEdit()
{
	if (m_suppressChangeEvent || m_formattingActive)
		return;
	if (m_loadingActive)
		return;

	NoteUserActivity(true);
	RefreshDirtyStateFromEditor();

	if (m_bMarkdownMode)
	{
		// Delay highlighting to avoid frequent updates
		ScheduleMarkdownFormatting(500);
	}

	UpdateOutlineSelectionByCaret();
}

void CMarkdownEditorDlg::OnEnChangeFileSearch()
{
	if (!::IsWindow(m_editFileSearch.GetSafeHwnd()))
		return;
	if (m_suppressFileSearchEvent)
		return;

	CString q;
	m_editFileSearch.GetWindowText(q);
	q.Trim();
	if (m_sidebarActiveTab == 0)
		m_fileSearchQuery = q;
	else if (m_sidebarActiveTab == 1)
		m_outlineSearchQuery = q;

	if (m_sidebarActiveTab == 0)
	{
		if (q.IsEmpty())
		{
			// Empty query means show the full file list (and recent files).
			UpdateSidebar();
			return;
		}
		FilterFileTreeByQuery(q);
		return;
	}

	if (m_sidebarActiveTab == 1)
	{
		SearchOutlineByQuery(q);
		return;
	}
}

void CMarkdownEditorDlg::OnTimer(UINT_PTR nIDEvent)
{
	try
	{
		SehTranslatorScope sehScope;
		if (nIDEvent == kDebounceTimerId)
		{
			KillTimer(kDebounceTimerId);
			if (m_bMarkdownMode)
				StartMarkdownParsingAsync();
		}
		else if (nIDEvent == kFormatTimerId)
		{
			ApplyMarkdownFormattingChunk();
		}
		else if (nIDEvent == kInsertTimerId)
		{
			if (m_textModeResetActive)
			{
				StartTextModeResetAsync();
			}
			else if (m_loadingActive)
				StartInsertLoadedText();
			else
				KillTimer(kInsertTimerId);
		}
		else if (nIDEvent == kStatusPulseTimerId)
		{
			// Keep the status text responsive while parsing/formatting/loading.
			if (m_renderBusy)
				UpdateStatusText();
			else
				KillTimer(kStatusPulseTimerId);
		}
		else if (nIDEvent == kAutoSaveTimerId)
		{
			TryAutoSaveIfIdle();
		}
		else if (nIDEvent == kFileWatchTimerId)
		{
			if (DetectExternalFileChange())
				HandleExternalFileChange();
		}
		else if (nIDEvent == kMermaidOverlayTimerId)
		{
			// Keep overlay rectangles in sync even when they are currently off-screen.
			if (!m_windowSizingMove && !m_loadingActive && !m_parsingActive && !m_formattingActive)
				UpdateMermaidOverlayRegion(false);
		}
		else if (nIDEvent == kOutlineFollowTimerId)
		{
			// Avoid selection jitter while opening/parsing/formatting.
			if (!m_outlineAutoSelectFirst && !m_loadingActive && !m_parsingActive && !m_formattingActive && !m_renderBusy)
				UpdateOutlineSelectionByViewTop();
		}
		else if (nIDEvent == kOverlaySyncTimerId)
		{
			KillTimer(kOverlaySyncTimerId);
			if (!m_windowSizingMove)
			{
				UpdateMermaidOverlayRegion(true);
				if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
					m_mermaidOverlay.Invalidate(FALSE);
			}
			// RichEdit may still be settling after a programmatic scroll; do a few retries.
			if (m_overlaySyncAttempts > 0)
			{
				--m_overlaySyncAttempts;
				if (m_overlaySyncAttempts > 0)
					SetTimer(kOverlaySyncTimerId, 60, NULL);
			}
		}
		CDialogEx::OnTimer(nIDEvent);
	}
	catch (const SehException& ex)
	{
		KillTimer(kDebounceTimerId);
		KillTimer(kFormatTimerId);
		KillTimer(kInsertTimerId);
		m_editContent.SetRedraw(TRUE);
		m_suppressChangeEvent = false;
		m_loadingActive = false;
		m_parsingActive = false;
		m_formattingActive = false;
		m_textModeResetActive = false;
		CString msg;
		msg.Format(_T("运行失败（异常码=0x%08X），已停止后台任务。"), static_cast<DWORD>(ex.code));
		SetRenderBusy(false, msg);
	}
	catch (...)
	{
		KillTimer(kDebounceTimerId);
		KillTimer(kFormatTimerId);
		KillTimer(kInsertTimerId);
		m_editContent.SetRedraw(TRUE);
		m_suppressChangeEvent = false;
		m_loadingActive = false;
		m_parsingActive = false;
		m_formattingActive = false;
		m_textModeResetActive = false;
		SetRenderBusy(false, _T("运行失败（未知异常），已停止后台任务。"));
	}
}

void CMarkdownEditorDlg::OnSize(UINT nType, int cx, int cy)
{
	CDialogEx::OnSize(nType, cx, cy);

	if (!::IsWindow(m_editContent.GetSafeHwnd()) || cx <= 0 || cy <= 0)
		return;

	UpdateLayout(cx, cy);
	if (!m_windowSizingMove)
	{
		// 表格覆盖层自适应，不再在窗口变化时重跑整文档格式化。
		if (m_bMarkdownMode && !m_tocOverlayBlocks.empty())
			UpdateTocSpacing();
		UpdateMermaidOverlayRegion(true);
	}
	// Fix: maximize/restore can leave static controls with stale paint.
	// During live resizing, avoid synchronous erase/update (causes visible flicker).
	if (!m_windowSizingMove && (nType == SIZE_MAXIMIZED || nType == SIZE_RESTORED))
	{
		if (::IsWindow(m_staticStatus.GetSafeHwnd()))
			m_staticStatus.RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
		if (::IsWindow(m_progressStatus.GetSafeHwnd()))
			m_progressStatus.RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
		if (::IsWindow(m_staticProgressText.GetSafeHwnd()))
			m_staticProgressText.RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
		RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE | RDW_ALLCHILDREN);
	}
	// On some systems, maximize/restore can leave the RichEdit with a stale invalid region
	// (content appears blank until the next scroll). Force a repaint.
	if (!m_windowSizingMove && ::IsWindow(m_editContent.GetSafeHwnd()))
	{
		m_editContent.Invalidate(FALSE);
		m_editContent.UpdateWindow();
	}
}

void CMarkdownEditorDlg::OnEnterSizeMove()
{
	CDialogEx::OnEnterSizeMove();
	m_windowSizingMove = true;
	KillTimer(kOverlaySyncTimerId);
	if (m_mermaidOverlayTimer != 0)
	{
		KillTimer(kMermaidOverlayTimerId);
		m_mermaidOverlayTimer = 0;
	}
	if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		m_mermaidOverlay.ShowWindow(SW_HIDE);
}

void CMarkdownEditorDlg::OnExitSizeMove()
{
	CDialogEx::OnExitSizeMove();
	m_windowSizingMove = false;
	UpdateEditorWrapWidth();
	if (!m_bMarkdownMode)
		return;
	if (m_tocOverlayBlocks.empty())
	{
		if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
			m_mermaidOverlay.ShowWindow(SW_SHOW);
		UpdateMermaidOverlayRegion(true);
		return;
	}
	if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		m_mermaidOverlay.ShowWindow(SW_SHOW);
	UpdateTocSpacing();
	UpdateMermaidOverlayRegion(true);
}

void CMarkdownEditorDlg::BeginSidebarResize(int xClient)
{
	m_sidebarResizing = true;
	m_sidebarResizeStartX = xClient;
	m_sidebarResizeStartWidth = m_sidebarWidth;
	// Keep the content view live (avoid blank/corruption during resize).
	if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		m_mermaidOverlay.SetRedraw(FALSE);
	if (::IsWindow(m_staticOverlay.GetSafeHwnd()))
		m_staticOverlay.SetRedraw(FALSE);
	if (::IsWindow(m_treeSidebar.GetSafeHwnd()))
		m_treeSidebar.SetRedraw(FALSE);
	if (::IsWindow(m_tabSidebar.GetSafeHwnd()))
		m_tabSidebar.SetRedraw(FALSE);

	// Prevent overlay region updates from causing flicker during live-resize.
	if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		m_mermaidOverlay.ShowWindow(SW_HIDE);
	if (m_mermaidOverlayTimer != 0)
	{
		KillTimer(kMermaidOverlayTimerId);
		m_mermaidOverlayTimer = 0;
	}
}

void CMarkdownEditorDlg::UpdateSidebarResize(int xClient)
{
	if (!m_sidebarResizing)
		return;
	int dx = xClient - m_sidebarResizeStartX;
	m_sidebarWidth = max(180, min(520, m_sidebarResizeStartWidth + dx));

	CRect rc;
	GetClientRect(&rc);
	UpdateLayout(rc.Width(), rc.Height());

	// Keep diagrams/HR in sync visually while dragging.
	UpdateMermaidOverlayRegion(true);
	if (::IsWindow(m_editContent.GetSafeHwnd()))
		m_editContent.Invalidate(FALSE);
}

void CMarkdownEditorDlg::EndSidebarResize()
{
	m_sidebarResizing = false;
	if (::IsWindow(m_tabSidebar.GetSafeHwnd()))
		m_tabSidebar.SetRedraw(TRUE);
	if (::IsWindow(m_treeSidebar.GetSafeHwnd()))
		m_treeSidebar.SetRedraw(TRUE);
	if (::IsWindow(m_staticOverlay.GetSafeHwnd()))
		m_staticOverlay.SetRedraw(TRUE);
	if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		m_mermaidOverlay.SetRedraw(TRUE);
	if (::IsWindow(m_editContent.GetSafeHwnd()))
		m_editContent.SetRedraw(TRUE);
	// Repaint after live resize (avoid partial blank areas).
	RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
	if (::IsWindow(m_editContent.GetSafeHwnd()))
	{
		m_editContent.Invalidate(FALSE);
		m_editContent.UpdateWindow();
	}
		UpdateMermaidDiagrams();
		UpdateMermaidSpacing();
		UpdateTocSpacing();
		UpdateMermaidOverlayRegion(true);
	}

void CMarkdownEditorDlg::UpdateEditorWrapWidth()
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return;

	CRect clientRect;
	m_editContent.GetClientRect(&clientRect);
	const int clientWidthPx = clientRect.Width();
	if (clientWidthPx <= 0)
	{
		m_editContent.SetTargetDevice(m_wrapTargetHdc, 0);
		return;
	}

	if (m_wrapTargetHdc == nullptr)
		m_wrapTargetHdc = ::GetDC(nullptr);

	int dpiX = 96;
	if (m_wrapTargetHdc != nullptr)
	{
		const int dcDpi = ::GetDeviceCaps(m_wrapTargetHdc, LOGPIXELSX);
		if (dcDpi >= 96)
			dpiX = dcDpi;
	}

	// Keep auto-wrap while avoiding temporary HDC lifetime risks.
	// A concrete target width is more stable for hidden marker rendering than "0" in this app.
	const int wrapWidthPx = max(clientWidthPx - 8, 1);
	const int wrapWidthTwips = MulDiv(wrapWidthPx, 1440, dpiX);
	m_editContent.SetTargetDevice(m_wrapTargetHdc, wrapWidthTwips);
}

void CMarkdownEditorDlg::UpdateLayout(int cx, int cy)
{
	if (cx <= 0 || cy <= 0)
		return;

	const int padding = 8;
	const int buttonHeight = 24;
	// 固定侧栏宽度，不允许动态调整。
	constexpr int kFixedSidebarWidth = 240;
	const int sidebarWidth = min(kFixedSidebarWidth, max(cx - padding * 2, 0));
	const int gap = 8;
	const int statusHeight = 22;

	auto calcButtonWidth = [&](CButton& button, int minWidth) -> int {
		if (!::IsWindow(button.GetSafeHwnd()))
			return minWidth;
		CString text;
		button.GetWindowText(text);
		CClientDC dc(this);
		CFont* old = dc.SelectObject(&m_fontButton);
		CSize size = dc.GetTextExtent(text);
		dc.SelectObject(old);
		return max(minWidth, size.cx + 22);
	};

	const int openButtonWidth = calcButtonWidth(m_btnOpenFile, 54);
	const int modeButtonWidth = calcButtonWidth(m_btnModeSwitch, 74);
	const int themeButtonWidth = calcButtonWidth(m_btnThemeSwitch, 54);

	// Put the action buttons into the sidebar header (less intrusive than top-right).
	int editTop = padding;
	int editBottom = cy - statusHeight - padding;
	if (editBottom < editTop)
		editBottom = editTop;

	int sidebarLeft = padding;
	int sidebarTop = editTop;
	int sidebarRight = min(sidebarLeft + sidebarWidth, cx - padding);
	int sidebarBottom = editBottom;
	int sidebarHeight = max(sidebarBottom - sidebarTop, 0);
	int sidebarActualWidth = max(sidebarRight - sidebarLeft, 0);

	const int searchHeight = 22;
	const int gapSmall = 4;
			// Sidebar header row: buttons.
		{
			int headerLeft = sidebarLeft + 2;
			int headerTop = sidebarTop;
			int headerWidth = max(sidebarActualWidth - 4, 0);

			const int minSmallW = 44;
			const int minModeW = 70;
			int openW = (headerWidth > 0) ? min(openButtonWidth, max(minSmallW, headerWidth / 4)) : openButtonWidth;
			int themeW = (headerWidth > 0) ? min(themeButtonWidth, max(minSmallW, headerWidth / 4)) : themeButtonWidth;
			if (headerWidth > 0)
			{
				int maxSmallW = max(1, (headerWidth - gapSmall * 2 - minModeW) / 2);
				maxSmallW = max(minSmallW, maxSmallW);
				openW = min(openW, maxSmallW);
				themeW = min(themeW, maxSmallW);
			}
			int modeX = headerLeft + openW + gapSmall + themeW + gapSmall;
			int modeW = max(1, headerLeft + headerWidth - modeX);

			if (::IsWindow(m_btnOpenFile.GetSafeHwnd()))
				m_btnOpenFile.MoveWindow(headerLeft, headerTop, openW, buttonHeight);
			if (::IsWindow(m_btnThemeSwitch.GetSafeHwnd()))
				m_btnThemeSwitch.MoveWindow(headerLeft + openW + gapSmall, headerTop, themeW, buttonHeight);
			if (::IsWindow(m_btnModeSwitch.GetSafeHwnd()))
				m_btnModeSwitch.MoveWindow(modeX, headerTop, modeW, buttonHeight);
		}
	if (::IsWindow(m_editFileSearch.GetSafeHwnd()))
		m_editFileSearch.MoveWindow(sidebarLeft + 2, sidebarTop + buttonHeight + gapSmall, max(sidebarActualWidth - 4, 0), searchHeight);

	int tabTop = sidebarTop + buttonHeight + gapSmall + searchHeight + gapSmall;
	int tabHeight = max(sidebarHeight - (buttonHeight + gapSmall + searchHeight + gapSmall), 0);
	if (::IsWindow(m_tabSidebar.GetSafeHwnd()))
	{
		m_tabSidebar.MoveWindow(sidebarLeft, tabTop, sidebarActualWidth, tabHeight);
		// Make the two tabs behave like a full-width segmented control (Typora-like).
		m_tabSidebar.ModifyStyle(0, TCS_FIXEDWIDTH);
		const int tabHeaderHeight = 26;
		int count = m_tabSidebar.GetItemCount();
		if (count <= 0)
			count = 1;
		// Leave a small safety margin so the common control won't show scroll arrows.
		int itemWidth = max(1, (sidebarActualWidth - 30) / count);
		m_tabSidebar.SendMessage(TCM_SETITEMSIZE, 0, MAKELPARAM(itemWidth, max(1, tabHeaderHeight - 6)));
		// Hide tab scroll arrows (up-down control) if they were created.
		HWND hUpDown = ::FindWindowEx(m_tabSidebar.GetSafeHwnd(), nullptr, _T("msctls_updown32"), nullptr);
		if (hUpDown)
		{
			::ShowWindow(hUpDown, SW_HIDE);
			::EnableWindow(hUpDown, FALSE);
		}
	}

	const int tabHeaderHeight = 26;
	int listTop = tabTop + tabHeaderHeight;
	int listHeight = max(sidebarBottom - listTop, 0);
	if (::IsWindow(m_treeSidebar.GetSafeHwnd()))
	{
		m_treeSidebar.MoveWindow(sidebarLeft + 2, listTop + 2, max(sidebarActualWidth - 4, 0), max(listHeight - 4, 0));
		// Ensure the TreeView is visible and above the TabCtrl background.
		m_treeSidebar.ShowWindow(SW_SHOW);
		m_treeSidebar.BringWindowToTop();
	}

	// Ensure correct Z-order: the tab control is a background host for the header,
	// while the TreeView renders the actual list content.
	HWND hwndTab = m_tabSidebar.GetSafeHwnd();
	HWND hwndTree = m_treeSidebar.GetSafeHwnd();
	HWND hwndSearch = m_editFileSearch.GetSafeHwnd();
	// Put the TabCtrl to the bottom; TreeView must stay above it.
	if (hwndTab)
		::SetWindowPos(hwndTab, HWND_BOTTOM, 0, 0, 0, 0,
			SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
	if (hwndTree)
		::SetWindowPos(hwndTree, HWND_TOP, 0, 0, 0, 0,
			SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
	if (hwndSearch)
		::SetWindowPos(hwndSearch, HWND_TOP, 0, 0, 0, 0,
			SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

	// Splitter hit-test region (between sidebar and editor)
	m_sidebarSplitterX = sidebarRight + gap / 2;
	m_sidebarSplitterTop = sidebarTop;
	m_sidebarSplitterBottom = sidebarBottom;
	if (::IsWindow(m_sidebarSplitter.GetSafeHwnd()))
	{
		const int splitWidth = max(6, gap);
		int splitLeft = sidebarRight;
		m_sidebarSplitter.MoveWindow(splitLeft, sidebarTop, splitWidth, sidebarBottom - sidebarTop);
		m_sidebarSplitter.Invalidate(FALSE);
	}

	int editLeft = sidebarRight + gap;
	// Keep the editor close to the right frame so the scrollbar aligns with the window edge.
	const int editRightPadding = 1;
int editWidth = max(cx - editLeft - editRightPadding, 0);
int editHeight = max(editBottom - editTop, 0);
if (::IsWindow(m_editContent.GetSafeHwnd()))
	m_editContent.MoveWindow(editLeft, editTop, editWidth, editHeight);
// Live window resize: defer expensive RichEdit wrap-target reflow to OnExitSizeMove.
if (!m_windowSizingMove)
	UpdateEditorWrapWidth();
if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
	m_mermaidOverlay.MoveWindow(editLeft, editTop, editWidth, editHeight);
	if (::IsWindow(m_staticOverlay.GetSafeHwnd()))
		m_staticOverlay.MoveWindow(editLeft, editTop, editWidth, editHeight);

	int statusWidth = max(cx - padding * 2, 0);
	int statusTop = cy - statusHeight - padding / 2;
	int minStatusTop = editBottom + padding / 2;
	if (statusTop < minStatusTop)
		statusTop = minStatusTop;
	if (statusTop + statusHeight > cy)
		statusTop = max(cy - statusHeight, editTop);

	// Status text uses natural content width; progress area fills from text-right to window-right.
	const int statusGap = 6;
	const int progressTextWidth = 72;
	const int minStatusTextWidth = 40;
	const int minProgressBarWidth = 80;

	int statusTextWidth = minStatusTextWidth;
	CString statusText;
	if (::IsWindow(m_staticStatus.GetSafeHwnd()))
		m_staticStatus.GetWindowText(statusText);
	if (statusText.IsEmpty())
		statusText = m_lastStatusText;

	if (!statusText.IsEmpty())
	{
		CClientDC dc(this);
		CFont* old = nullptr;
		if (::IsWindow(m_staticStatus.GetSafeHwnd()))
		{
			CFont* statusFont = m_staticStatus.GetFont();
			if (statusFont != nullptr)
				old = dc.SelectObject(statusFont);
		}
		CSize textSize = dc.GetTextExtent(statusText);
		if (old != nullptr)
			dc.SelectObject(old);
		statusTextWidth = max(minStatusTextWidth, textSize.cx + 12);
	}

	int maxStatusTextWidth = statusWidth - (statusGap + minProgressBarWidth + statusGap + progressTextWidth);
	maxStatusTextWidth = max(minStatusTextWidth, maxStatusTextWidth);
	statusTextWidth = min(statusTextWidth, maxStatusTextWidth);

	int progressTextLeft = padding + statusWidth - progressTextWidth;
	int progressBarLeft = padding + statusTextWidth + statusGap;
	int progressBarWidth = max(1, progressTextLeft - statusGap - progressBarLeft);

	if (progressBarWidth < minProgressBarWidth)
	{
		int requiredStatusWidth = statusWidth - (statusGap + minProgressBarWidth + statusGap + progressTextWidth);
		requiredStatusWidth = max(minStatusTextWidth, requiredStatusWidth);
		statusTextWidth = min(statusTextWidth, requiredStatusWidth);
		progressBarLeft = padding + statusTextWidth + statusGap;
		progressBarWidth = max(1, progressTextLeft - statusGap - progressBarLeft);
	}

	if (::IsWindow(m_staticStatus.GetSafeHwnd()))
		m_staticStatus.MoveWindow(padding, statusTop, statusTextWidth, statusHeight);
	if (::IsWindow(m_progressStatus.GetSafeHwnd()))
		m_progressStatus.MoveWindow(progressBarLeft, statusTop + 2, progressBarWidth, max(1, statusHeight - 4));
	if (::IsWindow(m_staticProgressText.GetSafeHwnd()))
		m_staticProgressText.MoveWindow(progressTextLeft, statusTop, progressTextWidth, statusHeight);
}

void CMarkdownEditorDlg::ApplyTheme(bool dark)
{
	m_themeDark = dark;
	m_themeBackground = dark ? RGB(34, 34, 34) : RGB(255, 255, 255);
	m_themeForeground = dark ? RGB(230, 230, 230) : RGB(51, 51, 51);
	m_themeSidebarBackground = dark ? RGB(28, 28, 28) : RGB(248, 248, 248);
	ApplyImmersiveDarkTitleBar(dark);

	// Mermaid bitmaps depend on theme colors. Drop cache to prevent stale "white blocks".
	for (auto& kv : m_mermaidBitmapCache)
	{
		if (kv.second)
			::DeleteObject(kv.second);
	}
	for (auto& kv : m_mermaidTransientBitmapCache)
	{
		if (kv.second)
			::DeleteObject(kv.second);
	}
	m_mermaidBitmapCache.clear();
	m_mermaidBitmapSize.clear();
	m_mermaidTransientBitmapCache.clear();
	m_mermaidTransientBitmapSize.clear();
	m_mermaidBitmapLastUseSeq.clear();
	m_mermaidBitmapUseSeq = 0;
	m_mermaidFetchInFlight.clear();

	// Recreate brushes used by OnCtlColor.
	if (m_brushDialogBackground.GetSafeHandle() != nullptr)
		m_brushDialogBackground.DeleteObject();
	m_brushDialogBackground.CreateSolidBrush(m_themeBackground);
	if (m_brushSidebarBackground.GetSafeHandle() != nullptr)
		m_brushSidebarBackground.DeleteObject();
	m_brushSidebarBackground.CreateSolidBrush(m_themeSidebarBackground);

	// Background brushes for controls (best-effort, keep it simple).
	m_editContent.SetBackgroundColor(FALSE, m_themeBackground);

	// Best-effort dark scrollbars / themed controls.
	ApplyDarkModeToWindow(m_editContent.GetSafeHwnd(), dark);
	ApplyDarkModeToWindow(m_treeSidebar.GetSafeHwnd(), dark);
	ApplyDarkModeToWindow(m_tabSidebar.GetSafeHwnd(), dark);
	ApplyDarkModeToWindow(m_editFileSearch.GetSafeHwnd(), dark);
	ApplyDarkModeToWindow(m_sidebarSplitter.GetSafeHwnd(), dark);
	RefreshDarkModeState();

	if (::IsWindow(m_treeSidebar.GetSafeHwnd()))
	{
		m_treeSidebar.SetBkColor(m_themeSidebarBackground);
		m_treeSidebar.SetTextColor(m_themeForeground);
		m_treeSidebar.SendMessage(WM_THEMECHANGED, 0, 0);
		m_treeSidebar.Invalidate(FALSE);
	}
	if (::IsWindow(m_tabSidebar.GetSafeHwnd()))
	{
		m_tabSidebar.SendMessage(WM_THEMECHANGED, 0, 0);
		m_tabSidebar.Invalidate();
		m_tabSidebar.UpdateWindow();
	}
	if (::IsWindow(m_editFileSearch.GetSafeHwnd()))
	{
		m_editFileSearch.Invalidate();
		m_editFileSearch.UpdateWindow();
	}
	if (::IsWindow(m_editContent.GetSafeHwnd()))
	{
		m_editContent.SendMessage(WM_THEMECHANGED, 0, 0);
		m_editContent.Invalidate();
		m_editContent.UpdateWindow();
	}

	// Reset normal/base formats so next formatting pass uses new theme colors.
	m_normalFormat = CHARFORMAT2{};
	m_hiddenFormat = CHARFORMAT2{};
	PrepareMarkdownFormats();

	// Apply base colors immediately.
	m_editContent.SetSel(0, -1);
	CHARFORMAT2 base = m_normalFormat;
	base.crBackColor = m_themeBackground;
	base.crTextColor = m_themeForeground;
	base.dwMask |= CFM_COLOR | CFM_BACKCOLOR;
	base.dwEffects = 0;
	m_editContent.SetSelectionCharFormat(base);
	m_editContent.SetSel(0, 0);

	// Make the "hidden" spans match background so markers truly disappear.
	CHARFORMAT2 hidden = m_hiddenFormat;
	hidden.crTextColor = m_themeBackground;
	hidden.crBackColor = m_themeBackground;
	m_hiddenFormat = hidden;

	// Force repaint for all child windows (buttons/tree/tab/search) and erase background.
	RedrawWindow(nullptr, nullptr,
		RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);

	if (m_bMarkdownMode)
	{
		ScheduleMarkdownFormatting(1);
		UpdateMermaidDiagrams();
		UpdateMermaidSpacing();
		UpdateTocSpacing();
		UpdateMermaidOverlayRegion(true);
	}
}

void CMarkdownEditorDlg::ApplyImmersiveDarkTitleBar(bool enable)
{
	HWND hwnd = GetSafeHwnd();
	if (hwnd == nullptr)
		return;

	HMODULE hDwm = ::LoadLibraryW(L"dwmapi.dll");
	if (!hDwm)
		return;

	typedef HRESULT(WINAPI* DwmSetWindowAttributeFn)(HWND, DWORD, LPCVOID, DWORD);
	auto fn = reinterpret_cast<DwmSetWindowAttributeFn>(::GetProcAddress(hDwm, "DwmSetWindowAttribute"));
	if (!fn)
	{
		::FreeLibrary(hDwm);
		return;
	}

	BOOL value = enable ? TRUE : FALSE;
	// DWMWA_USE_IMMERSIVE_DARK_MODE is 20 on 1903+, 19 on 1809.
	fn(hwnd, 20, &value, sizeof(value));
	fn(hwnd, 19, &value, sizeof(value));
	::FreeLibrary(hDwm);
}

void CMarkdownEditorDlg::OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	if (!lpDrawItemStruct)
		return;

	const UINT id = lpDrawItemStruct->CtlID;
	if (id != IDC_BUTTON_OPEN && id != IDC_BUTTON_MODE && id != IDC_BUTTON_THEME)
	{
		CDialogEx::OnDrawItem(nIDCtl, lpDrawItemStruct);
		return;
	}

	CDC dc;
	dc.Attach(lpDrawItemStruct->hDC);
	CRect rc = lpDrawItemStruct->rcItem;

	const bool disabled = (lpDrawItemStruct->itemState & ODS_DISABLED) != 0;
	const bool selected = (lpDrawItemStruct->itemState & ODS_SELECTED) != 0;

	COLORREF bg = m_themeDark ? RGB(44, 44, 44) : RGB(245, 245, 245);
	COLORREF border = m_themeDark ? RGB(70, 70, 70) : RGB(210, 210, 210);
	COLORREF text = m_themeForeground;
	if (selected)
		bg = m_themeDark ? RGB(60, 60, 60) : RGB(230, 230, 230);
	if (disabled)
		text = m_themeDark ? RGB(140, 140, 140) : RGB(150, 150, 150);

	CBrush brush(bg);
	dc.FillRect(rc, &brush);

	CPen pen(PS_SOLID, 1, border);
	CPen* oldPen = dc.SelectObject(&pen);
	CBrush* oldBrush = (CBrush*)dc.SelectStockObject(NULL_BRUSH);
	dc.RoundRect(rc, CPoint(6, 6));
	dc.SelectObject(oldBrush);
	dc.SelectObject(oldPen);

	CString caption;
	GetDlgItemText(id, caption);
	dc.SetBkMode(TRANSPARENT);
	dc.SetTextColor(text);

	CRect textRc = rc;
	if (selected)
		textRc.OffsetRect(1, 1);
	dc.DrawText(caption, textRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

	dc.Detach();
}

HBRUSH CMarkdownEditorDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);
	if (pWnd == nullptr)
		return hbr;

	UINT id = static_cast<UINT>(pWnd->GetDlgCtrlID());
	if (id == IDC_EDIT_FILE_SEARCH)
	{
		pDC->SetBkMode(OPAQUE);
		pDC->SetBkColor(m_themeSidebarBackground);
		pDC->SetTextColor(m_themeForeground);
		return static_cast<HBRUSH>(m_brushSidebarBackground.GetSafeHandle());
	}
	if (id == IDC_TAB_SIDEBAR || id == IDC_TREE_SIDEBAR)
	{
		pDC->SetBkMode(OPAQUE);
		pDC->SetBkColor(m_themeSidebarBackground);
		pDC->SetTextColor(m_themeForeground);
		return static_cast<HBRUSH>(m_brushSidebarBackground.GetSafeHandle());
	}

	if (id == IDC_STATIC_OVERLAY)
	{
		if (m_brushOverlay.GetSafeHandle() != nullptr)
			m_brushOverlay.DeleteObject();
		COLORREF overlayColor = m_themeBackground;
		m_brushOverlay.CreateSolidBrush(overlayColor);
		pDC->SetBkMode(OPAQUE);
		pDC->SetBkColor(overlayColor);
		pDC->SetTextColor(m_themeForeground);
		return static_cast<HBRUSH>(m_brushOverlay.GetSafeHandle());
	}

	if (nCtlColor == CTLCOLOR_DLG)
	{
		pDC->SetBkColor(m_themeBackground);
		return static_cast<HBRUSH>(m_brushDialogBackground.GetSafeHandle());
	}

	if (id == IDC_STATIC_STATUS || id == IDC_STATIC_PROGRESS_TEXT)
	{
		pDC->SetBkMode(OPAQUE);
		pDC->SetBkColor(m_themeBackground);
		pDC->SetTextColor(m_themeForeground);
		return static_cast<HBRUSH>(m_brushDialogBackground.GetSafeHandle());
	}

	return hbr;
}

void CMarkdownEditorDlg::FilterFileTreeByQuery(const CString& query)
{
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;

	CString q = query;
	q.Trim();
	q.MakeLower();

	m_treeSidebar.SetRedraw(FALSE);
	m_treeSidebar.DeleteAllItems();
	m_fileTreePaths.clear();

	if (q.IsEmpty())
	{
		BuildFileTree(m_strCurrentFile);
		// Recent files are shown under the folder root via UpdateSidebar() to avoid duplicate roots.
		if (m_treeSidebar.GetRootItem() == nullptr)
		{
			m_treeSidebar.InsertItem(_T("(暂无文件)"), TVI_ROOT, TVI_LAST);
		}

		m_treeSidebar.SetRedraw(TRUE);
		m_treeSidebar.Invalidate();
		return;
	}

	// Search markdown files in current folder by partial match.
	CString folder = m_strCurrentFile;
	int sep = folder.ReverseFind('\\');
	if (sep == -1)
		folder.Empty();
	else
		folder = folder.Left(sep);

	const CString folderCaption = folder.IsEmpty() ? _T("(当前目录)") : folder;
	HTREEITEM root = nullptr;
	CFileFind finder;
	CString pattern = folder.IsEmpty() ? _T("*.*") : (folder + _T("\\*.*"));
	BOOL working = finder.FindFile(pattern);
	int added = 0;
	while (working)
	{
		working = finder.FindNextFile();
		if (finder.IsDots() || finder.IsDirectory())
			continue;

		CString path = finder.GetFilePath();
		if (!IsMarkdownFile(path))
			continue;

		CString name = finder.GetFileName();
		CString lower = name;
		lower.MakeLower();
		if (lower.Find(q) == -1)
			continue;

		if (root == nullptr)
			root = m_treeSidebar.InsertItem(folderCaption, TVI_ROOT, TVI_LAST);
		HTREEITEM item = m_treeSidebar.InsertItem(name, root, TVI_LAST);
		m_treeSidebar.SetItemData(item, MakeTreeData(kTreeDataFolderTag, m_fileTreePaths.size()));
		m_fileTreePaths.push_back(path);
		++added;
		if (added >= 200)
			break;
	}
	finder.Close();

	// Also search recent files (even if outside current folder).
	if (!m_recentFiles.empty())
	{
		HTREEITEM recentRoot = m_treeSidebar.InsertItem(_T("最近"), TVI_ROOT, TVI_LAST);
		int recentAdded = 0;
		for (size_t i = 0; i < m_recentFiles.size(); ++i)
		{
			CString path = m_recentFiles[i];
			if (!IsMarkdownFile(path))
				continue;
			CString name = path;
			int pos = name.ReverseFind('\\');
			if (pos != -1)
				name = name.Mid(pos + 1);
			CString lower = name;
			lower.MakeLower();
			if (lower.Find(q) == -1)
				continue;

			HTREEITEM item = m_treeSidebar.InsertItem(name, recentRoot, TVI_LAST);
			m_treeSidebar.SetItemData(item, MakeTreeData(kTreeDataRecentTag, m_fileTreePaths.size()));
			m_fileTreePaths.push_back(path);
			++recentAdded;
			if (recentAdded >= 50)
				break;
		}
		if (recentAdded > 0)
			m_treeSidebar.Expand(recentRoot, TVE_EXPAND);
		else
			m_treeSidebar.InsertItem(_T("(无匹配文件)"), recentRoot, TVI_LAST);
	}

	if (added > 0 && root != nullptr)
		m_treeSidebar.Expand(root, TVE_EXPAND);

	m_treeSidebar.SetRedraw(TRUE);
	m_treeSidebar.Invalidate();
}

void CMarkdownEditorDlg::LoadFileContent(const CString& filePath)
{
	// Legacy sync load path; kept for small files or fallback.
	CFile file;
	if (file.Open(filePath, CFile::modeRead | CFile::typeBinary))
	{
		ULONGLONG fileSize = file.GetLength();
		if (fileSize > 0)
		{
			if (fileSize > static_cast<ULONGLONG>(SIZE_MAX - 1) || fileSize > static_cast<ULONGLONG>(INT_MAX))
			{
				file.Close();
				AfxMessageBox(_T("文件过大，无法使用旧的同步加载路径。"), MB_OK | MB_ICONWARNING);
				return;
			}

			// Read file as binary
			const size_t size = static_cast<size_t>(fileSize);
			std::vector<BYTE> buffer(size + 1);
			const UINT readBytes = file.Read(buffer.data(), static_cast<UINT>(size));
			buffer[readBytes] = 0;
			file.Close();
			if (readBytes != size)
			{
				AfxMessageBox(_T("读取文件失败。"), MB_OK | MB_ICONWARNING);
				return;
			}

			// Detect UTF-8 BOM first; otherwise validate UTF-8 and fallback to ACP
			bool hasUtf8Bom = false;
			const BYTE* actualData = buffer.data();
			size_t actualSize = size;

			if (size >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
			{
				// UTF-8 with BOM
				hasUtf8Bom = true;
				actualData = buffer.data() + 3;
				actualSize = size - 3;
			}

			CString content;
			int wideLen = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, (LPCSTR)actualData,
				(int)actualSize, NULL, 0);
			bool converted = false;
			if (wideLen > 0)
			{
				std::vector<WCHAR> wideBuffer(static_cast<size_t>(wideLen) + 1);
				MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)actualData, (int)actualSize, wideBuffer.data(), wideLen);
				wideBuffer[wideLen] = 0;
				content = wideBuffer.data();
				converted = true;
			}

			if (!converted && !hasUtf8Bom)
			{
				// Fallback for non-UTF8 text (ANSI/GBK on CN systems)
				wideLen = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)actualData, (int)actualSize, NULL, 0);
				if (wideLen > 0)
				{
					std::vector<WCHAR> wideBuffer(static_cast<size_t>(wideLen) + 1);
					MultiByteToWideChar(CP_ACP, 0, (LPCSTR)actualData, (int)actualSize, wideBuffer.data(), wideLen);
					wideBuffer[wideLen] = 0;
					content = wideBuffer.data();
				}
			}

			m_suppressChangeEvent = true;
			m_editContent.SetWindowText(content);
			m_suppressChangeEvent = false;

			if (m_bMarkdownMode)
			{
				ScheduleMarkdownFormatting(1);
			}
		}
		else
		{
			file.Close();
			m_editContent.SetWindowText(_T(""));
		}
	}
	else
	{
		AfxMessageBox(_T("无法打开文件！"));
	}
}

void CMarkdownEditorDlg::ApplyMarkdownHighlight()
{
	if (!m_bMarkdownMode) return;
	ScheduleMarkdownFormatting(1);
}

void CMarkdownEditorDlg::StartLoadDocumentAsync(const CString& filePath)
{
	if (filePath.IsEmpty())
		return;

	const bool treatAsMarkdown = (m_bMarkdownMode && IsMarkdownFile(filePath));

	CancelMarkdownFormatting();
	// Always hide previous content immediately when opening a new document
	// to avoid flashing stale text before the loading overlay appears.
	m_hideDocumentUntilRenderReady = true;
	m_loadBytesRead = 0;
	m_loadBytesTotal = 0;
	m_loadingActive = true;
	SetRenderBusy(true, _T("正在加载中..."));
	m_textDirty = false;
	m_loadingText.clear();
	m_loadingTextPos = 0;

	const unsigned long long generation = ++m_loadGeneration;
	HWND hwnd = GetSafeHwnd();
	CString filePathCopy = filePath;
	if (hwnd == nullptr)
		return;
	auto cancelToken = m_asyncCancelToken;
	auto workerCount = m_asyncWorkerCount;
	std::thread([hwnd, generation, filePathCopy, treatAsMarkdown, cancelToken, workerCount]() {
		AsyncWorkerScope workerScope(workerCount);
		if (cancelToken && cancelToken->load(std::memory_order_relaxed))
			return;

		auto payload = std::make_unique<MarkdownParsePayload>();
		payload->generation = generation;

		try
		{
			CFile file;
			if (!file.Open(filePathCopy, CFile::modeRead | CFile::typeBinary))
			{
				SafePostMessageUnique(hwnd, WM_DOCUMENT_LOAD_COMPLETE, 0, std::move(payload));
				return;
			}

			ULONGLONG fileSize = file.GetLength();
			if (fileSize == 0)
			{
				file.Close();
				SafePostMessageUnique(hwnd, WM_DOCUMENT_LOAD_COMPLETE, 0, std::move(payload));
				return;
			}
			if (fileSize > kMaxDocumentLoadBytes)
			{
				file.Close();
				payload->loadFailed = true;
				CString msg;
				msg.Format(_T("文件过大（上限 %llu MB），已拒绝加载。"),
					static_cast<unsigned long long>(kMaxDocumentLoadBytes / (1024ull * 1024ull)));
				payload->loadErrorMessage = msg.GetString();
				SafePostMessageUnique(hwnd, WM_DOCUMENT_LOAD_COMPLETE, 0, std::move(payload));
				return;
			}
			if (fileSize > static_cast<ULONGLONG>((std::numeric_limits<size_t>::max)() - 1))
			{
				file.Close();
				payload->loadFailed = true;
				payload->loadErrorMessage = L"文件尺寸超出当前进程可处理范围。";
				SafePostMessageUnique(hwnd, WM_DOCUMENT_LOAD_COMPLETE, 0, std::move(payload));
				return;
			}

			auto postLoadProgress = [&](ULONGLONG bytesRead, ULONGLONG bytesTotal) {
				if (cancelToken && cancelToken->load(std::memory_order_relaxed))
					return;
				auto progress = std::make_unique<DocumentLoadProgressPayload>();
				progress->generation = generation;
				progress->bytesRead = static_cast<unsigned long long>(bytesRead);
				progress->bytesTotal = static_cast<unsigned long long>(bytesTotal);
				SafePostMessageUnique(hwnd, WM_DOCUMENT_LOAD_PROGRESS, 0, std::move(progress));
			};
			postLoadProgress(0, fileSize);

			const size_t size = static_cast<size_t>(fileSize);
			std::unique_ptr<BYTE[]> buffer(new BYTE[size + 1]);
			ULONGLONG totalRead = 0;
			int lastPostedPercent = -1;
			constexpr UINT kReadChunkBytes = 1024 * 1024;
			while (totalRead < fileSize)
			{
				if (cancelToken && cancelToken->load(std::memory_order_relaxed))
					return;

				const ULONGLONG remain = fileSize - totalRead;
				const UINT toRead = static_cast<UINT>((remain < static_cast<ULONGLONG>(kReadChunkBytes))
					? remain : static_cast<ULONGLONG>(kReadChunkBytes));
				const UINT readNow = file.Read(buffer.get() + static_cast<size_t>(totalRead), toRead);
				if (readNow == 0)
					break;
				totalRead += readNow;

				const int percent = (fileSize > 0) ? static_cast<int>((totalRead * 100ULL) / fileSize) : 100;
				if (percent != lastPostedPercent)
				{
					lastPostedPercent = percent;
					postLoadProgress(totalRead, fileSize);
				}
			}
			if (totalRead != fileSize)
			{
				file.Close();
				payload->loadFailed = true;
				payload->loadErrorMessage = L"读取文件失败，文件可能在加载过程中被占用或变更。";
				payload->text.clear();
				payload->displayText.clear();
				SafePostMessageUnique(hwnd, WM_DOCUMENT_LOAD_COMPLETE, 0, std::move(payload));
				return;
			}
			buffer[size] = 0;
			file.Close();
			postLoadProgress(fileSize, fileSize);

			bool hasUtf8Bom = false;
			BYTE* actualData = buffer.get();
			ULONGLONG actualSize = fileSize;
			if (fileSize >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
			{
				hasUtf8Bom = true;
				actualData = buffer.get() + 3;
				actualSize = fileSize - 3;
			}

			bool converted = false;
			if (actualSize > static_cast<ULONGLONG>((std::numeric_limits<int>::max)()))
			{
				payload->loadFailed = true;
				payload->loadErrorMessage = L"文件过大，文本解码长度超出 Windows API 支持范围。";
				SafePostMessageUnique(hwnd, WM_DOCUMENT_LOAD_COMPLETE, 0, std::move(payload));
				return;
			}
			const int actualSizeInt = static_cast<int>(actualSize);

			if (actualSizeInt == 0)
			{
				converted = true;
			}
			else
			{
				int wideLen = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, (LPCSTR)actualData,
					actualSizeInt, NULL, 0);
				if (wideLen > 0)
				{
					payload->text.resize(wideLen);
					MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)actualData, actualSizeInt,
						reinterpret_cast<LPWSTR>(&payload->text[0]), wideLen);
					converted = true;
				}

				if (!converted && !hasUtf8Bom)
				{
					wideLen = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)actualData, actualSizeInt, NULL, 0);
					if (wideLen > 0)
					{
						payload->text.resize(wideLen);
						MultiByteToWideChar(CP_ACP, 0, (LPCSTR)actualData, actualSizeInt,
							reinterpret_cast<LPWSTR>(&payload->text[0]), wideLen);
						converted = true;
					}
				}
			}
			if (!converted)
			{
				payload->loadFailed = true;
				payload->loadErrorMessage = L"无法识别文件编码，当前仅支持 UTF-8/系统 ANSI。";
				SafePostMessageUnique(hwnd, WM_DOCUMENT_LOAD_COMPLETE, 0, std::move(payload));
				return;
			}

			payload->displayText = payload->text;
			if (treatAsMarkdown)
			{
				auto transformed = RenderText::TransformWithMapping(payload->text);
				payload->displayText = std::move(transformed.displayText);
				payload->displayToSourceIndexMap = std::move(transformed.mapping.displayToSource);
				payload->sourceToDisplayIndexMap = std::move(transformed.mapping.sourceToDisplay);
			}
			else
			{
				BuildIdentityIndexMap(payload->text.size(),
					payload->displayToSourceIndexMap,
					payload->sourceToDisplayIndexMap);
			}
		}
		catch (...)
		{
			payload->loadFailed = true;
			payload->loadErrorMessage = L"加载文件时发生未处理异常。";
			payload->text.clear();
			payload->displayText.clear();
			payload->displayToSourceIndexMap.clear();
			payload->sourceToDisplayIndexMap.clear();
		}

		if (!cancelToken || !cancelToken->load(std::memory_order_relaxed))
			SafePostMessageUnique(hwnd, WM_DOCUMENT_LOAD_COMPLETE, 0, std::move(payload));
	}).detach();
}

LRESULT CMarkdownEditorDlg::OnDocumentLoadComplete(WPARAM /*wParam*/, LPARAM lParam)
{
	try
	{
		SehTranslatorScope sehScope;
		auto payload = reinterpret_cast<MarkdownParsePayload*>(lParam);
		if (!payload)
		{
			// No payload means we are continuing chunked insertion via timer.
			if (m_destroying)
				return 0;
			StartInsertLoadedText();
			return 0;
		}
		std::unique_ptr<MarkdownParsePayload> holder(payload);
		if (m_destroying)
			return 0;
		if (payload->exceptionCode != 0)
		{
			CString msg;
			msg.Format(_T("解析失败（异常码=0x%08X），已跳过渲染。"), payload->exceptionCode);
			m_loadBytesRead = 0;
			m_loadBytesTotal = 0;
			SetRenderBusy(false, msg);
			m_parsingActive = false;
			m_parseReschedule = false;
			UpdateStatusText();
			return 0;
		}

		if (payload->generation != m_loadGeneration.load())
			return 0;

		if (payload->text.empty())
		{
			m_loadingActive = false;
			m_loadBytesRead = 0;
			m_loadBytesTotal = 0;
			m_hideDocumentUntilRenderReady = false;
			m_rawDocumentText.clear();
			m_displayToSourceIndexMap.clear();
			m_sourceToDisplayIndexMap.clear();
			m_richToSourceIndexMap.clear();
			m_loadingText.clear();
			m_loadingTextPos = 0;
			m_fullTextReplaceTask = FullTextReplaceTask::None;
			m_startFormattingAfterFullTextReplace = false;
			m_startTextModeResetAfterFullTextReplace = false;
			m_editContent.SetWindowText(_T(""));
			UpdateCurrentFileSignatureFromDisk();
			if (payload->loadFailed && !payload->loadErrorMessage.empty())
			{
				SetRenderBusy(false, payload->loadErrorMessage.c_str());
			}
			else
			{
				SetRenderBusy(false, _T(""));
			}
			return 0;
		}

		m_loadBytesRead = m_loadBytesTotal;
		m_rawDocumentText = std::move(payload->text);
		m_loadingText = std::move(payload->displayText);
		m_displayToSourceIndexMap = std::move(payload->displayToSourceIndexMap);
		m_sourceToDisplayIndexMap = std::move(payload->sourceToDisplayIndexMap);
		EnsureValidSourceDisplayMapping(m_rawDocumentText,
			m_sourceToDisplayIndexMap,
			m_displayToSourceIndexMap,
			_T("OnDocumentLoadComplete"));
		m_loadingTextPos = 0;
		m_pendingSourceText.clear();
		m_fullTextReplaceTask = FullTextReplaceTask::LoadFile;
		m_startFormattingAfterFullTextReplace = m_bMarkdownMode;
		m_startTextModeResetAfterFullTextReplace = false;
		StartInsertLoadedText();
		return 0;
	}
	catch (const SehException& ex)
	{
		KillTimer(kInsertTimerId);
		m_editContent.SetRedraw(TRUE);
		m_suppressChangeEvent = false;
		m_loadingActive = false;
		m_loadBytesRead = 0;
		m_loadBytesTotal = 0;
		m_hideDocumentUntilRenderReady = false;
		m_parsingActive = false;
		m_formattingActive = false;
		CString msg;
		msg.Format(_T("打开文件失败（异常码=0x%08X）。"), static_cast<DWORD>(ex.code));
		SetRenderBusy(false, msg);
		return 0;
	}
	catch (...)
	{
		KillTimer(kInsertTimerId);
		m_editContent.SetRedraw(TRUE);
		m_suppressChangeEvent = false;
		m_loadingActive = false;
		m_loadBytesRead = 0;
		m_loadBytesTotal = 0;
		m_hideDocumentUntilRenderReady = false;
		m_parsingActive = false;
		m_formattingActive = false;
		SetRenderBusy(false, _T("打开文件失败（未知异常）。"));
		return 0;
	}
}

LRESULT CMarkdownEditorDlg::OnDocumentLoadProgress(WPARAM /*wParam*/, LPARAM lParam)
{
	auto payload = reinterpret_cast<DocumentLoadProgressPayload*>(lParam);
	if (!payload)
		return 0;

	std::unique_ptr<DocumentLoadProgressPayload> holder(payload);
	if (m_destroying)
		return 0;
	if (payload->generation != m_loadGeneration.load())
		return 0;

	m_loadBytesRead = payload->bytesRead;
	m_loadBytesTotal = payload->bytesTotal;
	return 0;
}

LRESULT CMarkdownEditorDlg::OnPrepareDisplayTextComplete(WPARAM /*wParam*/, LPARAM lParam)
{
	auto payload = reinterpret_cast<DisplayTextPayload*>(lParam);
	if (!payload)
		return 0;
	std::unique_ptr<DisplayTextPayload> holder(payload);

	if (m_destroying)
		return 0;
	if (payload->generation != m_displayGeneration.load())
		return 0;
	if (!m_bMarkdownMode)
		return 0;
	if (m_loadingActive)
		return 0;

	m_displayToSourceIndexMap = std::move(payload->displayToSourceIndexMap);
	m_sourceToDisplayIndexMap = std::move(payload->sourceToDisplayIndexMap);
	EnsureValidSourceDisplayMapping(m_rawDocumentText,
		m_sourceToDisplayIndexMap,
		m_displayToSourceIndexMap,
		_T("OnPrepareDisplayTextComplete"));

	StartFullTextReplaceAsync(std::move(payload->displayText),
		FullTextReplaceTask::ApplyDisplayTextBeforeFormat,
		/*startFormattingAfterReplace*/ true,
		/*startTextModeResetAfterReplace*/ false);
	// Sidebar should remain stable during mode switches.
	m_modeSwitchInProgress = false;
	return 0;
}

LRESULT CMarkdownEditorDlg::OnFindReplaceCmd(WPARAM /*wParam*/, LPARAM lParam)
{
	CFindReplaceDialog* dlg = CFindReplaceDialog::GetNotifier(lParam);
	if (!dlg)
		return 0;

	if (dlg->IsTerminating())
	{
		if (dlg == m_findReplaceDialog)
		{
			m_findReplaceDialog = nullptr;
			m_findReplaceMode = false;
		}
		return 0;
	}

	const CString findText = dlg->GetFindString();
	const CString replaceText = dlg->GetReplaceString();
	const bool matchCase = (dlg->MatchCase() != FALSE);
	const bool wholeWord = (dlg->MatchWholeWord() != FALSE);
	const bool searchDown = (dlg->SearchDown() != FALSE);

	m_lastFindText = findText;
	m_lastReplaceText = replaceText;
	m_lastFindMatchCase = matchCase;
	m_lastFindWholeWord = wholeWord;

	if (dlg->FindNext())
	{
		if (!FindNextInEditor(findText, matchCase, wholeWord, searchDown, true))
			AfxMessageBox(_T("未找到匹配内容。"), MB_OK | MB_ICONINFORMATION);
		return 0;
	}

	if (dlg->ReplaceCurrent())
	{
		if ((m_editContent.GetStyle() & ES_READONLY) != 0)
		{
			PromptSwitchToTextModeForEdit();
			return 0;
		}

		if (!ReplaceCurrentInEditor(findText, replaceText, matchCase, wholeWord))
		{
			AfxMessageBox(_T("未找到可替换内容。"), MB_OK | MB_ICONINFORMATION);
		}
		else
		{
			NoteUserActivity(true);
		}
		return 0;
	}

	if (dlg->ReplaceAll())
	{
		if ((m_editContent.GetStyle() & ES_READONLY) != 0)
		{
			PromptSwitchToTextModeForEdit();
			return 0;
		}

		const int count = ReplaceAllInEditor(findText, replaceText, matchCase, wholeWord);
		if (count > 0)
			NoteUserActivity(true);

		CString msg;
		msg.Format(_T("已替换 %d 处内容。"), count);
		AfxMessageBox(msg, MB_OK | MB_ICONINFORMATION);
		return 0;
	}

	return 0;
}

void CMarkdownEditorDlg::StartInsertLoadedText()
{
	if (!m_loadingActive)
		return;
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return;

	// Insert in chunks to keep UI responsive.
	const DWORD startTick = GetTickCount();
	const DWORD budgetMs = 8;
	const size_t chunkChars = 64 * 1024;

	if (m_loadingTextPos == 0)
	{
		m_suppressChangeEvent = true;
		m_editContent.SetRedraw(FALSE);
		m_editContent.SetWindowText(_T(""));
		m_editContent.SetSel(0, 0);
		m_suppressChangeEvent = false;
	}

	while (GetTickCount() - startTick < budgetMs && m_loadingTextPos < m_loadingText.size())
	{
		size_t end = min(m_loadingText.size(), m_loadingTextPos + chunkChars);
		CString chunk(m_loadingText.substr(m_loadingTextPos, end - m_loadingTextPos).c_str());
		m_suppressChangeEvent = true;
		m_editContent.ReplaceSel(chunk, FALSE);
		m_suppressChangeEvent = false;
		m_loadingTextPos = end;
	}

	if (m_loadingTextPos < m_loadingText.size())
	{
		const CString hint = _T("正在加载中...");
		if (!m_renderBusy || m_renderBusyHint != hint)
			SetRenderBusy(true, hint);
		SetTimer(kInsertTimerId, 1, NULL);
		return;
	}

	// Finish insert
	m_loadingActive = false;
	m_textDirty = false;
	KillTimer(kInsertTimerId);
	m_editContent.SetRedraw(TRUE);
	const FullTextReplaceTask finishedTask = m_fullTextReplaceTask;
	const bool startFormatting = m_startFormattingAfterFullTextReplace;
	const bool startTextReset = m_startTextModeResetAfterFullTextReplace;
	// Ensure newly opened documents always start at the beginning.
	{
		m_suppressChangeEvent = true;
		m_editContent.SetSel(0, 0);
		m_editContent.SendMessage(EM_SCROLLCARET, 0, 0);
		m_suppressChangeEvent = false;
	}
	m_editContent.Invalidate(FALSE);
	if (!startFormatting)
		m_editContent.UpdateWindow();

	// Ensure emoji/symbols (e.g. ✅🔧🧪) render even before markdown parsing finishes (or if parsing fails).
	// This is especially important for headings, which often start with such symbols.
	if (m_bMarkdownMode && !startFormatting && !m_rawDocumentText.empty())
	{
		GETTEXTLENGTHEX textLenEx{};
		textLenEx.flags = GTL_NUMCHARS;
		textLenEx.codepage = 1200;
		const long len = static_cast<long>(m_editContent.SendMessage(EM_GETTEXTLENGTHEX,
			reinterpret_cast<WPARAM>(&textLenEx), 0));
		if (len > 0)
			ApplyEmojiFontFallbackToEdit(m_rawDocumentText, len);
	}

	m_fullTextReplaceTask = FullTextReplaceTask::None;
	m_startFormattingAfterFullTextReplace = false;
	m_startTextModeResetAfterFullTextReplace = false;
	m_loadingText.clear();
	if (finishedTask == FullTextReplaceTask::LoadFile)
		UpdateCurrentFileSignatureFromDisk();

	if (finishedTask == FullTextReplaceTask::RestoreRawTextBeforeTextMode)
		m_editContent.SetReadOnly(FALSE);
	else if (m_bMarkdownMode)
		m_editContent.SetReadOnly(TRUE);

	if (startTextReset)
	{
		m_hideDocumentUntilRenderReady = false;
		StartTextModeResetAsync();
		return;
	}
	if (startFormatting)
	{
		ScheduleMarkdownFormatting(1);
		return;
	}

	m_hideDocumentUntilRenderReady = false;
	SetRenderBusy(false, _T(""));
}

void CMarkdownEditorDlg::ApplyTextMode()
{
	CancelMarkdownFormatting();
	UpdateTocSpacing();
	m_tocOverlayBlocks.clear();
	m_tocHitRegions.clear();

	m_codeBlockRanges.clear();
	m_headingRanges.clear();
	m_mermaidBlocks.clear();
	m_mermaidDiagrams.clear();
	m_lastMermaidOverlayRects.clear();
	m_mermaidSpaceAfterTwips.clear();
	m_mermaidSpanIndex = 0;

	// 切到文本模式后不再需要 Mermaid 覆盖层，立即释放位图缓存避免高水位内存驻留。
	for (auto& kv : m_mermaidBitmapCache)
	{
		if (kv.second)
			::DeleteObject(kv.second);
	}
	for (auto& kv : m_mermaidTransientBitmapCache)
	{
		if (kv.second)
			::DeleteObject(kv.second);
	}
	m_mermaidBitmapCache.clear();
	m_mermaidBitmapSize.clear();
	m_mermaidTransientBitmapCache.clear();
	m_mermaidTransientBitmapSize.clear();
	m_mermaidBitmapLastUseSeq.clear();
	m_mermaidBitmapUseSeq = 0;
	m_mermaidFetchInFlight.clear();
	if (m_mermaidOverlayTimer != 0)
	{
		KillTimer(kMermaidOverlayTimerId);
		m_mermaidOverlayTimer = 0;
	}
	if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		m_mermaidOverlay.ShowWindow(SW_HIDE);

	if (!m_rawDocumentText.empty())
	{
		BuildIdentityIndexMap(m_rawDocumentText.size(),
			m_displayToSourceIndexMap,
			m_sourceToDisplayIndexMap);
		m_richToSourceIndexMap.clear();
		StartFullTextReplaceAsync(m_rawDocumentText,
			FullTextReplaceTask::RestoreRawTextBeforeTextMode,
			/*startFormattingAfterReplace*/ false,
			/*startTextModeResetAfterReplace*/ true);
		return;
	}

	m_displayToSourceIndexMap.clear();
	m_sourceToDisplayIndexMap.clear();
	m_richToSourceIndexMap.clear();
	m_editContent.SetReadOnly(FALSE);
	StartTextModeResetAsync();
}

void CMarkdownEditorDlg::StartFullTextReplaceAsync(std::wstring text,
	FullTextReplaceTask task,
	bool startFormattingAfterReplace,
	bool startTextModeResetAfterReplace)
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return;
	if (m_loadingActive)
		return;

	CancelMarkdownFormatting();
	KillTimer(kInsertTimerId);

	m_loadingActive = true;
	m_loadingText = std::move(text);
	m_loadingTextPos = 0;
	m_fullTextReplaceTask = task;
	m_startFormattingAfterFullTextReplace = startFormattingAfterReplace;
	m_startTextModeResetAfterFullTextReplace = startTextModeResetAfterReplace;

	SetRenderBusy(true, _T("正在切换..."));
	StartInsertLoadedText();
}

void CMarkdownEditorDlg::StartTextModeResetAsync()
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return;

	// 仅在首次启动重置时初始化进度，避免定时器续跑时反复从 0 开始导致长期卡住。
	const bool firstStart = !m_textModeResetActive;
	if (firstStart)
	{
		m_textModeResetActive = true;
		m_textModeResetPos = 0;
	}
	if (!m_renderBusy || m_renderBusyHint != _T("正在切换..."))
		SetRenderBusy(true, _T("正在切换..."));

	// Restore unified font to display original text (chunked to avoid blocking on large docs)
	CHARFORMAT2 format{};
	format.cbSize = sizeof(format);
	format.dwMask = CFM_FACE | CFM_SIZE | CFM_COLOR | CFM_CHARSET |
		CFM_BOLD | CFM_ITALIC | CFM_UNDERLINE | CFM_STRIKEOUT |
		CFM_LINK | CFM_BACKCOLOR | CFM_HIDDEN;
	format.dwEffects = 0;
	format.yHeight = 220;
	format.crTextColor = m_themeForeground;
	format.crBackColor = m_themeBackground;
	format.bCharSet = DEFAULT_CHARSET;
	lstrcpy(format.szFaceName, L"Consolas");

	GETTEXTLENGTHEX textLenEx{};
	textLenEx.flags = GTL_NUMCHARS;
	textLenEx.codepage = 1200;
	const long len = static_cast<long>(m_editContent.SendMessage(EM_GETTEXTLENGTHEX,
		reinterpret_cast<WPARAM>(&textLenEx), 0));

	const DWORD startTick = GetTickCount();
	const DWORD budgetMs = 10;
	const long chunk = 8000;
	while (GetTickCount() - startTick < budgetMs && m_textModeResetPos < len)
	{
		long start = m_textModeResetPos;
		long end = min(len, start + chunk);
		m_editContent.SetSel(start, end);
		m_editContent.SetSelectionCharFormat(format);
		m_textModeResetPos = end;
	}

	if (m_textModeResetPos < len)
	{
		SetTimer(kInsertTimerId, 1, NULL);
		return;
	}

	// Emoji/icon fallback: in text mode we still want emoji icons in headings to render correctly.
	{
		std::wstring source = m_rawDocumentText;
		if (source.empty())
		{
			CString content;
			m_editContent.GetWindowText(content);
			source.assign(content.GetString());
		}
		ApplyEmojiFontFallbackToEdit(source, len);
	}

	m_textModeResetActive = false;
	m_editContent.SetSel(0, 0);
	SetRenderBusy(false, _T(""));
	// Sidebar should remain stable during mode switches.
	m_modeSwitchInProgress = false;
}

void CMarkdownEditorDlg::ApplyEmojiFontFallbackToEdit(const std::wstring& sourceText, long richEditLength)
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return;
	if (sourceText.empty())
		return;
	if (richEditLength <= 0)
		return;

	// Preserve current selection.
	LONG selStart = 0;
	LONG selEnd = 0;
	m_editContent.GetSel(selStart, selEnd);

	// Rebuild raw->RichEdit index mapping for CRLF.
	std::vector<long> extraLfPrefix;
	extraLfPrefix.resize(sourceText.length() + 1);
	extraLfPrefix[0] = 0;
	for (size_t i = 0; i < sourceText.length(); ++i)
	{
		long extra = 0;
		if (sourceText[i] == L'\n' && i > 0 && sourceText[i - 1] == L'\r')
			extra = 1;
		extraLfPrefix[i + 1] = extraLfPrefix[i] + extra;
	}
	const long mappedLength = static_cast<long>(sourceText.length()) - extraLfPrefix[sourceText.length()];
	const bool treatCrlfAsSingle = (richEditLength == mappedLength && mappedLength != static_cast<long>(sourceText.length()));

	auto mapIndex = [&](long index) -> long {
		if (!treatCrlfAsSingle)
			return index;
		if (index <= 0)
			return 0;
		if (index > static_cast<long>(sourceText.length()))
			index = static_cast<long>(sourceText.length());
		return index - extraLfPrefix[static_cast<size_t>(index)];
	};

	auto isSymbolBmp = [](wchar_t ch) -> bool {
		// Common icon/symbol ranges used in markdown docs (BMP).
		// These include some emoji-like symbols (e.g. ✅). We will apply a dedicated font fallback later.
		return (ch >= 0x2600 && ch <= 0x27BF) || // Misc symbols + dingbats
			(ch == 0x00A9) || (ch == 0x00AE) || (ch == 0x3030) || (ch == 0x303D);
	};
	auto isEmojiCodePoint = [](unsigned int cp) -> bool {
		return (cp >= 0x1F000 && cp <= 0x1FAFF) || (cp >= 0x1FC00 && cp <= 0x1FFFF);
	};

	struct Range { long s; long e; };
	std::vector<Range> emojiRanges;   // surrogate-pair emoji -> Segoe UI Emoji
	std::vector<Range> symbolRanges;  // BMP symbols -> Segoe UI Symbol
	emojiRanges.reserve(32);
	symbolRanges.reserve(32);

	for (size_t i = 0; i < sourceText.size();)
	{
		wchar_t ch = sourceText[i];
		long rawStart = -1;
		long rawEnd = -1;
		bool isSymbol = false;
		bool isEmoji = false;
		if (ch >= 0xD800 && ch <= 0xDBFF && i + 1 < sourceText.size())
		{
			wchar_t lo = sourceText[i + 1];
			if (lo >= 0xDC00 && lo <= 0xDFFF)
			{
				unsigned int cp = 0x10000 + (((unsigned int)ch - 0xD800) << 10) + ((unsigned int)lo - 0xDC00);
				if (isEmojiCodePoint(cp))
				{
					rawStart = (long)i;
					rawEnd = (long)(i + 2);
					isEmoji = true;
					// Include VS16 if present.
					if (i + 2 < sourceText.size() && sourceText[i + 2] == 0xFE0F)
						rawEnd = (long)(i + 3);
				}
				i += 2;
				if (rawStart < 0)
					continue;
			}
		}
		else if (isSymbolBmp(ch))
		{
			rawStart = (long)i;
			rawEnd = (long)(i + 1);
			isSymbol = true;
			if (i + 1 < sourceText.size() && sourceText[i + 1] == 0xFE0F)
				rawEnd = (long)(i + 2);
			i += 1;
		}
		else
		{
			++i;
		}

		if (rawStart >= 0 && rawEnd > rawStart)
		{
			auto& ranges = isEmoji ? emojiRanges : symbolRanges;
			// Merge contiguous ranges.
			if (!ranges.empty() && ranges.back().e == rawStart)
				ranges.back().e = rawEnd;
			else
				ranges.push_back({ rawStart, rawEnd });
		}
	}

	auto applyRanges = [&](const std::vector<Range>& ranges, const wchar_t* faceName) {
		if (ranges.empty())
			return;

		CHARFORMAT2 fmt{};
		fmt.cbSize = sizeof(fmt);
		fmt.dwMask = CFM_FACE | CFM_CHARSET;
		fmt.bCharSet = DEFAULT_CHARSET;
		wcscpy_s(fmt.szFaceName, faceName);

		for (const auto& rg : ranges)
		{
			long s = mapIndex(rg.s);
			long e = mapIndex(rg.e);
			if (s < 0) s = 0;
			if (e > richEditLength) e = richEditLength;
			if (s >= e)
				continue;
			m_editContent.SetSel(s, e);
			m_editContent.SetSelectionCharFormat(fmt);
		}
	};

	// Prefer Segoe UI Emoji for both BMP symbols (e.g. ✅) and non-BMP emoji (e.g. 🔧).
	// Note: RichEdit may still render them as monochrome depending on OS/control capabilities.
	applyRanges(symbolRanges, L"Segoe UI Emoji");
	applyRanges(emojiRanges, L"Segoe UI Emoji");

	// Restore selection.
	if (selStart < 0) selStart = 0;
	if (selEnd < 0) selEnd = 0;
	if (selStart > richEditLength) selStart = richEditLength;
	if (selEnd > richEditLength) selEnd = richEditLength;
	m_editContent.SetSel(selStart, selEnd);
}


void CMarkdownEditorDlg::InitSidebar()
{
	if (!::IsWindow(m_tabSidebar.GetSafeHwnd()) || !::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;

	m_tabSidebar.DeleteAllItems();
	TCITEM item{};
	item.mask = TCIF_TEXT;
	item.pszText = const_cast<LPTSTR>(_T("文件"));
	m_tabSidebar.InsertItem(0, &item);
	item.pszText = const_cast<LPTSTR>(_T("大纲"));
	m_tabSidebar.InsertItem(1, &item);

	SetSidebarTab(0);
}

void CMarkdownEditorDlg::SetSidebarTab(int tabIndex)
{
	if (tabIndex < 0)
		tabIndex = 0;
	if (tabIndex > 1)
		tabIndex = 1;

	// Preserve per-tab search queries.
	if (::IsWindow(m_editFileSearch.GetSafeHwnd()))
	{
		CString current;
		m_editFileSearch.GetWindowText(current);
		current.Trim();
		if (m_sidebarActiveTab == 0)
			m_fileSearchQuery = current;
		else if (m_sidebarActiveTab == 1)
			m_outlineSearchQuery = current;
	}

	m_sidebarActiveTab = tabIndex;
	if (::IsWindow(m_tabSidebar.GetSafeHwnd()))
	{
		m_tabSidebar.SetCurSel(tabIndex);
		m_tabSidebar.Invalidate(FALSE);
		m_tabSidebar.UpdateWindow();
	}
	if (::IsWindow(m_editFileSearch.GetSafeHwnd()))
	{
		// Typora-like: the search box stays visible for both "文件" and "大纲".
		m_editFileSearch.ShowWindow(SW_SHOW);
		// Restore query for the new tab.
		m_suppressFileSearchEvent = true;
		CString target = (tabIndex == 0) ? m_fileSearchQuery : m_outlineSearchQuery;
		m_editFileSearch.SetWindowText(target);
		m_editFileSearch.SetSel(target.GetLength(), target.GetLength());
		m_suppressFileSearchEvent = false;
	}
	UpdateSidebar();
	// Ensure the initial selection highlight is correct on the first Outline switch.
	if (tabIndex == 1)
		EnsureOutlineInitialSelection();

	// Keep outline highlight synced with editor scroll position.
	if (tabIndex == 1)
	{
		SetTimer(kOutlineFollowTimerId, 120, NULL);
		UpdateOutlineSelectionByViewTop();
	}
	else
	{
		KillTimer(kOutlineFollowTimerId);
	}

	// Fix occasional incomplete painting when switching tabs quickly.
	// Force a full redraw so both sidebar and editor refresh immediately.
	if (::IsWindow(m_treeSidebar.GetSafeHwnd()))
	{
		m_treeSidebar.Invalidate(FALSE);
		m_treeSidebar.UpdateWindow();
	}
	if (::IsWindow(m_tabSidebar.GetSafeHwnd()))
	{
		m_tabSidebar.Invalidate(FALSE);
		m_tabSidebar.UpdateWindow();
	}
	if (::IsWindow(m_editFileSearch.GetSafeHwnd()))
	{
		m_editFileSearch.Invalidate(FALSE);
		m_editFileSearch.UpdateWindow();
	}
	if (::IsWindow(m_editContent.GetSafeHwnd()))
	{
		m_editContent.Invalidate(FALSE);
		m_editContent.UpdateWindow();
	}
	if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
	{
		m_mermaidOverlay.Invalidate(FALSE);
		m_mermaidOverlay.UpdateWindow();
	}
	Invalidate(FALSE);
	UpdateWindow();
}

void CMarkdownEditorDlg::LoadRecentFiles()
{
	m_recentFiles.clear();

	CString data = AfxGetApp()->GetProfileString(_T("RecentFiles"), _T("List"), _T(""));
	int start = 0;
	while (start >= 0 && start < data.GetLength())
	{
		int end = data.Find(_T('|'), start);
		CString item = (end == -1) ? data.Mid(start) : data.Mid(start, end - start);
		item.Trim();
		if (!item.IsEmpty())
		{
			m_recentFiles.push_back(item);
			if ((int)m_recentFiles.size() >= kMaxRecentFiles)
				break;
		}
		if (end == -1)
			break;
		start = end + 1;
	}
}

void CMarkdownEditorDlg::SaveRecentFiles() const
{
	CString data;
	for (size_t i = 0; i < m_recentFiles.size(); ++i)
	{
		if (i > 0)
			data += _T("|");
		data += m_recentFiles[i];
	}
	AfxGetApp()->WriteProfileString(_T("RecentFiles"), _T("List"), data);
}

void CMarkdownEditorDlg::AddRecentFile(const CString& filePath)
{
	if (filePath.IsEmpty())
		return;

	for (auto it = m_recentFiles.begin(); it != m_recentFiles.end(); ++it)
	{
		if (it->CompareNoCase(filePath) == 0)
		{
			m_recentFiles.erase(it);
			break;
		}
	}

	m_recentFiles.insert(m_recentFiles.begin(), filePath);
	if ((int)m_recentFiles.size() > kMaxRecentFiles)
		m_recentFiles.resize(kMaxRecentFiles);

	SaveRecentFiles();
}

bool CMarkdownEditorDlg::RemoveRecentFile(const CString& filePath)
{
	if (filePath.IsEmpty())
		return false;

	bool removed = false;
	for (auto it = m_recentFiles.begin(); it != m_recentFiles.end();)
	{
		if (it->CompareNoCase(filePath) == 0)
		{
			it = m_recentFiles.erase(it);
			removed = true;
			continue;
		}
		++it;
	}

	if (removed)
		SaveRecentFiles();
	return removed;
}

void CMarkdownEditorDlg::PromptSetDefaultMdAssociationIfNeeded()
{
	CWinApp* app = AfxGetApp();
	if (app == nullptr)
		return;

	// Compatibility: previous build may have set PromptedDefaultMd before user saw the prompt.
	// Use a new key to ensure one real, user-visible prompt can still happen.
	const CString promptFlagKey = _T("PromptedDefaultMdV2");
	if (app->GetProfileInt(_T("Settings"), promptFlagKey, 0) != 0)
		return;

	if (IsCurrentAppMdAssociationRegistered())
	{
		app->WriteProfileInt(_T("Settings"), promptFlagKey, 1);
		return;
	}

	const int answer = AfxMessageBox(
		_T("是否将本程序设为 .md 文件默认打开工具？"),
		MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2);
	if (answer == IDYES)
	{
		CString error;
		if (SetAsDefaultMdFileHandler(&error))
		{
			if (IsCurrentAppMdAssociationRegistered())
				AfxMessageBox(_T("已设置为 .md 文件默认打开工具。"), MB_OK | MB_ICONINFORMATION);
			else
				AfxMessageBox(_T("系统限制未能直接完成设置，已打开系统默认应用设置，请将 .md 关联到本程序。"), MB_OK | MB_ICONINFORMATION);
		}
		else
		{
			CString msg = _T("设置默认打开工具失败。");
			if (!error.IsEmpty())
			{
				msg += _T("\r\n");
				msg += error;
			}
			AfxMessageBox(msg, MB_OK | MB_ICONWARNING);
		}
	}

	app->WriteProfileInt(_T("Settings"), promptFlagKey, 1);
}

bool CMarkdownEditorDlg::SetAsDefaultMdFileHandler(CString* errorMessage)
{
	TCHAR exePath[MAX_PATH] = { 0 };
	const DWORD pathLen = ::GetModuleFileName(nullptr, exePath, _countof(exePath));
	if (pathLen == 0 || pathLen >= _countof(exePath))
	{
		if (errorMessage)
			*errorMessage = _T("无法获取程序路径。");
		return false;
	}

	const CString progId = _T("MyTypora.Markdown");
	const CString appName = _T("MyTypora Markdown Document");
	const CString iconValue = CString(_T("\"")) + exePath + _T("\",0");
	const CString commandValue = CString(_T("\"")) + exePath + _T("\" \"%1\"");

	const bool okExt = WriteRegistryStringValue(HKEY_CURRENT_USER,
		_T("Software\\Classes\\.md"), _T(""), progId);
	const bool okProg = WriteRegistryStringValue(HKEY_CURRENT_USER,
		_T("Software\\Classes\\MyTypora.Markdown"), _T(""), appName);
	const bool okIcon = WriteRegistryStringValue(HKEY_CURRENT_USER,
		_T("Software\\Classes\\MyTypora.Markdown\\DefaultIcon"), _T(""), iconValue);
	const bool okOpenCmd = WriteRegistryStringValue(HKEY_CURRENT_USER,
		_T("Software\\Classes\\MyTypora.Markdown\\shell\\open\\command"), _T(""), commandValue);
	const bool okOpenWith = WriteRegistryEmptyValue(HKEY_CURRENT_USER,
		_T("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.md\\OpenWithProgids"),
		_T("MyTypora.Markdown"));

	if (!(okExt && okProg && okIcon && okOpenCmd))
	{
		if (errorMessage)
			*errorMessage = _T("写入注册表失败，请检查当前用户权限。");
		return false;
	}
	UNREFERENCED_PARAMETER(okOpenWith);

	// Clear per-user override so HKCU\\Software\\Classes\\.md can take effect.
	(void)DeleteRegistryTreeIfExists(HKEY_CURRENT_USER,
		_T("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.md\\UserChoice"));

	::SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);
	if (IsCurrentAppMdAssociationRegistered())
		return true;

	// If Windows still keeps protected UserChoice, guide user to system UI.
	if (OpenSystemDefaultAppsSettings())
		return true;

	if (errorMessage)
		*errorMessage = _T("无法打开系统默认应用设置，请在 Windows 设置中手动将 .md 关联到本程序。");
	return false;
}

void CMarkdownEditorDlg::UpdateSidebar()
{
	if (m_modeSwitchInProgress)
		return;
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;

	CString searchText = (m_sidebarActiveTab == 0) ? m_fileSearchQuery : m_outlineSearchQuery;
	searchText.Trim();

	// Ensure the TreeView is visible (fix: blank area under tabs).
	m_treeSidebar.ShowWindow(SW_SHOW);
	m_treeSidebar.BringWindowToTop();
	// Batch update to avoid flicker and to guarantee a full repaint.
	m_treeSidebar.SetRedraw(FALSE);
	m_treeSidebar.DeleteAllItems();
	m_fileTreePaths.clear();

	if (m_sidebarActiveTab == 0)
	{
		if (!searchText.IsEmpty())
		{
			FilterFileTreeByQuery(searchText);
			m_treeSidebar.SetRedraw(TRUE);
			m_treeSidebar.RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
			return;
		}

		BuildFileTree(m_strCurrentFile);
		SelectCurrentFileInFileTree();
		if (!m_recentFiles.empty())
		{
			// Keep "最近" as a top-level root item; do not nest under folder tree.
			HTREEITEM recentRoot = m_treeSidebar.InsertItem(_T("最近"), TVI_ROOT, TVI_LAST);
			for (size_t i = 0; i < m_recentFiles.size(); ++i)
			{
				CString fileName = m_recentFiles[i];
				int pos = fileName.ReverseFind('\\');
				if (pos != -1)
					fileName = fileName.Mid(pos + 1);
				HTREEITEM item = m_treeSidebar.InsertItem(fileName, recentRoot, TVI_LAST);
				m_treeSidebar.SetItemData(item, MakeTreeData(kTreeDataRecentTag, m_fileTreePaths.size()));
				m_fileTreePaths.push_back(m_recentFiles[i]);
			}
			m_treeSidebar.Expand(recentRoot, TVE_EXPAND);
		}
		else
		{
			m_treeSidebar.InsertItem(_T("(暂无文件)"), TVI_ROOT, TVI_LAST);
		}
	}
	else
	{
		if (!searchText.IsEmpty())
			SearchOutlineByQuery(searchText);
		else
		{
			BuildOutlineTree();
			ExpandAllOutlineItems();
			if (m_outlineAutoSelectFirst)
				EnsureOutlineInitialSelection();
			else
				UpdateOutlineSelection();
		}
	}

	m_treeSidebar.SetRedraw(TRUE);
	// Ensure the tree is actually repainted (fix: items clickable but not visible).
	m_treeSidebar.RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
}

void CMarkdownEditorDlg::SelectCurrentFileInFileTree()
{
	if (m_sidebarActiveTab != 0)
		return;
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;
	if (m_strCurrentFile.IsEmpty())
		return;

	HTREEITEM root = m_treeSidebar.GetRootItem();
	if (!root)
		return;

	CString currentName = m_strCurrentFile;
	int pos = currentName.ReverseFind('\\');
	if (pos >= 0)
		currentName = currentName.Mid(pos + 1);
	currentName.MakeLower();

	for (HTREEITEM item = m_treeSidebar.GetChildItem(root); item != nullptr; item = m_treeSidebar.GetNextSiblingItem(item))
	{
		CString name = m_treeSidebar.GetItemText(item);
		CString lower = name;
		lower.MakeLower();
		if (lower == currentName)
		{
			m_sidebarSelectionUpdating = true;
			m_treeSidebar.SelectItem(item);
			m_treeSidebar.EnsureVisible(item);
			m_sidebarSelectionUpdating = false;
			m_treeSidebar.Invalidate(FALSE);
			break;
		}
	}

	// Do not force focus here; it can cause repaint/focus jitter when switching tabs.
}

void CMarkdownEditorDlg::UpdateOutline()
{
	if (m_sidebarActiveTab != 1)
		return;
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;
	if (m_loadingActive)
		return;

	CString q = m_outlineSearchQuery;
	q.Trim();
	if (!q.IsEmpty())
	{
		SearchOutlineByQuery(q);
		return;
	}

	m_treeSidebar.SetRedraw(FALSE);
	m_treeSidebar.DeleteAllItems();
	BuildOutlineTree();
	ExpandAllOutlineItems();
	if (m_outlineAutoSelectFirst)
		EnsureOutlineInitialSelection();
	else
		UpdateOutlineSelection();
	m_treeSidebar.SetRedraw(TRUE);
	m_treeSidebar.Invalidate();
}

void CMarkdownEditorDlg::SearchOutlineByQuery(const CString& query)
{
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;
	if (m_sidebarActiveTab != 1)
		return;

	CString q = query;
	q.Trim();
	if (q.IsEmpty())
	{
		UpdateOutline();
		return;
	}

	CString qLower = q;
	qLower.MakeLower();

	m_treeSidebar.SetRedraw(FALSE);
	m_treeSidebar.DeleteAllItems();
	m_outlineLineIndex.clear();
	m_outlineSourceIndex.clear();
	m_outlineItemSourceIndex.clear();

	if (m_outlineEntries.empty())
	{
		m_treeSidebar.InsertItem(m_parsingActive ? _T("(正在生成大纲...)") : _T("(暂无大纲)"), TVI_ROOT, TVI_LAST);
		m_treeSidebar.SetRedraw(TRUE);
		m_treeSidebar.Invalidate();
		return;
	}

	std::vector<bool> include(m_outlineEntries.size(), false);
	std::vector<size_t> stack;
	stack.reserve(8);
	for (size_t i = 0; i < m_outlineEntries.size(); ++i)
	{
		int lvl = m_outlineEntries[i].level;
		if (lvl < 1) lvl = 1;
		if (lvl > 6) lvl = 6;
		while (!stack.empty())
		{
			int prevLvl = m_outlineEntries[stack.back()].level;
			if (prevLvl < 1) prevLvl = 1;
			if (prevLvl > 6) prevLvl = 6;
			if (lvl > prevLvl)
				break;
			stack.pop_back();
		}

		CString title(m_outlineEntries[i].title.c_str());
		CString titleLower = title;
		titleLower.MakeLower();
		if (titleLower.Find(qLower) >= 0)
		{
			include[i] = true;
			for (size_t ancestor : stack)
				include[ancestor] = true;
		}
		stack.push_back(i);
	}

	HTREEITEM parents[7] = { nullptr };
	HTREEITEM firstMatch = nullptr;
	for (size_t i = 0; i < m_outlineEntries.size(); ++i)
	{
		if (!include[i])
			continue;
		const auto& entry = m_outlineEntries[i];
		int level = entry.level;
		if (level < 1) level = 1;
		if (level > 6) level = 6;

		CString title(entry.title.c_str());
		HTREEITEM parent = (level == 1) ? TVI_ROOT : (parents[level - 1] ? parents[level - 1] : TVI_ROOT);
		HTREEITEM item = m_treeSidebar.InsertItem(title, parent, TVI_LAST);
		m_treeSidebar.SetItemData(item, static_cast<DWORD_PTR>(entry.lineNo));
		m_outlineLineIndex.emplace_back(entry.lineNo, item);
		m_outlineSourceIndex.emplace_back(entry.sourceIndex, item);
		m_outlineItemSourceIndex[item] = entry.sourceIndex;
		parents[level] = item;
		for (int l = level + 1; l <= 6; ++l)
			parents[l] = nullptr;

		if (!firstMatch)
		{
			CString titleLower = title;
			titleLower.MakeLower();
			if (titleLower.Find(qLower) >= 0)
				firstMatch = item;
		}
	}

	std::sort(m_outlineLineIndex.begin(), m_outlineLineIndex.end(),
		[](const auto& lhs, const auto& rhs) { return lhs.first < rhs.first; });
	std::sort(m_outlineSourceIndex.begin(), m_outlineSourceIndex.end(),
		[](const auto& lhs, const auto& rhs) { return lhs.first < rhs.first; });

	if (m_treeSidebar.GetRootItem() == nullptr)
	{
		m_treeSidebar.InsertItem(_T("(未找到匹配标题)"), TVI_ROOT, TVI_LAST);
	}
	else
	{
		ExpandAllOutlineItems();
		if (firstMatch)
		{
			// Highlight the first match but don't auto-jump while the user is typing.
			m_sidebarSelectionUpdating = true;
			m_treeSidebar.SelectItem(firstMatch);
			m_treeSidebar.EnsureVisible(firstMatch);
			m_sidebarSelectionUpdating = false;
		}
	}

	m_treeSidebar.SetRedraw(TRUE);
	m_treeSidebar.Invalidate();
}

void CMarkdownEditorDlg::BuildFileTree(const CString& filePath)
{
	if (filePath.IsEmpty())
		return;

	CString folder = filePath;
	int pos = folder.ReverseFind('\\');
	if (pos == -1)
		return;
	folder = folder.Left(pos);

	HTREEITEM root = m_treeSidebar.InsertItem(folder, TVI_ROOT, TVI_LAST);
	CFileFind finder;
	CString pattern = folder + _T("\\*.*");
	BOOL working = finder.FindFile(pattern);
	int idx = 0;
	while (working)
	{
		working = finder.FindNextFile();
		if (finder.IsDots() || finder.IsDirectory())
			continue;
		CString path = finder.GetFilePath();
		if (!IsMarkdownFile(path))
			continue;
		CString name = finder.GetFileName();
		HTREEITEM item = m_treeSidebar.InsertItem(name, root, TVI_LAST);
		m_treeSidebar.SetItemData(item, MakeTreeData(kTreeDataFolderTag, m_fileTreePaths.size()));
		m_fileTreePaths.push_back(path);
		++idx;
		if (idx >= 200)
			break;
	}
	finder.Close();
	if (idx > 0)
		m_treeSidebar.Expand(root, TVE_EXPAND);
}

void CMarkdownEditorDlg::BuildOutlineTree()
{
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;

	m_outlineLineIndex.clear();
	m_outlineSourceIndex.clear();
	m_outlineItemSourceIndex.clear();

	if (m_outlineEntries.empty())
	{
		m_treeSidebar.InsertItem(m_parsingActive ? _T("(正在生成大纲...)") : _T("(暂无大纲)"), TVI_ROOT, TVI_LAST);
		return;
	}

	HTREEITEM parents[7] = { nullptr };
	for (const auto& entry : m_outlineEntries)
	{
		int level = entry.level;
		if (level < 1)
			level = 1;
		if (level > 6)
			level = 6;

		CString title(entry.title.c_str());
		HTREEITEM parent = (level == 1) ? TVI_ROOT : (parents[level - 1] ? parents[level - 1] : TVI_ROOT);
		HTREEITEM item = m_treeSidebar.InsertItem(title, parent, TVI_LAST);
		m_treeSidebar.SetItemData(item, static_cast<DWORD_PTR>(entry.lineNo));
		m_outlineLineIndex.emplace_back(entry.lineNo, item);
		m_outlineSourceIndex.emplace_back(entry.sourceIndex, item);
		m_outlineItemSourceIndex[item] = entry.sourceIndex;
		parents[level] = item;
		for (int l = level + 1; l <= 6; ++l)
			parents[l] = nullptr;
	}

	std::sort(m_outlineLineIndex.begin(), m_outlineLineIndex.end(),
		[](const auto& lhs, const auto& rhs) { return lhs.first < rhs.first; });
	std::sort(m_outlineSourceIndex.begin(), m_outlineSourceIndex.end(),
		[](const auto& lhs, const auto& rhs) { return lhs.first < rhs.first; });
}

void CMarkdownEditorDlg::UpdateOutlineSelection()
{
	// Default behavior: follow the caret/selection (for click/selection sync).
	UpdateOutlineSelectionByCaret();
}

void CMarkdownEditorDlg::UpdateOutlineSelectionByViewTop()
{
	if (m_sidebarActiveTab != 1)
		return;
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;
	if (m_sidebarSelectionUpdating)
		return;
	if (m_outlineAutoSelectFirst || m_loadingActive || m_parsingActive || m_formattingActive || m_renderBusy)
		return;
	if (m_outlineLineIndex.empty())
		return;

	int viewTopLine = static_cast<int>(m_editContent.SendMessage(EM_GETFIRSTVISIBLELINE, 0, 0));
	if (viewTopLine < 0)
		return;

	HTREEITEM bestItem = nullptr;
	int viewTopChar = m_editContent.LineIndex(viewTopLine);
	if (viewTopChar >= 0 && !m_outlineSourceIndex.empty())
	{
		int viewTopSource = MapRichIndexToSourceIndex(viewTopChar);
		for (size_t i = 0; i < m_outlineSourceIndex.size(); ++i)
		{
			int sourceIndex = m_outlineSourceIndex[i].first;
			if (sourceIndex > viewTopSource)
				break;
			bestItem = m_outlineSourceIndex[i].second;
		}
	}

	if (bestItem == nullptr)
	{
		for (size_t i = 0; i < m_outlineLineIndex.size(); ++i)
		{
			int line = m_outlineLineIndex[i].first;
			if (line > viewTopLine)
				break;
			bestItem = m_outlineLineIndex[i].second;
		}
	}

	if (bestItem != nullptr)
	{
		HTREEITEM current = m_treeSidebar.GetSelectedItem();
		if (current != bestItem)
		{
			m_sidebarSelectionUpdating = true;
			m_treeSidebar.SelectItem(bestItem);
			m_sidebarSelectionUpdating = false;
			m_treeSidebar.Invalidate(FALSE);
		}
	}
}

void CMarkdownEditorDlg::OnSidebarTreeCustomDraw(NMHDR* pNMHDR, LRESULT* pResult)
{
	if (!pNMHDR || !pResult)
		return;
	LPNMTVCUSTOMDRAW draw = reinterpret_cast<LPNMTVCUSTOMDRAW>(pNMHDR);
	*pResult = CDRF_DODEFAULT;
	if (!draw)
		return;

	if (draw->nmcd.dwDrawStage == CDDS_PREPAINT)
	{
		*pResult = CDRF_NOTIFYITEMDRAW;
		return;
	}
	if (draw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT))
	{
		// NOTE: When the TreeView does not have focus, CDIS_SELECTED can be unreliable.
		// Query the real TreeView item state so auto-follow selection and mouse selection
		// share the exact same highlight colors.
		HTREEITEM item = reinterpret_cast<HTREEITEM>(draw->nmcd.dwItemSpec);
		DWORD state = static_cast<DWORD>(::SendMessage(m_treeSidebar.GetSafeHwnd(), TVM_GETITEMSTATE,
			reinterpret_cast<WPARAM>(item), TVIS_SELECTED | TVIS_DROPHILITED));
		const bool selected = (state & (TVIS_SELECTED | TVIS_DROPHILITED)) != 0;

		COLORREF textColor = m_themeForeground;
		COLORREF bkColor = m_themeSidebarBackground;
		if (selected)
		{
			// Use one consistent highlight color for both mouse selection and scroll-follow selection.
			bkColor = m_themeDark ? RGB(60, 90, 150) : RGB(210, 228, 255);
			textColor = m_themeDark ? RGB(255, 255, 255) : RGB(20, 20, 20);
		}
		draw->clrText = textColor;
		draw->clrTextBk = bkColor;
		*pResult = CDRF_DODEFAULT;
		return;
	}
}

void CMarkdownEditorDlg::UpdateOutlineSelectionByCaret()
{
	if (m_sidebarActiveTab != 1)
		return;
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;
	if (m_sidebarSelectionUpdating)
		return;
	if (m_outlineAutoSelectFirst || m_loadingActive || m_parsingActive || m_formattingActive || m_renderBusy)
		return;
	if (m_outlineLineIndex.empty())
		return;
	// Avoid sidebar jitter while the user is dragging a selection.
	if (m_mouseSelecting)
		return;

	LONG selStart = 0;
	LONG selEnd = 0;
	m_editContent.GetSel(selStart, selEnd);
	int caretLine = m_editContent.LineFromChar(selStart);
	if (caretLine < 0)
		return;

	HTREEITEM bestItem = nullptr;
	if (!m_outlineSourceIndex.empty())
	{
		int caretSource = MapRichIndexToSourceIndex(static_cast<int>(selStart));
		for (size_t i = 0; i < m_outlineSourceIndex.size(); ++i)
		{
			int sourceIndex = m_outlineSourceIndex[i].first;
			if (sourceIndex > caretSource)
				break;
			bestItem = m_outlineSourceIndex[i].second;
		}
	}

	if (bestItem == nullptr)
	{
		for (size_t i = 0; i < m_outlineLineIndex.size(); ++i)
		{
			int line = m_outlineLineIndex[i].first;
			if (line > caretLine)
				break;
			bestItem = m_outlineLineIndex[i].second;
		}
	}

	if (bestItem != nullptr)
	{
		HTREEITEM current = m_treeSidebar.GetSelectedItem();
		if (current != bestItem)
		{
			m_sidebarSelectionUpdating = true;
			m_treeSidebar.SelectItem(bestItem);
			m_sidebarSelectionUpdating = false;
			m_treeSidebar.Invalidate(FALSE);
		}
	}
}

void CMarkdownEditorDlg::OnSidebarTabChanged(NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(pNMHDR);
	if (pResult)
		*pResult = 0;

	int tabIndex = m_tabSidebar.GetCurSel();
	SetSidebarTab(tabIndex);
}

bool CMarkdownEditorDlg::JumpToOutlineTreeItem(HTREEITEM item)
{
	if (item == nullptr)
		return false;

	int charIndex = -1;
	auto it = m_outlineItemSourceIndex.find(item);
	if (it != m_outlineItemSourceIndex.end() && it->second >= 0)
		charIndex = MapSourceIndexToRichIndex(it->second);

	if (charIndex < 0)
	{
		DWORD_PTR lineNo = m_treeSidebar.GetItemData(item);
		const int targetLineFallback = static_cast<int>(lineNo);
		if (targetLineFallback < 0)
			return false;
		charIndex = m_editContent.LineIndex(targetLineFallback);
	}

	if (charIndex < 0)
		return false;

	m_editContent.SetSel(charIndex, charIndex);
	const int targetLine = m_editContent.LineFromChar(charIndex);
	if (targetLine >= 0)
	{
		// Scroll so the heading line becomes the first visible line (Typora-like).
		const int firstVisible = static_cast<int>(m_editContent.SendMessage(EM_GETFIRSTVISIBLELINE, 0, 0));
		const int delta = targetLine - firstVisible;
		if (delta != 0)
			m_editContent.SendMessage(EM_LINESCROLL, 0, delta);
	}
	m_editContent.SendMessage(EM_SCROLLCARET, 0, 0);
	m_editContent.SetFocus();

	// Jumping via outline changes the first visible line; the RichEdit may not have finished
	// scrolling/layout when the selection changes. Schedule a short sync so overlays always
	// re-align after the jump (mermaid/HR/table).
	if (m_bMarkdownMode)
	{
		m_overlaySyncAttempts = 3;
		SetTimer(kOverlaySyncTimerId, 30, NULL);
	}

	return true;
}

void CMarkdownEditorDlg::OnSidebarItemChanged(NMHDR* pNMHDR, LRESULT* pResult)
{
	if (pResult)
		*pResult = 0;
	const NMTREEVIEW* pNmtv = reinterpret_cast<const NMTREEVIEW*>(pNMHDR);
	if (!pNmtv)
		return;
	if (m_sidebarSelectionUpdating)
		return;
	if (pNmtv->itemNew.hItem == nullptr)
		return;

	if (m_sidebarActiveTab == 0)
	{
		DWORD_PTR value = m_treeSidebar.GetItemData(pNmtv->itemNew.hItem);
		DWORD_PTR tag = 0;
		size_t index = 0;
		if (DecodeTreeData(value, tag, index))
		{
			if (index < m_fileTreePaths.size())
			{
				const CString path = m_fileTreePaths[index];
				if (tag == kTreeDataRecentTag && !FileExistsOnDisk(path))
				{
					CString msg;
					msg.Format(_T("文件已不存在，已从最近记录中删除：\r\n%s"), (LPCTSTR)path);
					AfxMessageBox(msg, MB_OK | MB_ICONWARNING);
					RemoveRecentFile(path);

					// Refresh sidebar safely (avoid re-entrant selection notifications).
					m_sidebarSelectionUpdating = true;
					UpdateSidebar();
					m_sidebarSelectionUpdating = false;
					return;
				}

				OpenDocument(path);
			}
		}
		return;
	}

	if (m_sidebarActiveTab != 1)
		return;

	JumpToOutlineTreeItem(pNmtv->itemNew.hItem);
}

void CMarkdownEditorDlg::OnSidebarTreeClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(pNMHDR);
	if (pResult)
		*pResult = 0;
	if (m_sidebarActiveTab != 1)
		return;
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;

	// When the user clicks an already-selected outline item, TVN_SELCHANGED does not fire.
	// Trigger the same jump behavior on repeated clicks.
	HTREEITEM item = m_treeSidebar.GetSelectedItem();
	if (!item)
		return;

	DWORD tick = GetTickCount();
	if (item == m_lastSidebarTreeSelItem)
	{
		// allow immediate re-jump even without timing constraints
		JumpToOutlineTreeItem(item);
	}
	else
	{
		// Track selection so a later same-item click can be detected.
		m_lastSidebarTreeSelItem = item;
		m_lastSidebarTreeSelTick = tick;
	}
}

void CMarkdownEditorDlg::OnSidebarTreeRClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(pNMHDR);
	if (pResult)
		*pResult = 0;
	if (m_sidebarActiveTab != 0)
		return;
	if (!::IsWindow(m_treeSidebar.GetSafeHwnd()))
		return;

	CPoint screenPt;
	::GetCursorPos(&screenPt);

	CPoint clientPt = screenPt;
	m_treeSidebar.ScreenToClient(&clientPt);

	TVHITTESTINFO hit{};
	hit.pt = clientPt;
	HTREEITEM item = m_treeSidebar.HitTest(&hit);
	if (item == nullptr)
		return;
	if ((hit.flags & (TVHT_ONITEMLABEL | TVHT_ONITEMICON | TVHT_ONITEMRIGHT | TVHT_ONITEMSTATEICON)) == 0)
		return;

	// Right click should not open the file; select for visual feedback only.
	m_sidebarSelectionUpdating = true;
	m_treeSidebar.SelectItem(item);
	m_sidebarSelectionUpdating = false;

	DWORD_PTR value = m_treeSidebar.GetItemData(item);
	DWORD_PTR tag = 0;
	size_t index = 0;
	if (!DecodeTreeData(value, tag, index))
		return;
	if (index >= m_fileTreePaths.size())
		return;

	const CString path = m_fileTreePaths[index];
	if (path.IsEmpty())
		return;

	CMenu menu;
	if (!menu.CreatePopupMenu())
		return;
	constexpr UINT kCmdCopyName = 1;
	constexpr UINT kCmdCopyPath = 2;
	menu.AppendMenu(MF_STRING, kCmdCopyName, _T("复制文件名"));
	menu.AppendMenu(MF_STRING, kCmdCopyPath, _T("复制完整路径"));

	UINT cmd = menu.TrackPopupMenu(TPM_RETURNCMD | TPM_NONOTIFY | TPM_RIGHTBUTTON,
		screenPt.x, screenPt.y, this);
	if (cmd == 0)
		return;

	if (cmd == kCmdCopyName)
	{
		const CString name = ExtractFileNameFromPath(path);
		CopyTextToClipboard(GetSafeHwnd(), name);
	}
	else if (cmd == kCmdCopyPath)
	{
		CopyTextToClipboard(GetSafeHwnd(), path);
	}
}

void CMarkdownEditorDlg::OnEditContentVScroll()
{
	if (m_bMarkdownMode)
	{
		if (!m_windowSizingMove && ::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		{
			UpdateMermaidOverlayRegion(false);
			m_mermaidOverlay.Invalidate(FALSE);
		}
	}
	UpdateOutlineSelectionByViewTop();
}

void CMarkdownEditorDlg::OnEditContentHScroll()
{
	if (!m_bMarkdownMode || m_windowSizingMove)
		return;
	if (!::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		return;
	UpdateMermaidOverlayRegion(false);
	m_mermaidOverlay.Invalidate(FALSE);
}

BOOL CMarkdownEditorDlg::IsMarkdownFile(const CString& filePath)
{
	if (filePath.IsEmpty())
		return FALSE;

	CString ext = ::PathFindExtension(filePath);
	ext.MakeLower();
	return (ext == _T(".md") || ext == _T(".markdown") || ext == _T(".mdown") ||
		ext == _T(".mkd") || ext == _T(".mkdn") || ext == _T(".mdwn"));
}

void CMarkdownEditorDlg::HighlightMarkdownSyntax()
{
	// Basic Markdown syntax highlighting implementation
	// In actual applications, Rich Edit control can be used for more complex syntax highlighting

	CString content;
	m_editContent.GetWindowText(content);

	// Detect Markdown syntax and display info in status bar
	int headerCount = 0;
	int boldCount = 0;
	int linkCount = 0;
	int codeBlockCount = 0;

	// Count Markdown elements
	CString line;
	int pos = 0;
	while (pos < content.GetLength())
	{
		int nextPos = content.Find(_T('\n'), pos);
		if (nextPos == -1) nextPos = content.GetLength();

		line = content.Mid(pos, nextPos - pos);
		line.Trim();

		// Detect headers (#, ##, ### etc.)
		if (line.GetLength() > 0 && line[0] == _T('#'))
		{
			headerCount++;
		}

		// Detect bold (**text**)
		if (line.Find(_T("**")) != -1)
		{
			boldCount++;
		}

		// Detect links [text](url)
		if (line.Find(_T("[")) != -1 && line.Find(_T("](")) != -1)
		{
			linkCount++;
		}

		// Detect code blocks ```
		if (line.Find(_T("```")) != -1)
		{
			codeBlockCount++;
		}

		pos = nextPos + 1;
	}

	// Update status info
	CString syntaxInfo;
	syntaxInfo.Format(_T(" | 标题:%d 粗体:%d 链接:%d 代码块:%d"),
		headerCount, boldCount, linkCount, codeBlockCount / 2); // 代码块成对出现

	CString currentStatus;
	m_staticStatus.GetWindowText(currentStatus);

	// 移除之前的语法信息
	int pipePos = currentStatus.Find(_T(" | 标题:"));
	if (pipePos != -1)
	{
		currentStatus = currentStatus.Left(pipePos);
	}

	currentStatus += syntaxInfo;
	m_staticStatus.SetWindowText(currentStatus);
}

void CMarkdownEditorDlg::ApplyModernStyling()
{
    // Styles are set uniformly in ApplyMarkdownFormatting, placeholder kept here
}

void CMarkdownEditorDlg::ApplyMarkdownFormatting()
{
	StartMarkdownParsingAsync();
}

void CMarkdownEditorDlg::CancelMarkdownFormatting()
{
	KillTimer(kDebounceTimerId);
	KillTimer(kFormatTimerId);
	KillTimer(kInsertTimerId);
	KillTimer(kStatusPulseTimerId);
	// 取消时提升代次，确保旧解析结果不会在切换后回灌 UI 状态。
	++m_parseGeneration;
	m_formattingActive = false;
	m_formattingReschedule = false;
	m_parsingActive = false;
	m_parseReschedule = false;
	m_parseProgressPercent = 0;
	m_richToSourceIndexMap.clear();
	m_loadingActive = false;
	m_hideDocumentUntilRenderReady = false;
	m_textModeResetActive = false;
	SetRenderBusy(false, _T(""));
}

void CMarkdownEditorDlg::ScheduleMarkdownFormatting(UINT delayMs)
{
	KillTimer(kDebounceTimerId);
	SetTimer(kDebounceTimerId, delayMs, NULL);
	SetRenderBusy(true, _T("正在加载中..."));
}

void CMarkdownEditorDlg::PrepareMarkdownFormats()
{
	if (m_normalFormat.cbSize == sizeof(CHARFORMAT2))
		return;

	const bool isDark = m_themeDark;
	const COLORREF textColor = m_themeForeground;
	const COLORREF backgroundColor = m_themeBackground;
	const COLORREF subtleTextColor = isDark ? RGB(170, 170, 170) : RGB(102, 102, 102);

	m_normalFormat.cbSize = sizeof(CHARFORMAT2);
	m_normalFormat.dwMask = CFM_FACE | CFM_SIZE | CFM_COLOR | CFM_BOLD | CFM_ITALIC |
		CFM_UNDERLINE | CFM_BACKCOLOR | CFM_CHARSET | CFM_STRIKEOUT | CFM_LINK;
	m_normalFormat.dwEffects = 0;
	m_normalFormat.yHeight = 220;
	m_normalFormat.crTextColor = textColor;
	m_normalFormat.crBackColor = backgroundColor;
	m_normalFormat.bCharSet = DEFAULT_CHARSET;
	lstrcpy(m_normalFormat.szFaceName, L"Microsoft YaHei");

	m_hiddenFormat.cbSize = sizeof(CHARFORMAT2);
	m_hiddenFormat.dwMask = CFM_HIDDEN | CFM_COLOR | CFM_SIZE | CFM_BACKCOLOR;
	m_hiddenFormat.dwEffects = CFE_HIDDEN;
	m_hiddenFormat.crTextColor = backgroundColor;
	m_hiddenFormat.crBackColor = backgroundColor;
	m_hiddenFormat.yHeight = 20;

	m_blockMarkerElisionFormat = CHARFORMAT2{};
	m_blockMarkerElisionFormat.cbSize = sizeof(CHARFORMAT2);
	m_blockMarkerElisionFormat.dwMask = CFM_COLOR | CFM_BACKCOLOR | CFM_SIZE;
	m_blockMarkerElisionFormat.dwEffects = 0;
	m_blockMarkerElisionFormat.crTextColor = backgroundColor;
	m_blockMarkerElisionFormat.crBackColor = backgroundColor;
	// 1pt surrogate for block markers to avoid CFE_HIDDEN wrap instability while staying visually hidden.
	m_blockMarkerElisionFormat.yHeight = 20;

	auto initHeading = [&](CHARFORMAT2& target, LONG height, COLORREF color) {
		target = CHARFORMAT2{};
		target.cbSize = sizeof(CHARFORMAT2);
		target.dwMask = CFM_BOLD | CFM_SIZE | CFM_COLOR;
		target.dwEffects = CFE_BOLD;
		target.yHeight = height;
		target.crTextColor = color;
	};

	if (isDark)
	{
		initHeading(m_h1Format, 480, RGB(245, 245, 245));
		initHeading(m_h2Format, 400, RGB(225, 225, 225));
		initHeading(m_h3Format, 360, RGB(210, 210, 210));
		initHeading(m_h4Format, 320, RGB(198, 198, 198));
		initHeading(m_h5Format, 280, RGB(186, 186, 186));
		initHeading(m_h6Format, 240, RGB(176, 176, 176));
	}
	else
	{
		initHeading(m_h1Format, 480, RGB(13, 27, 62));
		initHeading(m_h2Format, 400, RGB(36, 60, 126));
		initHeading(m_h3Format, 360, RGB(57, 90, 159));
		initHeading(m_h4Format, 320, RGB(84, 114, 174));
		initHeading(m_h5Format, 280, RGB(111, 138, 189));
		initHeading(m_h6Format, 240, RGB(138, 162, 204));
	}

	m_boldFormat = CHARFORMAT2{};
	m_boldFormat.cbSize = sizeof(CHARFORMAT2);
	m_boldFormat.dwMask = CFM_BOLD;
	m_boldFormat.dwEffects = CFE_BOLD;

	m_italicFormat = CHARFORMAT2{};
	m_italicFormat.cbSize = sizeof(CHARFORMAT2);
	m_italicFormat.dwMask = CFM_ITALIC;
	m_italicFormat.dwEffects = CFE_ITALIC;

	m_strikeFormat = CHARFORMAT2{};
	m_strikeFormat.cbSize = sizeof(CHARFORMAT2);
	m_strikeFormat.dwMask = CFM_STRIKEOUT | CFM_COLOR;
	m_strikeFormat.dwEffects = CFE_STRIKEOUT;
	m_strikeFormat.crTextColor = subtleTextColor;

	m_codeBlockFormat = m_normalFormat;
	m_codeBlockFormat.dwMask |= CFM_BACKCOLOR | CFM_FACE | CFM_COLOR |
		CFM_BOLD | CFM_ITALIC | CFM_UNDERLINE | CFM_STRIKEOUT | CFM_LINK;
	m_codeBlockFormat.dwEffects = 0;
	m_codeBlockFormat.crBackColor = isDark ? RGB(45, 45, 45) : RGB(246, 246, 246);
	m_codeBlockFormat.crTextColor = isDark ? RGB(230, 230, 230) : RGB(212, 73, 80);
	wcscpy_s(m_codeBlockFormat.szFaceName, L"Consolas");

	m_inlineCodeFormat = m_codeBlockFormat;
	m_inlineCodeFormat.yHeight = 200;

	m_linkFormat = CHARFORMAT2{};
	m_linkFormat.cbSize = sizeof(CHARFORMAT2);
	m_linkFormat.dwMask = CFM_UNDERLINE | CFM_COLOR | CFM_LINK;
	m_linkFormat.dwEffects = CFE_UNDERLINE | CFE_LINK;
	m_linkFormat.crTextColor = isDark ? RGB(120, 180, 255) : RGB(30, 100, 200);

	m_mermaidPlaceholderFormat = m_normalFormat;
	// Mermaid placeholder: keep paragraph structure but make the (invisible) code take minimal height.
	// The real diagram height is reserved via PFM_SPACEAFTER in UpdateMermaidSpacing().
	m_mermaidPlaceholderFormat.dwMask = CFM_COLOR | CFM_BACKCOLOR | CFM_SIZE;
	m_mermaidPlaceholderFormat.dwEffects = 0;
	// Mermaid code block source should be invisible in Markdown mode,
	// but must keep its paragraph space so later content flows below the diagram.
	m_mermaidPlaceholderFormat.crTextColor = backgroundColor;
	m_mermaidPlaceholderFormat.crBackColor = backgroundColor;
	// 1pt (20 twips) to reduce excessive blank space for multi-line mermaid blocks.
	m_mermaidPlaceholderFormat.yHeight = 20;

	m_quoteFormat = CHARFORMAT2{};
	m_quoteFormat.cbSize = sizeof(CHARFORMAT2);
	m_quoteFormat.dwMask = CFM_COLOR;
	m_quoteFormat.dwEffects = 0;
	m_quoteFormat.crTextColor = isDark ? RGB(185, 185, 185) : RGB(118, 118, 118);

	m_listFormat = CHARFORMAT2{};
	m_listFormat.cbSize = sizeof(CHARFORMAT2);
	// Do not force list text color; otherwise quoted list items lose quote gray style.
	m_listFormat.dwMask = 0;
	m_listFormat.dwEffects = 0;
	m_listFormat.crTextColor = textColor;

	// Word-like table appearance (Phase 1): zebra rows + hidden pipes + hidden separator line.
	m_tableHeaderFormat = CHARFORMAT2{};
	m_tableHeaderFormat.cbSize = sizeof(CHARFORMAT2);
	m_tableHeaderFormat.dwMask = CFM_BOLD | CFM_COLOR | CFM_BACKCOLOR;
	m_tableHeaderFormat.dwEffects = CFE_BOLD;
	m_tableHeaderFormat.crTextColor = RGB(255, 255, 255);
	m_tableHeaderFormat.crBackColor = isDark ? RGB(55, 85, 145) : RGB(68, 114, 196);

	m_tableRowEvenFormat = CHARFORMAT2{};
	m_tableRowEvenFormat.cbSize = sizeof(CHARFORMAT2);
	m_tableRowEvenFormat.dwMask = CFM_COLOR | CFM_BACKCOLOR;
	m_tableRowEvenFormat.dwEffects = 0;
	m_tableRowEvenFormat.crTextColor = textColor;
	m_tableRowEvenFormat.crBackColor = backgroundColor;

	m_tableRowOddFormat = CHARFORMAT2{};
	m_tableRowOddFormat.cbSize = sizeof(CHARFORMAT2);
	m_tableRowOddFormat.dwMask = CFM_COLOR | CFM_BACKCOLOR;
	m_tableRowOddFormat.dwEffects = 0;
	m_tableRowOddFormat.crTextColor = textColor;
	m_tableRowOddFormat.crBackColor = isDark ? RGB(38, 38, 38) : RGB(242, 245, 250);

	// Horizontal rule: hide source markers and draw a real line via overlay.
	m_hrFormat = CHARFORMAT2{};
	m_hrFormat.cbSize = sizeof(CHARFORMAT2);
	m_hrFormat.dwMask = CFM_COLOR;
	m_hrFormat.dwEffects = 0;
	m_hrFormat.crTextColor = backgroundColor;
}

void CMarkdownEditorDlg::BeginMarkdownFormatting()
{
	if (!m_bMarkdownMode)
		return;
	// Legacy path kept for backward compatibility.
	StartMarkdownParsingAsync();
}

void CMarkdownEditorDlg::StartMarkdownParsingAsync()
{
	if (!m_bMarkdownMode)
		return;
	if (m_loadingActive)
		return;
	if (m_formattingActive)
	{
		m_formattingReschedule = true;
		return;
	}
	if (m_parsingActive)
	{
		m_parseReschedule = true;
		return;
	}

	m_parsingActive = true;
	m_parseReschedule = false;
	m_parseProgressPercent = 0;
	SetRenderBusy(true, _T("正在加载中..."));

	// Always parse from the raw document text (unmodified), to avoid
	// transformations done for display (e.g., table pipes -> tabs) breaking the parser.
	std::wstring text = m_rawDocumentText;
	if (text.empty())
	{
		// Fallback for edge cases (e.g., user typed without loading a file).
		CString content;
		m_editContent.GetWindowText(content);
		text.assign(content.GetString());
	}
	m_textDirty = false;

	GETTEXTLENGTHEX textLenEx{};
	textLenEx.flags = GTL_NUMCHARS;
	textLenEx.codepage = 1200;
	const long richEditLength = static_cast<long>(m_editContent.SendMessage(EM_GETTEXTLENGTHEX,
		reinterpret_cast<WPARAM>(&textLenEx), 0));

	const unsigned long long generation = ++m_parseGeneration;
	HWND hwnd = GetSafeHwnd();
	if (hwnd == nullptr)
	{
		m_parsingActive = false;
		SetRenderBusy(false, _T(""));
		return;
	}
	auto cancelToken = m_asyncCancelToken;
	auto workerCount = m_asyncWorkerCount;
	std::thread([hwnd, generation, richEditLength, text = std::move(text), cancelToken, workerCount]() mutable {
		AsyncWorkerScope workerScope(workerCount);
		if (cancelToken && cancelToken->load(std::memory_order_relaxed))
			return;

		auto payload = std::make_unique<MarkdownParsePayload>();
		payload->generation = generation;
		payload->richEditLength = richEditLength;
		payload->text = std::move(text);
		auto postParseProgress = [&](int percent) {
			if (cancelToken && cancelToken->load(std::memory_order_relaxed))
				return;
			auto progress = std::make_unique<MarkdownParseProgressPayload>();
			progress->generation = generation;
			progress->percent = max(0, min(100, percent));
			SafePostMessageUnique(hwnd, WM_MARKDOWN_PARSE_PROGRESS, 0, std::move(progress));
		};

		try
		{
			SehTranslatorScope sehScope;
			MarkdownParser parser;
			payload->result = parser.Parse(payload->text, postParseProgress);
			if (payload->result.blockSpans.empty() && payload->result.inlineSpans.empty() && payload->result.hiddenSpans.empty())
			{
				// Fallback: if full parsing fails (e.g., md4c rejects input), keep at least block-level rendering.
				payload->result = parser.ParseBlocksOnly(payload->text);
			}

			// Build outline entries from the parsed block spans.
			// This avoids false positives like '# xxx' inside indented code blocks.
			payload->outlineEntries.clear();
			if (!payload->result.blockSpans.empty() && !payload->text.empty())
			{
				std::vector<long> lineStarts;
				lineStarts.reserve(512);
				lineStarts.push_back(0);
				for (size_t i = 0; i < payload->text.size(); ++i)
				{
					if (payload->text[i] == L'\n')
						lineStarts.push_back(static_cast<long>(i + 1));
				}

				auto getLineNoForIndex = [&](long index) -> int {
					if (index <= 0)
						return 0;
					if (index >= static_cast<long>(payload->text.size()))
						index = static_cast<long>(payload->text.size());
					auto it = std::upper_bound(lineStarts.begin(), lineStarts.end(), index);
					if (it == lineStarts.begin())
						return 0;
					return static_cast<int>(it - lineStarts.begin() - 1);
				};

				auto extractTitle = [&](long start, long end) -> std::wstring {
					start = max(0L, start);
					end = min(end, static_cast<long>(payload->text.size()));
					if (end <= start)
						return std::wstring();
					std::wstring title = payload->text.substr(static_cast<size_t>(start), static_cast<size_t>(end - start));
					// Trim whitespace/newlines.
					size_t s = 0;
					while (s < title.size() && (title[s] == L' ' || title[s] == L'\t' || title[s] == L'\r' || title[s] == L'\n'))
						++s;
					size_t e = title.size();
					while (e > s && (title[e - 1] == L' ' || title[e - 1] == L'\t' || title[e - 1] == L'\r' || title[e - 1] == L'\n'))
						--e;
					if (e <= s)
						return std::wstring();
					title = title.substr(s, e - s);
					return title;
				};

				for (const auto& span : payload->result.blockSpans)
				{
					int level = 0;
					switch (span.type)
					{
					case MarkdownBlockType::Heading1: level = 1; break;
					case MarkdownBlockType::Heading2: level = 2; break;
					case MarkdownBlockType::Heading3: level = 3; break;
					case MarkdownBlockType::Heading4: level = 4; break;
					case MarkdownBlockType::Heading5: level = 5; break;
					case MarkdownBlockType::Heading6: level = 6; break;
					default: break;
					}
					if (level <= 0)
						continue;

					std::wstring title = extractTitle(span.start, span.end);
					if (title.empty())
						continue;

					MarkdownOutlineEntry entry;
					entry.level = level;
					entry.lineNo = getLineNoForIndex(span.start);
					entry.sourceIndex = static_cast<int>(span.start);
					entry.title = std::move(title);
					payload->outlineEntries.push_back(std::move(entry));
				}
			}
		}
		catch (const SehException& ex)
		{
			payload->exceptionCode = static_cast<DWORD>(ex.code);
			payload->result = {};
			payload->outlineEntries.clear();
		}
		catch (...)
		{
			payload->exceptionCode = 0xE0000001;
			payload->result = {};
			payload->outlineEntries.clear();
		}

		if (!cancelToken || !cancelToken->load(std::memory_order_relaxed))
			SafePostMessageUnique(hwnd, WM_MARKDOWN_PARSE_COMPLETE, 0, std::move(payload));
	}).detach();
}

LRESULT CMarkdownEditorDlg::OnMarkdownParseComplete(WPARAM /*wParam*/, LPARAM lParam)
{
	try
	{
		SehTranslatorScope sehScope;
		auto payload = reinterpret_cast<MarkdownParsePayload*>(lParam);
		if (!payload)
			return 0;

		std::unique_ptr<MarkdownParsePayload> holder(payload);
		if (m_destroying)
			return 0;

		if (payload->generation != m_parseGeneration.load())
		{
			// Outdated parse result.
			return 0;
		}

		m_parsingActive = false;
		if (payload->exceptionCode != 0)
		{
			m_parseProgressPercent = 0;
			m_richToSourceIndexMap.clear();
			CString msg;
			msg.Format(_T("解析失败（异常码=0x%08X），已跳过渲染。"), payload->exceptionCode);
			m_hideDocumentUntilRenderReady = false;
			SetRenderBusy(false, msg);
			m_formattingActive = false;
			m_formattingReschedule = false;
			m_parseReschedule = false;
			UpdateStatusText();
			return 0;
		}

		m_parseProgressPercent = 100;
		SetRenderBusy(true, _T("正在加载中..."));
		PrepareMarkdownFormats();

		// Adopt parse result for incremental formatting.
		m_pendingSourceText = std::move(payload->text);
		m_pendingParseResult = std::move(payload->result);
		EnsureValidSourceDisplayMapping(m_pendingSourceText,
			m_sourceToDisplayIndexMap,
			m_displayToSourceIndexMap,
			_T("OnMarkdownParseComplete"));
		m_headingRanges.clear();
		m_horizontalRuleStarts.clear();
		m_tableOverlayRows.clear();
		ClearTableListViewOverlays();
		m_tocOverlayBlocks.clear();
		m_tocHitRegions.clear();
		m_outlineEntries = std::move(payload->outlineEntries);
		m_outlineEntriesGeneration = payload->generation;
		if (m_sidebarActiveTab == 1)
			UpdateOutline();
		m_pendingTextLength = payload->richEditLength;
		if (m_pendingTextLength < 0)
			m_pendingTextLength = 0;

		// Convert parser source indices to RichEdit indices using:
		// source -> display mapping, then display -> RichEdit (CRLF collapse) mapping.
		const long sourceLength = static_cast<long>(m_pendingSourceText.length());
		std::vector<std::pair<long, long>> tocSourceRanges;
		{
			auto isTrimSpace = [](wchar_t ch) -> bool {
				return ch == L' ' || ch == L'\t' || ch == L'\r';
			};
			auto equalsTocMarkerNoCase = [&](const std::wstring& text, long start, long end) -> bool {
				if (start < 0) start = 0;
				if (end < start) end = start;
				if (end > static_cast<long>(text.size())) end = static_cast<long>(text.size());
				if (end - start != 5)
					return false;
				return (text[static_cast<size_t>(start)] == L'[' &&
					(towlower(text[static_cast<size_t>(start + 1)]) == L't') &&
					(towlower(text[static_cast<size_t>(start + 2)]) == L'o') &&
					(towlower(text[static_cast<size_t>(start + 3)]) == L'c') &&
					text[static_cast<size_t>(start + 4)] == L']');
			};

			bool inFence = false;
			wchar_t fenceChar = 0;
			long fenceCount = 0;
			long lineStart = 0;
			while (lineStart <= sourceLength)
			{
				long lineEnd = lineStart;
				while (lineEnd < sourceLength && m_pendingSourceText[static_cast<size_t>(lineEnd)] != L'\n')
					++lineEnd;
				long contentEnd = lineEnd;
				if (contentEnd > lineStart && m_pendingSourceText[static_cast<size_t>(contentEnd - 1)] == L'\r')
					--contentEnd;

				long trimmedStart = lineStart;
				int leadingIndentColumns = 0;
				while (trimmedStart < contentEnd)
				{
					const wchar_t ch = m_pendingSourceText[static_cast<size_t>(trimmedStart)];
					if (ch == L' ')
					{
						++leadingIndentColumns;
						++trimmedStart;
						continue;
					}
					if (ch == L'\t')
					{
						const int rem = leadingIndentColumns % 4;
						leadingIndentColumns += (rem == 0) ? 4 : (4 - rem);
						++trimmedStart;
						continue;
					}
					break;
				}
				long trimmedEnd = contentEnd;
				while (trimmedEnd > trimmedStart && isTrimSpace(m_pendingSourceText[static_cast<size_t>(trimmedEnd - 1)]))
					--trimmedEnd;
				const bool isIndentedCodeLine = (trimmedStart < trimmedEnd) && (leadingIndentColumns >= 4);

				bool isFenceLine = false;
				wchar_t thisFenceChar = 0;
				long thisFenceCount = 0;
				if (trimmedStart < trimmedEnd && leadingIndentColumns <= 3)
				{
					const wchar_t ch = m_pendingSourceText[static_cast<size_t>(trimmedStart)];
					if (ch == L'`' || ch == L'~')
					{
						thisFenceChar = ch;
						long p = trimmedStart;
						while (p < trimmedEnd && m_pendingSourceText[static_cast<size_t>(p)] == ch)
						{
							++thisFenceCount;
							++p;
						}
						if (thisFenceCount >= 3)
							isFenceLine = true;
					}
				}

				if (!inFence && !isIndentedCodeLine &&
					equalsTocMarkerNoCase(m_pendingSourceText, trimmedStart, trimmedEnd))
				{
					tocSourceRanges.emplace_back(lineStart, contentEnd);
				}

				if (isFenceLine)
				{
					if (!inFence)
					{
						inFence = true;
						fenceChar = thisFenceChar;
						fenceCount = thisFenceCount;
					}
					else if (thisFenceChar == fenceChar && thisFenceCount >= fenceCount)
					{
						inFence = false;
						fenceChar = 0;
						fenceCount = 0;
					}
				}

				if (lineEnd >= sourceLength)
					break;
				lineStart = lineEnd + 1;
			}
		}
		for (const auto& tocRange : tocSourceRanges)
		{
			if (tocRange.second > tocRange.first)
			{
				MarkdownHiddenSpan hidden{};
				hidden.start = tocRange.first;
				hidden.end = tocRange.second;
				hidden.kind = MarkdownHiddenSpanKind::Block;
				m_pendingParseResult.hiddenSpans.push_back(hidden);
			}
		}
		long mappedDisplayLength = sourceLength;
		const bool hasSourceDisplayMap =
			(m_sourceToDisplayIndexMap.size() == (m_pendingSourceText.length() + 1));
		if (hasSourceDisplayMap && !m_sourceToDisplayIndexMap.empty())
			mappedDisplayLength = max(0L, m_sourceToDisplayIndexMap.back());

		auto clampSourceIndex = [&](long index) -> long {
			if (index <= 0)
				return 0;
			if (index > sourceLength)
				return sourceLength;
			return index;
		};

		auto mapSourceToDisplay = [&](long sourceIndex) -> long {
			sourceIndex = clampSourceIndex(sourceIndex);
			if (!hasSourceDisplayMap)
				return sourceIndex;

			long mapped = m_sourceToDisplayIndexMap[static_cast<size_t>(sourceIndex)];
			if (mapped < 0)
				mapped = 0;
			if (mapped > mappedDisplayLength)
				mapped = mappedDisplayLength;
			return mapped;
		};

		// Build CRLF prefix map on display indices (RichEdit treats CRLF as one char).
		std::vector<long> displayExtraLfPrefix(static_cast<size_t>(mappedDisplayLength) + 1, 0);
		if (mappedDisplayLength > 0)
		{
			std::vector<long> diff(static_cast<size_t>(mappedDisplayLength) + 2, 0);
			for (long i = 1; i < sourceLength; ++i)
			{
				if (m_pendingSourceText[static_cast<size_t>(i)] != L'\n' ||
					m_pendingSourceText[static_cast<size_t>(i - 1)] != L'\r')
				{
					continue;
				}

				const long displayPrev = mapSourceToDisplay(i - 1);
				const long displayCur = mapSourceToDisplay(i);
				if (displayCur == displayPrev + 1 &&
					displayCur >= 0 &&
					displayCur < mappedDisplayLength)
				{
					diff[static_cast<size_t>(displayCur + 1)] += 1;
				}
			}

			long run = 0;
			for (size_t i = 0; i < displayExtraLfPrefix.size(); ++i)
			{
				run += diff[i];
				displayExtraLfPrefix[i] = run;
			}
		}

		const long displayExtraLfCount = displayExtraLfPrefix.empty()
			? 0
			: displayExtraLfPrefix.back();
		const long richEquivalentLength = mappedDisplayLength - displayExtraLfCount;
		const bool treatDisplayCrlfAsSingle =
			(m_pendingTextLength == richEquivalentLength && mappedDisplayLength != m_pendingTextLength);

		auto mapDisplayToRich = [&](long displayIndex) -> long {
			if (displayIndex <= 0)
				return 0;
			if (displayIndex > mappedDisplayLength)
				displayIndex = mappedDisplayLength;

			long mapped = displayIndex;
			if (treatDisplayCrlfAsSingle &&
				!displayExtraLfPrefix.empty() &&
				displayIndex >= 0 &&
				displayIndex < static_cast<long>(displayExtraLfPrefix.size()))
			{
				mapped = displayIndex - displayExtraLfPrefix[static_cast<size_t>(displayIndex)];
			}

			if (mapped < 0)
				mapped = 0;
			if (mapped > m_pendingTextLength)
				mapped = m_pendingTextLength;
			return mapped;
		};

		auto mapIndex = [&](long sourceIndex) -> long {
			return mapDisplayToRich(mapSourceToDisplay(sourceIndex));
		};

		m_tocOverlayBlocks.clear();
		m_tocHitRegions.clear();
		if (!tocSourceRanges.empty())
		{
			std::vector<TocOverlayLine> tocLines;
			tocLines.reserve(256);
			TocOverlayLine titleLine{};
			titleLine.text = L"\x76EE\x5F55"; // 目录
			titleLine.targetSourceIndex = -1;
			tocLines.push_back(std::move(titleLine));
			if (m_outlineEntries.empty())
			{
				TocOverlayLine emptyLine{};
				emptyLine.text = L"(\x65E0\x6807\x9898)"; // (无标题)
				emptyLine.targetSourceIndex = -1;
				tocLines.push_back(std::move(emptyLine));
			}
			else
			{
				const size_t maxItems = 200;
				size_t written = 0;
				for (const auto& entry : m_outlineEntries)
				{
					if (entry.level <= 0 || entry.title.empty())
						continue;
						if (written >= maxItems)
							break;
						const int indentCount = max(0, min(10, entry.level - 1)) * 2;
						TocOverlayLine line{};
						line.text.append(static_cast<size_t>(indentCount), L' ');
						line.text.append(L"\x2022 ");
						for (wchar_t ch : entry.title)
						{
							if (ch == L'\r' || ch == L'\n')
								continue;
							line.text.push_back(ch);
						}
						line.targetSourceIndex = entry.sourceIndex;
						tocLines.push_back(std::move(line));
						++written;
					}
					if (written == 0)
					{
						TocOverlayLine emptyLine{};
						emptyLine.text = L"(\x65E0\x6807\x9898)"; // (无标题)
						emptyLine.targetSourceIndex = -1;
						tocLines.push_back(std::move(emptyLine));
					}
				}

				for (const auto& tocRange : tocSourceRanges)
				{
					TocOverlayBlock block{};
					block.mappedStart = mapIndex(tocRange.first);
					block.mappedEnd = mapIndex(tocRange.second);
					if (block.mappedEnd < block.mappedStart)
						std::swap(block.mappedStart, block.mappedEnd);
					block.lines = tocLines;
					m_tocOverlayBlocks.push_back(std::move(block));
				}
			}

		// Build RichEdit -> source boundary map for reverse lookups
		// (for example, link extraction on click positions).
		{
			std::vector<long> sourceToRichBoundary(static_cast<size_t>(sourceLength) + 1, 0);
			for (long i = 0; i <= sourceLength; ++i)
			{
				long mapped = mapIndex(i);
				if (i > 0)
					mapped = max(mapped, sourceToRichBoundary[static_cast<size_t>(i - 1)]);
				sourceToRichBoundary[static_cast<size_t>(i)] = mapped;
			}

			m_richToSourceIndexMap.assign(static_cast<size_t>(m_pendingTextLength) + 1, 0);
			size_t sourceCursor = 0;
			for (long rich = 0; rich <= m_pendingTextLength; ++rich)
			{
				while (sourceCursor + 1 < sourceToRichBoundary.size() &&
					sourceToRichBoundary[sourceCursor + 1] <= rich)
				{
					++sourceCursor;
				}
				m_richToSourceIndexMap[static_cast<size_t>(rich)] = static_cast<long>(sourceCursor);
			}
			ValidateRichToSourceMapping(m_richToSourceIndexMap,
				m_pendingTextLength,
				sourceLength,
				_T("OnMarkdownParseComplete"));
		}

		auto convertSpan = [&](auto& span) {
			span.start = mapIndex(span.start);
			span.end = mapIndex(span.end);
			if (span.end < span.start)
				std::swap(span.start, span.end);
		};

		for (auto& span : m_pendingParseResult.hiddenSpans)
			convertSpan(span);
		for (auto& span : m_pendingParseResult.blockSpans)
			convertSpan(span);
		for (auto& span : m_pendingParseResult.inlineSpans)
			convertSpan(span);
		for (auto& range : m_pendingParseResult.codeRanges)
		{
			range.first = mapIndex(range.first);
			range.second = mapIndex(range.second);
			if (range.second < range.first)
				std::swap(range.first, range.second);
		}

	// Cache mapped mermaid blocks for click hit-testing.
	m_mermaidBlocks.clear();
	m_mermaidBlocks.reserve(m_pendingParseResult.mermaidBlocks.size());
	for (const auto& block : m_pendingParseResult.mermaidBlocks)
	{
		MermaidBlock mb;
		mb.rawStart = block.first;
		mb.rawEnd = block.second;
		mb.mappedStart = mapIndex(block.first);
		mb.mappedEnd = mapIndex(block.second);
		if (mb.mappedEnd < mb.mappedStart)
			std::swap(mb.mappedStart, mb.mappedEnd);
		m_mermaidBlocks.push_back(mb);
	}
	// Mermaid formatting loop uses mapped mermaid blocks.
	m_mermaidSpanIndex = 0;

	// Final safety clamp: never allow spans outside the current RichEdit length.
	auto clampIndex = [this](long value) -> long {
		if (value < 0)
			return 0;
		if (value > m_pendingTextLength)
			return m_pendingTextLength;
		return value;
	};

	auto clampSpan = [&](auto& span) -> bool {
		span.start = clampIndex(span.start);
		span.end = clampIndex(span.end);
		if (span.end < span.start)
			std::swap(span.start, span.end);
		return span.start < span.end;
	};

	auto clampRange = [&](std::pair<long, long>& range) -> bool {
		range.first = clampIndex(range.first);
		range.second = clampIndex(range.second);
		if (range.second < range.first)
			std::swap(range.first, range.second);
		return range.first < range.second;
	};

	{
		auto& spans = m_pendingParseResult.hiddenSpans;
		spans.erase(std::remove_if(spans.begin(), spans.end(), [&](auto& span) { return !clampSpan(span); }), spans.end());
	}
	{
		auto& spans = m_pendingParseResult.blockSpans;
		spans.erase(std::remove_if(spans.begin(), spans.end(), [&](auto& span) { return !clampSpan(span); }), spans.end());
	}
	{
		auto& spans = m_pendingParseResult.inlineSpans;
		spans.erase(std::remove_if(spans.begin(), spans.end(), [&](auto& span) { return !clampSpan(span); }), spans.end());
	}
		{
			auto& ranges = m_pendingParseResult.codeRanges;
			ranges.erase(std::remove_if(ranges.begin(), ranges.end(), [&](auto& range) { return !clampRange(range); }), ranges.end());
		}
			{
				auto& blocks = m_tocOverlayBlocks;
				blocks.erase(std::remove_if(blocks.begin(), blocks.end(), [&](auto& block) {
					block.mappedStart = clampIndex(block.mappedStart);
					block.mappedEnd = clampIndex(block.mappedEnd);
					if (block.mappedEnd < block.mappedStart)
						std::swap(block.mappedStart, block.mappedEnd);
					block.lines.erase(std::remove_if(block.lines.begin(), block.lines.end(),
						[](const TocOverlayLine& line) { return line.text.empty(); }),
						block.lines.end());
					return block.lines.empty();
					}), blocks.end());
			}

		// Snapshot selection/scroll for formatting pass.
		m_editContent.GetSel(m_pendingSelStart, m_pendingSelEnd);
		m_pendingFirstVisible = m_editContent.GetFirstVisibleLine();

		m_formattingActive = true;
		m_formattingReschedule = false;
		m_blockSpanIndex = 0;
		m_inlineSpanIndex = 0;
		m_hiddenSpanIndex = 0;
		m_mermaidSpanIndex = 0;
		m_baseFormatApplied = false;
		m_baseFormatPos = 0;

		m_suppressChangeEvent = true;
		m_editContent.SetRedraw(FALSE);
		SetTimer(kFormatTimerId, 1, NULL);

		return 0;
	}
	catch (const SehException& ex)
	{
		KillTimer(kFormatTimerId);
		m_editContent.SetRedraw(TRUE);
		m_suppressChangeEvent = false;
		m_parsingActive = false;
		m_parseProgressPercent = 0;
		m_richToSourceIndexMap.clear();
		m_formattingActive = false;
		m_hideDocumentUntilRenderReady = false;
		CString msg;
		msg.Format(_T("渲染失败（异常码=0x%08X）。"), static_cast<DWORD>(ex.code));
		SetRenderBusy(false, msg);
		return 0;
	}
	catch (...)
	{
		KillTimer(kFormatTimerId);
		m_editContent.SetRedraw(TRUE);
		m_suppressChangeEvent = false;
		m_parsingActive = false;
		m_parseProgressPercent = 0;
		m_richToSourceIndexMap.clear();
		m_formattingActive = false;
		m_hideDocumentUntilRenderReady = false;
		SetRenderBusy(false, _T("渲染失败（未知异常）。"));
		return 0;
	}
}

LRESULT CMarkdownEditorDlg::OnMarkdownParseProgress(WPARAM /*wParam*/, LPARAM lParam)
{
	auto payload = reinterpret_cast<MarkdownParseProgressPayload*>(lParam);
	if (!payload)
		return 0;

	std::unique_ptr<MarkdownParseProgressPayload> holder(payload);
	if (m_destroying)
		return 0;
	if (payload->generation != m_parseGeneration.load())
		return 0;

	int percent = max(0, min(100, payload->percent));
	if (percent > m_parseProgressPercent)
		m_parseProgressPercent = percent;

	if (m_parsingActive && m_renderBusy)
		UpdateStatusText(false);
	return 0;
}

void CMarkdownEditorDlg::ApplyMarkdownFormattingChunk()
{
	try
	{
		SehTranslatorScope sehScope;
		if (!m_formattingActive)
		{
			KillTimer(kFormatTimerId);
			return;
		}

	auto clampValue = [this](long value) -> long {
		if (value < 0)
			return 0;
		if (value > m_pendingTextLength)
			return m_pendingTextLength;
		return value;
	};

	auto applyRange = [&](long start, long end, const CHARFORMAT2& format) {
		start = clampValue(start);
		end = clampValue(end);
		if (start >= end)
			return;

		m_editContent.SetSel(static_cast<LONG>(start), static_cast<LONG>(end));
		CHARFORMAT2 localFormat = format;
		m_editContent.SetSelectionCharFormat(localFormat);
	};

	auto recordHeadingRange = [&](long start, long end, const CHARFORMAT2& headingFormat) {
		start = clampValue(start);
		end = clampValue(end);
		if (start >= end)
			return;
		if (headingFormat.cbSize != sizeof(CHARFORMAT2))
			return;
		HeadingRange hr;
		hr.start = start;
		hr.end = end;
		hr.yHeight = headingFormat.yHeight;
		m_headingRanges.push_back(hr);
	};

		auto isSupportedMermaidBlock = [&](const MermaidBlock& block) -> bool {
		long rawStart = block.rawStart;
		long rawEnd = block.rawEnd;
		if (rawStart < 0)
			rawStart = 0;
		if (rawEnd < rawStart)
			rawEnd = rawStart;
		if (rawEnd > static_cast<long>(m_pendingSourceText.size()))
			rawEnd = static_cast<long>(m_pendingSourceText.size());
		if (rawEnd <= rawStart)
			return false;
		std::wstring code = m_pendingSourceText.substr(static_cast<size_t>(rawStart),
			static_cast<size_t>(rawEnd - rawStart));
		code = TrimW(code);
		if (code.empty())
			return false;

		bool hasGraph = false;
		bool hasFlowchart = false;
		bool hasDirVertical = false;   // TD / TB / BT
		bool hasDirHorizontal = false; // LR / RL
		bool hasSequence = false;
		bool hasEdge = false;
		bool hasNode = false;
		size_t p = 0;
		while (p < code.size())
		{
			size_t e = code.find(L'\n', p);
			if (e == std::wstring::npos)
				e = code.size();
			std::wstring line = code.substr(p, e - p);
			line = TrimW(line);
			if (!line.empty())
			{
				std::wstring lower = line;
				for (auto& ch : lower)
					ch = static_cast<wchar_t>(towlower(ch));
				if (lower.rfind(L"sequencediagram", 0) == 0)
					hasSequence = true;
				auto hasPrefix = [&](const wchar_t* prefix) {
					return lower.rfind(prefix, 0) == 0;
				};
				if (hasPrefix(L"graph "))
				{
					hasGraph = true;
					if (lower.find(L" td") != std::wstring::npos || lower.find(L" td;") != std::wstring::npos) hasDirVertical = true;
					if (lower.find(L" tb") != std::wstring::npos || lower.find(L" tb;") != std::wstring::npos) hasDirVertical = true;
					if (lower.find(L" bt") != std::wstring::npos || lower.find(L" bt;") != std::wstring::npos) hasDirVertical = true;
					if (lower.find(L" lr") != std::wstring::npos || lower.find(L" lr;") != std::wstring::npos) hasDirHorizontal = true;
					if (lower.find(L" rl") != std::wstring::npos || lower.find(L" rl;") != std::wstring::npos) hasDirHorizontal = true;
				}
				if (hasPrefix(L"flowchart "))
				{
					hasFlowchart = true;
					if (lower.find(L" td") != std::wstring::npos || lower.find(L" td;") != std::wstring::npos) hasDirVertical = true;
					if (lower.find(L" tb") != std::wstring::npos || lower.find(L" tb;") != std::wstring::npos) hasDirVertical = true;
					if (lower.find(L" bt") != std::wstring::npos || lower.find(L" bt;") != std::wstring::npos) hasDirVertical = true;
					if (lower.find(L" lr") != std::wstring::npos || lower.find(L" lr;") != std::wstring::npos) hasDirHorizontal = true;
					if (lower.find(L" rl") != std::wstring::npos || lower.find(L" rl;") != std::wstring::npos) hasDirHorizontal = true;
				}
				if (line.find(L"-->") != std::wstring::npos)
					hasEdge = true;
				if ((line.find(L'[') != std::wstring::npos && line.find(L']') != std::wstring::npos) ||
					(line.find(L'{') != std::wstring::npos && line.find(L'}') != std::wstring::npos) ||
					(line.find(L'(') != std::wstring::npos && line.find(L')') != std::wstring::npos))
					hasNode = true;
			}
			p = (e < code.size()) ? (e + 1) : code.size();
		}
		if (hasSequence)
			return true;
		return (hasGraph || hasFlowchart) && (hasDirVertical || hasDirHorizontal) && hasEdge && hasNode;
	};

	auto getTableAvailableWidthTwips = [&]() -> LONG {
		CRect formatRect;
		m_editContent.GetRect(&formatRect);
		int widthPx = formatRect.Width();
		if (widthPx <= 0)
		{
			CRect clientRect;
			m_editContent.GetClientRect(&clientRect);
			widthPx = clientRect.Width();
		}
		if (widthPx <= 0)
			return 0;

		CClientDC dpiDc(&m_editContent);
		int dpiX = dpiDc.GetDeviceCaps(LOGPIXELSX);
		if (dpiX <= 0)
			dpiX = 96;

		LONG availableTwips = static_cast<LONG>(MulDiv(max(80, widthPx - 8), 1440, dpiX));
		const LONG kSidePaddingTwips = 240;
		if (availableTwips > kSidePaddingTwips)
			availableTwips -= kSidePaddingTwips;
		return (std::max)(availableTwips, static_cast<LONG>(1));
	};

	// 统一表格列宽自适应策略：始终按当前可视宽度回流，避免“中间伪空列/空白带”。
	auto buildAdjustedTableStops = [&](const MarkdownBlockSpan& sourceSpan, LONG outStops[32], int& outCount) -> bool {
		outCount = min(sourceSpan.tabStopCount, 32);
		if (outCount <= 0)
			return false;

		LONG prevStop = 0;
		for (int i = 0; i < outCount; ++i)
		{
			LONG stop = sourceSpan.tabStopsTwips[i];
			if (stop <= prevStop)
				stop = prevStop + 1;
			outStops[i] = stop;
			prevStop = stop;
		}

		const LONG availableTwips = getTableAvailableWidthTwips();
		if (availableTwips <= 0)
			return true;

		LONG sourceWidths[32]{};
		LONG prevEdge = 0;
		LONG sourceTotal = 0;
		for (int i = 0; i < outCount; ++i)
		{
			LONG width = outStops[i] - prevEdge;
			if (width <= 0)
				width = 1;
			sourceWidths[i] = width;
			sourceTotal += width;
			prevEdge = outStops[i];
		}
		if (sourceTotal <= 0)
			sourceTotal = outCount;

		const LONG kPreferredMinColTwips = 420;
		const LONG kSoftMinColTwips = 180;
		LONG adaptivePerCol = availableTwips / outCount;
		if (adaptivePerCol < 1)
			adaptivePerCol = 1;

		LONG minColTwips = kPreferredMinColTwips;
		if (availableTwips < kPreferredMinColTwips * outCount)
			minColTwips = adaptivePerCol;
		else if (adaptivePerCol < minColTwips)
			minColTwips = adaptivePerCol;
		if (minColTwips < kSoftMinColTwips && availableTwips >= kSoftMinColTwips * outCount)
			minColTwips = kSoftMinColTwips;
		if (minColTwips > adaptivePerCol)
			minColTwips = adaptivePerCol;
		if (minColTwips < 1)
			minColTwips = 1;

		LONG remainingTotal = availableTwips;
		LONG remainingWeight = sourceTotal;
		LONG accum = 0;
		for (int i = 0; i < outCount; ++i)
		{
			const int remainingCols = outCount - i;
			const LONG minReservedForOthers = minColTwips * (remainingCols - 1);
			const LONG distributable = (std::max)(static_cast<LONG>(0), remainingTotal - minColTwips * remainingCols);

			LONG part = 0;
			if (remainingWeight > 0 && distributable > 0)
			{
				if (i == outCount - 1)
					part = distributable;
				else
					part = static_cast<LONG>(MulDiv(distributable, sourceWidths[i], remainingWeight));
			}

			LONG colWidth = minColTwips + part;
			LONG maxAllowed = remainingTotal - minReservedForOthers;
			if (maxAllowed < minColTwips)
				maxAllowed = minColTwips;

			// 防止单列占用过宽导致中间出现明显空白带。
			if (remainingCols > 1)
			{
				LONG softCap = (std::max)(minColTwips, (availableTwips * 70) / 100);
				if (maxAllowed > softCap)
					maxAllowed = softCap;
			}

			if (colWidth > maxAllowed)
				colWidth = maxAllowed;
			if (colWidth < minColTwips)
				colWidth = minColTwips;

			accum += colWidth;
			outStops[i] = accum;

			remainingTotal -= colWidth;
			if (remainingTotal < 0)
				remainingTotal = 0;
			remainingWeight -= sourceWidths[i];
			if (remainingWeight < 0)
				remainingWeight = 0;
		}

		// 舍入修正：保证最后一个停靠点与目标总宽一致，避免右边界漂移。
		LONG drift = availableTwips - outStops[outCount - 1];
		if (drift != 0)
		{
			for (int i = 0; i < outCount; ++i)
			{
				outStops[i] += drift;
				if (outStops[i] < i + 1)
					outStops[i] = i + 1;
			}
		}
		return true;
	};

	auto appendTableOverlayRow = [&](const MarkdownBlockSpan& span, bool isHeader, COLORREF backColor) {
		TableOverlayRow row{};
		row.start = span.start;
		row.end = span.end;
		row.tabStopCount = min(span.tabStopCount, 32);
		LONG adjustedStops[32]{};
		int adjustedStopCount = 0;
		if (buildAdjustedTableStops(span, adjustedStops, adjustedStopCount))
		{
			row.tabStopCount = adjustedStopCount > 0 ? adjustedStopCount : row.tabStopCount;
			for (int i = 0; i < row.tabStopCount; ++i)
				row.tabStopsTwips[i] = (adjustedStopCount > 0) ? adjustedStops[i] : span.tabStopsTwips[i];
		}
		row.isHeader = isHeader;
		row.backColor = backColor;
		m_tableOverlayRows.push_back(row);
	};

	// 表格文本由 GridCtrl 覆盖层统一绘制：
	// 底层 RichEdit 仅隐藏文字，不再保留旧表格底色，避免滚动时出现底层色块干扰。
	auto buildTableBaseTextHiddenFormat = [&]() -> CHARFORMAT2 {
		CHARFORMAT2 fmt = m_normalFormat;
		fmt.dwMask |= CFM_COLOR;
		fmt.dwMask &= ~CFM_BACKCOLOR;
		fmt.dwEffects &= ~CFE_AUTOCOLOR;
		fmt.crTextColor = m_themeBackground;
		return fmt;
	};

	auto isInsideTableBlock = [&](long pos) -> bool {
		for (const auto& block : m_pendingParseResult.blockSpans)
		{
			if (block.type != MarkdownBlockType::TableHeader &&
				block.type != MarkdownBlockType::TableRowEven &&
				block.type != MarkdownBlockType::TableRowOdd)
			{
				continue;
			}
			if (pos >= block.start && pos < block.end)
				return true;
		}
		return false;
	};

	auto isRangeIntersectTableBlock = [&](long start, long end) -> bool {
		if (end <= start)
			return isInsideTableBlock(start);
		for (const auto& block : m_pendingParseResult.blockSpans)
		{
			if (block.type != MarkdownBlockType::TableHeader &&
				block.type != MarkdownBlockType::TableRowEven &&
				block.type != MarkdownBlockType::TableRowOdd)
			{
				continue;
			}
			if (end > block.start && start < block.end)
				return true;
		}
		return false;
	};

	DWORD startTick = GetTickCount();
	const DWORD budgetMs = 10;

	if (!m_baseFormatApplied)
	{
		const DWORD baseBudgetMs = (budgetMs > 20) ? budgetMs : 20;
		DWORD baseStart = GetTickCount();
		const long chunkChars = 8000;
		while (!m_baseFormatApplied && GetTickCount() - baseStart < baseBudgetMs)
		{
			long start = m_baseFormatPos;
			long end = min(m_pendingTextLength, start + chunkChars);
			m_editContent.SetSel(static_cast<LONG>(start), static_cast<LONG>(end));
			PARAFORMAT2 paraReset{};
			paraReset.cbSize = sizeof(paraReset);
			paraReset.dwMask = PFM_STARTINDENT | PFM_OFFSET | PFM_TABSTOPS;
			paraReset.dxStartIndent = 0;
			paraReset.dxOffset = 0;
			paraReset.cTabCount = 0;
			m_editContent.SetParaFormat(paraReset);

			CHARFORMAT2 baseFormat = m_normalFormat;
			m_editContent.SetSelectionCharFormat(baseFormat);

			m_baseFormatPos = end;
			if (m_baseFormatPos >= m_pendingTextLength)
				m_baseFormatApplied = true;
		}

		if (!m_baseFormatApplied)
			return;
	}

	while (GetTickCount() - startTick < budgetMs)
	{
		// Apply hidden spans first so Markdown markers are suppressed as early as possible.
		// This avoids visible marker leakage when a formatting pass is interrupted mid-way.
		if (m_hiddenSpanIndex < m_pendingParseResult.hiddenSpans.size())
		{
			const size_t hiddenIndex = m_hiddenSpanIndex++;
			const auto& span = m_pendingParseResult.hiddenSpans[hiddenIndex];
			const bool useHiddenEffect = !m_enableWrapSafeBlockMarkerElision ||
				(span.kind != MarkdownHiddenSpanKind::Block);
			applyRange(span.start, span.end, useHiddenEffect ? m_hiddenFormat : m_blockMarkerElisionFormat);
			continue;
		}

		if (m_blockSpanIndex < m_pendingParseResult.blockSpans.size())
		{
			const auto& span = m_pendingParseResult.blockSpans[m_blockSpanIndex++];
			switch (span.type)
			{
			case MarkdownBlockType::Heading1:
				applyRange(span.start, span.end, m_h1Format);
				recordHeadingRange(span.start, span.end, m_h1Format);
				break;
			case MarkdownBlockType::Heading2:
				applyRange(span.start, span.end, m_h2Format);
				recordHeadingRange(span.start, span.end, m_h2Format);
				break;
			case MarkdownBlockType::Heading3:
				applyRange(span.start, span.end, m_h3Format);
				recordHeadingRange(span.start, span.end, m_h3Format);
				break;
			case MarkdownBlockType::Heading4:
				applyRange(span.start, span.end, m_h4Format);
				recordHeadingRange(span.start, span.end, m_h4Format);
				break;
			case MarkdownBlockType::Heading5:
				applyRange(span.start, span.end, m_h5Format);
				recordHeadingRange(span.start, span.end, m_h5Format);
				break;
			case MarkdownBlockType::Heading6:
				applyRange(span.start, span.end, m_h6Format);
				recordHeadingRange(span.start, span.end, m_h6Format);
				break;
			case MarkdownBlockType::BlockQuote:
				applyRange(span.start, span.end, m_quoteFormat);
				if (span.paraStartIndentTwips != 0 || span.paraFirstLineOffsetTwips != 0)
				{
					long selStart = clampValue(span.start);
					long selEnd = clampValue(span.end);
					if (selStart >= selEnd)
						break;

					PARAFORMAT2 para{};
					para.cbSize = sizeof(para);
					para.dwMask = PFM_STARTINDENT | PFM_OFFSET;
					para.dxStartIndent = span.paraStartIndentTwips;
					para.dxOffset = span.paraFirstLineOffsetTwips;
					m_editContent.SetSel(static_cast<LONG>(selStart), static_cast<LONG>(selEnd));
					m_editContent.SetParaFormat(para);
				}
				break;
			case MarkdownBlockType::ListItem:
				applyRange(span.start, span.end, m_listFormat);
				if (span.paraStartIndentTwips != 0 || span.paraFirstLineOffsetTwips != 0)
				{
					long selStart = clampValue(span.start);
					long selEnd = clampValue(span.end);
					if (selStart >= selEnd)
						break;

					PARAFORMAT2 para{};
					para.cbSize = sizeof(para);
					para.dwMask = PFM_STARTINDENT | PFM_OFFSET;
					para.dxStartIndent = span.paraStartIndentTwips;
					para.dxOffset = span.paraFirstLineOffsetTwips;
					m_editContent.SetSel(static_cast<LONG>(selStart), static_cast<LONG>(selEnd));
					m_editContent.SetParaFormat(para);
				}
				break;
			case MarkdownBlockType::CodeBlock:
				applyRange(span.start, span.end, m_codeBlockFormat);
				if (span.paraStartIndentTwips != 0 || span.paraFirstLineOffsetTwips != 0)
				{
					long selStart = clampValue(span.start);
					long selEnd = clampValue(span.end);
					if (selStart >= selEnd)
						break;

					PARAFORMAT2 para{};
					para.cbSize = sizeof(para);
					para.dwMask = PFM_STARTINDENT | PFM_OFFSET | PFM_SPACEBEFORE | PFM_SPACEAFTER;
					para.dxStartIndent = span.paraStartIndentTwips;
					para.dxOffset = span.paraFirstLineOffsetTwips;
					para.dySpaceBefore = 80;
					para.dySpaceAfter = 80;
					m_editContent.SetSel(static_cast<LONG>(selStart), static_cast<LONG>(selEnd));
					m_editContent.SetParaFormat(para);
				}
				break;
			case MarkdownBlockType::TableHeader:
				applyRange(span.start, span.end, buildTableBaseTextHiddenFormat());
				appendTableOverlayRow(span, true, m_tableHeaderFormat.crBackColor);
				break;
			case MarkdownBlockType::TableRowEven:
				applyRange(span.start, span.end, buildTableBaseTextHiddenFormat());
				appendTableOverlayRow(span, false, m_tableRowEvenFormat.crBackColor);
				break;
			case MarkdownBlockType::TableRowOdd:
				applyRange(span.start, span.end, buildTableBaseTextHiddenFormat());
				appendTableOverlayRow(span, false, m_tableRowOddFormat.crBackColor);
				break;
			case MarkdownBlockType::HorizontalRule:
				// Hide the source marker line, draw a real horizontal rule via overlay.
				applyRange(span.start, span.end, m_hrFormat);
				m_horizontalRuleStarts.push_back(span.start);
				break;
			default: break;
			}
			continue;
		}

		if (m_inlineSpanIndex < m_pendingParseResult.inlineSpans.size())
		{
			const auto& span = m_pendingParseResult.inlineSpans[m_inlineSpanIndex++];
			// 表格正文由 Grid 覆盖层统一绘制，底层 RichEdit 的内联样式全部跳过，
			// 避免重新显色导致漏字/色块伪影。
			if (isRangeIntersectTableBlock(span.start, span.end))
				continue;
			switch (span.type)
			{
			case MarkdownInlineType::Bold: applyRange(span.start, span.end, m_boldFormat); break;
			case MarkdownInlineType::BoldItalic:
			{
				CHARFORMAT2 fmt = m_normalFormat;
				fmt.dwMask |= CFM_BOLD | CFM_ITALIC;
				fmt.dwEffects = CFE_BOLD | CFE_ITALIC;
				applyRange(span.start, span.end, fmt);
				break;
			}
			case MarkdownInlineType::Italic: applyRange(span.start, span.end, m_italicFormat); break;
			case MarkdownInlineType::Mark:
			{
				CHARFORMAT2 fmt = m_normalFormat;
				fmt.dwMask |= CFM_BACKCOLOR;
				fmt.dwEffects = 0;
				fmt.crBackColor = m_themeDark ? RGB(70, 60, 20) : RGB(255, 245, 150);
				applyRange(span.start, span.end, fmt);
				break;
			}
			case MarkdownInlineType::Underline:
			{
				CHARFORMAT2 fmt = m_normalFormat;
				fmt.dwMask |= CFM_UNDERLINE;
				fmt.dwEffects = CFE_UNDERLINE;
				applyRange(span.start, span.end, fmt);
				break;
			}
			case MarkdownInlineType::Strikethrough: applyRange(span.start, span.end, m_strikeFormat); break;
			case MarkdownInlineType::InlineCode: applyRange(span.start, span.end, m_inlineCodeFormat); break;
			case MarkdownInlineType::Link: applyRange(span.start, span.end, m_linkFormat); break;
			case MarkdownInlineType::AutoLink: applyRange(span.start, span.end, m_linkFormat); break;
			default: break;
			}
			continue;
		}

		if (m_mermaidSpanIndex < m_mermaidBlocks.size())
		{
			const auto& block = m_mermaidBlocks[m_mermaidSpanIndex++];
			// If mermaid format is supported by our minimal renderer, hide the source text
			// (already blanked in display text) and keep placeholder paragraph space.
			if (isSupportedMermaidBlock(block))
			{
				// IMPORTANT:
				// Do NOT apply CFE_HIDDEN to the entire block.
				// Some RichEdit versions (notably on Win7) will collapse hidden text
				// from layout, which breaks EM_POSFROMCHAR mapping and causes overlay
				// diagrams to overlap with each other and subsequent content.
				applyRange(block.mappedStart, block.mappedEnd, m_mermaidPlaceholderFormat);
			}
			continue;
		}

		FinishMarkdownFormatting();
		return;
	}
	}
	catch (const SehException& ex)
	{
		KillTimer(kFormatTimerId);
		m_editContent.SetRedraw(TRUE);
		m_suppressChangeEvent = false;
		m_formattingActive = false;
		m_hideDocumentUntilRenderReady = false;
		m_formattingReschedule = false;
		m_parseReschedule = false;
		CString msg;
		msg.Format(_T("渲染失败（异常码=0x%08X），已停止渲染。"), static_cast<DWORD>(ex.code));
		SetRenderBusy(false, msg);
	}
	catch (...)
	{
		KillTimer(kFormatTimerId);
		m_editContent.SetRedraw(TRUE);
		m_suppressChangeEvent = false;
		m_formattingActive = false;
		m_hideDocumentUntilRenderReady = false;
		m_formattingReschedule = false;
		m_parseReschedule = false;
		SetRenderBusy(false, _T("渲染失败（未知异常），已停止渲染。"));
	}
}

void CMarkdownEditorDlg::FinishMarkdownFormatting()
{
	KillTimer(kFormatTimerId);

	// Preserve user selection for post-processing passes.
	LONG restoreSelStart = m_pendingSelStart;
	LONG restoreSelEnd = m_pendingSelEnd;

	m_codeBlockRanges.clear();
	m_codeBlockRanges.reserve(m_pendingParseResult.codeRanges.size());
	for (const auto& range : m_pendingParseResult.codeRanges)
	{
		long start = range.first;
		long end = range.second;
		if (start < 0) start = 0;
		if (end > m_pendingTextLength) end = m_pendingTextLength;
		if (start < end)
			m_codeBlockRanges.emplace_back(static_cast<int>(start), static_cast<int>(end));
	}

	int currentFirstVisible = m_editContent.GetFirstVisibleLine();
	m_editContent.LineScroll(m_pendingFirstVisible - currentFirstVisible);
	{
		long selStart = m_pendingSelStart;
		long selEnd = m_pendingSelEnd;
		if (selStart < 0) selStart = 0;
		if (selEnd < 0) selEnd = 0;
		if (selStart > m_pendingTextLength) selStart = m_pendingTextLength;
		if (selEnd > m_pendingTextLength) selEnd = m_pendingTextLength;
		m_editContent.SetSel(selStart, selEnd);
	}
	// Keep the restored selection values consistent.
	restoreSelStart = m_pendingSelStart;
	restoreSelEnd = m_pendingSelEnd;
	if (restoreSelStart < 0) restoreSelStart = 0;
	if (restoreSelEnd < 0) restoreSelEnd = 0;
	if (restoreSelStart > m_pendingTextLength) restoreSelStart = m_pendingTextLength;
	if (restoreSelEnd > m_pendingTextLength) restoreSelEnd = m_pendingTextLength;

		m_editContent.SetRedraw(TRUE);
		m_editContent.Invalidate();
		m_editContent.UpdateWindow();
		m_suppressChangeEvent = false;
		UpdateMermaidDiagrams();
		UpdateMermaidSpacing();
		UpdateTocSpacing();
		UpdateMermaidOverlayRegion(true);

		// Ensure inline code inside headings keeps code styling
		// while also inheriting the heading font size.
		if (!m_headingRanges.empty() && !m_pendingParseResult.inlineSpans.empty())
		{
			for (const auto& hr : m_headingRanges)
			{
				if (hr.start >= hr.end || hr.yHeight <= 0)
					continue;
				for (const auto& span : m_pendingParseResult.inlineSpans)
				{
					if (span.type != MarkdownInlineType::InlineCode)
						continue;
					if (span.end <= hr.start)
						continue;
					if (span.start >= hr.end)
						break;

					const long s = max(span.start, hr.start);
					const long e = min(span.end, hr.end);
					if (s >= e)
						continue;

					CHARFORMAT2 fmt = m_inlineCodeFormat;
					fmt.dwMask |= CFM_SIZE;
					fmt.yHeight = hr.yHeight;
					m_editContent.SetSel(s, e);
					m_editContent.SetSelectionCharFormat(fmt);
				}
			}
			// Restore selection later.
		}

		// Emoji/icon fallback: force emoji characters to use Segoe UI Emoji,
		// otherwise some icons render as tofu (empty squares) with the default font.
		ApplyEmojiFontFallbackToEdit(m_pendingSourceText, m_pendingTextLength);

		// Restore selection (some post-processing passes changed it).
		m_editContent.SetSel(restoreSelStart, restoreSelEnd);

	if (m_sidebarActiveTab == 1)
	{
		UpdateOutline();
	}

	m_formattingActive = false;
	m_parseProgressPercent = 100;
	m_hideDocumentUntilRenderReady = false;
	SetRenderBusy(false, _T(""));

	if (m_parsingActive)
		return;

	if (m_formattingReschedule)
	{
		m_formattingReschedule = false;
		ScheduleMarkdownFormatting(100);
	}
	else if (m_parseReschedule)
	{
		m_parseReschedule = false;
		StartMarkdownParsingAsync();
	}
}

int CMarkdownEditorDlg::MapRichIndexToSourceIndex(int richIndex) const
{
	if (richIndex < 0)
		return -1;

	if (!m_richToSourceIndexMap.empty())
	{
		if (richIndex >= static_cast<int>(m_richToSourceIndexMap.size()))
			return m_richToSourceIndexMap.back();
		return m_richToSourceIndexMap[static_cast<size_t>(richIndex)];
	}

	if (!m_pendingSourceText.empty())
	{
		const int sourceLen = static_cast<int>(m_pendingSourceText.size());
		if (richIndex > sourceLen)
			richIndex = sourceLen;
	}

	return richIndex;
}

int CMarkdownEditorDlg::MapSourceIndexToRichIndex(int sourceIndex) const
{
	if (sourceIndex < 0)
		return -1;

	if (!m_richToSourceIndexMap.empty())
	{
		auto it = std::lower_bound(m_richToSourceIndexMap.begin(), m_richToSourceIndexMap.end(),
			static_cast<long>(sourceIndex));
		if (it == m_richToSourceIndexMap.end())
			return static_cast<int>(m_richToSourceIndexMap.size() - 1);
		return static_cast<int>(it - m_richToSourceIndexMap.begin());
	}

	if (!m_sourceToDisplayIndexMap.empty() &&
		sourceIndex < static_cast<int>(m_sourceToDisplayIndexMap.size()))
	{
		long mapped = m_sourceToDisplayIndexMap[static_cast<size_t>(sourceIndex)];
		if (mapped < 0)
			mapped = 0;
		if (mapped > m_pendingTextLength)
			mapped = m_pendingTextLength;
		return static_cast<int>(mapped);
	}

	if (sourceIndex > m_pendingTextLength)
		sourceIndex = m_pendingTextLength;
	return sourceIndex;
}

CString CMarkdownEditorDlg::ExtractLinkFromPosition(int position)
{
    CString parsed = ExtractLinkFromParseResult(position);
    if (!parsed.IsEmpty())
        return parsed;

    CString text;
	int probePos = position;
    if (!m_pendingSourceText.empty())
    {
        text = m_pendingSourceText.c_str();
		probePos = MapRichIndexToSourceIndex(probePos);
    }
    else
    {
        m_editContent.GetWindowText(text);
    }

    int length = text.GetLength();
    if (probePos < 0 || probePos >= length)
        return CString();

    int directStart = probePos;
    while (directStart > 0 && !IsLinkBoundaryChar(text[directStart - 1]))
        --directStart;

    int directEnd = probePos;
    while (directEnd < length && !IsLinkBoundaryChar(text[directEnd]))
        ++directEnd;

    if (directEnd > directStart)
    {
        CString candidate = text.Mid(directStart, directEnd - directStart);
        SanitizeLinkToken(candidate);

        CString lowerCandidate = candidate;
        lowerCandidate.MakeLower();
        if (lowerCandidate.Find(_T("http://")) == 0 || lowerCandidate.Find(_T("https://")) == 0)
        {
            return candidate;
        }
    }

    // Find the opening '[' preceding the position
    int anchorStart = probePos;
    while (anchorStart >= 0 && text[anchorStart] != _T('['))
    {
        if (text[anchorStart] == _T('\n') || text[anchorStart] == _T('\r'))
            return CString();
        --anchorStart;
    }
    if (anchorStart < 0 || text[anchorStart] != _T('['))
        return CString();

    // Find the closing ']' for the anchor text
    int closingBracket = text.Find(_T(']'), probePos);
    if (closingBracket == -1)
        return CString();

    // Ensure a following '(' exists
    int openParenIndex = closingBracket + 1;
    while (openParenIndex < text.GetLength() && text[openParenIndex] == _T(' '))
        ++openParenIndex;

    if (openParenIndex >= text.GetLength() || text[openParenIndex] != _T('('))
        return CString();

    // Extract until the matching ')'
    int urlStart = openParenIndex + 1;
    int urlEnd = urlStart;
    while (urlEnd < text.GetLength() && text[urlEnd] != _T(')'))
    {
        if (text[urlEnd] == _T('\r') || text[urlEnd] == _T('\n'))
            return CString();
        ++urlEnd;
    }

    if (urlEnd >= text.GetLength())
        return CString();

    CString url = text.Mid(urlStart, urlEnd - urlStart);
    SanitizeLinkToken(url);
    return url;
}

CString CMarkdownEditorDlg::ExtractLinkFromParseResult(int position)
{
    if (position < 0)
        return CString();

    for (const auto& span : m_pendingParseResult.inlineSpans)
    {
        if ((span.type == MarkdownInlineType::Link || span.type == MarkdownInlineType::AutoLink) &&
            position >= span.start && position < span.end)
        {
            CString url(span.href.c_str());
            SanitizeLinkToken(url);
            return url;
        }
    }

    return CString();
}

void CMarkdownEditorDlg::OpenLinkWithDefaultTool(const CString& url)
{
    CString safeUrl = url;
    SanitizeLinkToken(safeUrl);
    if (safeUrl.IsEmpty())
        return;

    CString scheme;
    if (!TryExtractUrlScheme(safeUrl, scheme) || !IsAllowedExternalLinkScheme(scheme))
    {
        AfxMessageBox(_T("已阻止打开非白名单协议链接（仅允许 http/https/mailto/ftp）。"),
            MB_OK | MB_ICONWARNING);
        return;
    }

    HINSTANCE openResult = ShellExecute(nullptr, _T("open"), safeUrl, nullptr, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(openResult) <= 32)
    {
        AfxMessageBox(_T("打开链接失败。"), MB_OK | MB_ICONWARNING);
    }
}

void CMarkdownEditorDlg::OnEditContentLink(NMHDR* pNMHDR, LRESULT* pResult)
{
    if (pResult == nullptr)
        return;

    *pResult = 0;

    if (!m_bMarkdownMode)
        return;

    const ENLINK* pLink = reinterpret_cast<ENLINK*>(pNMHDR);
    if (pLink == nullptr)
        return;

    if (pLink->msg == WM_SETCURSOR || pLink->msg == WM_MOUSEMOVE)
    {
        ::SetCursor(::LoadCursor(nullptr, IDC_HAND));
        *pResult = TRUE;
        return;
    }

    if (pLink->msg == WM_LBUTTONUP)
    {
        CString url = ExtractLinkFromPosition(pLink->chrg.cpMin);
        if (url.IsEmpty())
            url = ExtractLinkFromPosition(pLink->chrg.cpMax - 1);

        if (url.IsEmpty() && pLink->chrg.cpMax > pLink->chrg.cpMin)
        {
            LONG span = pLink->chrg.cpMax - pLink->chrg.cpMin;
            CString anchor;
            LPTSTR buffer = anchor.GetBuffer(span + 1);
            TEXTRANGE tr{};
            tr.chrg = pLink->chrg;
            tr.lpstrText = buffer;
            LRESULT copied = m_editContent.SendMessage(EM_GETTEXTRANGE, 0, reinterpret_cast<LPARAM>(&tr));
            if (copied < 0)
                copied = 0;
            anchor.ReleaseBuffer(static_cast<int>(copied));
            SanitizeLinkToken(anchor);
            url = anchor;
        }

        SanitizeLinkToken(url);

        if (!url.IsEmpty())
        {
            if (url.Find(_T("://")) == -1 && url.Left(4).CompareNoCase(_T("www.")) == 0)
                url = _T("http://") + url;
            OpenLinkWithDefaultTool(url);
            *pResult = TRUE;
        }
    }
}

void CMarkdownEditorDlg::OnEditContentSelChanged(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
    if (pResult)
        *pResult = 0;

    // While the user is dragging to select text, avoid auto-syncing outline selection,
    // which can cause visible jitter and unexpected scrolling.
    m_mouseSelecting = ((::GetKeyState(VK_LBUTTON) & 0x8000) != 0);

    UpdateOutlineSelectionByCaret();
    // Selection changes (including right-click caret moves) can cause the RichEdit to repaint
    // and temporarily obscure overlay-drawn elements like HR lines. Force an overlay redraw.
    if (m_bMarkdownMode && ::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
    {
        UpdateMermaidOverlayRegion(false);
        m_mermaidOverlay.Invalidate(FALSE);
    }
}

BOOL CMarkdownEditorDlg::IsPositionInCodeBlock(int position)
{
    if (position < 0)
        return FALSE;

    return IsIndexInsideRanges(m_codeBlockRanges, position) ? TRUE : FALSE;
}


