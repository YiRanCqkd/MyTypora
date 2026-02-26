#include "pch.h"
#include "framework.h"
#include "MarkdownEditor.h"
#include "MarkdownEditorDlg.h"

#include <algorithm>
#include <memory>
#include <gdiplus.h>

namespace
{
    constexpr UINT_PTR kMermaidOverlayTimerId = 5;

    inline size_t HashWStringForMermaid(const std::wstring& s)
    {
        size_t h = static_cast<size_t>(1469598103934665603ull);
        for (wchar_t ch : s)
        {
            h ^= static_cast<size_t>(ch);
            h *= static_cast<size_t>(1099511628211ull);
        }
        return h;
    }

    inline std::wstring TrimWForMermaid(const std::wstring& s)
    {
        size_t a = 0;
        while (a < s.size() && (s[a] == L' ' || s[a] == L'\t' || s[a] == L'\r' || s[a] == L'\n'))
            ++a;
        size_t b = s.size();
        while (b > a && (s[b - 1] == L' ' || s[b - 1] == L'\t' || s[b - 1] == L'\r' || s[b - 1] == L'\n'))
            --b;
        return (b > a) ? s.substr(a, b - a) : std::wstring();
    }

    inline std::unique_ptr<Gdiplus::Font> CreateUiFontForMermaid(float sizePx, INT style)
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

    inline std::unique_ptr<Gdiplus::Font> CreateMonoFontForMermaid(float sizePx, INT style)
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

	inline int MeasureWrappedTextHeightPx(CDC& dc, const std::wstring& text, int widthPx)
	{
		if (text.empty() || widthPx <= 1)
			return 0;
		CRect calcRect(0, 0, widthPx, 1);
		dc.DrawText(text.c_str(), static_cast<int>(text.size()), &calcRect,
			DT_LEFT | DT_TOP | DT_WORDBREAK | DT_NOPREFIX | DT_EDITCONTROL | DT_CALCRECT);
		return max(0, calcRect.Height());
	}

	inline CFont* SelectTocFontByZoom(CDC& dc, int zoomPercent, const CHARFORMAT2& baseFormat, CFont& tocFont)
	{
		LOGFONT lf{};
		const bool hasBaseFace = (baseFormat.cbSize == sizeof(CHARFORMAT2) && baseFormat.szFaceName[0] != L'\0');
		if (hasBaseFace)
		{
			ZeroMemory(&lf, sizeof(lf));
			lf.lfCharSet = (baseFormat.bCharSet != 0) ? baseFormat.bCharSet : DEFAULT_CHARSET;
			wcscpy_s(lf.lfFaceName, baseFormat.szFaceName);
			lf.lfWeight = (baseFormat.dwEffects & CFE_BOLD) ? FW_BOLD : FW_NORMAL;
			lf.lfItalic = (baseFormat.dwEffects & CFE_ITALIC) ? TRUE : FALSE;
			lf.lfUnderline = (baseFormat.dwEffects & CFE_UNDERLINE) ? TRUE : FALSE;
			lf.lfStrikeOut = (baseFormat.dwEffects & CFE_STRIKEOUT) ? TRUE : FALSE;
		}
		else
		{
			CFont* currentFont = dc.GetCurrentFont();
			if (currentFont == nullptr || !currentFont->GetLogFont(&lf))
			{
				ZeroMemory(&lf, sizeof(lf));
				lf.lfCharSet = DEFAULT_CHARSET;
				wcscpy_s(lf.lfFaceName, L"Microsoft YaHei");
				lf.lfHeight = -16;
				lf.lfWeight = FW_NORMAL;
			}
		}

		if (zoomPercent < 10) zoomPercent = 10;
		if (zoomPercent > 500) zoomPercent = 500;
		const int dpiY = max(96, dc.GetDeviceCaps(LOGPIXELSY));
		LONG baseHeight = lf.lfHeight;
		if (hasBaseFace)
		{
			LONG twips = baseFormat.yHeight;
			if (twips <= 0)
				twips = 220;
			baseHeight = -static_cast<LONG>(MulDiv(twips, dpiY, 1440));
		}
		if (baseHeight == 0)
			baseHeight = -16;
		LONG scaledHeight = static_cast<LONG>(MulDiv(baseHeight, zoomPercent, 100));
		if (scaledHeight == 0)
			scaledHeight = (baseHeight < 0) ? -1 : 1;
		lf.lfHeight = scaledHeight;

		if (!tocFont.CreateFontIndirect(&lf))
			return nullptr;
		return dc.SelectObject(&tocFont);
	}
}

void CMarkdownEditorDlg::UpdateMermaidOverlayRegion(bool forceInvalidate)
{
	if (!::IsWindow(m_mermaidOverlay.GetSafeHwnd()) || !::IsWindow(m_editContent.GetSafeHwnd()))
		return;
	if (m_sidebarResizing)
		return;

	const bool wasVisible = (m_mermaidOverlay.IsWindowVisible() != FALSE);

	if (!m_bMarkdownMode || (m_mermaidBlocks.empty() && m_horizontalRuleStarts.empty() && m_tableOverlayRows.empty() && m_tocOverlayBlocks.empty()))
	{
		m_tocHitRegions.clear();
		m_mermaidOverlay.ShowWindow(SW_HIDE);
		if (m_mermaidOverlayTimer != 0)
		{
			KillTimer(kMermaidOverlayTimerId);
			m_mermaidOverlayTimer = 0;
		}
		return;
	}

	std::vector<CRect> newRects;
	newRects.reserve(m_mermaidDiagrams.size() + m_horizontalRuleStarts.size() + m_tableOverlayRows.size() + m_tocOverlayBlocks.size());

	CRect overlayClient;
	m_mermaidOverlay.GetClientRect(&overlayClient);
	const double zoomScale = max(0.1, min(5.0, (double)GetZoomPercent() / 100.0));

	for (const auto& diagram : m_mermaidDiagrams)
	{
		POINTL pt{};
		LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)diagram.mappedStart);
		if (r == -1)
			continue;

		CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
		m_editContent.ClientToScreen(&topLeft);
		m_mermaidOverlay.ScreenToClient(&topLeft);

		int srcWidthPx = 0;
		int srcHeightPx = 0;
		auto itSize = m_mermaidBitmapSize.find(diagram.hash);
		if (itSize != m_mermaidBitmapSize.end())
		{
			srcWidthPx = itSize->second.first;
			srcHeightPx = itSize->second.second;
		}
		else
		{
			auto itTransientSize = m_mermaidTransientBitmapSize.find(diagram.hash);
			if (itTransientSize != m_mermaidTransientBitmapSize.end())
			{
				srcWidthPx = itTransientSize->second.first;
				srcHeightPx = itTransientSize->second.second;
			}
		}
		if (srcWidthPx <= 0) srcWidthPx = 520;
		if (srcHeightPx <= 0) srcHeightPx = max(140, min(320, diagram.lineCount * 18 + 40));

		int drawWidthPx = (int)(srcWidthPx * zoomScale);
		int drawHeightPx = (int)(srcHeightPx * zoomScale);
		// Do NOT fit-to-width here. Oversized diagrams should be scrollable.

		CRect rc(topLeft.x, topLeft.y, topLeft.x + drawWidthPx, topLeft.y + drawHeightPx);
		if (!CRect().IntersectRect(&rc, &overlayClient))
			continue;
		newRects.push_back(rc);
	}

	// Horizontal rules: include the entire HR line height so the 1px rule line never gets clipped.
	for (long start : m_horizontalRuleStarts)
	{
		POINTL pt{};
		LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)start);
		if (r == -1)
			continue;
		CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
		m_editContent.ClientToScreen(&topLeft);
		m_mermaidOverlay.ScreenToClient(&topLeft);

		int line = m_editContent.LineFromChar(start);
		int rowHeight = 20;
		if (line >= 0)
		{
			int nextLine = line + 1;
			long nextLineStart = m_editContent.LineIndex(nextLine);
			if (nextLineStart >= 0)
			{
				POINTL ptNext{};
				LRESULT rNext = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&ptNext, (LPARAM)nextLineStart);
				if (rNext != -1)
					rowHeight = max(18, (int)ptNext.y - (int)pt.y);
			}
		}
		if (rowHeight <= 0)
			rowHeight = 20;

		CRect rc(overlayClient.left, topLeft.y, overlayClient.right, topLeft.y + rowHeight);
		if (!CRect().IntersectRect(&rc, &overlayClient))
			continue;
		newRects.push_back(rc);
	}

	// Table grid lines: include each table row region.
	if (!m_tableOverlayRows.empty())
	{
		CClientDC dpiDc(this);
		int dpiX = dpiDc.GetDeviceCaps(LOGPIXELSX);
		if (dpiX <= 0) dpiX = 96;
		auto twipsToPxX = [&](LONG twips) -> int {
			return (int)MulDiv(twips, dpiX, 1440);
		};

		for (const auto& row : m_tableOverlayRows)
		{
			if (row.tabStopCount <= 0)
				continue;
			POINTL pt{};
			LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)row.start);
			if (r == -1)
				continue;
			CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
			m_editContent.ClientToScreen(&topLeft);
			m_mermaidOverlay.ScreenToClient(&topLeft);

			int line = m_editContent.LineFromChar(row.start);
			if (line < 0)
				continue;
			int nextLine = line + 1;
			long nextLineStart = m_editContent.LineIndex(nextLine);
			int rowHeight = 20;
			if (nextLineStart >= 0)
			{
				POINTL ptNext{};
				LRESULT rNext = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&ptNext, (LPARAM)nextLineStart);
				if (rNext != -1)
					rowHeight = max(18, (int)ptNext.y - (int)pt.y);
			}
			if (rowHeight <= 0)
				rowHeight = 20;

			int left = max(overlayClient.left + 2, topLeft.x);
			LONG lastStopTwips = row.tabStopsTwips[min(row.tabStopCount, 32) - 1];
			int right = left + twipsToPxX(lastStopTwips);
			right = min(right, overlayClient.right - 2);
			if (right <= left + 40)
				continue;

			CRect rc(left - 2, topLeft.y - 1, right + 2, topLeft.y + rowHeight + 1);
			if (!CRect().IntersectRect(&rc, &overlayClient))
				continue;
			newRects.push_back(rc);
		}
	}

	// TOC overlays: include the expanded text region.
	if (!m_tocOverlayBlocks.empty())
	{
		CClientDC tocMeasureDc(&m_mermaidOverlay);
		CFont* oldTocFont = nullptr;
		if (CFont* editFont = m_editContent.GetFont())
			oldTocFont = tocMeasureDc.SelectObject(editFont);
		CFont tocMeasureFont;
		CFont* oldMeasureZoomFont = SelectTocFontByZoom(tocMeasureDc, GetZoomPercent(), m_normalFormat, tocMeasureFont);

		for (const auto& block : m_tocOverlayBlocks)
		{
			if (block.lines.empty())
				continue;
			POINTL pt{};
			LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)block.mappedStart);
			if (r == -1)
				continue;
			CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
			m_editContent.ClientToScreen(&topLeft);
			m_mermaidOverlay.ScreenToClient(&topLeft);

			int line = m_editContent.LineFromChar(block.mappedStart);
			int rowHeight = 20;
			if (line >= 0)
			{
				long nextLineStart = m_editContent.LineIndex(line + 1);
				if (nextLineStart >= 0)
				{
					POINTL ptNext{};
					LRESULT rNext = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&ptNext, (LPARAM)nextLineStart);
					if (rNext != -1)
						rowHeight = max(18, (int)ptNext.y - (int)pt.y);
				}
			}

			const int textLeft = max(overlayClient.left + 10, topLeft.x);
			const int textRight = overlayClient.right - 12;
			const int textWidth = max(40, textRight - textLeft);
			int measuredTextH = 0;
			for (const auto& lineItem : block.lines)
			{
				if (lineItem.text.empty())
					continue;
				int oneLineH = MeasureWrappedTextHeightPx(tocMeasureDc, lineItem.text, textWidth);
				if (oneLineH <= 0)
					oneLineH = max(rowHeight, 18);
				measuredTextH += oneLineH;
			}
			if (measuredTextH <= 0)
				measuredTextH = rowHeight;
			const int drawHeight = max(rowHeight + 6, measuredTextH + 8);
			CRect rc(max(overlayClient.left + 8, textLeft - 2), topLeft.y,
				min(overlayClient.right - 8, textRight + 2), topLeft.y + drawHeight);

			const bool validRect = (rc.right > rc.left) && CRect().IntersectRect(&rc, &overlayClient);
			if (!validRect)
				continue;

			newRects.push_back(rc);
		}

		if (oldMeasureZoomFont != nullptr)
			tocMeasureDc.SelectObject(oldMeasureZoomFont);

		if (oldTocFont != nullptr)
			tocMeasureDc.SelectObject(oldTocFont);
	}

	// If nothing to show, hide overlay.
	if (newRects.empty())
	{
		m_tocHitRegions.clear();
		m_mermaidOverlay.ShowWindow(SW_HIDE);
		m_lastMermaidOverlayRects.clear();
		// Keep a lightweight timer running as long as the document contains overlay-drawn
		// elements (mermaid/HR/table). Otherwise, when the user scrolls to a diagram for the
		// first time, the overlay won't re-appear until another interaction triggers a repaint.
		if (m_mermaidOverlayTimer == 0)
		{
			m_mermaidOverlayTimer = kMermaidOverlayTimerId;
			SetTimer(kMermaidOverlayTimerId, 200, NULL);
		}
		return;
	}

	auto rectLess = [](const CRect& a, const CRect& b) {
		if (a.top != b.top) return a.top < b.top;
		if (a.left != b.left) return a.left < b.left;
		if (a.bottom != b.bottom) return a.bottom < b.bottom;
		return a.right < b.right;
	};
	std::sort(newRects.begin(), newRects.end(), rectLess);

	bool regionChanged = forceInvalidate;
	if (!regionChanged)
	{
		if (newRects.size() != m_lastMermaidOverlayRects.size())
			regionChanged = true;
		else
		{
			for (size_t i = 0; i < newRects.size(); ++i)
			{
				if (newRects[i] != m_lastMermaidOverlayRects[i])
				{
					regionChanged = true;
					break;
				}
			}
		}
	}

	if (regionChanged)
	{
		CRgn rgn;
		rgn.CreateRectRgn(0, 0, 0, 0);
		for (const auto& rc : newRects)
		{
			CRgn part;
			part.CreateRectRgn(rc.left, rc.top, rc.right, rc.bottom);
			rgn.CombineRgn(&rgn, &part, RGN_OR);
		}

		m_mermaidOverlay.SetWindowRgn((HRGN)rgn.Detach(), TRUE);
		m_lastMermaidOverlayRects = std::move(newRects);
	}

	if (m_mermaidOverlayTimer == 0)
	{
		m_mermaidOverlayTimer = kMermaidOverlayTimerId;
		SetTimer(kMermaidOverlayTimerId, 250, NULL);
	}

	m_mermaidOverlay.ShowWindow(SW_SHOW);
	if (forceInvalidate || regionChanged || !wasVisible)
		m_mermaidOverlay.Invalidate(FALSE);
}

void CMarkdownEditorDlg::UpdateMermaidSpacing()
{
	// Ensure the placeholder paragraphs reserve enough vertical space for the rendered diagram,
	// otherwise the overlay will paint over subsequent content.
	if (!m_bMarkdownMode)
		return;
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return;
	if (m_mermaidDiagrams.empty())
		return;

	// Reset previous spacing.
	for (const auto& diagram : m_mermaidDiagrams)
	{
		m_mermaidSpaceAfterTwips[diagram.hash] = 0;
	}

	CClientDC dc(this);
	int dpiY = dc.GetDeviceCaps(LOGPIXELSY);
	if (dpiY <= 0) dpiY = 96;
	const double zoomScale = max(0.1, min(5.0, (double)GetZoomPercent() / 100.0));
	auto pxToTwipsY = [&](int px) -> LONG {
		// twips = 1/1440 inch
		return (LONG)MulDiv(px, 1440, dpiY);
	};

	for (const auto& diagram : m_mermaidDiagrams)
	{
		int srcWidthPx = 0;
		int srcHeightPx = 0;
		auto itSize = m_mermaidBitmapSize.find(diagram.hash);
		if (itSize != m_mermaidBitmapSize.end())
		{
			srcWidthPx = itSize->second.first;
			srcHeightPx = itSize->second.second;
		}
		else
		{
			auto itTransientSize = m_mermaidTransientBitmapSize.find(diagram.hash);
			if (itTransientSize != m_mermaidTransientBitmapSize.end())
			{
				srcWidthPx = itTransientSize->second.first;
				srcHeightPx = itTransientSize->second.second;
			}
		}
		if (srcWidthPx <= 0) srcWidthPx = 520;
		if (srcHeightPx <= 0) srcHeightPx = max(140, min(320, diagram.lineCount * 18 + 40));

		// Compute drawn height after fit-to-width.
		CRect overlayClient;
		m_mermaidOverlay.GetClientRect(&overlayClient);
		POINTL pt{};
		LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)diagram.mappedStart);
		if (r == -1)
			continue;
		CPoint topLeft((int)pt.x, (int)pt.y);
		m_editContent.ClientToScreen(&topLeft);
		m_mermaidOverlay.ScreenToClient(&topLeft);
		int drawW = (int)(srcWidthPx * zoomScale);
		int drawH = (int)(srcHeightPx * zoomScale);
		// Do NOT force-fit mermaid diagrams to the current viewport width.
		// For LR diagrams, fitting makes nodes look tiny. Allow overflow.

		// Placeholder current height: difference between start and end position.
		POINTL ptEnd{};
		LRESULT r2 = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&ptEnd, (LPARAM)diagram.mappedEnd);
		if (r2 == -1)
			continue;
		int placeholderH = max(0, (int)ptEnd.y - (int)pt.y);
		// The mermaid source is hidden but still includes multiple blank lines; cap the
		// placeholder height to avoid excessive empty space for large diagrams.
		if (placeholderH > drawH)
			placeholderH = drawH;
		int neededExtraPx = drawH - placeholderH;
		if (neededExtraPx < 0)
			neededExtraPx = 0;
		// Add a small padding to separate from next heading.
		neededExtraPx += 12;

		LONG spaceAfter = pxToTwipsY(neededExtraPx);
		if (spaceAfter <= 0)
			continue;

		PARAFORMAT2 para{};
		para.cbSize = sizeof(para);
		para.dwMask = PFM_SPACEAFTER;
		para.dySpaceAfter = spaceAfter;
		m_editContent.SetSel(diagram.mappedEnd, diagram.mappedEnd);
		m_editContent.SetParaFormat(para);
		m_mermaidSpaceAfterTwips[diagram.hash] = spaceAfter;
	}
}

void CMarkdownEditorDlg::TouchMermaidBitmapCache(size_t hash)
{
	m_mermaidBitmapLastUseSeq[hash] = ++m_mermaidBitmapUseSeq;
}

void CMarkdownEditorDlg::TrimMermaidBitmapCache(size_t reserveForNewEntry, const std::set<size_t>* protectedHashes)
{
	while (m_mermaidBitmapCache.size() + reserveForNewEntry > kMaxMermaidBitmapCacheEntries)
	{
		bool foundCandidate = false;
		size_t oldestHash = 0;
		unsigned long long oldestSeq = 0;

		for (const auto& kv : m_mermaidBitmapCache)
		{
			const size_t hash = kv.first;
			if (m_mermaidFetchInFlight.find(hash) != m_mermaidFetchInFlight.end())
				continue;
			if (protectedHashes != nullptr && protectedHashes->find(hash) != protectedHashes->end())
				continue;

			unsigned long long seq = 0;
			auto itSeq = m_mermaidBitmapLastUseSeq.find(hash);
			if (itSeq != m_mermaidBitmapLastUseSeq.end())
				seq = itSeq->second;

			if (!foundCandidate || seq < oldestSeq)
			{
				foundCandidate = true;
				oldestHash = hash;
				oldestSeq = seq;
			}
		}

		if (!foundCandidate)
			break;

		auto itBmp = m_mermaidBitmapCache.find(oldestHash);
		if (itBmp != m_mermaidBitmapCache.end())
		{
			if (itBmp->second)
				::DeleteObject(itBmp->second);
			m_mermaidBitmapCache.erase(itBmp);
		}
		m_mermaidBitmapSize.erase(oldestHash);
		m_mermaidBitmapLastUseSeq.erase(oldestHash);
		m_mermaidSpaceAfterTwips.erase(oldestHash);
	}
}

void CMarkdownEditorDlg::UpdateMermaidDiagrams()
{
	m_mermaidDiagrams.clear();
	for (auto& kv : m_mermaidTransientBitmapCache)
	{
		if (kv.second)
			::DeleteObject(kv.second);
	}
	m_mermaidTransientBitmapCache.clear();
	m_mermaidTransientBitmapSize.clear();
	++m_mermaidRenderGeneration;
	if (!m_bMarkdownMode)
		return;
	if (!::IsWindow(m_editContent.GetSafeHwnd()) || !::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		return;
	if (m_mermaidBlocks.empty())
	{
		m_mermaidOverlay.Invalidate(FALSE);
		return;
	}

	// Do NOT force-show the overlay here.
	// The overlay window region may be stale (e.g. after sidebar live-resize), and showing it
	// before UpdateMermaidOverlayRegion() runs can temporarily cover the editor.
	if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()) && m_mermaidOverlay.IsWindowVisible())
	{
		m_mermaidOverlay.Invalidate(FALSE);
	}

	// Mermaid rendering strategy: render in-process (no external dependencies).
	// Keep a reasonable upper bound to avoid insane allocations on malformed input.
	const int maxDiagramWidth = 4096;
	std::set<size_t> currentDiagramHashes;
	// NOTE: m_mermaidRenderGeneration is still used to invalidate caches on re-format.
	for (const auto& block : m_mermaidBlocks)
	{
		// Extract raw mermaid code.
		long rawStart = block.rawStart;
		long rawEnd = block.rawEnd;
		if (rawStart < 0) rawStart = 0;
		if (rawEnd < rawStart) rawEnd = rawStart;
		if (rawEnd > static_cast<long>(m_pendingSourceText.size()))
			rawEnd = static_cast<long>(m_pendingSourceText.size());
		std::wstring code = m_pendingSourceText.substr(static_cast<size_t>(rawStart), static_cast<size_t>(rawEnd - rawStart));
		code = TrimWForMermaid(code);
		if (code.empty())
			continue;
		// sequenceDiagram is handled by a minimal local renderer below.

		size_t hash = HashWStringForMermaid(code);
		currentDiagramHashes.insert(hash);
		MermaidDiagram d;
		d.mappedStart = block.mappedStart;
		d.mappedEnd = block.mappedEnd;
		d.hash = hash;
		// approximate line count for placeholder height before bitmap arrives
		d.lineCount = 1;
		for (wchar_t ch : code)
			if (ch == L'\n')
				++d.lineCount;
		m_mermaidDiagrams.push_back(d);

		if (m_mermaidBitmapCache.find(hash) != m_mermaidBitmapCache.end())
		{
			TouchMermaidBitmapCache(hash);
			continue;
		}
		if (m_mermaidTransientBitmapCache.find(hash) != m_mermaidTransientBitmapCache.end())
			continue;
		if (m_mermaidFetchInFlight.find(hash) != m_mermaidFetchInFlight.end())
			continue;
		m_mermaidFetchInFlight.insert(hash);

		// Local flowchart render (minimal): graph TD/LR with A[Text] and -->.
		HBITMAP bmp = nullptr;
		int w = 0;
		int h = 0;
		{
			// Parse lines
			std::wstring src = code;
			for (auto& ch : src)
				if (ch == L'\r') ch = L'\n';
			std::vector<std::wstring> lines;
			{
				size_t p = 0;
				while (p < src.size())
				{
					size_t e = src.find(L'\n', p);
					if (e == std::wstring::npos) e = src.size();
					std::wstring line = src.substr(p, e - p);
					line = TrimWForMermaid(line);
					if (!line.empty() && !(line.size() >= 2 && line[0] == L'%' && line[1] == L'%'))
						lines.push_back(line);
					p = (e < src.size()) ? (e + 1) : src.size();
				}
			}

			auto startsWithNoCase = [&](const std::wstring& s, const std::wstring& prefix) {
				if (s.size() < prefix.size()) return false;
				for (size_t i = 0; i < prefix.size(); ++i)
				{
					wchar_t a = static_cast<wchar_t>(towlower(s[i]));
					wchar_t b = static_cast<wchar_t>(towlower(prefix[i]));
					if (a != b) return false;
				}
				return true;
			};

			bool isGraphTD = false;
			bool isGraphLR = false;
			bool isReverseDir = false; // BT / RL
			for (const auto& line : lines)
			{
				// Mermaid v10+ also supports "flowchart". Treat it as an alias.
				// TD/TB: top-down; BT: bottom-up.
				if (startsWithNoCase(line, L"graph TD") || startsWithNoCase(line, L"graph TD;") ||
					startsWithNoCase(line, L"graph TB") || startsWithNoCase(line, L"graph TB;") ||
					startsWithNoCase(line, L"flowchart TD") || startsWithNoCase(line, L"flowchart TD;") ||
					startsWithNoCase(line, L"flowchart TB") || startsWithNoCase(line, L"flowchart TB;"))
				{
					isGraphTD = true;
					isReverseDir = false;
					break;
				}
				if (startsWithNoCase(line, L"graph BT") || startsWithNoCase(line, L"graph BT;") ||
					startsWithNoCase(line, L"flowchart BT") || startsWithNoCase(line, L"flowchart BT;"))
				{
					isGraphTD = true;
					isReverseDir = true;
					break;
				}
				// LR: left-to-right; RL: right-to-left.
				if (startsWithNoCase(line, L"graph LR") || startsWithNoCase(line, L"graph LR;") ||
					startsWithNoCase(line, L"flowchart LR") || startsWithNoCase(line, L"flowchart LR;"))
				{
					isGraphLR = true;
					isReverseDir = false;
					break;
				}
				if (startsWithNoCase(line, L"graph RL") || startsWithNoCase(line, L"graph RL;") ||
					startsWithNoCase(line, L"flowchart RL") || startsWithNoCase(line, L"flowchart RL;"))
				{
					isGraphLR = true;
					isReverseDir = true;
					break;
				}
			}

			struct Node { CString id; CString label; MermaidNodeShape shape = MermaidNodeShape::Rect; };
			std::vector<Node> nodes;
			std::vector<std::pair<int, int>> edges;
			std::vector<CString> edgeLabels;
			auto getNodeIndex = [&](const CString& id) -> int {
				for (int i = 0; i < (int)nodes.size(); ++i)
					if (nodes[i].id.CompareNoCase(id) == 0)
						return i;
				nodes.push_back({ id, id, MermaidNodeShape::Rect });
				return (int)nodes.size() - 1;
			};

			auto parseNodeToken = [&](const std::wstring& token, CString& outId, CString& outLabel, MermaidNodeShape& outShape) -> bool {
				// Supported minimal subset:
				// - A[Text]  : rectangle
				// - A{Text}  : diamond (decision)
				// - A(Text)  : rounded rectangle
				// - A([Text]) : stadium (terminator)
				// - A((Text)) : circle
				// Note: for compatibility we accept nested forms like Start([Text]).
				std::wstring t = TrimWForMermaid(token);
				if (t.empty())
					return false;

				// Special-case nested stadium/circle forms first, otherwise the generic parser
				// would treat the leading '(' as the label delimiter and keep extra brackets.
				size_t pb2 = t.find(L'(');
				size_t pe2 = t.rfind(L')');
				if (pb2 != std::wstring::npos && pe2 != std::wstring::npos && pe2 > pb2 + 3)
				{
					// Stadium: A([Text])
					if (t[pb2 + 1] == L'[' && t[pe2 - 1] == L']')
					{
						std::wstring id = TrimWForMermaid(t.substr(0, pb2));
						std::wstring label = t.substr(pb2 + 2, (pe2 - 1) - (pb2 + 2));
						id = TrimWForMermaid(id);
						label = TrimWForMermaid(label);
						if (!id.empty())
						{
							outId = id.c_str();
							outLabel = label.empty() ? outId : CString(label.c_str());
							outShape = MermaidNodeShape::Stadium;
							return true;
						}
					}
					// Circle: A((Text))
					if (t[pb2 + 1] == L'(' && t[pe2 - 1] == L')' && pe2 > pb2 + 4)
					{
						std::wstring id = TrimWForMermaid(t.substr(0, pb2));
						std::wstring label = t.substr(pb2 + 2, (pe2 - 1) - (pb2 + 2));
						id = TrimWForMermaid(id);
						label = TrimWForMermaid(label);
						if (!id.empty())
						{
							outId = id.c_str();
							outLabel = label.empty() ? outId : CString(label.c_str());
							outShape = MermaidNodeShape::Circle;
							return true;
						}
					}
				}

				size_t lb = t.find(L'[');
				size_t rb = t.rfind(L']');
				size_t cb = t.find(L'{');
				size_t ce = t.rfind(L'}');
				size_t pb = t.find(L'(');
				size_t pe = t.rfind(L')');

				const bool hasRect = (lb != std::wstring::npos && rb != std::wstring::npos && rb > lb);
				const bool hasDiamond = (cb != std::wstring::npos && ce != std::wstring::npos && ce > cb);
				const bool hasRound = (pb != std::wstring::npos && pe != std::wstring::npos && pe > pb);

				if (!hasRect && !hasDiamond && !hasRound)
				{
					std::wstring id = TrimWForMermaid(t);
					if (id.empty()) return false;
					outId = id.c_str();
					outLabel = outId;
					outShape = MermaidNodeShape::Rect;
					return true;
				}

				size_t openPos = std::wstring::npos;
				size_t closePos = std::wstring::npos;
				MermaidNodeShape shape = MermaidNodeShape::Rect;
				auto pick = [&](size_t pos, size_t close, MermaidNodeShape s) {
					if (pos == std::wstring::npos || close == std::wstring::npos || close <= pos)
						return;
					if (openPos == std::wstring::npos || pos < openPos)
					{
						openPos = pos;
						closePos = close;
						shape = s;
					}
				};
				pick(lb, rb, MermaidNodeShape::Rect);
				pick(cb, ce, MermaidNodeShape::Diamond);
				pick(pb, pe, MermaidNodeShape::RoundRect);
				if (openPos == std::wstring::npos || closePos == std::wstring::npos)
					return false;

				std::wstring id = TrimWForMermaid(t.substr(0, openPos));
				std::wstring label = t.substr(openPos + 1, closePos - openPos - 1);
				id = TrimWForMermaid(id);
				label = TrimWForMermaid(label);
				if (id.empty()) return false;
				outId = id.c_str();
				outLabel = label.empty() ? outId : CString(label.c_str());
				outShape = shape;
				return true;
			};

			auto parseEdge = [&](const std::wstring& line, std::wstring& outLeft, std::wstring& outRight, std::wstring& outEdgeLabel) -> bool {
				outLeft.clear();
				outRight.clear();
				outEdgeLabel.clear();
				std::wstring s = TrimWForMermaid(line);
				if (s.empty())
					return false;
				size_t arrow = s.find(L"-->");
				if (arrow == std::wstring::npos)
					return false;

				std::wstring left = TrimWForMermaid(s.substr(0, arrow));
				std::wstring right = TrimWForMermaid(s.substr(arrow + 3));
				if (left.empty() || right.empty())
					return false;

				// Right-side label form: A -->|One| B
				if (!right.empty() && right[0] == L'|')
				{
					size_t second = right.find(L'|', 1);
					if (second != std::wstring::npos)
					{
						outEdgeLabel = TrimWForMermaid(right.substr(1, second - 1));
						right = TrimWForMermaid(right.substr(second + 1));
					}
				}

				// Left-side label form: A -- text --> B (rare, but appears in examples)
				size_t labelSep = left.find(L"--");
				if (labelSep != std::wstring::npos)
				{
					std::wstring maybeFrom = TrimWForMermaid(left.substr(0, labelSep));
					std::wstring maybeLabel = TrimWForMermaid(left.substr(labelSep + 2));
					if (!maybeFrom.empty() && !maybeLabel.empty())
					{
						left = maybeFrom;
						if (outEdgeLabel.empty())
							outEdgeLabel = maybeLabel;
					}
				}

				outLeft = left;
				outRight = right;
				return !outLeft.empty() && !outRight.empty();
			};

			bool isSequenceDiagram = false;
			for (const auto& line : lines)
			{
				if (startsWithNoCase(line, L"sequenceDiagram"))
				{
					isSequenceDiagram = true;
					break;
				}
			}

			if (isSequenceDiagram)
			{
				// Minimal sequence diagram renderer:
				// - participants inferred from declarations/messages
				// - note over/left of/right of
				// - alt/else/end/opt/loop/par blocks
				// - messages/notes rendered in source order
				struct SeqMsg { int fromIdx = -1; int toIdx = -1; bool dashed = false; CString text; };
				enum class SeqNoteKind { Over, LeftOf, RightOf };
				struct SeqNote
				{
					SeqNoteKind kind = SeqNoteKind::Over;
					int p1 = -1;
					int p2 = -1;
					CString text;
				};
				struct SeqEvent
				{
					enum class Type { Message, Note };
					Type type = Type::Message;
					int index = -1;
				};
				struct SeqBranch { int startRowIndex = 0; CString label; };
				struct SeqBlock
				{
					CString kind; // alt/opt/loop/par
					CString label;
					int startRowIndex = 0;
					int endRowIndexExclusive = 0;
					int depth = 0;
					std::vector<SeqBranch> branches; // alt/par branches
				};
				struct SeqParticipant
				{
					CString key;
					CString label;
					bool actor = false;
				};
				std::vector<SeqParticipant> participants;
				std::map<std::wstring, int> participantLookup;
				std::vector<SeqMsg> msgs;
				std::vector<SeqNote> notes;
				std::vector<SeqEvent> events;
				std::vector<SeqBlock> blocks;
				std::vector<int> openBlockStack;

				auto toLower = [](std::wstring s) -> std::wstring {
					for (auto& ch : s)
						ch = static_cast<wchar_t>(towlower(ch));
					return s;
				};
				auto normalizeKey = [&](const std::wstring& src) -> std::wstring {
					return toLower(TrimWForMermaid(src));
				};
				auto normalizeParticipantToken = [&](const std::wstring& src) -> std::wstring {
					std::wstring t = TrimWForMermaid(src);
					while (!t.empty() && (t.back() == L';' || t.back() == L','))
						t.pop_back();
					t = TrimWForMermaid(t);
					if (t.size() >= 2)
					{
						const wchar_t first = t.front();
						const wchar_t last = t.back();
						if ((first == L'"' && last == L'"') || (first == L'\'' && last == L'\'') ||
							(first == L'`' && last == L'`'))
						{
							t = TrimWForMermaid(t.substr(1, t.size() - 2));
						}
					}
					return t;
				};
				auto findOrAddParticipant = [&](const std::wstring& name, const std::wstring& displayName = std::wstring(),
					bool actorDeclared = false, bool explicitDeclaration = false) -> int {
					std::wstring key = normalizeParticipantToken(name);
					if (key.empty())
						return -1;
					const std::wstring keyLower = normalizeKey(key);
					auto it = participantLookup.find(keyLower);
					if (it != participantLookup.end())
					{
						const int idx = it->second;
						std::wstring label = normalizeParticipantToken(displayName);
						if (idx >= 0 && idx < (int)participants.size())
						{
							if (explicitDeclaration && !label.empty())
							{
								participants[idx].label = CString(label.c_str());
							}
							if (actorDeclared)
								participants[idx].actor = true;
						}
						return idx;
					}

					std::wstring label = normalizeParticipantToken(displayName);
					if (label.empty())
						label = key;
					const int idx = (int)participants.size();
					SeqParticipant p;
					p.key = CString(key.c_str());
					p.label = CString(label.c_str());
					p.actor = actorDeclared;
					participants.push_back(p);

					participantLookup[keyLower] = idx;
					return idx;
				};

				auto tryParseArrow = [&](const std::wstring& s, std::wstring& outFrom, std::wstring& outTo, std::wstring& outText, bool& outDashed) -> bool {
					outFrom.clear();
					outTo.clear();
					outText.clear();
					outDashed = false;
					if (s.empty())
						return false;
					// Find an arrow token. Support common forms: ->, ->>, -->, -->>, etc.
					size_t posLong = s.find(L"-->");
					size_t posShort = s.find(L"->");
					size_t arrow = std::wstring::npos;
					size_t baseLen = 0;
					if (posLong != std::wstring::npos)
					{
						arrow = posLong;
						baseLen = 3;
					}
					if (posShort != std::wstring::npos && (arrow == std::wstring::npos || posShort < arrow))
					{
						arrow = posShort;
						baseLen = 2;
					}
					if (arrow == std::wstring::npos)
						return false;

					std::wstring left = TrimWForMermaid(s.substr(0, arrow));
					if (left.empty())
						return false;
					size_t rightStart = arrow + baseLen;
					// Skip extra arrow decorations: > - + = x o <
					while (rightStart < s.size())
					{
						wchar_t ch = s[rightStart];
						if (ch == L'>' || ch == L'-' || ch == L'+' || ch == L'=' || ch == L'x' || ch == L'o' || ch == L'<')
							++rightStart;
						else
							break;
					}
					if (rightStart >= s.size())
						return false;
					std::wstring arrowToken = s.substr(arrow, rightStart - arrow);
					outDashed = (arrowToken.find(L"--") != std::wstring::npos);
					std::wstring right = TrimWForMermaid(s.substr(rightStart));
					if (right.empty())
						return false;

					size_t colon = right.find(L':');
					std::wstring to = (colon == std::wstring::npos) ? TrimWForMermaid(right) : TrimWForMermaid(right.substr(0, colon));
					std::wstring text = (colon == std::wstring::npos) ? L"" : TrimWForMermaid(right.substr(colon + 1));
					while (!to.empty() && (to.front() == L'+' || to.front() == L'-'))
						to.erase(to.begin());
					while (!left.empty() && (left.back() == L'+' || left.back() == L'-'))
						left.pop_back();
					left = normalizeParticipantToken(left);
					to = normalizeParticipantToken(to);
					if (left.empty())
						return false;
					if (to.empty())
						return false;
					outFrom = left;
					outTo = to;
					outText = text;
					return true;
				};

				auto startsKeywordLine = [&](const std::wstring& lower, const wchar_t* keyword) -> bool {
					const size_t n = wcslen(keyword);
					if (lower.size() < n)
						return false;
					if (lower.compare(0, n, keyword) != 0)
						return false;
					return (lower.size() == n) || iswspace(lower[n]) || lower[n] == 0x3000;
				};

				auto extractKeywordTail = [&](const std::wstring& srcLine, size_t keywordLen) -> CString {
					if (srcLine.size() <= keywordLen)
						return CString();
					std::wstring tail = TrimWForMermaid(srcLine.substr(keywordLen));
					return tail.empty() ? CString() : CString(tail.c_str());
				};

				auto normalizeSeqLine = [&](const std::wstring& src) -> std::wstring {
					std::wstring t = src;
					for (auto& ch : t)
					{
						// full-width / non-breaking space -> regular space
						if (ch == 0x3000 || ch == 0x00A0)
							ch = L' ';
					}
					return TrimWForMermaid(t);
				};

				auto findColonPos = [&](const std::wstring& s) -> size_t {
					size_t p1 = s.find(L':');
					size_t p2 = s.find(L'：');
					if (p1 == std::wstring::npos) return p2;
					if (p2 == std::wstring::npos) return p1;
					return min(p1, p2);
				};

				auto isWsEx = [&](wchar_t ch) -> bool {
					return iswspace(ch) || ch == 0x3000;
				};
				auto findAsKeywordPos = [&](const std::wstring& src) -> size_t {
					if (src.size() < 4)
						return std::wstring::npos;
					std::wstring lower = toLower(src);
					for (size_t i = 1; i + 2 < lower.size(); ++i)
					{
						if (lower[i] == L'a' && lower[i + 1] == L's' &&
							isWsEx(lower[i - 1]) && isWsEx(lower[i + 2]))
						{
							return i;
						}
					}
					return std::wstring::npos;
				};

				for (const auto& line : lines)
				{
					std::wstring normalizedLine = normalizeSeqLine(line);
					if (normalizedLine.empty())
						continue;

					if (startsWithNoCase(normalizedLine, L"sequenceDiagram"))
						continue;
					// explicit participants
					std::wstring lowerForKeyword = normalizedLine;
					for (auto& ch : lowerForKeyword) ch = static_cast<wchar_t>(towlower(ch));
					if (startsKeywordLine(lowerForKeyword, L"participant") || startsKeywordLine(lowerForKeyword, L"actor"))
					{
						const bool isActor = startsKeywordLine(lowerForKeyword, L"actor");
						const size_t keywordLen = isActor ? 5 : 11;
						std::wstring name = normalizeSeqLine(normalizedLine.substr(keywordLen));
						if (!name.empty())
						{
							size_t asPos = findAsKeywordPos(name);
							if (asPos != std::wstring::npos)
							{
								// Mermaid/Typora style: participant <id> as <display>.
								std::wstring leftPart = normalizeParticipantToken(name.substr(0, asPos));
								std::wstring rightPart = normalizeParticipantToken(name.substr(asPos + 2));
								if (!leftPart.empty())
								{
									std::wstring idPart = leftPart;
									std::wstring displayPart = rightPart.empty() ? leftPart : rightPart;
									findOrAddParticipant(idPart, displayPart, isActor, true);
								}
								else
								{
									findOrAddParticipant(name, std::wstring(), isActor, true);
								}
							}
							else
							{
								findOrAddParticipant(name, std::wstring(), isActor, true);
							}
						}
						continue;
					}

					std::wstring lower = normalizedLine;
					for (auto& ch : lower) ch = static_cast<wchar_t>(towlower(ch));

					const bool isAlt = startsKeywordLine(lower, L"alt");
					const bool isOpt = startsKeywordLine(lower, L"opt");
					const bool isLoop = startsKeywordLine(lower, L"loop");
					const bool isPar = startsKeywordLine(lower, L"par");
					const bool isElse = startsKeywordLine(lower, L"else");
					const bool isAnd = startsKeywordLine(lower, L"and");
					const bool isEnd = startsKeywordLine(lower, L"end");

					if (isAlt || isOpt || isLoop || isPar)
					{
						SeqBlock block;
						if (isAlt)
							block.kind = _T("alt");
						else if (isOpt)
							block.kind = _T("opt");
						else if (isLoop)
							block.kind = _T("loop");
						else
							block.kind = _T("par");
						block.label = extractKeywordTail(line, block.kind.GetLength());
						block.startRowIndex = (int)events.size();
						block.endRowIndexExclusive = (int)events.size();
						block.depth = (int)openBlockStack.size();
						if (block.kind.CompareNoCase(_T("alt")) == 0 || block.kind.CompareNoCase(_T("par")) == 0)
						{
							SeqBranch firstBranch;
							firstBranch.startRowIndex = block.startRowIndex;
							firstBranch.label = block.label;
							block.branches.push_back(firstBranch);
						}
						blocks.push_back(block);
						openBlockStack.push_back((int)blocks.size() - 1);
						continue;
					}

					if (isElse || isAnd)
					{
						if (!openBlockStack.empty())
						{
							SeqBlock& current = blocks[openBlockStack.back()];
							const bool allowElse = (current.kind.CompareNoCase(_T("alt")) == 0);
							const bool allowAnd = (current.kind.CompareNoCase(_T("par")) == 0);
							if ((isElse && allowElse) || (isAnd && allowAnd))
							{
								const size_t keywordLen = isElse ? 4 : 3;
								SeqBranch branch;
								branch.startRowIndex = (int)events.size();
								branch.label = extractKeywordTail(line, keywordLen);
								if (branch.label.IsEmpty())
									branch.label = isElse ? _T("else") : _T("and");
								if (current.branches.empty() || current.branches.back().startRowIndex != branch.startRowIndex)
									current.branches.push_back(branch);
							}
						}
						continue;
					}

					if (isEnd)
					{
						if (!openBlockStack.empty())
						{
							SeqBlock& current = blocks[openBlockStack.back()];
							current.endRowIndexExclusive = (int)events.size();
							openBlockStack.pop_back();
						}
						continue;
					}

					if (startsKeywordLine(lower, L"note"))
					{
						std::wstring tail = normalizeSeqLine(normalizedLine.substr(4));
						std::wstring lowerTail = toLower(tail);
						SeqNote note;
						bool noteParsed = false;

						auto parseSingleTarget = [&](const std::wstring& target, int& outIdx) -> bool {
							std::wstring t = normalizeParticipantToken(target);
							if (t.empty())
								return false;
							outIdx = findOrAddParticipant(t);
							return outIdx >= 0;
						};

						if (lowerTail.rfind(L"over ", 0) == 0)
						{
							note.kind = SeqNoteKind::Over;
							std::wstring rest = normalizeSeqLine(tail.substr(5));
							size_t colon = findColonPos(rest);
							std::wstring targetPart = (colon == std::wstring::npos) ? rest : TrimWForMermaid(rest.substr(0, colon));
							std::wstring textPart = (colon == std::wstring::npos) ? std::wstring() : TrimWForMermaid(rest.substr(colon + 1));
							size_t comma = targetPart.find(L',');
							if (comma == std::wstring::npos)
								comma = targetPart.find(L'，');
							if (comma == std::wstring::npos)
							{
								if (parseSingleTarget(targetPart, note.p1))
								{
									note.p2 = note.p1;
									note.text = CString(textPart.c_str());
									noteParsed = true;
								}
							}
							else
							{
								std::wstring t1 = TrimWForMermaid(targetPart.substr(0, comma));
								std::wstring t2 = TrimWForMermaid(targetPart.substr(comma + 1));
								if (parseSingleTarget(t1, note.p1) && parseSingleTarget(t2, note.p2))
								{
									note.text = CString(textPart.c_str());
									noteParsed = true;
								}
							}
						}
						else if (lowerTail.rfind(L"left of ", 0) == 0 || lowerTail.rfind(L"right of ", 0) == 0)
						{
							const bool leftOf = (lowerTail.rfind(L"left of ", 0) == 0);
							note.kind = leftOf ? SeqNoteKind::LeftOf : SeqNoteKind::RightOf;
							std::wstring rest = normalizeSeqLine(tail.substr(leftOf ? 8 : 9));
							size_t colon = findColonPos(rest);
							std::wstring targetPart = (colon == std::wstring::npos) ? rest : TrimWForMermaid(rest.substr(0, colon));
							std::wstring textPart = (colon == std::wstring::npos) ? std::wstring() : TrimWForMermaid(rest.substr(colon + 1));
							if (parseSingleTarget(targetPart, note.p1))
							{
								note.p2 = note.p1;
								note.text = CString(textPart.c_str());
								noteParsed = true;
							}
						}

						if (noteParsed)
						{
							notes.push_back(note);
							SeqEvent ev;
							ev.type = SeqEvent::Type::Note;
							ev.index = (int)notes.size() - 1;
							events.push_back(ev);
							continue;
						}
					}

					std::wstring from;
					std::wstring to;
					std::wstring text;
					bool dashed = false;
					if (!tryParseArrow(normalizedLine, from, to, text, dashed))
						continue;

					int fromIdx = findOrAddParticipant(from);
					int toIdx = findOrAddParticipant(to);
					if (fromIdx < 0 || toIdx < 0)
						continue;
					SeqMsg m;
					m.fromIdx = fromIdx;
					m.toIdx = toIdx;
					m.dashed = dashed;
					m.text = CString(text.c_str());
					msgs.push_back(m);
					SeqEvent ev;
					ev.type = SeqEvent::Type::Message;
					ev.index = (int)msgs.size() - 1;
					events.push_back(ev);
				}

				while (!openBlockStack.empty())
				{
					SeqBlock& current = blocks[openBlockStack.back()];
					current.endRowIndexExclusive = (int)events.size();
					openBlockStack.pop_back();
				}

				if (participants.empty())
				{
					// Fall back to readable text image.
					const int targetW = min(640, maxDiagramWidth);
					const int targetH = max(140, min(460, d.lineCount * 18 + 60));
					Gdiplus::Bitmap gdiBmp(targetW, targetH, PixelFormat32bppARGB);
					Gdiplus::Graphics g(&gdiBmp);
					g.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
					g.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
					g.Clear(m_themeDark ? Gdiplus::Color(255, 36, 36, 36) : Gdiplus::Color(255, 250, 250, 250));
					Gdiplus::Pen pen(m_themeDark ? Gdiplus::Color(255, 80, 80, 80) : Gdiplus::Color(255, 210, 210, 210), 1.0f);
					g.DrawRectangle(&pen, 0, 0, targetW - 1, targetH - 1);
					auto font = CreateMonoFontForMermaid(12.0f, Gdiplus::FontStyleRegular);
					Gdiplus::SolidBrush brush(m_themeDark ? Gdiplus::Color(255, 220, 220, 220) : Gdiplus::Color(255, 60, 60, 60));
					Gdiplus::RectF layout(12.0f, 10.0f, (Gdiplus::REAL)targetW - 24.0f, (Gdiplus::REAL)targetH - 20.0f);
					std::wstring title = L"Mermaid(最小渲染)\nsequenceDiagram 渲染失败，已降级显示源码\n\n";
					std::wstring text = title + code;
					g.DrawString(text.c_str(), (INT)text.size(), font.get(), layout, nullptr, &brush);
					gdiBmp.GetHBITMAP(Gdiplus::Color(0, 0, 0, 0), &bmp);
					w = targetW;
					h = targetH;
				}
				else
				{
					// Layout: participants as columns, events (messages/notes) as rows.
					const int padding = 18;
					const int headerH = 44;
					const int footerH = 44;
					const int rowH = 40;
					const int colW = 180;
					const int maxH = 2600;
					const int eventCount = max(2, (int)events.size());

					int width = padding * 2 + (int)participants.size() * colW;
					int height = padding + headerH + eventCount * rowH + footerH + padding;
					if (height > maxH) height = maxH;
					if (width < 320) width = 320;
					if (width > maxDiagramWidth) width = maxDiagramWidth;
					const int lifelineTop = padding + headerH;
					const int footerTop = height - padding - footerH;
					const int lifelineBottom = max(lifelineTop + rowH, footerTop);

					Gdiplus::Bitmap gdiBmp(width, height, PixelFormat32bppARGB);
					Gdiplus::Graphics g(&gdiBmp);
					g.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
					g.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
					g.Clear(m_themeDark ? Gdiplus::Color(255, 36, 36, 36) : Gdiplus::Color(255, 250, 250, 250));
					Gdiplus::Pen borderPen(m_themeDark ? Gdiplus::Color(255, 90, 90, 90) : Gdiplus::Color(255, 200, 200, 200), 1.0f);
					Gdiplus::Pen linePen(m_themeDark ? Gdiplus::Color(255, 150, 150, 150) : Gdiplus::Color(255, 130, 130, 130), 1.2f);
					Gdiplus::Pen dashedLinePen(m_themeDark ? Gdiplus::Color(255, 160, 160, 160) : Gdiplus::Color(255, 120, 120, 120), 1.2f);
					dashedLinePen.SetDashStyle(Gdiplus::DashStyleDash);
					Gdiplus::SolidBrush textBrush(m_themeDark ? Gdiplus::Color(255, 230, 230, 230) : Gdiplus::Color(255, 50, 50, 50));
					Gdiplus::SolidBrush boxFill(m_themeDark ? Gdiplus::Color(255, 48, 48, 48) : Gdiplus::Color(255, 255, 255, 255));
					auto font = CreateUiFontForMermaid(12.0f, Gdiplus::FontStyleRegular);

					auto drawParticipantNode = [&](int x, int y, const SeqParticipant& p) {
						if (!p.actor)
						{
							Gdiplus::RectF box((Gdiplus::REAL)x + 10.0f, (Gdiplus::REAL)y, (Gdiplus::REAL)colW - 20.0f, (Gdiplus::REAL)headerH - 10.0f);
							g.FillRectangle(&boxFill, box);
							g.DrawRectangle(&borderPen, box);
							std::wstring name(p.label.GetString());
							Gdiplus::StringFormat sf;
							sf.SetAlignment(Gdiplus::StringAlignmentCenter);
							sf.SetLineAlignment(Gdiplus::StringAlignmentCenter);
							g.DrawString(name.c_str(), (INT)name.size(), font.get(), box, &sf, &textBrush);
							return;
						}

						// Simple actor icon + label.
						const Gdiplus::REAL cx = (Gdiplus::REAL)(x + colW / 2);
						const Gdiplus::REAL top = (Gdiplus::REAL)y + 2.0f;
						const Gdiplus::REAL r = 8.0f;
						g.DrawEllipse(&borderPen, cx - r, top, r * 2.0f, r * 2.0f);
						g.DrawLine(&borderPen, Gdiplus::PointF(cx, top + r * 2.0f), Gdiplus::PointF(cx, top + r * 2.0f + 16.0f));
						g.DrawLine(&borderPen, Gdiplus::PointF(cx - 12.0f, top + r * 2.0f + 7.0f), Gdiplus::PointF(cx + 12.0f, top + r * 2.0f + 7.0f));
						g.DrawLine(&borderPen, Gdiplus::PointF(cx, top + r * 2.0f + 16.0f), Gdiplus::PointF(cx - 10.0f, top + r * 2.0f + 26.0f));
						g.DrawLine(&borderPen, Gdiplus::PointF(cx, top + r * 2.0f + 16.0f), Gdiplus::PointF(cx + 10.0f, top + r * 2.0f + 26.0f));
						std::wstring name(p.label.GetString());
						Gdiplus::RectF rc(cx - (Gdiplus::REAL)colW / 2.0f + 10.0f, top + 28.0f, (Gdiplus::REAL)colW - 20.0f, 14.0f);
						Gdiplus::StringFormat sf;
						sf.SetAlignment(Gdiplus::StringAlignmentCenter);
						sf.SetLineAlignment(Gdiplus::StringAlignmentNear);
						g.DrawString(name.c_str(), (INT)name.size(), font.get(), rc, &sf, &textBrush);
					};

					// Participants header + footer + lifelines.
					for (int i = 0; i < (int)participants.size(); ++i)
					{
						int x = padding + i * colW;
						drawParticipantNode(x, padding, participants[i]);
						drawParticipantNode(x, footerTop, participants[i]);

						Gdiplus::REAL lx = (Gdiplus::REAL)(x + colW / 2);
						g.DrawLine(&dashedLinePen, Gdiplus::PointF(lx, (Gdiplus::REAL)lifelineTop), Gdiplus::PointF(lx, (Gdiplus::REAL)lifelineBottom));
					}

					// Control blocks (alt/else/end/opt/loop/par) visual layer.
					if (!blocks.empty())
					{
						const int blockInsetBase = 8;
						const int blockInsetPerDepth = 14;
						auto blockFont = CreateUiFontForMermaid(10.0f, Gdiplus::FontStyleBold);
						auto blockTextFont = CreateUiFontForMermaid(10.0f, Gdiplus::FontStyleRegular);

						auto chooseBlockStyle = [&](const CString& kind, Gdiplus::Color& fill, Gdiplus::Color& stroke) {
							if (m_themeDark)
							{
								if (kind.CompareNoCase(_T("alt")) == 0) { fill = Gdiplus::Color(54, 76, 124, 192); stroke = Gdiplus::Color(200, 124, 172, 240); return; }
								if (kind.CompareNoCase(_T("opt")) == 0) { fill = Gdiplus::Color(50, 76, 140, 94); stroke = Gdiplus::Color(200, 124, 204, 148); return; }
								if (kind.CompareNoCase(_T("loop")) == 0) { fill = Gdiplus::Color(50, 170, 128, 60); stroke = Gdiplus::Color(200, 220, 178, 104); return; }
								if (kind.CompareNoCase(_T("par")) == 0) { fill = Gdiplus::Color(50, 146, 118, 182); stroke = Gdiplus::Color(200, 196, 168, 228); return; }
								fill = Gdiplus::Color(44, 92, 92, 92);
								stroke = Gdiplus::Color(180, 150, 150, 150);
							}
							else
							{
								if (kind.CompareNoCase(_T("alt")) == 0) { fill = Gdiplus::Color(38, 96, 148, 212); stroke = Gdiplus::Color(210, 82, 134, 196); return; }
								if (kind.CompareNoCase(_T("opt")) == 0) { fill = Gdiplus::Color(34, 112, 186, 132); stroke = Gdiplus::Color(210, 88, 156, 108); return; }
								if (kind.CompareNoCase(_T("loop")) == 0) { fill = Gdiplus::Color(36, 236, 196, 128); stroke = Gdiplus::Color(210, 184, 140, 68); return; }
								if (kind.CompareNoCase(_T("par")) == 0) { fill = Gdiplus::Color(34, 176, 150, 224); stroke = Gdiplus::Color(210, 128, 100, 188); return; }
								fill = Gdiplus::Color(28, 188, 188, 188);
								stroke = Gdiplus::Color(180, 160, 160, 160);
							}
						};

						for (const auto& block : blocks)
						{
							int startRow = max(0, block.startRowIndex);
							int endRow = max(startRow + 1, block.endRowIndexExclusive);
							Gdiplus::REAL top = (Gdiplus::REAL)(lifelineTop + startRow * rowH + 2);
							Gdiplus::REAL bottom = (Gdiplus::REAL)(lifelineTop + endRow * rowH + 2);
							const Gdiplus::REAL clampBottom = (Gdiplus::REAL)(lifelineBottom - 2);
							if (top > clampBottom - 8.0f)
								continue;
							bottom = min(bottom, clampBottom);
							if (bottom <= top + 6.0f)
								bottom = top + 6.0f;

							int inset = blockInsetBase + block.depth * blockInsetPerDepth;
							Gdiplus::REAL left = (Gdiplus::REAL)(padding + inset);
							Gdiplus::REAL right = (Gdiplus::REAL)(padding + (int)participants.size() * colW - inset);
							if (right - left < 80.0f)
								continue;

							Gdiplus::RectF rc(left, top, right - left, bottom - top);
							Gdiplus::Color fillColor;
							Gdiplus::Color strokeColor;
							chooseBlockStyle(block.kind, fillColor, strokeColor);
							Gdiplus::SolidBrush blockFill(fillColor);
							Gdiplus::Pen blockPen(strokeColor, 1.0f);
							blockPen.SetDashStyle(Gdiplus::DashStyleDash);
							g.FillRectangle(&blockFill, rc);
							g.DrawRectangle(&blockPen, rc);

							CString blockTitle = block.kind;
							if (!block.label.IsEmpty())
							{
								blockTitle += _T(" ");
								blockTitle += block.label;
							}
							std::wstring titleW(blockTitle.GetString());
							Gdiplus::RectF titleRc(left + 6.0f, top + 3.0f, max(20.0f, right - left - 12.0f), 16.0f);
							g.DrawString(titleW.c_str(), (INT)titleW.size(), blockFont.get(), titleRc, nullptr, &textBrush);

							if (block.branches.size() > 1)
							{
								for (size_t bi = 1; bi < block.branches.size(); ++bi)
								{
									int splitRow = block.branches[bi].startRowIndex;
									if (splitRow <= startRow || splitRow >= endRow)
										continue;
									Gdiplus::REAL splitY = (Gdiplus::REAL)(lifelineTop + splitRow * rowH + 2);
									g.DrawLine(&blockPen, Gdiplus::PointF(left, splitY), Gdiplus::PointF(right, splitY));

									CString branchTitle = block.branches[bi].label;
									if (branchTitle.IsEmpty())
										branchTitle = (block.kind.CompareNoCase(_T("alt")) == 0) ? _T("else") : _T("and");
									std::wstring branchW(branchTitle.GetString());
									Gdiplus::RectF branchRc(left + 10.0f, splitY + 2.0f, max(20.0f, right - left - 20.0f), 15.0f);
									g.DrawString(branchW.c_str(), (INT)branchW.size(), blockTextFont.get(), branchRc, nullptr, &textBrush);
								}
							}
						}
					}

					Gdiplus::SolidBrush noteFill(m_themeDark ? Gdiplus::Color(220, 64, 72, 84) : Gdiplus::Color(230, 255, 244, 209));
					Gdiplus::Pen notePen(m_themeDark ? Gdiplus::Color(220, 136, 156, 186) : Gdiplus::Color(220, 196, 148, 92), 1.0f);

					for (int row = 0; row < (int)events.size(); ++row)
					{
						int rowTop = lifelineTop + row * rowH;
						if (rowTop > lifelineBottom - rowH)
							break;
						int y = rowTop + 12;

						const SeqEvent& ev = events[row];
						Gdiplus::REAL yy = (Gdiplus::REAL)y;
						if (ev.type == SeqEvent::Type::Message)
						{
							if (ev.index < 0 || ev.index >= (int)msgs.size())
								continue;
							const SeqMsg& msg = msgs[ev.index];
							int fromIdx = msg.fromIdx;
							int toIdx = msg.toIdx;
							if (fromIdx < 0 || toIdx < 0 || fromIdx >= (int)participants.size() || toIdx >= (int)participants.size())
								continue;
							Gdiplus::REAL x1 = (Gdiplus::REAL)(padding + fromIdx * colW + colW / 2);
							Gdiplus::REAL x2 = (Gdiplus::REAL)(padding + toIdx * colW + colW / 2);
							Gdiplus::Pen* msgPen = msg.dashed ? &dashedLinePen : &linePen;
							g.DrawLine(msgPen, Gdiplus::PointF(x1, yy), Gdiplus::PointF(x2, yy));
							// Arrow head.
							if (x2 >= x1)
							{
								g.DrawLine(msgPen, Gdiplus::PointF(x2 - 8, yy - 4), Gdiplus::PointF(x2, yy));
								g.DrawLine(msgPen, Gdiplus::PointF(x2 - 8, yy + 4), Gdiplus::PointF(x2, yy));
							}
							else
							{
								g.DrawLine(msgPen, Gdiplus::PointF(x2 + 8, yy - 4), Gdiplus::PointF(x2, yy));
								g.DrawLine(msgPen, Gdiplus::PointF(x2 + 8, yy + 4), Gdiplus::PointF(x2, yy));
							}

							CString msgText = msg.text;
							msgText.Trim();
							if (!msgText.IsEmpty())
							{
								std::wstring wt(msgText.GetString());
								Gdiplus::REAL textWidth = max(40.0f, (Gdiplus::REAL)abs((int)(x2 - x1)) - 16.0f);
								Gdiplus::RectF textRc(min(x1, x2) + 8.0f, yy - 18.0f, textWidth, 18.0f);
								Gdiplus::StringFormat sf;
								sf.SetAlignment(Gdiplus::StringAlignmentCenter);
								sf.SetLineAlignment(Gdiplus::StringAlignmentCenter);
								g.DrawString(wt.c_str(), (INT)wt.size(), font.get(), textRc, &sf, &textBrush);
							}
						}
						else
						{
							if (ev.index < 0 || ev.index >= (int)notes.size())
								continue;
							const SeqNote& note = notes[ev.index];
							if (note.p1 < 0 || note.p1 >= (int)participants.size() || note.p2 < 0 || note.p2 >= (int)participants.size())
								continue;

							Gdiplus::REAL noteLeft = 0.0f;
							Gdiplus::REAL noteWidth = 0.0f;
							if (note.kind == SeqNoteKind::Over)
							{
								int leftIdx = min(note.p1, note.p2);
								int rightIdx = max(note.p1, note.p2);
								noteLeft = (Gdiplus::REAL)(padding + leftIdx * colW + 14);
								Gdiplus::REAL noteRight = (Gdiplus::REAL)(padding + (rightIdx + 1) * colW - 14);
								noteWidth = max(60.0f, noteRight - noteLeft);
							}
							else
							{
								int pIdx = note.p1;
								Gdiplus::REAL centerX = (Gdiplus::REAL)(padding + pIdx * colW + colW / 2);
								noteWidth = (Gdiplus::REAL)min(220, max(120, colW - 24));
								if (note.kind == SeqNoteKind::LeftOf)
									noteLeft = centerX - noteWidth - 22.0f;
								else
									noteLeft = centerX + 22.0f;
								noteLeft = max((Gdiplus::REAL)(padding + 4), min(noteLeft, (Gdiplus::REAL)(width - padding - 4) - noteWidth));
							}

							Gdiplus::RectF noteRc(noteLeft, yy - 16.0f, noteWidth, 30.0f);
							g.FillRectangle(&noteFill, noteRc);
							g.DrawRectangle(&notePen, noteRc);
							if (!note.text.IsEmpty())
							{
								std::wstring wt(note.text.GetString());
								Gdiplus::StringFormat sf;
								sf.SetAlignment(Gdiplus::StringAlignmentNear);
								sf.SetLineAlignment(Gdiplus::StringAlignmentCenter);
								Gdiplus::RectF textRc = noteRc;
								textRc.X += 6.0f;
								textRc.Width = max(20.0f, textRc.Width - 12.0f);
								g.DrawString(wt.c_str(), (INT)wt.size(), font.get(), textRc, &sf, &textBrush);
							}
						}
					}

					gdiBmp.GetHBITMAP(Gdiplus::Color(0, 0, 0, 0), &bmp);
					w = width;
					h = height;
				}
			}
			else if (isGraphTD || isGraphLR)
			{
				for (const auto& line : lines)
				{
					if (startsWithNoCase(line, L"graph") || startsWithNoCase(line, L"flowchart"))
						continue;
					std::wstring s = line;
					// strip trailing ';'
					while (!s.empty() && (s.back() == L';' || s.back() == L' ' || s.back() == L'\t'))
						s.pop_back();
					s = TrimWForMermaid(s);
					if (s.empty())
						continue;
					std::wstring left;
					std::wstring right;
					std::wstring edgeLabel;
					if (!parseEdge(s, left, right, edgeLabel))
						continue;
					CString leftId, leftLabel, rightId, rightLabel;
					MermaidNodeShape leftShape = MermaidNodeShape::Rect;
					MermaidNodeShape rightShape = MermaidNodeShape::Rect;
					const bool leftExplicit = ((left.find(L'[') != std::wstring::npos && left.rfind(L']') != std::wstring::npos) ||
						(left.find(L'{') != std::wstring::npos && left.rfind(L'}') != std::wstring::npos) ||
						(left.find(L'(') != std::wstring::npos && left.rfind(L')') != std::wstring::npos));
					const bool rightExplicit = ((right.find(L'[') != std::wstring::npos && right.rfind(L']') != std::wstring::npos) ||
						(right.find(L'{') != std::wstring::npos && right.rfind(L'}') != std::wstring::npos) ||
						(right.find(L'(') != std::wstring::npos && right.rfind(L')') != std::wstring::npos));
					if (!parseNodeToken(left, leftId, leftLabel, leftShape))
						continue;
					if (!parseNodeToken(right, rightId, rightLabel, rightShape))
						continue;
					int li = getNodeIndex(leftId);
					int ri = getNodeIndex(rightId);
					// Do not let an ID-only token (e.g. "B") override a previously parsed label
					// from "B[xxx]". Only update when explicit [label] exists.
					if (leftExplicit)
					{
						nodes[li].label = leftLabel;
						nodes[li].shape = leftShape;
					}
					if (rightExplicit)
					{
						nodes[ri].label = rightLabel;
						nodes[ri].shape = rightShape;
					}
					edges.push_back({ li, ri });
					edgeLabels.push_back(edgeLabel.empty() ? CString() : CString(edgeLabel.c_str()));
				}
			}

			const bool supportedFlowchart = (isGraphTD || isGraphLR) && !edges.empty() && !nodes.empty();

			// If unsupported, render as a readable fallback (no hyperlink).
			if (!isSequenceDiagram && !supportedFlowchart)
			{
					const int targetW = min(640, maxDiagramWidth);
					const int targetH = max(140, min(460, d.lineCount * 18 + 60));
					Gdiplus::Bitmap gdiBmp(targetW, targetH, PixelFormat32bppARGB);
					Gdiplus::Graphics g(&gdiBmp);
					g.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
					g.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
					g.Clear(m_themeDark ? Gdiplus::Color(255, 36, 36, 36) : Gdiplus::Color(255, 250, 250, 250));
					Gdiplus::Pen pen(m_themeDark ? Gdiplus::Color(255, 80, 80, 80) : Gdiplus::Color(255, 210, 210, 210), 1.0f);
					g.DrawRectangle(&pen, 0, 0, targetW - 1, targetH - 1);
					auto font = CreateMonoFontForMermaid(12.0f, Gdiplus::FontStyleRegular);
					Gdiplus::SolidBrush brush(m_themeDark ? Gdiplus::Color(255, 220, 220, 220) : Gdiplus::Color(255, 60, 60, 60));
					Gdiplus::RectF layout(12.0f, 10.0f, (Gdiplus::REAL)targetW - 24.0f, (Gdiplus::REAL)targetH - 20.0f);
					std::wstring title = L"Mermaid(最小渲染回退)\n仅支持: graph/flowchart TD/TB/BT/LR/RL; A[文本]/A{判断}/A(圆角)/A([起止])/A((圆));  -->;  支持边标签 |xx| 或 -- xx -->\n\n";
					std::wstring text = title + code;
					g.DrawString(text.c_str(), (INT)text.size(), font.get(), layout, nullptr, &brush);
					gdiBmp.GetHBITMAP(Gdiplus::Color(0, 0, 0, 0), &bmp);
					w = targetW;
					h = targetH;
			}
			else if (!isSequenceDiagram)
			{
			// Simple layout: graph TD => top-down, graph LR => left-to-right.
			// IMPORTANT: Mermaid graphs may contain cycles; do NOT use longest-path relaxation
			// (it diverges on cycles). Use a bounded BFS layering instead.
			const int basePadding = 14;
			const int baseNodeW = 176;
			const int baseNodeH = 50;
			const int baseHGap = 56;
			const int baseVGap = 42;
			const int maxDiagramHeight = 3000;

			std::vector<std::vector<int>> outgoing(nodes.size());
			std::vector<std::vector<int>> incoming(nodes.size());
			std::vector<int> indegree(nodes.size(), 0);
			for (auto& e : edges)
			{
				int a = e.first;
				int b = e.second;
				if (a < 0 || b < 0 || a >= (int)nodes.size() || b >= (int)nodes.size())
					continue;
				outgoing[a].push_back(b);
				incoming[b].push_back(a);
				if (a != b)
					++indegree[b];
			}

			std::vector<int> level(nodes.size(), -1);
			std::vector<int> queue;
			queue.reserve(nodes.size());
			for (int i = 0; i < (int)nodes.size(); ++i)
			{
				if (indegree[i] == 0)
				{
					level[i] = 0;
					queue.push_back(i);
				}
			}
			if (queue.empty() && !nodes.empty())
			{
				level[0] = 0;
				queue.push_back(0);
			}

			for (size_t qi = 0; qi < queue.size(); ++qi)
			{
				int u = queue[qi];
				int nextLevel = (level[u] >= 0) ? (level[u] + 1) : 0;
				for (int v : outgoing[u])
				{
					if (v < 0 || v >= (int)nodes.size())
						continue;
					if (level[v] != -1)
						continue;
					level[v] = min(nextLevel, 50);
					queue.push_back(v);
				}
			}
			int maxAssigned = 0;
			for (int v : level)
				maxAssigned = max(maxAssigned, (v >= 0) ? v : 0);
			for (int i = 0; i < (int)nodes.size(); ++i)
			{
				if (level[i] == -1)
				{
					level[i] = ++maxAssigned;
				}
			}
			int maxLevel = 0;
			for (int v : level) maxLevel = max(maxLevel, v);
			std::vector<std::vector<int>> buckets(maxLevel + 1);
			for (int i = 0; i < (int)nodes.size(); ++i)
				buckets[level[i]].push_back(i);

			// Improve ordering inside each layer (tree-like) using a lightweight barycenter heuristic.
			// This helps multi-branch graphs look more like a (binary) tree, reducing crossings.
			if (maxLevel > 0 && nodes.size() > 1)
			{
				std::vector<int> posInLevel(nodes.size(), 0);
				auto updatePos = [&](int lv) {
					if (lv < 0 || lv > maxLevel)
						return;
					auto& b = buckets[lv];
					for (int i = 0; i < (int)b.size(); ++i)
						posInLevel[b[i]] = i;
				};
				for (int lv = 0; lv <= maxLevel; ++lv)
					updatePos(lv);

				auto parentBary = [&](int node) -> double {
					double sum = 0.0;
					int cnt = 0;
					for (int p : incoming[node])
					{
						if (p < 0 || p >= (int)posInLevel.size())
							continue;
						sum += (double)posInLevel[p];
						++cnt;
					}
					return (cnt > 0) ? (sum / (double)cnt) : (double)posInLevel[node];
				};
				auto childBary = [&](int node) -> double {
					double sum = 0.0;
					int cnt = 0;
					for (int c : outgoing[node])
					{
						if (c < 0 || c >= (int)posInLevel.size())
							continue;
						sum += (double)posInLevel[c];
						++cnt;
					}
					return (cnt > 0) ? (sum / (double)cnt) : (double)posInLevel[node];
				};

				for (int iter = 0; iter < 2; ++iter)
				{
					// Top-down sweep: order nodes based on parents.
					for (int lv = 1; lv <= maxLevel; ++lv)
					{
						auto& b = buckets[lv];
						std::stable_sort(b.begin(), b.end(), [&](int a, int c) {
							double ba = parentBary(a);
							double bc = parentBary(c);
							if (ba == bc)
								return posInLevel[a] < posInLevel[c];
							return ba < bc;
						});
						updatePos(lv);
					}
					// Bottom-up sweep: order nodes based on children.
					for (int lv = maxLevel - 1; lv >= 0; --lv)
					{
						auto& b = buckets[lv];
						std::stable_sort(b.begin(), b.end(), [&](int a, int c) {
							double ba = childBary(a);
							double bc = childBary(c);
							if (ba == bc)
								return posInLevel[a] < posInLevel[c];
							return ba < bc;
						});
						updatePos(lv);
					}
				}
			}

			d.nodeRects.clear();
			d.nodeLabels.clear();
			d.edges.clear();
			d.edgeLabels.clear();
			d.nodeRects.resize(nodes.size());
			d.nodeLabels.resize(nodes.size());
			d.nodeShapes.clear();
			d.nodeShapes.resize(nodes.size());
			for (int i = 0; i < (int)nodes.size(); ++i)
			{
				d.nodeLabels[i] = nodes[i].label;
				// Start/End nodes (source/sink) are rendered as rounded rectangles.
				// This keeps visual style consistent with common Mermaid flowcharts.
				MermaidNodeShape shape = nodes[i].shape;
				const bool isStartNode = (i >= 0 && i < (int)indegree.size() && indegree[i] == 0);
				const bool isEndNode = (i >= 0 && i < (int)outgoing.size() && outgoing[i].empty());
				if (isStartNode || isEndNode)
					shape = MermaidNodeShape::RoundRect;
				d.nodeShapes[i] = shape;
			}
			for (size_t ei = 0; ei < edges.size(); ++ei)
			{
				d.edges.push_back({ edges[ei].first, edges[ei].second });
				d.edgeLabels.push_back((ei < edgeLabels.size()) ? edgeLabels[ei] : CString());
			}

			int maxCols = 0;
			for (auto& b : buckets) maxCols = max(maxCols, (int)b.size());

			// Natural size.
			int naturalW = 0;
			int naturalH = 0;
			if (isGraphLR)
			{
				// In LR mode, "level" is X (column), and bucket row is Y.
				naturalW = basePadding * 2 + (maxLevel + 1) * baseNodeW + maxLevel * baseHGap;
				naturalH = basePadding * 2 + maxCols * baseNodeH + max(0, maxCols - 1) * baseVGap;
			}
			else
			{
				// TD mode: "level" is Y (row), and bucket row is X.
				naturalW = basePadding * 2 + maxCols * baseNodeW + max(0, maxCols - 1) * baseHGap;
				naturalH = basePadding * 2 + (maxLevel + 1) * baseNodeH + maxLevel * baseVGap;
			}
			if (naturalW < 240) naturalW = 240;
			if (naturalH < 120) naturalH = 120;

			// Scale policy:
			// - Prefer preserving node size.
			// - Only scale down when the natural size exceeds bounds.
			double scale = 1.0;
			if (naturalW > 0 && naturalH > 0)
			{
				double fitW = (maxDiagramWidth > 0) ? ((double)maxDiagramWidth / (double)naturalW) : 1.0;
				double fitH = (maxDiagramHeight > 0) ? ((double)maxDiagramHeight / (double)naturalH) : 1.0;
				double fit = min(fitW, fitH);
				if (fit < 1.0)
				{
					scale = max(0.25, fit);
				}
			}

			const int padding = max(6, (int)(basePadding * scale));
			int nodeW = max(64, (int)(baseNodeW * scale));
			int nodeH = max(22, (int)(baseNodeH * scale));
			int hGap = max(14, (int)(baseHGap * scale));
			int vGap = max(14, (int)(baseVGap * scale));

			int width = 0;
			int height = 0;
			if (isGraphLR)
			{
				width = padding * 2 + (maxLevel + 1) * nodeW + maxLevel * hGap;
				height = padding * 2 + maxCols * nodeH + max(0, maxCols - 1) * vGap;
			}
			else
			{
				width = padding * 2 + maxCols * nodeW + max(0, maxCols - 1) * hGap;
				height = padding * 2 + (maxLevel + 1) * nodeH + maxLevel * vGap;
			}
			if (width < 240) width = 240;
			if (height < 120) height = 120;
			// Avoid cropping: if layout still exceeds bounds (e.g. because of min node sizes),
			// proportionally shrink the layout rather than clamping bitmap size.
			if (width > maxDiagramWidth || height > maxDiagramHeight)
			{
				double shrinkW = (width > 0) ? ((double)maxDiagramWidth / (double)width) : 1.0;
				double shrinkH = (height > 0) ? ((double)maxDiagramHeight / (double)height) : 1.0;
				double shrink = min(shrinkW, shrinkH);
				shrink = min(1.0, max(0.2, shrink));
				nodeW = max(44, (int)(nodeW * shrink));
				nodeH = max(18, (int)(nodeH * shrink));
				hGap = max(10, (int)(hGap * shrink));
				vGap = max(10, (int)(vGap * shrink));
				int padding2 = max(6, (int)(padding * shrink));
				if (isGraphLR)
				{
					width = padding2 * 2 + (maxLevel + 1) * nodeW + maxLevel * hGap;
					height = padding2 * 2 + maxCols * nodeH + max(0, maxCols - 1) * vGap;
				}
				else
				{
					width = padding2 * 2 + maxCols * nodeW + max(0, maxCols - 1) * hGap;
					height = padding2 * 2 + (maxLevel + 1) * nodeH + maxLevel * vGap;
				}
				if (width > maxDiagramWidth) width = maxDiagramWidth;
				if (height > maxDiagramHeight) height = maxDiagramHeight;
				// Recompute rects with the shrunken padding.
				for (int lv = 0; lv <= maxLevel; ++lv)
				{
					auto& b = buckets[lv];
					for (int row = 0; row < (int)b.size(); ++row)
					{
						int idx = b[row];
						int x = 0;
						int y = 0;
						if (isGraphLR)
						{
							int col = isReverseDir ? (maxLevel - lv) : lv;
							x = padding2 + col * (nodeW + hGap);
							y = padding2 + row * (nodeH + vGap);
						}
						else
						{
							x = padding2 + row * (nodeW + hGap);
							int r = isReverseDir ? (maxLevel - lv) : lv;
							y = padding2 + r * (nodeH + vGap);
						}
						d.nodeRects[idx] = CRect(x, y, x + nodeW, y + nodeH);
					}
				}
			}
			d.diagramWidth = width;
			d.diagramHeight = height;

			for (int lv = 0; lv <= maxLevel; ++lv)
			{
				auto& b = buckets[lv];
				for (int row = 0; row < (int)b.size(); ++row)
				{
					int idx = b[row];
					int x = 0;
					int y = 0;
					if (isGraphLR)
					{
						// level -> x; within level bucket -> y
						int col = isReverseDir ? (maxLevel - lv) : lv;
						x = padding + col * (nodeW + hGap);
						y = padding + row * (nodeH + vGap);
					}
					else
					{
						// level -> y; within level bucket -> x
						x = padding + row * (nodeW + hGap);
						int r = isReverseDir ? (maxLevel - lv) : lv;
						y = padding + r * (nodeH + vGap);
					}
					d.nodeRects[idx] = CRect(x, y, x + nodeW, y + nodeH);
				}
			}

			Gdiplus::Bitmap gdiBmp(d.diagramWidth, d.diagramHeight, PixelFormat32bppARGB);
			Gdiplus::Graphics g(&gdiBmp);
			g.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
			g.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
			const Gdiplus::Color bg(m_themeDark ? Gdiplus::Color(255, 36, 36, 36) : Gdiplus::Color(255, 250, 250, 250));
			g.Clear(bg);
			const Gdiplus::Color nodeFillColor(m_themeDark ? Gdiplus::Color(255, 48, 48, 48) : Gdiplus::Color(255, 255, 255, 255));
			const Gdiplus::Color nodeBorderColor(m_themeDark ? Gdiplus::Color(255, 90, 90, 90) : Gdiplus::Color(255, 200, 200, 200));
			const Gdiplus::Color edgeColor(m_themeDark ? Gdiplus::Color(255, 150, 150, 150) : Gdiplus::Color(255, 130, 130, 130));
			const Gdiplus::Color textColor(m_themeDark ? Gdiplus::Color(255, 230, 230, 230) : Gdiplus::Color(255, 50, 50, 50));
			const Gdiplus::Color edgeLabelColor(textColor);

			g.Clear(bg);
			Gdiplus::Pen borderPen(nodeBorderColor, 1.0f);
			Gdiplus::Pen edgePen(edgeColor, 1.2f);
			Gdiplus::SolidBrush textBrush(textColor);
			Gdiplus::SolidBrush edgeLabelBrush(edgeLabelColor);
			Gdiplus::SolidBrush nodeFill(nodeFillColor);
			const float fontPx = (float)max(10.0, min(14.0, 13.0 * scale));
			auto font = CreateUiFontForMermaid(fontPx, Gdiplus::FontStyleRegular);
			auto edgeLabelFont = CreateUiFontForMermaid(max(9.0f, fontPx - 1.0f), Gdiplus::FontStyleRegular);

			auto drawArrowHead = [&](const Gdiplus::PointF& from, const Gdiplus::PointF& tip) {
				Gdiplus::REAL vx = tip.X - from.X;
				Gdiplus::REAL vy = tip.Y - from.Y;
				Gdiplus::REAL len = (Gdiplus::REAL)sqrt(vx * vx + vy * vy);
				if (len < 0.01f)
					return;
				vx /= len;
				vy /= len;
				const Gdiplus::REAL back = 8.0f;
				const Gdiplus::REAL wing = 4.0f;
				Gdiplus::PointF base(tip.X - vx * back, tip.Y - vy * back);
				Gdiplus::PointF perp(-vy, vx);
				Gdiplus::PointF p1(base.X + perp.X * wing, base.Y + perp.Y * wing);
				Gdiplus::PointF p2(base.X - perp.X * wing, base.Y - perp.Y * wing);
				g.DrawLine(&edgePen, p1, tip);
				g.DrawLine(&edgePen, p2, tip);
			};

			auto inflatedNodeRectF = [&](const CRect& rc, float pad) -> Gdiplus::RectF {
				return Gdiplus::RectF((Gdiplus::REAL)rc.left - pad,
					(Gdiplus::REAL)rc.top - pad,
					(Gdiplus::REAL)rc.Width() + pad * 2.0f,
					(Gdiplus::REAL)rc.Height() + pad * 2.0f);
			};

			auto segmentIntersectsRectF = [&](const Gdiplus::PointF& a,
				const Gdiplus::PointF& b,
				const Gdiplus::RectF& r) -> bool {
				// Quick reject.
				const float minX = min(a.X, b.X);
				const float maxX = max(a.X, b.X);
				const float minY = min(a.Y, b.Y);
				const float maxY = max(a.Y, b.Y);
				if (maxX < r.X || minX > r.X + r.Width || maxY < r.Y || minY > r.Y + r.Height)
					return false;

				auto contains = [&](const Gdiplus::PointF& p) {
					return p.X >= r.X && p.X <= r.X + r.Width && p.Y >= r.Y && p.Y <= r.Y + r.Height;
				};
				if (contains(a) || contains(b))
					return true;

				auto cross = [&](const Gdiplus::PointF& p, const Gdiplus::PointF& q, const Gdiplus::PointF& s) -> double {
					return (double)(q.X - p.X) * (double)(s.Y - p.Y) - (double)(q.Y - p.Y) * (double)(s.X - p.X);
				};
				auto onSeg = [&](const Gdiplus::PointF& p, const Gdiplus::PointF& q, const Gdiplus::PointF& s) -> bool {
					return min(p.X, q.X) - 0.01f <= s.X && s.X <= max(p.X, q.X) + 0.01f &&
						min(p.Y, q.Y) - 0.01f <= s.Y && s.Y <= max(p.Y, q.Y) + 0.01f;
				};
				auto segIntersects = [&](const Gdiplus::PointF& p1, const Gdiplus::PointF& p2,
					const Gdiplus::PointF& q1, const Gdiplus::PointF& q2) -> bool {
					double d1 = cross(p1, p2, q1);
					double d2 = cross(p1, p2, q2);
					double d3 = cross(q1, q2, p1);
					double d4 = cross(q1, q2, p2);
					if (((d1 > 0 && d2 < 0) || (d1 < 0 && d2 > 0)) && ((d3 > 0 && d4 < 0) || (d3 < 0 && d4 > 0)))
						return true;
					if (fabs(d1) < 1e-6 && onSeg(p1, p2, q1)) return true;
					if (fabs(d2) < 1e-6 && onSeg(p1, p2, q2)) return true;
					if (fabs(d3) < 1e-6 && onSeg(q1, q2, p1)) return true;
					if (fabs(d4) < 1e-6 && onSeg(q1, q2, p2)) return true;
					return false;
				};

				Gdiplus::PointF tl(r.X, r.Y);
				Gdiplus::PointF tr(r.X + r.Width, r.Y);
				Gdiplus::PointF br(r.X + r.Width, r.Y + r.Height);
				Gdiplus::PointF bl(r.X, r.Y + r.Height);
				return segIntersects(a, b, tl, tr) || segIntersects(a, b, tr, br) ||
					segIntersects(a, b, br, bl) || segIntersects(a, b, bl, tl);
			};

			auto deflateRectF = [&](Gdiplus::RectF r, float d) -> Gdiplus::RectF {
				r.X += d;
				r.Y += d;
				r.Width = max(0.0f, r.Width - d * 2.0f);
				r.Height = max(0.0f, r.Height - d * 2.0f);
				return r;
			};

			auto segmentIntersectsAnyNodeObstacle = [&](const Gdiplus::PointF& a,
				const Gdiplus::PointF& b,
				int fromNode,
				int toNode,
				float pad) -> bool {
				// Do not allow edges to pass through node bounding boxes.
				// For endpoints, use a smaller *interior* rect (no inflation) so:
				// - the segment can start/end near the border
				// - but it is still forbidden to cut through the node interior
				for (int i = 0; i < (int)d.nodeRects.size(); ++i)
				{
					const bool isEndpoint = (i == fromNode || i == toNode);
					Gdiplus::RectF r = inflatedNodeRectF(d.nodeRects[i], isEndpoint ? 0.0f : pad);
					if (isEndpoint)
						r = deflateRectF(r, 3.0f);
					if (segmentIntersectsRectF(a, b, r))
						return true;
				}
				return false;
			};

			auto segmentIntersectsAnyNodeGeneral = [&](const Gdiplus::PointF& a,
				const Gdiplus::PointF& b,
				int ignore1,
				int ignore2,
				float pad) -> bool {
				for (int i = 0; i < (int)d.nodeRects.size(); ++i)
				{
					if (i == ignore1 || i == ignore2)
						continue;
					Gdiplus::RectF r = inflatedNodeRectF(d.nodeRects[i], pad);
					if (segmentIntersectsRectF(a, b, r))
						return true;
				}
				return false;
			};

			auto segmentIntersectsAnyNode = [&](const Gdiplus::PointF& a,
				const Gdiplus::PointF& b,
				int ignore1,
				int ignore2,
				float pad) -> bool {
				const bool horiz = fabs(a.Y - b.Y) < 0.01f;
				const bool vert = fabs(a.X - b.X) < 0.01f;
				if (!horiz && !vert)
					return false;
				float minX = min(a.X, b.X);
				float maxX = max(a.X, b.X);
				float minY = min(a.Y, b.Y);
				float maxY = max(a.Y, b.Y);
				for (int i = 0; i < (int)d.nodeRects.size(); ++i)
				{
					if (i == ignore1 || i == ignore2)
						continue;
					Gdiplus::RectF r = inflatedNodeRectF(d.nodeRects[i], pad);
					if (horiz)
					{
						if (a.Y >= r.Y && a.Y <= r.Y + r.Height)
						{
							if (!(maxX < r.X || minX > r.X + r.Width))
								return true;
						}
					}
					else
					{
						if (a.X >= r.X && a.X <= r.X + r.Width)
						{
							if (!(maxY < r.Y || minY > r.Y + r.Height))
								return true;
						}
					}
				}
				return false;
			};

			auto routeEdgeOrthogonal = [&](int a, int b, const CRect& ra, const CRect& rb,
				const Gdiplus::PointF& pStart, const Gdiplus::PointF& pEnd) -> std::vector<Gdiplus::PointF> {
				const float margin = 10.0f;
				const float pad = 10.0f;
				std::vector<Gdiplus::PointF> pts;
				pts.reserve(6);
				pts.push_back(pStart);

				if (isGraphLR)
				{
					float dir = (pEnd.X >= pStart.X) ? 1.0f : -1.0f;
					Gdiplus::PointF startOut(pStart.X + dir * margin, pStart.Y);
					Gdiplus::PointF endOut(pEnd.X - dir * margin, pEnd.Y);
					auto isClearRouteY = [&](float y) -> bool {
						// Check all segments (including short out segments) against nodes.
						Gdiplus::PointF s0 = pStart;
						Gdiplus::PointF s1 = startOut;
						Gdiplus::PointF s2(startOut.X, y);
						Gdiplus::PointF s3(endOut.X, y);
						Gdiplus::PointF s4 = endOut;
						Gdiplus::PointF s5 = pEnd;
						if (segmentIntersectsAnyNodeObstacle(s0, s1, a, b, pad)) return false;
						if (segmentIntersectsAnyNodeObstacle(s1, s2, a, b, pad)) return false;
						if (segmentIntersectsAnyNodeObstacle(s2, s3, a, b, pad)) return false;
						if (segmentIntersectsAnyNodeObstacle(s3, s4, a, b, pad)) return false;
						if (segmentIntersectsAnyNodeObstacle(s4, s5, a, b, pad)) return false;
						return true;
					};

					std::vector<float> candidates;
					candidates.reserve(16);
					candidates.push_back((startOut.Y + endOut.Y) / 2.0f);
					candidates.push_back(startOut.Y);
					candidates.push_back(endOut.Y);
					// Prefer routing through gaps between rows.
					for (const auto& rc : d.nodeRects)
					{
						candidates.push_back((float)rc.top - pad - 6.0f);
						candidates.push_back((float)rc.bottom + pad + 6.0f);
					}
					float yRoute = (startOut.Y + endOut.Y) / 2.0f;
					bool found = false;
					double bestLen = 1e18;
					auto scoreY = [&](float y) -> double {
						// Manhattan length of the 5-segment route.
						return (double)fabs(pStart.X - startOut.X) + (double)fabs(pStart.Y - startOut.Y) +
							(double)fabs(startOut.Y - y) + (double)fabs(startOut.X - endOut.X) +
							(double)fabs(endOut.Y - y) + (double)fabs(pEnd.X - endOut.X) +
							(double)fabs(pEnd.Y - endOut.Y);
					};
					for (float y : candidates)
					{
						if (y < 6.0f || y > (float)d.diagramHeight - 6.0f)
							continue;
						if (!isClearRouteY(y))
							continue;
						double len = scoreY(y);
						if (len < bestLen)
						{
							bestLen = len;
							yRoute = y;
							found = true;
						}
					}
					if (!found)
					{
						// Last resort: scan from top/bottom towards center.
						const float topBound = pad + 6.0f;
						const float bottomBound = (float)d.diagramHeight - (pad + 6.0f);
						for (int i = 0; i < 40; ++i)
						{
							float yTop = topBound + (float)i * 20.0f;
							float yBottom = bottomBound - (float)i * 20.0f;
							if (yTop <= bottomBound && isClearRouteY(yTop))
							{
								double len = scoreY(yTop);
								if (len < bestLen)
								{
									bestLen = len;
									yRoute = yTop;
									found = true;
								}
							}
							if (yBottom >= topBound && isClearRouteY(yBottom))
							{
								double len = scoreY(yBottom);
								if (len < bestLen)
								{
									bestLen = len;
									yRoute = yBottom;
									found = true;
								}
							}
						}
						if (!found)
							yRoute = max(topBound, min(bottomBound, (startOut.Y + endOut.Y) / 2.0f));
					}

					pts.push_back(startOut);
					pts.push_back(Gdiplus::PointF(startOut.X, yRoute));
					pts.push_back(Gdiplus::PointF(endOut.X, yRoute));
					pts.push_back(endOut);
					pts.push_back(pEnd);
				}
				else
				{
					float dir = (pEnd.Y >= pStart.Y) ? 1.0f : -1.0f;
					Gdiplus::PointF startOut(pStart.X, pStart.Y + dir * margin);
					Gdiplus::PointF endOut(pEnd.X, pEnd.Y - dir * margin);
					auto isClearRouteX = [&](float x) -> bool {
						// Check all segments (including short out segments) against nodes.
						Gdiplus::PointF s0 = pStart;
						Gdiplus::PointF s1 = startOut;
						Gdiplus::PointF s2(x, startOut.Y);
						Gdiplus::PointF s3(x, endOut.Y);
						Gdiplus::PointF s4 = endOut;
						Gdiplus::PointF s5 = pEnd;
						if (segmentIntersectsAnyNodeObstacle(s0, s1, a, b, pad)) return false;
						if (segmentIntersectsAnyNodeObstacle(s1, s2, a, b, pad)) return false;
						if (segmentIntersectsAnyNodeObstacle(s2, s3, a, b, pad)) return false;
						if (segmentIntersectsAnyNodeObstacle(s3, s4, a, b, pad)) return false;
						if (segmentIntersectsAnyNodeObstacle(s4, s5, a, b, pad)) return false;
						return true;
					};

					std::vector<float> candidates;
					candidates.reserve(16);
					candidates.push_back((startOut.X + endOut.X) / 2.0f);
					candidates.push_back(startOut.X);
					candidates.push_back(endOut.X);
					// Prefer routing through gaps between columns.
					for (const auto& rc : d.nodeRects)
					{
						candidates.push_back((float)rc.left - pad - 6.0f);
						candidates.push_back((float)rc.right + pad + 6.0f);
					}
					float xRoute = (startOut.X + endOut.X) / 2.0f;
					bool found = false;
					double bestLen = 1e18;
					auto scoreX = [&](float x) -> double {
						return (double)fabs(pStart.X - startOut.X) + (double)fabs(pStart.Y - startOut.Y) +
							(double)fabs(startOut.X - x) + (double)fabs(startOut.Y - endOut.Y) +
							(double)fabs(endOut.X - x) + (double)fabs(pEnd.X - endOut.X) +
							(double)fabs(pEnd.Y - endOut.Y);
					};
					for (float x : candidates)
					{
						if (x < 6.0f || x > (float)d.diagramWidth - 6.0f)
							continue;
						if (!isClearRouteX(x))
							continue;
						double len = scoreX(x);
						if (len < bestLen)
						{
							bestLen = len;
							xRoute = x;
							found = true;
						}
					}
					if (!found)
					{
						// Last resort: scan from left/right towards center.
						const float leftBound = pad + 6.0f;
						const float rightBound = (float)d.diagramWidth - (pad + 6.0f);
						for (int i = 0; i < 40; ++i)
						{
							float xLeft = leftBound + (float)i * 20.0f;
							float xRight = rightBound - (float)i * 20.0f;
							if (xLeft <= rightBound && isClearRouteX(xLeft))
							{
								double len = scoreX(xLeft);
								if (len < bestLen)
								{
									bestLen = len;
									xRoute = xLeft;
									found = true;
								}
							}
							if (xRight >= leftBound && isClearRouteX(xRight))
							{
								double len = scoreX(xRight);
								if (len < bestLen)
								{
									bestLen = len;
									xRoute = xRight;
									found = true;
								}
							}
						}
						if (!found)
							xRoute = max(leftBound, min(rightBound, (startOut.X + endOut.X) / 2.0f));
					}

					pts.push_back(startOut);
					pts.push_back(Gdiplus::PointF(xRoute, startOut.Y));
					pts.push_back(Gdiplus::PointF(xRoute, endOut.Y));
					pts.push_back(Gdiplus::PointF(endOut.X, endOut.Y));
					pts.push_back(pEnd);
				}

				// Remove consecutive duplicates.
				std::vector<Gdiplus::PointF> out;
				out.reserve(pts.size());
				for (const auto& pt : pts)
				{
					if (!out.empty() && fabs(out.back().X - pt.X) < 0.01f && fabs(out.back().Y - pt.Y) < 0.01f)
						continue;
					out.push_back(pt);
				}
				return out;
			};

			auto routeEdgePreferStraight = [&](int a, int b, const CRect& ra, const CRect& rb,
				const Gdiplus::PointF& pStart, const Gdiplus::PointF& pEnd) -> std::vector<Gdiplus::PointF> {
				const float pad = 10.0f;
				auto compress = [&](std::vector<Gdiplus::PointF> pts) -> std::vector<Gdiplus::PointF> {
					// Remove consecutive duplicates and colinear intermediate points.
					std::vector<Gdiplus::PointF> out;
					out.reserve(pts.size());
					for (const auto& pt : pts)
					{
						if (!out.empty() && fabs(out.back().X - pt.X) < 0.01f && fabs(out.back().Y - pt.Y) < 0.01f)
							continue;
						out.push_back(pt);
					}
					if (out.size() <= 2)
						return out;
					std::vector<Gdiplus::PointF> out2;
					out2.reserve(out.size());
					out2.push_back(out[0]);
					for (size_t i = 1; i + 1 < out.size(); ++i)
					{
						auto p0 = out2.back();
						auto p1 = out[i];
						auto p2 = out[i + 1];
						const bool colinear = (fabs((p1.X - p0.X) * (p2.Y - p0.Y) - (p1.Y - p0.Y) * (p2.X - p0.X)) < 0.01f);
						if (colinear)
							continue;
						out2.push_back(p1);
					}
					out2.push_back(out.back());
					return out2;
				};

				// 1) Try a direct straight line (allow diagonal) when it doesn't pass through nodes.
				if (!segmentIntersectsAnyNodeObstacle(pStart, pEnd, a, b, pad))
					return compress({ pStart, pEnd });

				// 2) Try a single-bend polyline ("折") if direct isn't possible.
				auto routeIsClear = [&](const std::vector<Gdiplus::PointF>& pts) -> bool {
					for (size_t i = 1; i < pts.size(); ++i)
						if (segmentIntersectsAnyNodeObstacle(pts[i - 1], pts[i], a, b, pad))
							return false;
					return true;
				};

				// Add a small corner offset so the bend doesn't land inside the node padding.
				Gdiplus::PointF mid1(pStart.X, pEnd.Y);
				Gdiplus::PointF mid2(pEnd.X, pStart.Y);
				std::vector<Gdiplus::PointF> c1 = { pStart, mid1, pEnd };
				std::vector<Gdiplus::PointF> c2 = { pStart, mid2, pEnd };
				if (routeIsClear(c1))
					return compress(c1);
				if (routeIsClear(c2))
					return compress(c2);

				// 3) Fallback: detour routing through gaps.
				return compress(routeEdgeOrthogonal(a, b, ra, rb, pStart, pEnd));
			};

			struct EdgeRenderInfo
			{
				int fromNode = -1;
				int toNode = -1;
				std::vector<Gdiplus::PointF> pts;
				CString label;
			};
			std::vector<EdgeRenderInfo> edgeInfos;
			edgeInfos.reserve(d.edges.size());

			auto rectCenter = [&](const CRect& rc) -> Gdiplus::PointF {
				return Gdiplus::PointF((Gdiplus::REAL)rc.left + (Gdiplus::REAL)rc.Width() / 2.0f,
					(Gdiplus::REAL)rc.top + (Gdiplus::REAL)rc.Height() / 2.0f);
			};

			enum PortSide
			{
				PortLeft = 0,
				PortRight,
				PortTop,
				PortBottom,
			};

			struct PortCandidate
			{
				Gdiplus::PointF pt;
				int score = 0;
				PortSide side = PortRight;
			};

			auto buildPortCandidates = [&](const CRect& r, const Gdiplus::PointF& otherCenter) -> std::vector<PortCandidate> {
				// IMPORTANT: endpoints must attach ONLY at side midpoints (上/下/左/右边框中点).
				// So we provide exactly four candidates.
				const float xL = (float)r.left;
				const float xR = (float)r.right;
				const float yT = (float)r.top;
				const float yB = (float)r.bottom;
				const float xC = (float)r.left + (float)r.Width() / 2.0f;
				const float yC = (float)r.top + (float)r.Height() / 2.0f;
				const auto c = rectCenter(r);
				const float dx = otherCenter.X - c.X;
				const float dy = otherCenter.Y - c.Y;

				auto add = [&](float x, float y, int score, PortSide side) {
					PortCandidate pc;
					pc.pt = Gdiplus::PointF(x, y);
					pc.score = score;
					pc.side = side;
					return pc;
				};

				// Build an ordered set of candidates (best-first).
				std::vector<PortCandidate> out;
				out.reserve(4);
				const bool horizontalDominant = (fabs(dx) > fabs(dy));
				if (horizontalDominant)
				{
					if (dx >= 0)
					{
						out.push_back(add(xR, yC, 0, PortRight));
						out.push_back(add(xC, yB, 10, PortBottom));
						out.push_back(add(xC, yT, 10, PortTop));
						out.push_back(add(xL, yC, 20, PortLeft));
					}
					else
					{
						out.push_back(add(xL, yC, 0, PortLeft));
						out.push_back(add(xC, yB, 10, PortBottom));
						out.push_back(add(xC, yT, 10, PortTop));
						out.push_back(add(xR, yC, 20, PortRight));
					}
				}
				else
				{
					if (dy >= 0)
					{
						out.push_back(add(xC, yB, 0, PortBottom));
						out.push_back(add(xR, yC, 10, PortRight));
						out.push_back(add(xL, yC, 10, PortLeft));
						out.push_back(add(xC, yT, 20, PortTop));
					}
					else
					{
						out.push_back(add(xC, yT, 0, PortTop));
						out.push_back(add(xR, yC, 10, PortRight));
						out.push_back(add(xL, yC, 10, PortLeft));
						out.push_back(add(xC, yB, 20, PortBottom));
					}
				}
				return out;
			};

			auto chooseBestRoute = [&](int fromNode,
				int toNode,
				const CRect& ra,
				const CRect& rb,
				const CString& edgeLabel) -> std::vector<Gdiplus::PointF> {
				const auto ca = rectCenter(ra);
				const auto cb = rectCenter(rb);
				auto portsA = buildPortCandidates(ra, cb);
				auto portsB = buildPortCandidates(rb, ca);

				// Heuristic: if there's an edge label, prefer routes that have at least
				// one segment long enough to place it without covering the line.
				float labelMinSegment = 0.0f;
				if (!edgeLabel.IsEmpty())
				{
					std::wstring wl(edgeLabel.GetString());
					Gdiplus::RectF measureRc(0, 0, 1000.0f, 200.0f);
					Gdiplus::RectF bound;
					g.MeasureString(wl.c_str(), (INT)wl.size(), edgeLabelFont.get(), measureRc, &bound);
					labelMinSegment = bound.Width + 18.0f;
				}

				// Avoid huge detours: keep routes reasonably close to the two nodes.
				CRect bbox(min(ra.left, rb.left), min(ra.top, rb.top), max(ra.right, rb.right), max(ra.bottom, rb.bottom));
				bbox.InflateRect(120, 120);

				auto routeCost = [&](const std::vector<Gdiplus::PointF>& pts, int portScore) -> double {
					if (pts.size() < 2)
						return 1e18;
					double cost = 0.0;
					double maxSeg = 0.0;
					for (size_t i = 1; i < pts.size(); ++i)
					{
						double dx = (double)(pts[i].X - pts[i - 1].X);
						double dy = (double)(pts[i].Y - pts[i - 1].Y);
						double seg = fabs(dx) + fabs(dy);
						cost += seg;
						maxSeg = max(maxSeg, seg);
					}
					const int bends = max(0, (int)pts.size() - 2);
					// Strict priority: straight > single-bend > multi-bend.
					if (bends == 0)
						cost += 0.0;
					else if (bends == 1)
						cost += 200000.0;
					else
						cost += 400000.0 + (double)(bends - 2) * 70000.0;
					// Prefer direction-consistent ports.
					cost += (double)portScore * 3.0;
					// Penalize routes going too far away.
					for (const auto& p : pts)
					{
						if (!bbox.PtInRect(CPoint((int)p.X, (int)p.Y)))
							cost += 3000.0;
					}
					// Penalize too-short segments for labels.
					if (labelMinSegment > 0.0f && maxSeg < (double)labelMinSegment)
						cost += 5000.0 + ((double)labelMinSegment - maxSeg) * 15.0;
					return cost;
				};

				auto directionMatchesStartSide = [&](const Gdiplus::PointF& p0, const Gdiplus::PointF& p1, PortSide side) -> bool {
					const float eps = 0.2f;
					switch (side)
					{
					case PortLeft: return p1.X <= p0.X - eps;
					case PortRight: return p1.X >= p0.X + eps;
					case PortTop: return p1.Y <= p0.Y - eps;
					case PortBottom: return p1.Y >= p0.Y + eps;
					default: return true;
					}
				};
				auto directionMatchesEndSide = [&](const Gdiplus::PointF& prev, const Gdiplus::PointF& end, PortSide side) -> bool {
					const float eps = 0.2f;
					switch (side)
					{
					case PortLeft: return prev.X <= end.X - eps;
					case PortRight: return prev.X >= end.X + eps;
					case PortTop: return prev.Y <= end.Y - eps;
					case PortBottom: return prev.Y >= end.Y + eps;
					default: return true;
					}
				};
				auto routeHardValid = [&](const std::vector<Gdiplus::PointF>& pts, PortSide fromSide, PortSide toSide) -> bool {
					if (pts.size() < 2)
						return false;
					if (!directionMatchesStartSide(pts.front(), pts[1], fromSide))
						return false;
					if (!directionMatchesEndSide(pts[pts.size() - 2], pts.back(), toSide))
						return false;

					for (size_t si = 1; si < pts.size(); ++si)
					{
						const auto a = pts[si - 1];
						const auto b = pts[si];
						for (int ni = 0; ni < (int)d.nodeRects.size(); ++ni)
						{
							Gdiplus::RectF r = inflatedNodeRectF(d.nodeRects[ni], 2.0f);
							if (ni == fromNode)
							{
								if (si == 1)
								{
									// First segment may touch start border, but must not enter start interior.
									r = inflatedNodeRectF(d.nodeRects[ni], 0.0f);
									r = deflateRectF(r, 2.0f);
								}
							}
							if (ni == toNode)
							{
								if (si == pts.size() - 1)
								{
									// Last segment may touch end border, but must not cross end interior.
									r = inflatedNodeRectF(d.nodeRects[ni], 0.0f);
									r = deflateRectF(r, 2.0f);
								}
							}
							if (segmentIntersectsRectF(a, b, r))
								return false;
						}
					}
					return true;
				};
				auto detectPortSide = [&](const CRect& r, const Gdiplus::PointF& p) -> PortSide {
					const float cx = (float)r.left + (float)r.Width() / 2.0f;
					const float cy = (float)r.top + (float)r.Height() / 2.0f;
					const float dl = (float)fabs((double)(p.X - (float)r.left)) + (float)fabs((double)(p.Y - cy));
					const float dr = (float)fabs((double)(p.X - (float)r.right)) + (float)fabs((double)(p.Y - cy));
					const float dt = (float)fabs((double)(p.X - cx)) + (float)fabs((double)(p.Y - (float)r.top));
					const float db = (float)fabs((double)(p.X - cx)) + (float)fabs((double)(p.Y - (float)r.bottom));
					PortSide side = PortLeft;
					float best = dl;
					if (dr < best) { best = dr; side = PortRight; }
					if (dt < best) { best = dt; side = PortTop; }
					if (db < best) { best = db; side = PortBottom; }
					return side;
				};

				std::vector<Gdiplus::PointF> best;
				double bestCost = 1e18;
				// Limit combinations for performance (ports are already ordered).
				const int maxPorts = 4;
				for (int ia = 0; ia < (int)portsA.size() && ia < maxPorts; ++ia)
				{
					for (int ib = 0; ib < (int)portsB.size() && ib < maxPorts; ++ib)
					{
						const auto& pa = portsA[ia];
						const auto& pb = portsB[ib];
						auto pts = routeEdgePreferStraight(fromNode, toNode, ra, rb, pa.pt, pb.pt);
						if (pts.size() < 2)
							continue;
						if (!routeHardValid(pts, pa.side, pb.side))
							continue;
						double c = routeCost(pts, pa.score + pb.score);
						if (c < bestCost)
						{
							bestCost = c;
							best = std::move(pts);
						}
					}
				}
				if (best.size() >= 2)
					return best;

				// Last resort: keep previous behavior (midpoint ports).
				auto endpointsMid = [&]() -> std::pair<Gdiplus::PointF, Gdiplus::PointF> {
					float dx = cb.X - ca.X;
					float dy = cb.Y - ca.Y;
					if (fabs(dx) > fabs(dy))
					{
						if (dx >= 0)
							return { Gdiplus::PointF((Gdiplus::REAL)ra.right, ca.Y), Gdiplus::PointF((Gdiplus::REAL)rb.left, cb.Y) };
						return { Gdiplus::PointF((Gdiplus::REAL)ra.left, ca.Y), Gdiplus::PointF((Gdiplus::REAL)rb.right, cb.Y) };
					}
				if (dy >= 0)
						return { Gdiplus::PointF(ca.X, (Gdiplus::REAL)ra.bottom), Gdiplus::PointF(cb.X, (Gdiplus::REAL)rb.top) };
					return { Gdiplus::PointF(ca.X, (Gdiplus::REAL)ra.top), Gdiplus::PointF(cb.X, (Gdiplus::REAL)rb.bottom) };
				}();
				auto fallback = routeEdgePreferStraight(fromNode, toNode, ra, rb, endpointsMid.first, endpointsMid.second);
				if (fallback.size() >= 2)
				{
					PortSide fs = detectPortSide(ra, endpointsMid.first);
					PortSide ts = detectPortSide(rb, endpointsMid.second);
					if (routeHardValid(fallback, fs, ts))
						return fallback;
				}
				return {};
			};

			// edges: compute routes; draw after nodes (so nodes never hide lines).
			for (size_t ei = 0; ei < d.edges.size(); ++ei)
			{
				auto& e = d.edges[ei];
				int a = e.first;
				int b = e.second;
				if (a < 0 || b < 0 || a >= (int)d.nodeRects.size() || b >= (int)d.nodeRects.size())
					continue;
				CRect ra = d.nodeRects[a];
				CRect rb = d.nodeRects[b];

				CString edgeLabel = (ei < d.edgeLabels.size()) ? d.edgeLabels[ei] : CString();
				edgeLabel.Trim();

				// Edge routing:
				// - Try multiple ports (midpoints + corners) and pick the best route.
				// - Prefer a direct straight segment (including diagonal) when possible.
				// - Otherwise use a bend (折), with detour fallback.
				auto edgePts = chooseBestRoute(a, b, ra, rb, edgeLabel);
				EdgeRenderInfo info;
				info.fromNode = a;
				info.toNode = b;
				info.pts = std::move(edgePts);
				info.label = edgeLabel;
				edgeInfos.push_back(std::move(info));
			}

			// nodes
			for (int i = 0; i < (int)d.nodeRects.size(); ++i)
			{
				CRect rc = d.nodeRects[i];
				Gdiplus::RectF rcf((Gdiplus::REAL)rc.left, (Gdiplus::REAL)rc.top, (Gdiplus::REAL)rc.Width(), (Gdiplus::REAL)rc.Height());
				if (i < (int)d.nodeShapes.size() && d.nodeShapes[i] == MermaidNodeShape::Diamond)
				{
					Gdiplus::PointF pts[4] = {
						Gdiplus::PointF(rcf.X + rcf.Width / 2.0f, rcf.Y),
						Gdiplus::PointF(rcf.X + rcf.Width, rcf.Y + rcf.Height / 2.0f),
						Gdiplus::PointF(rcf.X + rcf.Width / 2.0f, rcf.Y + rcf.Height),
						Gdiplus::PointF(rcf.X, rcf.Y + rcf.Height / 2.0f)
					};
					g.FillPolygon(&nodeFill, pts, 4);
					g.DrawPolygon(&borderPen, pts, 4);
				}
				else if (i < (int)d.nodeShapes.size() && d.nodeShapes[i] == MermaidNodeShape::Stadium)
				{
					const float radius = max(10.0f, min(rcf.Height / 2.0f, rcf.Width / 2.5f));
					Gdiplus::GraphicsPath path;
					path.AddArc(rcf.X, rcf.Y, radius * 2, radius * 2, 90, 180);
					path.AddArc(rcf.X + rcf.Width - radius * 2, rcf.Y, radius * 2, radius * 2, 270, 180);
					path.CloseFigure();
					g.FillPath(&nodeFill, &path);
					g.DrawPath(&borderPen, &path);
				}
				else if (i < (int)d.nodeShapes.size() && d.nodeShapes[i] == MermaidNodeShape::Circle)
				{
					g.FillEllipse(&nodeFill, rcf);
					g.DrawEllipse(&borderPen, rcf);
				}
				else if (i < (int)d.nodeShapes.size() && d.nodeShapes[i] == MermaidNodeShape::RoundRect)
				{
					const float radius = max(6.0f, min(12.0f, rcf.Height / 3.0f));
					Gdiplus::GraphicsPath path;
					const float x = rcf.X;
					const float y = rcf.Y;
					const float w = rcf.Width;
					const float h = rcf.Height;
					path.AddArc(x, y, radius * 2, radius * 2, 180, 90);
					path.AddArc(x + w - radius * 2, y, radius * 2, radius * 2, 270, 90);
					path.AddArc(x + w - radius * 2, y + h - radius * 2, radius * 2, radius * 2, 0, 90);
					path.AddArc(x, y + h - radius * 2, radius * 2, radius * 2, 90, 90);
					path.CloseFigure();
					g.FillPath(&nodeFill, &path);
					g.DrawPath(&borderPen, &path);
				}
				else
				{
					g.FillRectangle(&nodeFill, rcf);
					g.DrawRectangle(&borderPen, rcf);
				}
				CString label = d.nodeLabels[i];
				std::wstring wl(label.GetString());
				Gdiplus::RectF tr = rcf;
				tr.Inflate(-8.0f, -6.0f);
				Gdiplus::StringFormat sf;
				sf.SetAlignment(Gdiplus::StringAlignmentCenter);
				sf.SetLineAlignment(Gdiplus::StringAlignmentCenter);
				sf.SetTrimming(Gdiplus::StringTrimmingEllipsisCharacter);
				sf.SetFormatFlags(Gdiplus::StringFormatFlagsLineLimit);
				g.DrawString(wl.c_str(), (INT)wl.size(), font.get(), tr, &sf, &textBrush);
			}

			// edges: draw on top, but clip to node interiors (so lines never enter nodes).
			{
				Gdiplus::Region clip(Gdiplus::Rect(0, 0, d.diagramWidth, d.diagramHeight));
				for (int i = 0; i < (int)d.nodeRects.size(); ++i)
				{
					Gdiplus::RectF r = inflatedNodeRectF(d.nodeRects[i], 0.0f);
					r = deflateRectF(r, 2.0f);
					clip.Exclude(r);
				}
				g.SetClip(&clip, Gdiplus::CombineModeReplace);
				for (const auto& edge : edgeInfos)
				{
					if (edge.pts.size() >= 2)
					{
						// Always draw straight segments (no curves).
						for (size_t pi = 1; pi < edge.pts.size(); ++pi)
							g.DrawLine(&edgePen, edge.pts[pi - 1], edge.pts[pi]);
					}
					if (edge.pts.size() >= 2)
					{
						// Find the last non-degenerate segment for arrow direction.
						Gdiplus::PointF tip = edge.pts.back();
						Gdiplus::PointF prev = edge.pts[edge.pts.size() - 2];
						for (size_t i = edge.pts.size() - 1; i > 0; --i)
						{
							Gdiplus::PointF a = edge.pts[i - 1];
							Gdiplus::PointF b = edge.pts[i];
							Gdiplus::REAL dx = b.X - a.X;
							Gdiplus::REAL dy = b.Y - a.Y;
							if (fabs(dx) + fabs(dy) >= 2.0f)
							{
								prev = a;
								tip = b;
								break;
							}
						}
					// Arrow direction is determined by the last visible segment.
					drawArrowHead(prev, tip);
					}
				}
				g.ResetClip();
			}

			// edge labels
			for (const auto& edge : edgeInfos)
			{
				CString edgeLabel = edge.label;
				edgeLabel.Trim();
				if (edgeLabel.IsEmpty())
					continue;

				std::wstring wl(edgeLabel.GetString());
				Gdiplus::RectF measureRc(0, 0, 1000.0f, 200.0f);
				Gdiplus::RectF bound;
				g.MeasureString(wl.c_str(), (INT)wl.size(), edgeLabelFont.get(), measureRc, &bound);
				const float labelW = bound.Width + 8.0f;
				const float labelH = bound.Height + 4.0f;

				// Choose a long segment to place the label (avoid tiny stubs near nodes).
				size_t segIndex = 0;
				double bestLen = -1.0;
				for (size_t i = 1; i < edge.pts.size(); ++i)
				{
					auto s = edge.pts[i - 1];
					auto t = edge.pts[i];
					double len = fabs((double)(t.X - s.X)) + fabs((double)(t.Y - s.Y));
					if (len > bestLen)
					{
						bestLen = len;
						segIndex = i - 1;
					}
				}
				auto s = edge.pts[segIndex];
				auto t = edge.pts[segIndex + 1];
				Gdiplus::PointF mid((s.X + t.X) / 2.0f, (s.Y + t.Y) / 2.0f);
				const bool horiz = fabs(t.Y - s.Y) <= fabs(t.X - s.X);

				auto labelOverlapsNode = [&](const Gdiplus::RectF& rr) -> bool {
					for (int i = 0; i < (int)d.nodeRects.size(); ++i)
					{
						Gdiplus::RectF nr = inflatedNodeRectF(d.nodeRects[i], 4.0f);
						if (rr.IntersectsWith(nr))
							return true;
					}
					return false;
				};
				auto labelOverlapsEdge = [&](const Gdiplus::RectF& rr) -> bool {
					for (size_t i = 1; i < edge.pts.size(); ++i)
					{
						if (segmentIntersectsRectF(edge.pts[i - 1], edge.pts[i], rr))
							return true;
					}
					return false;
				};

				auto makeRectAt = [&](float ox, float oy) -> Gdiplus::RectF {
					return Gdiplus::RectF(mid.X - labelW / 2.0f + ox,
						mid.Y - labelH / 2.0f + oy,
						labelW,
						labelH);
				};

				std::vector<std::pair<float, float>> offsets;
				offsets.reserve(12);
				if (horiz)
				{
					offsets.push_back({ 0.0f, -18.0f });
					offsets.push_back({ 0.0f, 18.0f });
					offsets.push_back({ 0.0f, -30.0f });
					offsets.push_back({ 0.0f, 30.0f });
					offsets.push_back({ 16.0f, -16.0f });
					offsets.push_back({ -16.0f, -16.0f });
				}
				else
				{
					offsets.push_back({ 18.0f, 0.0f });
					offsets.push_back({ -18.0f, 0.0f });
					offsets.push_back({ 30.0f, 0.0f });
					offsets.push_back({ -30.0f, 0.0f });
					offsets.push_back({ 16.0f, -16.0f });
					offsets.push_back({ 16.0f, 16.0f });
				}

				Gdiplus::RectF tr = makeRectAt(0.0f, horiz ? -18.0f : 18.0f);
				bool placed = false;
				for (const auto& off : offsets)
				{
					Gdiplus::RectF cand = makeRectAt(off.first, off.second);
					// Keep within image bounds.
					if (cand.X < 2.0f || cand.Y < 2.0f || cand.X + cand.Width > (float)d.diagramWidth - 2.0f || cand.Y + cand.Height > (float)d.diagramHeight - 2.0f)
						continue;
					if (labelOverlapsNode(cand))
						continue;
					if (labelOverlapsEdge(cand))
						continue;
					// Avoid covering the arrow head area.
					if (edge.pts.size() >= 2)
					{
						Gdiplus::PointF tip = edge.pts.back();
						Gdiplus::RectF tipRc(tip.X - 12.0f, tip.Y - 12.0f, 24.0f, 24.0f);
						if (cand.IntersectsWith(tipRc))
							continue;
					}
					tr = cand;
					placed = true;
					break;
				}
				if (!placed)
				{
					// As a fallback, draw with background so it remains readable.
					tr = makeRectAt(0.0f, horiz ? -16.0f : 16.0f);
				}

				Gdiplus::SolidBrush labelBg(bg);
				g.FillRectangle(&labelBg, tr);
				Gdiplus::StringFormat sf;
				sf.SetAlignment(Gdiplus::StringAlignmentCenter);
				sf.SetLineAlignment(Gdiplus::StringAlignmentCenter);
				g.DrawString(wl.c_str(), (INT)wl.size(), edgeLabelFont.get(), tr, &sf, &edgeLabelBrush);
			}

			gdiBmp.GetHBITMAP(Gdiplus::Color(0, 0, 0, 0), &bmp);
			w = d.diagramWidth;
			h = d.diagramHeight;
			}
		}
		if (bmp)
		{
			// Never evict bitmaps that are still needed in this update round; otherwise
			// large documents can leave some currently visible diagrams stuck in placeholder state.
			TrimMermaidBitmapCache(1, &currentDiagramHashes);
			if (m_mermaidBitmapCache.size() + 1 <= kMaxMermaidBitmapCacheEntries)
			{
				m_mermaidBitmapCache[hash] = bmp;
				m_mermaidBitmapSize[hash] = { w, h };
				TouchMermaidBitmapCache(hash);
			}
			else
			{
				// Keep overflow diagrams for this update round so 65+ diagrams still render,
				// while keeping the persistent LRU cache under the configured cap.
				m_mermaidTransientBitmapCache[hash] = bmp;
				m_mermaidTransientBitmapSize[hash] = { w, h };
			}
		}
		m_mermaidFetchInFlight.erase(hash);
	}

	UpdateMermaidOverlayRegion(true);
}

void CMarkdownEditorDlg::DrawMermaidOverlays(CDC& dc)
{
	if (!::IsWindow(m_mermaidOverlay.GetSafeHwnd()) || !::IsWindow(m_editContent.GetSafeHwnd()))
		return;
	CRect overlayClient;
	m_mermaidOverlay.GetClientRect(&overlayClient);
	const double zoomScale = max(0.1, min(5.0, (double)GetZoomPercent() / 100.0));

	// Do not paint outside diagram rectangles to avoid covering editor.

	for (const auto& diagram : m_mermaidDiagrams)
	{
		// Map character position to rect in RichEdit.
		POINTL pt{};
		pt.x = 0;
		pt.y = 0;
		// EM_POSFROMCHAR returns -1 for invalid positions.
		LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)diagram.mappedStart);
		if (r == -1)
			continue;

		CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
		// Convert from RichEdit client to dialog client then to overlay client.
		m_editContent.ClientToScreen(&topLeft);
		m_mermaidOverlay.ScreenToClient(&topLeft);

		int widthPx = 0;
		int heightPx = 0;
		HBITMAP bmp = nullptr;
		auto it = m_mermaidBitmapCache.find(diagram.hash);
		if (it != m_mermaidBitmapCache.end())
		{
			bmp = it->second;
			TouchMermaidBitmapCache(diagram.hash);
		}
		else
		{
			auto itTransient = m_mermaidTransientBitmapCache.find(diagram.hash);
			if (itTransient != m_mermaidTransientBitmapCache.end())
				bmp = itTransient->second;
		}
		auto itSize = m_mermaidBitmapSize.find(diagram.hash);
		if (itSize != m_mermaidBitmapSize.end())
		{
			widthPx = itSize->second.first;
			heightPx = itSize->second.second;
		}
		else
		{
			auto itTransientSize = m_mermaidTransientBitmapSize.find(diagram.hash);
			if (itTransientSize != m_mermaidTransientBitmapSize.end())
			{
				widthPx = itTransientSize->second.first;
				heightPx = itTransientSize->second.second;
			}
		}
		if (widthPx <= 0) widthPx = 520;
		if (heightPx <= 0) heightPx = max(140, min(320, diagram.lineCount * 18 + 40));

		int drawW = (int)(widthPx * zoomScale);
		int drawH = (int)(heightPx * zoomScale);
		// Do NOT fit-to-width: allow overflow and let RichEdit scrollbars handle it.
		// Ensure RichEdit uses only vertical scroll, so horizontal overflow can be handled
		// by the diagram bitmap itself (not by scrolling the entire document).

		CRect rc(topLeft.x, topLeft.y, topLeft.x + drawW, topLeft.y + drawH);
		if (!CRect().IntersectRect(&rc, &overlayClient))
			continue;

		COLORREF border = m_themeDark ? RGB(70, 70, 70) : RGB(210, 210, 210);
		CPen pen(PS_SOLID, 1, border);
		CPen* oldPen = dc.SelectObject(&pen);
		CBrush* oldBrush = (CBrush*)dc.SelectStockObject(NULL_BRUSH);
		dc.RoundRect(rc, CPoint(10, 10));
		dc.SelectObject(oldBrush);
		dc.SelectObject(oldPen);

		if (bmp)
		{
			CDC mem;
			mem.CreateCompatibleDC(&dc);
			HBITMAP oldBmp = (HBITMAP)mem.SelectObject(bmp);
			dc.StretchBlt(rc.left, rc.top, rc.Width(), rc.Height(), &mem, 0, 0, widthPx, heightPx, SRCCOPY);
			mem.SelectObject(oldBmp);
		}
		else
		{
			COLORREF fill = m_themeDark ? RGB(36, 36, 36) : RGB(250, 250, 250);
			CBrush brush(fill);
			dc.FillRect(rc, &brush);
			CString hint = _T("正在加载中...");
			dc.SetBkMode(TRANSPARENT);
			dc.SetTextColor(m_themeForeground);
			CRect textRc = rc;
			textRc.DeflateRect(12, 12);
			dc.DrawText(hint, textRc, DT_LEFT | DT_TOP | DT_WORDBREAK);
		}
	}
}

void CMarkdownEditorDlg::DrawHorizontalRules(CDC& dc)
{
	if (!::IsWindow(m_mermaidOverlay.GetSafeHwnd()) || !::IsWindow(m_editContent.GetSafeHwnd()))
		return;
	if (!m_bMarkdownMode)
		return;
	if (m_horizontalRuleStarts.empty())
		return;

	CRect overlayClient;
	m_mermaidOverlay.GetClientRect(&overlayClient);

	COLORREF color = m_themeDark ? RGB(90, 90, 90) : RGB(200, 200, 200);
	CPen pen(PS_SOLID, 1, color);
	CPen* oldPen = dc.SelectObject(&pen);

	for (long start : m_horizontalRuleStarts)
	{
		POINTL pt{};
		LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)start);
		if (r == -1)
			continue;
		CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
		m_editContent.ClientToScreen(&topLeft);
		m_mermaidOverlay.ScreenToClient(&topLeft);

		int line = m_editContent.LineFromChar(start);
		int rowHeight = 20;
		if (line >= 0)
		{
			int nextLine = line + 1;
			long nextLineStart = m_editContent.LineIndex(nextLine);
			if (nextLineStart >= 0)
			{
				POINTL ptNext{};
				LRESULT rNext = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&ptNext, (LPARAM)nextLineStart);
				if (rNext != -1)
					rowHeight = max(18, (int)ptNext.y - (int)pt.y);
			}
		}
		if (rowHeight <= 0)
			rowHeight = 20;

		int y = topLeft.y + rowHeight / 2;
		if (y < overlayClient.top || y >= overlayClient.bottom)
			continue;

		dc.MoveTo(overlayClient.left + 12, y);
		dc.LineTo(overlayClient.right - 12, y);
	}

	dc.SelectObject(oldPen);
}

void CMarkdownEditorDlg::DrawTableGrid(CDC& dc)
{
	if (!::IsWindow(m_mermaidOverlay.GetSafeHwnd()) || !::IsWindow(m_editContent.GetSafeHwnd()))
		return;
	if (!m_bMarkdownMode)
		return;
	if (m_tableOverlayRows.empty())
		return;

	CRect overlayClient;
	m_mermaidOverlay.GetClientRect(&overlayClient);

	COLORREF gridColor = m_themeDark ? RGB(75, 75, 75) : RGB(220, 220, 220);
	CPen pen(PS_SOLID, 1, gridColor);
	CPen* oldPen = dc.SelectObject(&pen);

	CClientDC dpiDc(this);
	int dpiX = dpiDc.GetDeviceCaps(LOGPIXELSX);
	if (dpiX <= 0) dpiX = 96;
	auto twipsToPxX = [&](LONG twips) -> int {
		return (int)MulDiv(twips, dpiX, 1440);
	};

	// Group table rows by tab-stop signature and compute a stable left edge for each group.
	// This avoids per-row left-edge jitter when some cells are very long.
	struct TableKey
	{
		int count = 0;
		LONG stops[32]{};
	};
	auto keyLess = [](const TableKey& a, const TableKey& b) {
		if (a.count != b.count) return a.count < b.count;
		for (int i = 0; i < a.count && i < 32; ++i)
		{
			if (a.stops[i] != b.stops[i]) return a.stops[i] < b.stops[i];
		}
		return false;
	};

	std::map<TableKey, int, decltype(keyLess)> groupLeftPx(keyLess);
	for (size_t rowIndex = 0; rowIndex < m_tableOverlayRows.size(); ++rowIndex)
	{
		const auto& row = m_tableOverlayRows[rowIndex];
		if (row.tabStopCount <= 0)
			continue;
		TableKey k{};
		k.count = min(row.tabStopCount, 32);
		for (int i = 0; i < k.count; ++i)
			k.stops[i] = row.tabStopsTwips[i];

		POINTL pt{};
		LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)row.start);
		if (r == -1)
			continue;
		CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
		m_editContent.ClientToScreen(&topLeft);
		m_mermaidOverlay.ScreenToClient(&topLeft);
		int candidate = max(overlayClient.left + 2, topLeft.x);
		auto it = groupLeftPx.find(k);
		if (it == groupLeftPx.end())
			groupLeftPx.emplace(k, candidate);
		else
			it->second = min(it->second, candidate);
	}

	for (size_t rowIndex = 0; rowIndex < m_tableOverlayRows.size(); ++rowIndex)
	{
		const auto& row = m_tableOverlayRows[rowIndex];
		if (row.tabStopCount <= 0)
			continue;
		TableKey k{};
		k.count = min(row.tabStopCount, 32);
		for (int i = 0; i < k.count; ++i)
			k.stops[i] = row.tabStopsTwips[i];

		POINTL pt{};
		LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)row.start);
		if (r == -1)
			continue;
		CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
		m_editContent.ClientToScreen(&topLeft);
		m_mermaidOverlay.ScreenToClient(&topLeft);

		int line = m_editContent.LineFromChar(row.start);
		if (line < 0)
			continue;
		int nextLine = line + 1;
		long nextLineStart = m_editContent.LineIndex(nextLine);
		int rowHeight = 20;
		if (nextLineStart >= 0)
		{
			POINTL ptNext{};
			LRESULT rNext = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&ptNext, (LPARAM)nextLineStart);
			if (rNext != -1)
				rowHeight = max(18, (int)ptNext.y - (int)pt.y);
		}
		// Prefer the next table row's top as current row bottom to avoid gaps caused by
		// hidden table-separator lines between header and first data row.
		if (rowIndex + 1 < m_tableOverlayRows.size())
		{
			const auto& nextRow = m_tableOverlayRows[rowIndex + 1];
			if (nextRow.tabStopCount > 0)
			{
				TableKey nextKey{};
				nextKey.count = min(nextRow.tabStopCount, 32);
				for (int i = 0; i < nextKey.count; ++i)
					nextKey.stops[i] = nextRow.tabStopsTwips[i];

				if (!keyLess(k, nextKey) && !keyLess(nextKey, k))
				{
					POINTL ptNextRow{};
					LRESULT rNextRow = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&ptNextRow, (LPARAM)nextRow.start);
					if (rNextRow != -1)
					{
						const int inferredHeight = (int)ptNextRow.y - (int)pt.y;
						if (inferredHeight > 0)
							rowHeight = max(18, inferredHeight);
					}
				}
			}
		}
		if (rowHeight <= 0)
			rowHeight = 20;
		const int yTop = topLeft.y;
		const int yBottom = yTop + rowHeight; // exclusive
		const int yBottomLine = max(yTop, yBottom - 1);
		if (yBottom < overlayClient.top || yTop > overlayClient.bottom)
			continue;

		// Align the table grid with actual text layout.
		int left = max(overlayClient.left + 2, topLeft.x);
		auto itLeft = groupLeftPx.find(k);
		if (itLeft != groupLeftPx.end())
			left = itLeft->second;
		LONG lastStopTwips = row.tabStopsTwips[min(row.tabStopCount, 32) - 1];
		int right = left + twipsToPxX(lastStopTwips);
		right = min(right, overlayClient.right - 2);
		if (right <= left + 40)
			continue;

		const int leftLine = max(overlayClient.left + 1, left - 2);
		const int rightLine = min(overlayClient.right - 1, max(leftLine + 1, right - 1));

		// Horizontal line at top + bottom of each row.
		dc.MoveTo(leftLine, yTop);
		dc.LineTo(rightLine + 1, yTop);
		dc.MoveTo(leftLine, yBottomLine);
		dc.LineTo(rightLine + 1, yBottomLine);

		// Vertical lines at tab stop positions (exclude the last stop; right border is drawn separately).
		const int separatorCount = max(0, row.tabStopCount - 1);
		for (int i = 0; i < separatorCount; ++i)
		{
			int x = left + twipsToPxX(row.tabStopsTwips[i]);
			x -= 2; // keep separator slightly left so it does not touch text edge
			if (x <= leftLine || x >= rightLine)
				continue;
			dc.MoveTo(x, yTop);
			dc.LineTo(x, yBottomLine + 1);
		}

		// Draw right/left borders for all rows (including header).
		dc.MoveTo(rightLine, yTop);
		dc.LineTo(rightLine, yBottomLine + 1);
		dc.MoveTo(leftLine, yTop);
		dc.LineTo(leftLine, yBottomLine + 1);
	}

	dc.SelectObject(oldPen);
}

void CMarkdownEditorDlg::DrawTocOverlay(CDC& dc)
{
	if (!::IsWindow(m_mermaidOverlay.GetSafeHwnd()) || !::IsWindow(m_editContent.GetSafeHwnd()))
		return;
	if (!m_bMarkdownMode)
		return;
	if (m_tocOverlayBlocks.empty())
		return;

	m_tocHitRegions.clear();

	CRect overlayClient;
	m_mermaidOverlay.GetClientRect(&overlayClient);
	dc.SetBkMode(TRANSPARENT);
	dc.SetTextColor(m_themeForeground);
	CFont* oldFont = nullptr;
	if (CFont* editFont = m_editContent.GetFont())
		oldFont = dc.SelectObject(editFont);
	CFont tocDrawFont;
	CFont* oldDrawZoomFont = SelectTocFontByZoom(dc, GetZoomPercent(), m_normalFormat, tocDrawFont);

	for (const auto& block : m_tocOverlayBlocks)
	{
		if (block.lines.empty())
			continue;
		POINTL pt{};
		LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)block.mappedStart);
		if (r == -1)
			continue;
		CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
		m_editContent.ClientToScreen(&topLeft);
		m_mermaidOverlay.ScreenToClient(&topLeft);

		int line = m_editContent.LineFromChar(block.mappedStart);
		int rowHeight = 20;
		if (line >= 0)
		{
			long nextLineStart = m_editContent.LineIndex(line + 1);
			if (nextLineStart >= 0)
			{
				POINTL ptNext{};
				LRESULT rNext = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&ptNext, (LPARAM)nextLineStart);
				if (rNext != -1)
					rowHeight = max(18, (int)ptNext.y - (int)pt.y);
			}
		}

		const int textLeft = max(overlayClient.left + 10, topLeft.x);
		const int textRight = overlayClient.right - 12;
		const int textWidth = max(40, textRight - textLeft);
		int y = topLeft.y + 2;
		for (const auto& lineItem : block.lines)
		{
			if (lineItem.text.empty())
				continue;
			int lineHeight = MeasureWrappedTextHeightPx(dc, lineItem.text, textWidth);
			if (lineHeight <= 0)
				lineHeight = max(rowHeight, 18);
			CRect lineRc(textLeft, y, textRight, y + lineHeight);
			if (lineRc.right > lineRc.left && lineRc.bottom > lineRc.top)
			{
				CRect drawRc = lineRc;
				if (CRect().IntersectRect(&drawRc, &overlayClient))
				{
					// Repaint line background before transparent text draw to avoid cumulative darkening.
					dc.FillSolidRect(&drawRc, m_themeBackground);
					dc.DrawText(lineItem.text.c_str(), static_cast<int>(lineItem.text.size()),
						&drawRc, DT_LEFT | DT_TOP | DT_WORDBREAK | DT_NOPREFIX);
				}
				if (lineItem.targetSourceIndex >= 0)
				{
					TocHitRegion hit{};
					hit.rect = lineRc;
					hit.targetSourceIndex = lineItem.targetSourceIndex;
					m_tocHitRegions.push_back(std::move(hit));
				}
			}
			y += lineHeight;
		}
	}

	if (oldDrawZoomFont != nullptr)
		dc.SelectObject(oldDrawZoomFont);
	if (oldFont != nullptr)
		dc.SelectObject(oldFont);
}

bool CMarkdownEditorDlg::HandleTocOverlayClick(CPoint overlayPoint)
{
	if (!m_bMarkdownMode)
		return false;
	if (m_tocHitRegions.empty())
		return false;
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return false;

	for (const auto& hit : m_tocHitRegions)
	{
		if (hit.targetSourceIndex < 0)
			continue;
		if (!hit.rect.PtInRect(overlayPoint))
			continue;

		int richStart = MapSourceIndexToRichIndex(hit.targetSourceIndex);
		if (richStart < 0)
			return false;

		int targetLine = m_editContent.LineFromChar(richStart);
		if (targetLine < 0)
			targetLine = 0;
		long caretPos = richStart;
		bool hasPreciseCaret = false;
		for (const auto& entry : m_outlineEntries)
		{
			if (entry.sourceIndex != hit.targetSourceIndex)
				continue;
			if (entry.title.empty())
				break;
			const int sourceEnd = entry.sourceIndex + static_cast<int>(entry.title.size());
			const int richEnd = MapSourceIndexToRichIndex(sourceEnd);
			if (richEnd >= 0)
			{
				caretPos = richEnd;
				hasPreciseCaret = true;
			}
			break;
		}
		if (!hasPreciseCaret)
		{
			long lineStart = m_editContent.LineIndex(targetLine);
			if (lineStart < 0)
				lineStart = richStart;
			int lineLen = m_editContent.LineLength(static_cast<int>(lineStart));
			caretPos = lineStart + max(0, lineLen);
		}

		m_editContent.SetSel(caretPos, caretPos);
		const int firstVisible = static_cast<int>(m_editContent.SendMessage(EM_GETFIRSTVISIBLELINE, 0, 0));
		const int delta = targetLine - firstVisible;
		if (delta != 0)
			m_editContent.SendMessage(EM_LINESCROLL, 0, delta);
		m_editContent.SendMessage(EM_SCROLLCARET, 0, 0);
		m_editContent.SetFocus();
		UpdateOutlineSelectionByCaret();

		if (m_bMarkdownMode)
		{
			m_overlaySyncAttempts = 0;
			UpdateMermaidOverlayRegion(true);
		}
		return true;
	}
	return false;
}

void CMarkdownEditorDlg::UpdateTocSpacing()
{
	if (!::IsWindow(m_editContent.GetSafeHwnd()))
		return;

	LONG originalSelStart = 0;
	LONG originalSelEnd = 0;
	m_editContent.GetSel(originalSelStart, originalSelEnd);
	const int originalFirstVisible = m_editContent.GetFirstVisibleLine();
	auto restoreEditorView = [&]() {
		if (!::IsWindow(m_editContent.GetSafeHwnd()))
			return;

		GETTEXTLENGTHEX textLenEx{};
		textLenEx.flags = GTL_NUMCHARS;
		textLenEx.codepage = 1200;
		const long textLen = static_cast<long>(m_editContent.SendMessage(EM_GETTEXTLENGTHEX,
			reinterpret_cast<WPARAM>(&textLenEx), 0));
		LONG selStart = originalSelStart;
		LONG selEnd = originalSelEnd;
		if (selStart < 0) selStart = 0;
		if (selEnd < 0) selEnd = 0;
		if (selStart > textLen) selStart = textLen;
		if (selEnd > textLen) selEnd = textLen;
		m_editContent.SetSel(selStart, selEnd);

		const int currentFirstVisible = m_editContent.GetFirstVisibleLine();
		const int delta = originalFirstVisible - currentFirstVisible;
		if (delta != 0)
			m_editContent.LineScroll(delta);
	};

	// Reset spacing from previous TOC render pass.
	for (long start : m_tocSpaceAfterStarts)
	{
		PARAFORMAT2 reset{};
		reset.cbSize = sizeof(reset);
		reset.dwMask = PFM_SPACEAFTER;
		reset.dySpaceAfter = 0;
		m_editContent.SetSel(start, start);
		m_editContent.SetParaFormat(reset);
	}
	m_tocSpaceAfterStarts.clear();

	if (!m_bMarkdownMode || m_tocOverlayBlocks.empty())
	{
		restoreEditorView();
		return;
	}

	CWnd* measureWnd = ::IsWindow(m_mermaidOverlay.GetSafeHwnd())
		? static_cast<CWnd*>(&m_mermaidOverlay)
		: static_cast<CWnd*>(&m_editContent);
	CClientDC dc(measureWnd);
	int dpiY = dc.GetDeviceCaps(LOGPIXELSY);
	if (dpiY <= 0) dpiY = 96;
	auto pxToTwipsY = [&](int px) -> LONG {
		return (LONG)MulDiv(px, 1440, dpiY);
	};
	CFont* oldFont = nullptr;
	if (CFont* editFont = m_editContent.GetFont())
		oldFont = dc.SelectObject(editFont);
	CFont tocSpacingFont;
	CFont* oldSpacingZoomFont = SelectTocFontByZoom(dc, GetZoomPercent(), m_normalFormat, tocSpacingFont);

	CRect overlayClient;
	if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
		m_mermaidOverlay.GetClientRect(&overlayClient);
	else
		m_editContent.GetClientRect(&overlayClient);

	for (const auto& block : m_tocOverlayBlocks)
	{
		if (block.lines.empty())
			continue;
		int line = m_editContent.LineFromChar(block.mappedStart);
		int rowHeight = 20;
		int textLeft = 10;
		int textRight = overlayClient.right - 12;
		if (line >= 0)
		{
			long nextLineStart = m_editContent.LineIndex(line + 1);
			if (nextLineStart >= 0)
			{
				POINTL pt{};
				POINTL ptNext{};
				LRESULT r = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&pt, (LPARAM)block.mappedStart);
				LRESULT rNext = m_editContent.SendMessage(EM_POSFROMCHAR, (WPARAM)&ptNext, (LPARAM)nextLineStart);
				if (r != -1 && rNext != -1)
				{
					CPoint topLeft(static_cast<int>(pt.x), static_cast<int>(pt.y));
					m_editContent.ClientToScreen(&topLeft);
					if (::IsWindow(m_mermaidOverlay.GetSafeHwnd()))
						m_mermaidOverlay.ScreenToClient(&topLeft);
					textLeft = max(overlayClient.left + 10, topLeft.x);
					textRight = overlayClient.right - 12;
					rowHeight = max(18, (int)ptNext.y - (int)pt.y);
				}
			}
		}

		const int textWidth = max(40, textRight - textLeft);
		int measuredTextH = 0;
		for (const auto& lineItem : block.lines)
		{
			if (lineItem.text.empty())
				continue;
			int oneLineH = MeasureWrappedTextHeightPx(dc, lineItem.text, textWidth);
			if (oneLineH <= 0)
				oneLineH = max(rowHeight, 18);
			measuredTextH += oneLineH;
		}
		if (measuredTextH <= 0)
			measuredTextH = rowHeight;
		int extraPx = max(0, measuredTextH - rowHeight) + 8;
		if (extraPx <= 0)
			continue;
		LONG spaceAfter = pxToTwipsY(extraPx);
		if (spaceAfter <= 0)
			continue;
		PARAFORMAT2 para{};
		para.cbSize = sizeof(para);
		para.dwMask = PFM_SPACEAFTER;
		para.dySpaceAfter = spaceAfter;
		m_editContent.SetSel(block.mappedStart, block.mappedStart);
		m_editContent.SetParaFormat(para);
		m_tocSpaceAfterStarts.push_back(block.mappedStart);
	}

	if (oldSpacingZoomFont != nullptr)
		dc.SelectObject(oldSpacingZoomFont);
	if (oldFont != nullptr)
		dc.SelectObject(oldFont);
	restoreEditorView();
}

