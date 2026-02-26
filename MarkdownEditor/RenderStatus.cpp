#include "pch.h"
#include "framework.h"
#include "MarkdownEditor.h"
#include "MarkdownEditorDlg.h"

namespace
{
    constexpr UINT_PTR kStatusPulseTimerId = 4;
}

int CMarkdownEditorDlg::CalculateOpenProgressPercent() const
{
	if (!m_loadingActive && !(m_loadBytesTotal > 0 && m_loadBytesRead < m_loadBytesTotal))
		return 100;

	if (m_loadBytesTotal > 0 && m_loadBytesRead < m_loadBytesTotal)
	{
		const unsigned long long value = (m_loadBytesRead * 100ULL) / m_loadBytesTotal;
		return static_cast<int>(max(0ULL, min(100ULL, value)));
	}

	if (!m_loadingText.empty())
	{
		const unsigned long long value = (static_cast<unsigned long long>(m_loadingTextPos) * 100ULL) /
			static_cast<unsigned long long>(m_loadingText.size());
		return static_cast<int>(max(0ULL, min(100ULL, value)));
	}

	return 0;
}

int CMarkdownEditorDlg::CalculateRenderProgressPercent() const
{
	if (m_parsingActive && !m_formattingActive)
	{
		const int parse = max(1, min(100, m_parseProgressPercent));
		const int mapped = 1 + (parse * 59) / 100; // 1~60
		return max(1, min(60, mapped));
	}
	if (!m_formattingActive)
		return m_renderBusy ? 0 : 100;

	double totalUnits = 0.0;
	double doneUnits = 0.0;

	totalUnits += 1.0;
	if (m_baseFormatApplied)
	{
		doneUnits += 1.0;
	}
	else if (m_pendingTextLength > 0)
	{
		const double base = static_cast<double>(min(m_baseFormatPos, m_pendingTextLength)) /
			static_cast<double>(m_pendingTextLength);
		doneUnits += max(0.0, min(1.0, base));
	}

	const size_t blockCount = m_pendingParseResult.blockSpans.size();
	const size_t inlineCount = m_pendingParseResult.inlineSpans.size();
	const size_t hiddenCount = m_pendingParseResult.hiddenSpans.size();
	const size_t mermaidCount = m_mermaidBlocks.size();

	totalUnits += static_cast<double>(blockCount + inlineCount + hiddenCount + mermaidCount);
	doneUnits += static_cast<double>(min(m_blockSpanIndex, blockCount));
	doneUnits += static_cast<double>(min(m_inlineSpanIndex, inlineCount));
	doneUnits += static_cast<double>(min(m_hiddenSpanIndex, hiddenCount));
	doneUnits += static_cast<double>(min(m_mermaidSpanIndex, mermaidCount));

	if (totalUnits <= 0.0)
		return 0;

	const int formatPercent = static_cast<int>((doneUnits * 100.0) / totalUnits + 0.5);
	const int normalized = 60 + (max(0, min(100, formatPercent)) * 40) / 100; // 60~100
	return max(60, min(100, normalized));
}

void CMarkdownEditorDlg::UpdateSidebarInteractivity()
{
	const bool disable = m_renderBusy || m_loadingActive || m_parsingActive ||
		m_formattingActive || m_textModeResetActive ||
		m_fullTextReplaceTask != FullTextReplaceTask::None;
	if (disable == m_sidebarControlsDisabled)
		return;

	m_sidebarControlsDisabled = disable;
	auto setEnable = [&](CWnd& wnd) {
		if (::IsWindow(wnd.GetSafeHwnd()))
			wnd.EnableWindow(disable ? FALSE : TRUE);
	};

	setEnable(m_btnOpenFile);
	setEnable(m_btnModeSwitch);
	setEnable(m_btnThemeSwitch);
	setEnable(m_editFileSearch);
	setEnable(m_tabSidebar);
	setEnable(m_treeSidebar);
	setEnable(m_sidebarSplitter);
}

void CMarkdownEditorDlg::UpdateStatusText(bool force)
{
	CString status;
	status.Format(_T("模式: %s"),
		m_bMarkdownMode ? _T("Markdown") : _T("文本"));
	if (m_documentDirty || m_editingHintVisible)
		status += _T(" | 编辑中...");
	else if (m_autoSavedStatusVisible)
		status += _T(" | 已自动保存");

	const bool unifiedLoading = m_renderBusy && (m_loadingActive || m_parsingActive || m_formattingActive);
	if (unifiedLoading)
	{
		status += _T(" | 正在加载中...");
	}
	else if (m_renderBusy && !m_renderBusyHint.IsEmpty())
	{
		status += _T(" | ");
		status += m_renderBusyHint;
	}

	bool progressActive = false;
	int progressPercent = 0;
	CString progressText;
	const bool loadingStage = m_loadingActive || (m_loadBytesTotal > 0 && m_loadBytesRead < m_loadBytesTotal);
	const bool renderingStage = m_parsingActive || m_formattingActive;
	if (loadingStage || renderingStage)
	{
		progressActive = true;
		progressPercent = loadingStage ? CalculateOpenProgressPercent() : CalculateRenderProgressPercent();
		progressText.Format(_T("%d%%"), progressPercent);
	}
	progressPercent = max(0, min(100, progressPercent));

	// Progress bar + numeric percentage in status area.
	if (::IsWindow(m_progressStatus.GetSafeHwnd()) && ::IsWindow(m_staticProgressText.GetSafeHwnd()))
	{
		if (progressActive)
		{
			if (!m_statusProgressVisible)
			{
				m_progressStatus.ShowWindow(SW_SHOW);
				m_staticProgressText.ShowWindow(SW_SHOW);
			}
			m_statusProgressVisible = true;
			m_progressStatus.SetPos(progressPercent);
			m_staticProgressText.SetWindowText(progressText);
		}
		else
		{
			if (m_statusProgressVisible)
			{
				m_progressStatus.SetPos(0);
				m_progressStatus.ShowWindow(SW_HIDE);
				m_staticProgressText.SetWindowText(_T(""));
				m_staticProgressText.ShowWindow(SW_HIDE);
			}
			m_statusProgressVisible = false;
		}
	}

	{
		CString zoom;
		zoom.Format(_T(" | 缩放:%d%%"), GetZoomPercent());
		status += zoom;
	}

	if (!force)
	{
		const bool transient = m_renderBusy || m_loadingActive || m_parsingActive || m_formattingActive;
		if (transient)
		{
			const unsigned long now = ::GetTickCount();
			if (m_lastStatusUpdateTick != 0 && (now - m_lastStatusUpdateTick) < 120)
				return;
		}
	}

	if (status != m_lastStatusText)
	{
		m_lastStatusText = status;
		m_lastStatusUpdateTick = ::GetTickCount();
		m_staticStatus.SetWindowText(status);
		if (GetSafeHwnd() != nullptr)
		{
			CRect rc;
			GetClientRect(&rc);
			UpdateLayout(rc.Width(), rc.Height());
		}
	}

	// Keep status area repaint stable during high-frequency updates to avoid ghosting.
	const bool transientBusy = m_renderBusy || m_loadingActive || m_parsingActive || m_formattingActive;
	if (force || transientBusy || progressActive)
	{
		if (::IsWindow(m_staticStatus.GetSafeHwnd()))
			m_staticStatus.RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
		if (::IsWindow(m_progressStatus.GetSafeHwnd()))
			m_progressStatus.RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
		if (::IsWindow(m_staticProgressText.GetSafeHwnd()))
			m_staticProgressText.RedrawWindow(nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
}

void CMarkdownEditorDlg::SetRenderBusy(bool busy, const CString& hint)
{
	m_renderBusy = busy;
	m_renderBusyHint = hint;

	bool holdContent = false;
	if (::IsWindow(m_editContent.GetSafeHwnd()))
	{
		holdContent = m_hideDocumentUntilRenderReady &&
			(busy || m_loadingActive || m_parsingActive || m_formattingActive || m_fullTextReplaceTask != FullTextReplaceTask::None);
		const BOOL editVisible = m_editContent.IsWindowVisible();
		if (holdContent && editVisible)
			m_editContent.ShowWindow(SW_HIDE);
		else if (!holdContent && !editVisible)
		{
			m_editContent.ShowWindow(SW_SHOW);
			m_editContent.Invalidate(FALSE);
		}
	}

	if (::IsWindow(m_staticOverlay.GetSafeHwnd()))
	{
		long docLen = 0;
		if (::IsWindow(m_editContent.GetSafeHwnd()))
		{
			GETTEXTLENGTHEX textLenEx{};
			textLenEx.flags = GTL_NUMCHARS;
			textLenEx.codepage = 1200;
			docLen = static_cast<long>(m_editContent.SendMessage(EM_GETTEXTLENGTHEX,
				reinterpret_cast<WPARAM>(&textLenEx), 0));
		}

		constexpr long kOverlayThresholdChars = 200000;
		const bool showOverlay = holdContent || (busy && (m_loadingActive || docLen >= kOverlayThresholdChars));
		CString overlayText;
		if (showOverlay)
		{
			overlayText = _T("正在加载中...");
		}
		CString currentOverlayText;
		m_staticOverlay.GetWindowText(currentOverlayText);
		const bool textChanged = (currentOverlayText != (showOverlay ? overlayText : _T("")));
		const bool visibleChanged = ((m_staticOverlay.IsWindowVisible() != FALSE) != showOverlay);
		if (textChanged)
			m_staticOverlay.SetWindowText(showOverlay ? overlayText : _T(""));
		if (visibleChanged)
		{
			m_staticOverlay.ShowWindow(showOverlay ? SW_SHOW : SW_HIDE);
			if (showOverlay)
				m_staticOverlay.BringWindowToTop();
		}
		if (textChanged || visibleChanged)
		{
			m_staticOverlay.Invalidate();
			if (showOverlay)
				m_staticOverlay.UpdateWindow();
		}
	}

	UpdateSidebarInteractivity();
	UpdateStatusText(true);
	if (busy)
		SetTimer(kStatusPulseTimerId, 350, NULL);
	else
		KillTimer(kStatusPulseTimerId);
}

