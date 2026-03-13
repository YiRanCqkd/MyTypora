#pragma once

#include "MarkdownParser.h"
#include "third_party/gridctrl_src/GridCtrl.h"

#include <atomic>
#include <map>
#include <memory>
#include <set>
#include <vector>
#include <utility>
#include <string>

struct MarkdownOutlineEntry
{
    int level = 0;
    int lineNo = 0;
    int sourceIndex = -1;
    std::wstring title;
};

class CMarkdownEditorDlg;
class CFindReplaceDialog;

class CMermaidOverlayWnd : public CWnd
{
public:
	void SetHost(CMarkdownEditorDlg* host) { m_host = host; }
	virtual BOOL OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult) override;

protected:
	afx_msg void OnPaint();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg LRESULT OnNcHitTest(CPoint point);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	DECLARE_MESSAGE_MAP()

private:
	CMarkdownEditorDlg* m_host = nullptr;
	bool m_tocClickSuppressUp = false;
};

class CTableGridOverlayCtrl : public CGridCtrl
{
public:
	void SetHost(CMarkdownEditorDlg* host) { m_host = host; }

protected:
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	DECLARE_MESSAGE_MAP()

private:
	void ForwardMouseToEditor(UINT msg, UINT nFlags, CPoint point);
	CMarkdownEditorDlg* m_host = nullptr;
};

class CSidebarSplitterWnd : public CWnd
{
public:
	void SetHost(CMarkdownEditorDlg* host) { m_host = host; }

protected:
	afx_msg void OnPaint();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	DECLARE_MESSAGE_MAP()

private:
	CMarkdownEditorDlg* m_host = nullptr;
	bool m_dragging = false;
};

class CSidebarTabCtrl : public CTabCtrl
{
public:
	void SetHost(CMarkdownEditorDlg* host) { m_host = host; }

protected:
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnNcPaint();
	afx_msg void OnPaint();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
	DECLARE_MESSAGE_MAP()

private:
	CMarkdownEditorDlg* m_host = nullptr;
};

class CSidebarSearchEdit : public CEdit
{
public:
	void SetHost(CMarkdownEditorDlg* host) { m_host = host; }

protected:
	afx_msg void OnNcPaint();
	afx_msg void OnPaint();
	DECLARE_MESSAGE_MAP()

private:
	CMarkdownEditorDlg* m_host = nullptr;
};

class CMarkdownEditorDlg : public CDialogEx
{
public:
    CMarkdownEditorDlg(CWnd* pParent = nullptr);
    void SetStartupOpenFilePath(const CString& filePath);

#ifdef AFX_DESIGN_TIME
    enum { IDD = IDD_MARKDOWNEDITOR_DIALOG };
#endif

protected:
    virtual void DoDataExchange(CDataExchange* pDX);
    virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void OnCancel();
	virtual void OnOK();

protected:
    HICON m_hIcon;

    virtual BOOL OnInitDialog();
    afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
    afx_msg void OnPaint();
    afx_msg HCURSOR OnQueryDragIcon();
    afx_msg void OnBnClickedOpenFile();
    afx_msg void OnBnClickedModeSwitch();
    afx_msg void OnBnClickedThemeSwitch();
    afx_msg void OnEnChangeEdit();
    afx_msg void OnTimer(UINT_PTR nIDEvent);
    afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnEnterSizeMove();
	afx_msg void OnExitSizeMove();
	afx_msg void OnDropFiles(HDROP hDropInfo);
	afx_msg void OnEditContentLink(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnEditContentSelChanged(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnEditContentVScroll();
	afx_msg void OnEditContentHScroll();
	afx_msg void OnSidebarTreeCustomDraw(NMHDR* pNMHDR, LRESULT* pResult);
	void SearchOutlineByQuery(const CString& query);
    afx_msg LRESULT OnInitDragDrop(WPARAM wParam, LPARAM lParam);
    afx_msg void OnSidebarTabChanged(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSidebarItemChanged(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSidebarTreeClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSidebarTreeRClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnEnChangeFileSearch();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnDestroy();
	afx_msg void OnClose();
	afx_msg LRESULT OnMarkdownParseProgress(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnMarkdownParseComplete(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnDocumentLoadComplete(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnDocumentLoadProgress(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnPrepareDisplayTextComplete(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnOpenStartupFile(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnFindReplaceCmd(WPARAM wParam, LPARAM lParam);
	// Mermaid rendering is implemented as an in-process (no external dependencies) renderer.

    DECLARE_MESSAGE_MAP()

private:
    friend class CMermaidOverlayWnd;
	friend class CTableGridOverlayCtrl;
    friend class CSidebarSplitterWnd;
    friend class CSidebarTabCtrl;
    friend class CSidebarSearchEdit;
	enum class MermaidNodeShape : unsigned char
	{
		Rect = 0,
		Diamond = 1,
		RoundRect = 2,
		Stadium = 3,
		Circle = 4
	};
    struct MermaidBlock
    {
        long rawStart = 0;
        long rawEnd = 0;
        long mappedStart = 0;
        long mappedEnd = 0;
    };

	struct MermaidDiagram
	{
		long mappedStart = 0;
		long mappedEnd = 0;
		int lineCount = 0;
		size_t hash = 0;
		std::vector<std::pair<int, int>> edges;
		std::vector<CString> edgeLabels;
		std::vector<CRect> nodeRects;
		std::vector<MermaidNodeShape> nodeShapes;
		std::vector<CString> nodeLabels;
		int diagramWidth = 0;
		int diagramHeight = 0;
	};

    enum class FullTextReplaceTask
    {
        None,
        LoadFile,
        ApplyDisplayTextBeforeFormat,
        RestoreRawTextBeforeTextMode
    };

    CRichEditCtrl m_editContent;
    CMermaidOverlayWnd m_mermaidOverlay;
	CSidebarSplitterWnd m_sidebarSplitter;
    CButton m_btnOpenFile;
    CButton m_btnModeSwitch;
    CButton m_btnThemeSwitch;
    CStatic m_staticStatus;
	CProgressCtrl m_progressStatus;
	CStatic m_staticProgressText;
    CStatic m_staticOverlay;

    CSidebarTabCtrl m_tabSidebar;
    CTreeCtrl m_treeSidebar;
    CSidebarSearchEdit m_editFileSearch;
    CString m_fileSearchQuery;
	CString m_outlineSearchQuery;

	int m_sidebarActiveTab = 0;
	bool m_sidebarSelectionUpdating = false;
	bool m_suppressFileSearchEvent = false;
	HTREEITEM m_lastSidebarTreeSelItem = nullptr;
	DWORD m_lastSidebarTreeSelTick = 0;
	HTREEITEM m_outlineAutoFollowSelItem = nullptr;
	bool m_modeSwitchInProgress = false;
	bool m_themeDark = false;
	COLORREF m_themeBackground = RGB(255, 255, 255);
	COLORREF m_themeForeground = RGB(51, 51, 51);
	COLORREF m_themeSidebarBackground = RGB(248, 248, 248);
	CBrush m_brushDialogBackground;
	CBrush m_brushSidebarBackground;

    std::vector<CString> m_recentFiles;
    static constexpr int kMaxRecentFiles = 10;

    std::vector<CString> m_fileTreePaths;

    std::vector<std::pair<int, HTREEITEM>> m_outlineLineIndex;
    std::vector<std::pair<int, HTREEITEM>> m_outlineSourceIndex;
    std::map<HTREEITEM, int> m_outlineItemSourceIndex;
    std::vector<MarkdownOutlineEntry> m_outlineEntries;
    unsigned long long m_outlineEntriesGeneration = 0;
	bool m_outlineAutoSelectFirst = false;

    BOOL m_bMarkdownMode;
    CString m_strCurrentFile;
	CString m_currentFileName;
    CString m_strOriginalContent;

    bool m_suppressChangeEvent = false;
    bool m_formattingActive = false;
    bool m_formattingReschedule = false;
    bool m_waitCursorActive = false;
	bool m_parsingActive = false;
	bool m_parseReschedule = false;
	std::atomic_ullong m_parseGeneration{ 0 };
	std::atomic_ullong m_loadGeneration{ 0 };
	std::atomic_ullong m_displayGeneration{ 0 };
	int m_parseProgressPercent = 0;
	unsigned long long m_loadBytesRead = 0;
	unsigned long long m_loadBytesTotal = 0;
	bool m_loadingActive = false;
	bool m_hideDocumentUntilRenderReady = false;
	bool m_textDirty = false;
	std::wstring m_loadingText;
	size_t m_loadingTextPos = 0;
	FullTextReplaceTask m_fullTextReplaceTask = FullTextReplaceTask::None;
	bool m_startFormattingAfterFullTextReplace = false;
	bool m_startTextModeResetAfterFullTextReplace = false;
	bool m_loadingRestoreOriginal = false;
	bool m_loadingAfterRestoreResetTextMode = false;
	bool m_textModeResetActive = false;
	int m_overlaySyncAttempts = 0;
	long m_textModeResetPos = 0;
	std::wstring m_rawDocumentText;
	std::vector<long> m_displayToSourceIndexMap;
	std::vector<long> m_sourceToDisplayIndexMap;
	std::vector<long> m_richToSourceIndexMap;
	bool m_enableWrapSafeBlockMarkerElision = true;
	MarkdownParseResult m_deferredParseResult;
	std::vector<MarkdownOutlineEntry> m_deferredOutlineEntries;
	unsigned long long m_deferredGeneration = 0;
	bool m_hasDeferredParseResult = false;

	bool m_renderBusy = false;
	CString m_renderBusyHint;
	CString m_lastStatusText;
	bool m_statusProgressVisible = false;
	unsigned long m_lastStatusUpdateTick = 0;
	bool m_documentDirty = false;
	DWORD m_lastReadonlyEditPromptTick = 0;
	DWORD m_lastUserActivityTick = 0;
	DWORD m_lastAutoSaveErrorPromptTick = 0;
	bool m_autoSavedStatusVisible = false;
	bool m_editingHintVisible = false;
	bool m_externalChangePromptActive = false;
	bool m_externalFileMissingDetected = false;
	bool m_fileSignatureValid = false;
	ULONGLONG m_knownFileSize = 0;
	FILETIME m_knownFileWriteTime{};
	bool m_sidebarControlsDisabled = false;
	CBrush m_brushOverlay;
	CFindReplaceDialog* m_findReplaceDialog = nullptr;
	bool m_findReplaceMode = false;
	CString m_lastFindText;
	CString m_lastReplaceText;
	bool m_lastFindMatchCase = false;
	bool m_lastFindWholeWord = false;

    MarkdownParseResult m_pendingParseResult;
    std::wstring m_pendingSourceText;
    long m_pendingTextLength = 0;
    LONG m_pendingSelStart = 0;
    LONG m_pendingSelEnd = 0;
    int m_pendingFirstVisible = 0;
    size_t m_blockSpanIndex = 0;
    size_t m_inlineSpanIndex = 0;
    size_t m_hiddenSpanIndex = 0;
    size_t m_mermaidSpanIndex = 0;
    std::vector<MermaidBlock> m_mermaidBlocks;
    std::vector<MermaidDiagram> m_mermaidDiagrams;
	bool m_baseFormatApplied = false;
	long m_baseFormatPos = 0;

	int m_sidebarWidth = 240;
	int m_sidebarSplitterX = 0;
	int m_sidebarSplitterTop = 0;
	int m_sidebarSplitterBottom = 0;
	bool m_sidebarResizing = false;
	bool m_windowSizingMove = false;
	int m_sidebarResizeStartX = 0;
	int m_sidebarResizeStartWidth = 0;
	WNDPROC m_oldRichEditProc = nullptr;
	HDC m_wrapTargetHdc = nullptr;
	UINT_PTR m_mermaidOverlayTimer = 0;
	bool m_mouseSelecting = false;
	bool m_destroying = false;
	std::shared_ptr<std::atomic_bool> m_asyncCancelToken = std::make_shared<std::atomic_bool>(false);
	std::shared_ptr<std::atomic_int> m_asyncWorkerCount = std::make_shared<std::atomic_int>(0);

    CHARFORMAT2 m_normalFormat{};
    CHARFORMAT2 m_hiddenFormat{};
	CHARFORMAT2 m_blockMarkerElisionFormat{};
    CHARFORMAT2 m_h1Format{};
    CHARFORMAT2 m_h2Format{};
    CHARFORMAT2 m_h3Format{};
    CHARFORMAT2 m_h4Format{};
    CHARFORMAT2 m_h5Format{};
    CHARFORMAT2 m_h6Format{};
    CHARFORMAT2 m_boldFormat{};
    CHARFORMAT2 m_italicFormat{};
    CHARFORMAT2 m_strikeFormat{};
    CHARFORMAT2 m_codeBlockFormat{};
    CHARFORMAT2 m_inlineCodeFormat{};
    CHARFORMAT2 m_linkFormat{};
    CHARFORMAT2 m_quoteFormat{};
    CHARFORMAT2 m_listFormat{};
    CHARFORMAT2 m_tableHeaderFormat{};
    CHARFORMAT2 m_tableRowEvenFormat{};
    CHARFORMAT2 m_tableRowOddFormat{};
	CHARFORMAT2 m_hrFormat{};
	CHARFORMAT2 m_mermaidPlaceholderFormat{};

    // Cached fonts
    CFont m_fontButton;
    CFont m_fontStatus;
    CFont m_fontSidebar;
    CFont m_fontOverlay;

    void LoadFileContent(const CString& filePath);
    void ApplyMarkdownHighlight();
    void ApplyMarkdownFormatting();
    void CancelMarkdownFormatting();
    void ScheduleMarkdownFormatting(UINT delayMs);
    void BeginMarkdownFormatting();
    void ApplyMarkdownFormattingChunk();
    void FinishMarkdownFormatting();
    void StartMarkdownParsingAsync();
    void SetRenderBusy(bool busy, const CString& hint);
    void StartLoadDocumentAsync(const CString& filePath);
    void StartInsertLoadedText();
    void StartTextModeResetAsync();
    void PrepareMarkdownFormats();
    void ApplyTextMode();
    void ApplyModernStyling();
    void UpdateStatusText(bool force = false);
	int CalculateOpenProgressPercent() const;
	int CalculateRenderProgressPercent() const;
	void UpdateSidebarInteractivity();
	void ApplyEmojiFontFallbackToEdit(const std::wstring& sourceText, long richEditLength);
	void UpdateWindowTitle();
	void MarkDocumentDirty(bool dirty);
	void RefreshDirtyStateFromEditor();
	bool PromptSwitchToTextModeForEdit();
	void NoteUserActivity(bool clearAutoSavedNotice = true);
	void TryAutoSaveIfIdle();
	bool SaveCurrentDocument(bool saveAs = false, CString* errorMessage = nullptr);
	bool EnsureCanDiscardUnsavedChanges(const CString& nextPath, bool closingApp);
	bool HandleEditorTabIndent(bool outdent);
	bool CopySelectionWithMarkdownPayload(bool cut);
	bool PasteSelectionWithMarkdownPayload();
	void ShowFindReplaceDialog(bool replaceMode);
	bool FindNextInEditor(const CString& findText, bool matchCase, bool wholeWord, bool searchDown, bool wrapSearch);
	bool ReplaceCurrentInEditor(const CString& findText, const CString& replaceText, bool matchCase, bool wholeWord);
	int ReplaceAllInEditor(const CString& findText, const CString& replaceText, bool matchCase, bool wholeWord);
	bool QueryFileSignature(const CString& filePath, ULONGLONG& size, FILETIME& writeTime);
	void UpdateCurrentFileSignatureFromDisk();
	bool DetectExternalFileChange();
	void HandleExternalFileChange();
	void BeginSidebarResize(int xClient);
	void UpdateSidebarResize(int xClient);
	void EndSidebarResize();
	void UpdateMermaidOverlayRegion(bool forceInvalidate);
	void StartFullTextReplaceAsync(std::wstring text,
		FullTextReplaceTask task,
		bool startFormattingAfterReplace,
		bool startTextModeResetAfterReplace);
	void AdoptParseResultAndBeginFormatting(unsigned long long generation,
		MarkdownParseResult parseResult,
		std::vector<MarkdownOutlineEntry> outlineEntries);
	void FilterFileTreeByQuery(const CString& query);
    BOOL IsMarkdownFile(const CString& filePath);
    void HighlightMarkdownSyntax();
    void UpdateLayout(int cx, int cy);
    void UpdateEditorWrapWidth();
    void ApplyTheme(bool dark);

    void InitSidebar();
    void UpdateSidebar();
    void UpdateOutline();
    void SetSidebarTab(int tabIndex);
    void UpdateOutlineSelection();
	void UpdateOutlineSelectionByViewTop();
	void UpdateOutlineSelectionByCaret();
    void ExpandAllOutlineItems();
	void EnsureOutlineInitialSelection();
    void ExpandTreeItemRecursive(HTREEITEM item);
	void SelectCurrentFileInFileTree();
	bool JumpToOutlineTreeItem(HTREEITEM item);

    void OpenDocument(const CString& filePath);
    void LoadRecentFiles();
    void SaveRecentFiles() const;
    void AddRecentFile(const CString& filePath);
	bool RemoveRecentFile(const CString& filePath);
	void PromptSetDefaultMdAssociationIfNeeded();
	bool SetAsDefaultMdFileHandler(CString* errorMessage = nullptr);

    void BuildFileTree(const CString& filePath);
    void BuildOutlineTree();

    // Helpers
	int MapRichIndexToSourceIndex(int richIndex) const;
	int MapSourceIndexToRichIndex(int sourceIndex) const;
    CString ExtractLinkFromPosition(int position);
    CString ExtractLinkFromParseResult(int position);
    void OpenLinkWithDefaultTool(const CString& url);
    BOOL IsPositionInCodeBlock(int position);
	void ApplyImmersiveDarkTitleBar(bool enable);
	int GetZoomPercent() const;
	void SetZoomPercent(int percent);
	void UpdateMermaidDiagrams();
	void DrawMermaidOverlays(CDC& dc);
	void TouchMermaidBitmapCache(size_t hash);
	void TrimMermaidBitmapCache(size_t reserveForNewEntry, const std::set<size_t>* protectedHashes);
	void DrawHorizontalRules(CDC& dc);
	void UpdateTableListViewOverlays();
	void ClearTableListViewOverlays();
	void DrawTocOverlay(CDC& dc);
	void UpdateMermaidSpacing();
	void UpdateTocSpacing();
	bool HandleTocOverlayClick(CPoint overlayPoint);
	bool IsPointOnSidebarSplitter(CPoint point) const;

	std::vector<std::pair<int, int>> m_codeBlockRanges;
	std::vector<long> m_horizontalRuleStarts;
	struct TableOverlayRow
	{
		long start = 0;
		long end = 0;
		int tabStopCount = 0;
		LONG tabStopsTwips[32]{};
		bool isHeader = false;
		COLORREF backColor = CLR_INVALID;
	};
	std::vector<TableOverlayRow> m_tableOverlayRows;
	struct TableListViewOverlay
	{
		std::unique_ptr<CTableGridOverlayCtrl> gridCtrl;
		long anchorRowStart = -1;
		int cachedLeft = 0;
		int cachedRight = 0;
		int cachedHeight = 0;
		bool active = false;
	};
	std::vector<TableListViewOverlay> m_tableListViewOverlays;
	CFont m_tableListViewFont;
	int m_tableListViewFontZoom = 0;
	struct TocOverlayLine
	{
		std::wstring text;
		int targetSourceIndex = -1; // -1 means non-clickable
	};
	struct TocOverlayBlock
	{
		long mappedStart = 0;
		long mappedEnd = 0;
		std::vector<TocOverlayLine> lines;
	};
	std::vector<TocOverlayBlock> m_tocOverlayBlocks;
	struct TocHitRegion
	{
		CRect rect;
		int targetSourceIndex = -1;
	};
	std::vector<TocHitRegion> m_tocHitRegions;
	std::vector<long> m_tocSpaceAfterStarts;
	std::vector<long> m_tableSpaceAfterStarts;
	int m_tableSpaceAfterWidthPx = 0;
	int m_tableSpaceAfterZoom = 0;
	size_t m_tableSpaceAfterRowStamp = 0;

	struct HeadingRange
	{
		long start = 0;
		long end = 0;
		LONG yHeight = 0;
	};
	std::vector<HeadingRange> m_headingRanges;

	std::map<size_t, HBITMAP> m_mermaidBitmapCache;
	std::map<size_t, std::pair<int, int>> m_mermaidBitmapSize;
	std::map<size_t, HBITMAP> m_mermaidTransientBitmapCache;
	std::map<size_t, std::pair<int, int>> m_mermaidTransientBitmapSize;
	std::map<size_t, unsigned long long> m_mermaidBitmapLastUseSeq;
	std::set<size_t> m_mermaidFetchInFlight;
	unsigned long long m_mermaidRenderGeneration = 0;
	unsigned long long m_mermaidBitmapUseSeq = 0;
	static constexpr size_t kMaxMermaidBitmapCacheEntries = 64;
	std::vector<CRect> m_lastMermaidOverlayRects;
	std::map<size_t, LONG> m_mermaidSpaceAfterTwips;
	ULONG_PTR m_gdiplusToken = 0;
	CString m_startupOpenFilePath;

};
