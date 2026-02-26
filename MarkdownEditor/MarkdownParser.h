#pragma once

#include <string>
#include <utility>
#include <vector>
#include <functional>

enum class MarkdownBlockType
{
    Heading1,
    Heading2,
    Heading3,
    Heading4,
    Heading5,
    Heading6,
    BlockQuote,
    ListItem,
    CodeBlock,
    TableHeader,
    TableRowEven,
    TableRowOdd,
    HorizontalRule
};

enum class MarkdownInlineType
{
    Bold,
    BoldItalic,
    Italic,
    Strikethrough,
    Underline,
    InlineCode,
    Mark,
    Link,
    AutoLink
};

enum class MarkdownHiddenSpanKind
{
    Inline = 0,
    Block = 1
};

struct MarkdownHiddenSpan
{
    long start = 0;
    long end = 0;
    MarkdownHiddenSpanKind kind = MarkdownHiddenSpanKind::Inline;
};

struct MarkdownBlockSpan
{
    long start = 0;
    long end = 0;
    MarkdownBlockType type;
    int paraStartIndentTwips = 0;
    int paraFirstLineOffsetTwips = 0;

    // Table column alignment: tab stops in twips for this paragraph.
    // Only filled for TableHeader/TableRowEven/TableRowOdd.
    int tabStopCount = 0;
    int tabStopsTwips[32] = {};
    // 0=left, 1=center, 2=right
    int tabAlignments[32] = {};
};

struct MarkdownInlineSpan
{
    long start = 0;
    long end = 0;
    MarkdownInlineType type;
    std::wstring href;
};

struct MarkdownParseResult
{
    std::vector<MarkdownHiddenSpan> hiddenSpans;
    std::vector<MarkdownBlockSpan> blockSpans;
    std::vector<MarkdownInlineSpan> inlineSpans;
    std::vector<std::pair<long, long>> codeRanges;

    // Fenced code blocks tagged as mermaid (content range only, excluding fences).
    std::vector<std::pair<long, long>> mermaidBlocks;
};

class MarkdownParser
{
public:
    using ParseProgressCallback = std::function<void(int /*percent*/)>; // 0~100
    MarkdownParseResult Parse(const std::wstring& text) const;
    MarkdownParseResult Parse(const std::wstring& text, const ParseProgressCallback& progressCallback) const;
    MarkdownParseResult ParseBlocksOnly(const std::wstring& text) const;

private:
    struct Range
    {
        long start = 0;
        long end = 0;
    };

    static bool IsWhitespace(wchar_t ch);
    static bool IsLineBreak(wchar_t ch);
    static long FindLineEnd(const std::wstring& text, long start);
    static bool RangeIntersects(const std::vector<Range>& ranges, long start, long end);
    static bool RangeContains(const std::vector<Range>& ranges, long position);
    static void AddProtectedRange(std::vector<Range>& ranges, long start, long end);

    void ParseBlocks(const std::wstring& text,
                     MarkdownParseResult& result,
                     std::vector<Range>& protectedRanges) const;

    void ParseInline(const std::wstring& text,
                     MarkdownParseResult& result,
                     std::vector<Range>& protectedRanges) const;

    void ParseBoldItalic(const std::wstring& text,
                         MarkdownParseResult& result,
                         std::vector<Range>& protectedRanges) const;

    void ParseInlineCode(const std::wstring& text,
                         MarkdownParseResult& result,
                         std::vector<Range>& protectedRanges) const;

    void ParseInlineHtmlTags(const std::wstring& text,
                             MarkdownParseResult& result,
                             std::vector<Range>& protectedRanges) const;

    void ParseStrongEmphasis(const std::wstring& text,
                             MarkdownParseResult& result,
                             std::vector<Range>& protectedRanges) const;

    void ParseEmphasis(const std::wstring& text,
                       MarkdownParseResult& result,
                       std::vector<Range>& protectedRanges) const;

    void ParseStrikethrough(const std::wstring& text,
                            MarkdownParseResult& result,
                            std::vector<Range>& protectedRanges) const;

    void ParseLinks(const std::wstring& text,
                    MarkdownParseResult& result,
                    std::vector<Range>& protectedRanges) const;

    void ParseAutoLinks(const std::wstring& text,
                        MarkdownParseResult& result,
                        std::vector<Range>& protectedRanges) const;

    void ResetProgressState(const ParseProgressCallback& callback) const;
    void BeginProgressPhase(int basePercent, int spanPercent, size_t totalUnits) const;
    void TickProgress(size_t processedUnits, bool force = false) const;
    void EmitProgress(int percent, bool force = false) const;

    mutable ParseProgressCallback m_progressCallback;
    mutable int m_progressLastPercent = -1;
    mutable int m_progressPhaseBase = 0;
    mutable int m_progressPhaseSpan = 0;
    mutable size_t m_progressPhaseTotal = 0;
    mutable size_t m_progressPhaseLastUnits = 0;
    mutable unsigned long m_progressLastTick = 0;

    // NOTE: Inline parsing is implemented via md4c (library-backed tokenizer).
};
