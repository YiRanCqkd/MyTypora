#include "pch.h"
#include "MarkdownParser.h"

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <eh.h>
#include <memory>

namespace
{
    constexpr wchar_t kTrailingPunct[] = L".,;:!?)]}";

    struct Md4cSpanFrame
    {
        MD_SPANTYPE type = MD_SPAN_EM;
        long start = -1;
        long end = -1;
        std::wstring href;
        bool isAutoLink = false;
        bool hasText = false;
    };

    struct Md4cCollector
    {
        const std::wstring* source = nullptr;
        const wchar_t* base = nullptr;
        size_t length = 0;
        MarkdownParseResult* result = nullptr;
        std::vector<Md4cSpanFrame> spanStack;
    };

    struct Md4cSehException
    {
        unsigned int code = 0;
        explicit Md4cSehException(unsigned int value) : code(value) {}
    };

    static void __cdecl Md4cSehTranslator(unsigned int code, EXCEPTION_POINTERS* /*info*/)
    {
        throw Md4cSehException(code);
    }

    struct Md4cSehTranslatorScope
    {
        _se_translator_function previous = nullptr;

        Md4cSehTranslatorScope() { previous = _set_se_translator(Md4cSehTranslator); }
        ~Md4cSehTranslatorScope() { _set_se_translator(previous); }
    };

    std::wstring Md4cAttributeToWString(const MD_ATTRIBUTE& attr)
    {
        if (attr.text == nullptr || attr.size == 0)
            return {};

        // Attributes can be backed by internal MD4C buffers (not necessarily the input string).
        // Only apply a size cap to avoid pathological allocations/overreads.
        constexpr size_t kMaxAttributeChars = 32 * 1024;
        const size_t safeLen = (attr.size > kMaxAttributeChars) ? kMaxAttributeChars : static_cast<size_t>(attr.size);
        if (safeLen == 0)
            return {};

        const auto* text = reinterpret_cast<const wchar_t*>(attr.text);
        return std::wstring(text, text + safeLen);
    }

    inline void AddHiddenSpan(MarkdownParseResult& result, long start, long end, MarkdownHiddenSpanKind kind)
    {
        if (start < end)
            result.hiddenSpans.push_back({ start, end, kind });
    }

    inline void AddHiddenSpanInline(MarkdownParseResult& result, long start, long end)
    {
        AddHiddenSpan(result, start, end, MarkdownHiddenSpanKind::Inline);
    }

    inline void AddHiddenSpanBlock(MarkdownParseResult& result, long start, long end)
    {
        AddHiddenSpan(result, start, end, MarkdownHiddenSpanKind::Block);
    }

    void Md4cAddHiddenSpan(MarkdownParseResult& result, long start, long end)
    {
        AddHiddenSpanInline(result, start, end);
    }

    void Md4cHideInlineMarkers(const std::wstring& source,
                               MarkdownParseResult& result,
                               const Md4cSpanFrame& frame)
    {
        const long length = static_cast<long>(source.length());
        if (frame.start <= 0 || frame.end < 0 || frame.start > length)
            return;

        auto hidePair = [&](long openStart, long openEnd, long closeStart, long closeEnd) {
            Md4cAddHiddenSpan(result, openStart, openEnd);
            Md4cAddHiddenSpan(result, closeStart, closeEnd);
        };

        switch (frame.type)
        {
        case MD_SPAN_EM:
        {
            const long openPos = frame.start - 1;
            if (openPos < 0 || openPos >= length)
                break;
            const wchar_t marker = source[static_cast<size_t>(openPos)];
            if (marker != L'*' && marker != L'_')
                break;
            if (frame.end >= 0 && frame.end < length && source[static_cast<size_t>(frame.end)] == marker)
                hidePair(openPos, openPos + 1, frame.end, frame.end + 1);
            break;
        }
        case MD_SPAN_STRONG:
        {
            const long openPos = frame.start - 2;
            if (openPos < 0 || openPos + 1 >= length)
                break;
            const wchar_t marker = source[static_cast<size_t>(frame.start - 1)];
            if ((marker != L'*' && marker != L'_') || source[static_cast<size_t>(openPos)] != marker)
                break;
            if (frame.end + 1 < length && source[static_cast<size_t>(frame.end)] == marker &&
                source[static_cast<size_t>(frame.end + 1)] == marker)
            {
                hidePair(openPos, frame.start, frame.end, frame.end + 2);
            }
            break;
        }
        case MD_SPAN_DEL:
        {
            const long openPos = frame.start - 2;
            if (openPos < 0 || openPos + 1 >= length)
                break;
            if (source[static_cast<size_t>(openPos)] != L'~' || source[static_cast<size_t>(openPos + 1)] != L'~')
                break;
            if (frame.end + 1 < length && source[static_cast<size_t>(frame.end)] == L'~' &&
                source[static_cast<size_t>(frame.end + 1)] == L'~')
            {
                hidePair(openPos, frame.start, frame.end, frame.end + 2);
            }
            break;
        }
        case MD_SPAN_CODE:
        {
            long openLen = 0;
            for (long pos = frame.start - 1; pos >= 0 && source[static_cast<size_t>(pos)] == L'`'; --pos)
                ++openLen;
            long closeLen = 0;
            for (long pos = frame.end; pos < length && source[static_cast<size_t>(pos)] == L'`'; ++pos)
                ++closeLen;
            const long fenceLen = (openLen > 0 && closeLen > 0) ? (std::min)(openLen, closeLen) : 0;
            if (fenceLen > 0)
                hidePair(frame.start - fenceLen, frame.start, frame.end, frame.end + fenceLen);
            break;
        }
        case MD_SPAN_A:
        {
            if (frame.start <= 0)
                break;
            const long openPos = frame.start - 1;
            const wchar_t openCh = source[static_cast<size_t>(openPos)];

            if (frame.isAutoLink)
            {
                if (openCh == L'<' && frame.end < length && source[static_cast<size_t>(frame.end)] == L'>')
                    hidePair(openPos, openPos + 1, frame.end, frame.end + 1);
                break;
            }

            // Standard inline link: [text](dest)
            if (openCh != L'[')
                break;
            Md4cAddHiddenSpan(result, openPos, openPos + 1);

            long closeBracket = frame.end;
            while (closeBracket < length && source[static_cast<size_t>(closeBracket)] != L']' &&
                   source[static_cast<size_t>(closeBracket)] != L'\r' && source[static_cast<size_t>(closeBracket)] != L'\n')
            {
                ++closeBracket;
            }
            if (closeBracket >= length || source[static_cast<size_t>(closeBracket)] != L']')
                break;

            long parenStart = closeBracket + 1;
            while (parenStart < length && source[static_cast<size_t>(parenStart)] == L' ')
                ++parenStart;

            if (parenStart >= length || source[static_cast<size_t>(parenStart)] != L'(')
            {
                Md4cAddHiddenSpan(result, closeBracket, closeBracket + 1);
                break;
            }

            int depth = 0;
            long parenEnd = parenStart;
            for (; parenEnd < length; ++parenEnd)
            {
                wchar_t ch = source[static_cast<size_t>(parenEnd)];
                if (ch == L'\r' || ch == L'\n')
                    break;
                if (ch == L'(')
                    ++depth;
                else if (ch == L')')
                {
                    --depth;
                    if (depth == 0)
                        break;
                }
            }

            if (parenEnd < length && source[static_cast<size_t>(parenEnd)] == L')')
                Md4cAddHiddenSpan(result, closeBracket, parenEnd + 1);
            else
                Md4cAddHiddenSpan(result, closeBracket, closeBracket + 1);
            break;
        }
        default:
            break;
        }
    }

    int Md4cEnterSpan(MD_SPANTYPE type, void* detail, void* userdata)
    {
        auto* collector = reinterpret_cast<Md4cCollector*>(userdata);
        if (!collector || !collector->source || !collector->result)
            return 0;

        Md4cSpanFrame frame;
        frame.type = type;

        if (type == MD_SPAN_A && detail != nullptr)
        {
            const auto* a = reinterpret_cast<const MD_SPAN_A_DETAIL*>(detail);
            frame.href = Md4cAttributeToWString(a->href);
            frame.isAutoLink = (a->is_autolink != 0);
        }

        collector->spanStack.push_back(std::move(frame));
        return 0;
    }

    int Md4cLeaveSpan(MD_SPANTYPE type, void* /*detail*/, void* userdata)
    {
        auto* collector = reinterpret_cast<Md4cCollector*>(userdata);
        if (!collector || !collector->source || !collector->result)
            return 0;
        if (collector->spanStack.empty())
            return 0;

        Md4cSpanFrame frame = std::move(collector->spanStack.back());
        collector->spanStack.pop_back();

        if (frame.type != type)
            frame.type = type;
        if (!frame.hasText || frame.start < 0 || frame.end <= frame.start)
            return 0;

        MarkdownInlineSpan span;
        span.start = frame.start;
        span.end = frame.end;

        bool emit = true;
        switch (frame.type)
        {
        case MD_SPAN_EM:
            span.type = MarkdownInlineType::Italic;
            break;
        case MD_SPAN_STRONG:
            span.type = MarkdownInlineType::Bold;
            break;
        case MD_SPAN_U:
            span.type = MarkdownInlineType::Underline;
            break;
        case MD_SPAN_DEL:
            span.type = MarkdownInlineType::Strikethrough;
            break;
        case MD_SPAN_CODE:
            span.type = MarkdownInlineType::InlineCode;
            collector->result->codeRanges.push_back({ span.start, span.end });
            break;
        case MD_SPAN_A:
            span.type = frame.isAutoLink ? MarkdownInlineType::AutoLink : MarkdownInlineType::Link;
            span.href = frame.href;
            break;
        default:
            emit = false;
            break;
        }

        if (emit)
        {
            collector->result->inlineSpans.push_back(std::move(span));
            Md4cHideInlineMarkers(*collector->source, *collector->result, frame);
        }

        return 0;
    }

    int Md4cText(MD_TEXTTYPE type, const MD_CHAR* text, MD_SIZE size, void* userdata)
    {
        auto* collector = reinterpret_cast<Md4cCollector*>(userdata);
        if (!collector || !collector->source || !collector->result)
            return 0;
        if (text == nullptr || size == 0)
            return 0;

        // Minimal text-level mapping.
        if (type == MD_TEXT_BR || type == MD_TEXT_SOFTBR)
        {
            // md4c may pass a pointer not belonging to the original buffer (e.g. a literal "\n").
            // We cannot reliably hide it; just ignore.
            return 0;
        }

        // Inline HTML tags: minimal support by hiding the raw tag text.
        // md4c reports them as MD_TEXT_HTML.
        if (type == MD_TEXT_HTML)
        {
            const auto* ptr = reinterpret_cast<const wchar_t*>(text);
            const wchar_t* base = collector->base;
            const wchar_t* limit = collector->base + collector->length;
            if (ptr >= base && ptr < limit)
            {
                long start = static_cast<long>(ptr - base);
                long end = start + static_cast<long>(size);
                if (start < 0) start = 0;
                if (end < start) end = start;
                if (end > static_cast<long>(collector->length))
                    end = static_cast<long>(collector->length);
                AddHiddenSpanInline(*collector->result, start, end);
            }
            return 0;
        }

        if (collector->spanStack.empty())
            return 0;

        const auto* ptr = reinterpret_cast<const wchar_t*>(text);
        const wchar_t* base = collector->base;
        const wchar_t* limit = collector->base + collector->length;

        // md4c sometimes passes pointers to string literals (e.g. "\n" for soft/hard breaks)
        // which are not part of the original input buffer. In that case we cannot map to a
        // stable offset and must ignore the callback.
        if (ptr < base || ptr >= limit)
            return 0;

        long start = static_cast<long>(ptr - base);
        long end = start + static_cast<long>(size);

        if (start < 0)
            start = 0;
        if (end < start)
            end = start;
        if (end > static_cast<long>(collector->length))
            end = static_cast<long>(collector->length);

        for (auto& frame : collector->spanStack)
        {
            if (frame.start < 0)
                frame.start = start;
            frame.end = (std::max)(frame.end, end);
            frame.hasText = true;
        }

        return 0;
    }

    void ParseInlineWithMd4c(const std::wstring& text, MarkdownParseResult& result)
    {
        Md4cCollector collector;
        collector.source = &text;
        collector.base = text.c_str();
        collector.length = text.length();
        collector.result = &result;

        MD_PARSER parser{};
        parser.abi_version = 0;
        parser.flags = (MD_DIALECT_GITHUB | MD_FLAG_UNDERLINE);
        parser.enter_span = &Md4cEnterSpan;
        parser.leave_span = &Md4cLeaveSpan;
        parser.text = &Md4cText;

        int rc = 0;
        try
        {
            // md4c is C code; be defensive in case it raises SEH exceptions on malformed input.
            Md4cSehTranslatorScope sehScope;
            rc = md_parse(reinterpret_cast<const MD_CHAR*>(text.c_str()),
                          static_cast<MD_SIZE>(text.length()),
                          &parser,
                          &collector);
        }
        catch (const Md4cSehException&)
        {
            return;
        }
        catch (...)
        {
            return;
        }

        if (rc != 0)
            return;
    }
}

// NOTE:
// The rest of this file contains the legacy inline parsing implementation.
// We keep it as the primary inline parser because it operates on source
// indices, which is required for RichEdit formatting.

MarkdownParseResult MarkdownParser::Parse(const std::wstring& text) const
{
    return Parse(text, ParseProgressCallback());
}

MarkdownParseResult MarkdownParser::Parse(const std::wstring& text, const ParseProgressCallback& progressCallback) const
{
    MarkdownParseResult result;
    if (text.empty())
        return result;

    ResetProgressState(progressCallback);
    EmitProgress(1, true);

    std::vector<Range> protectedRanges;
    const size_t safeTotalUnits = (std::max)(text.length(), static_cast<size_t>(1));

    // 阶段1：块级解析（1%~60%）
    BeginProgressPhase(1, 59, safeTotalUnits);
    ParseBlocks(text, result, protectedRanges);
    TickProgress(safeTotalUnits, true);

    // 阶段2：行内解析（60%~99%）
    BeginProgressPhase(60, 39, safeTotalUnits);
    ParseInline(text, result, protectedRanges);
    TickProgress(safeTotalUnits, true);

    auto sortByStart = [](auto& spans) {
        std::sort(spans.begin(), spans.end(), [](const auto& lhs, const auto& rhs) {
            if (lhs.start == rhs.start)
                return lhs.end < rhs.end;
            return lhs.start < rhs.start;
        });
    };

    sortByStart(result.hiddenSpans);
    sortByStart(result.blockSpans);
    sortByStart(result.inlineSpans);
    std::sort(result.codeRanges.begin(), result.codeRanges.end(), [](const auto& lhs, const auto& rhs) {
        if (lhs.first == rhs.first)
            return lhs.second < rhs.second;
        return lhs.first < rhs.first;
    });

    EmitProgress(100, true);
    ResetProgressState(ParseProgressCallback());
    return result;
}

MarkdownParseResult MarkdownParser::ParseBlocksOnly(const std::wstring& text) const
{
    MarkdownParseResult result;
    if (text.empty())
        return result;

    std::vector<Range> protectedRanges;
    ParseBlocks(text, result, protectedRanges);

    auto sortByStart = [](auto& spans) {
        std::sort(spans.begin(), spans.end(), [](const auto& lhs, const auto& rhs) {
            if (lhs.start == rhs.start)
                return lhs.end < rhs.end;
            return lhs.start < rhs.start;
        });
    };

    sortByStart(result.hiddenSpans);
    sortByStart(result.blockSpans);
    sortByStart(result.inlineSpans);
    std::sort(result.codeRanges.begin(), result.codeRanges.end(), [](const auto& lhs, const auto& rhs) {
        if (lhs.first == rhs.first)
            return lhs.second < rhs.second;
        return lhs.first < rhs.first;
    });

    return result;
}

bool MarkdownParser::IsWhitespace(wchar_t ch)
{
    return ch == L' ' || ch == L'\t' || std::iswspace(ch) != 0;
}

bool MarkdownParser::IsLineBreak(wchar_t ch)
{
    return ch == L'\r' || ch == L'\n';
}

long MarkdownParser::FindLineEnd(const std::wstring& text, long start)
{
    const long length = static_cast<long>(text.length());
    long pos = start;
    while (pos < length && !IsLineBreak(text[pos]))
        ++pos;
    return pos;
}

bool MarkdownParser::RangeIntersects(const std::vector<Range>& ranges, long start, long end)
{
    if (start >= end)
        return true;

    for (const auto& range : ranges)
    {
        if (start < range.end && end > range.start)
            return true;
    }
    return false;
}

bool MarkdownParser::RangeContains(const std::vector<Range>& ranges, long position)
{
    for (const auto& range : ranges)
    {
        if (position >= range.start && position < range.end)
            return true;
    }
    return false;
}

void MarkdownParser::AddProtectedRange(std::vector<Range>& ranges, long start, long end)
{
    if (start >= end)
        return;
    ranges.push_back({ start, end });
}

void MarkdownParser::ResetProgressState(const ParseProgressCallback& callback) const
{
    m_progressCallback = callback;
    m_progressLastPercent = -1;
    m_progressPhaseBase = 0;
    m_progressPhaseSpan = 0;
    m_progressPhaseTotal = 0;
    m_progressPhaseLastUnits = 0;
    m_progressLastTick = 0;
}

void MarkdownParser::BeginProgressPhase(int basePercent, int spanPercent, size_t totalUnits) const
{
    m_progressPhaseBase = (std::max)(0, (std::min)(100, basePercent));
    m_progressPhaseSpan = (std::max)(0, spanPercent);
    m_progressPhaseTotal = (std::max)(static_cast<size_t>(1), totalUnits);
    m_progressPhaseLastUnits = 0;
    EmitProgress(m_progressPhaseBase, false);
}

void MarkdownParser::TickProgress(size_t processedUnits, bool force) const
{
    if (!m_progressCallback)
        return;

    if (m_progressPhaseTotal == 0)
        m_progressPhaseTotal = 1;
    if (processedUnits > m_progressPhaseTotal)
        processedUnits = m_progressPhaseTotal;

    const unsigned long now = ::GetTickCount();
    if (!force)
    {
        const size_t deltaUnits = (processedUnits > m_progressPhaseLastUnits)
            ? (processedUnits - m_progressPhaseLastUnits) : 0;
        if (deltaUnits < 2048 && m_progressLastTick != 0 && now - m_progressLastTick < 80)
            return;
    }

    m_progressPhaseLastUnits = processedUnits;
    int percent = m_progressPhaseBase;
    if (m_progressPhaseSpan > 0)
    {
        const size_t numerator = processedUnits * static_cast<size_t>(m_progressPhaseSpan);
        percent += static_cast<int>(numerator / m_progressPhaseTotal);
    }
    EmitProgress(percent, force);
}

void MarkdownParser::EmitProgress(int percent, bool force) const
{
    if (!m_progressCallback)
        return;

    int clamped = (std::max)(0, (std::min)(100, percent));
    if (!force && clamped <= m_progressLastPercent)
        return;
    if (force && clamped < m_progressLastPercent)
        clamped = m_progressLastPercent;

    m_progressLastPercent = clamped;
    m_progressLastTick = ::GetTickCount();
    m_progressCallback(clamped);
}

void MarkdownParser::ParseBlocks(const std::wstring& text,
                                 MarkdownParseResult& result,
                                 std::vector<Range>& protectedRanges) const
{
    const size_t length = text.length();
    size_t pos = 0;

    struct PendingCodeBlock
    {
        long fenceStart = 0;
        long fenceEnd = 0;
        long contentStart = 0;
        wchar_t fenceChar = 0;
        bool active = false;
        bool isMermaid = false;
    } pending;

    while (pos < length)
    {
        size_t lineStartPos = pos;
        size_t cursor = pos;
        while (cursor < length && text[cursor] != L'\n' && text[cursor] != L'\r')
            ++cursor;

        size_t lineEndPos = cursor;
        size_t newlineLength = 0;
        if (cursor < length)
        {
            if (text[cursor] == L'\r' && cursor + 1 < length && text[cursor + 1] == L'\n')
            {
                newlineLength = 2;
            }
            else
            {
                newlineLength = 1;
            }
        }

        std::wstring line = text.substr(lineStartPos, lineEndPos - lineStartPos);
        size_t trimmedPos = 0;
        while (trimmedPos < line.size() && IsWhitespace(line[trimmedPos]))
            ++trimmedPos;

        long absoluteLineStart = static_cast<long>(lineStartPos);
        long absoluteLineEnd = static_cast<long>(lineEndPos);

        bool inCodeFence = pending.active;
        bool processed = false;

        auto indentColumnsBetween = [&](size_t startIndex, size_t endIndex) -> size_t {
            size_t columns = 0;
            if (endIndex > line.size())
                endIndex = line.size();
            for (size_t i = startIndex; i < endIndex; ++i)
            {
                wchar_t ch = line[i];
                if (ch == L' ')
                {
                    columns += 1;
                }
                else if (ch == L'\t')
                {
                    size_t step = 4 - (columns % 4);
                    columns += (step == 0) ? 4 : step;
                }
                else
                {
                    break;
                }
            }
            return columns;
        };

        // Handle code fences (``` or ~~~). For opening fences, hide the whole line
        // so language tags (```cpp) do not appear in the WYSIWYM view.
        if (!line.empty())
        {
            size_t fenceCheckPos = trimmedPos;
            wchar_t fenceChar = 0;
            size_t fenceCount = 0;
            if (fenceCheckPos < line.size() && (line[fenceCheckPos] == L'`' || line[fenceCheckPos] == L'~'))
            {
                fenceChar = line[fenceCheckPos];
                while (fenceCheckPos < line.size() && line[fenceCheckPos] == fenceChar)
                {
                    ++fenceCount;
                    ++fenceCheckPos;
                }
            }

            if (fenceCount >= 3)
            {
                if (!pending.active)
                {
                    // Determine fenced code language tag (e.g. ```cpp, ```mermaid).
                    bool mermaidFence = false;
                    {
                        size_t langStart = fenceCheckPos;
                        while (langStart < line.size() && IsWhitespace(line[langStart]))
                            ++langStart;
                        size_t langEnd = langStart;
                        while (langEnd < line.size() && !IsWhitespace(line[langEnd]))
                            ++langEnd;
                        if (langEnd > langStart)
                        {
                            std::wstring lang = line.substr(langStart, langEnd - langStart);
                            for (auto& ch : lang)
                                ch = static_cast<wchar_t>(towlower(ch));
                            if (lang == L"mermaid")
                                mermaidFence = true;
                        }
                    }

                    pending.active = true;
                    pending.fenceChar = fenceChar;
                    pending.fenceStart = absoluteLineStart + static_cast<long>(trimmedPos);
                    pending.fenceEnd = absoluteLineEnd;
                    pending.contentStart = static_cast<long>(lineEndPos + newlineLength);
                    pending.isMermaid = mermaidFence;
                    processed = true;
                }
                else
                {
                    // Only close with the same fence char.
                    if (fenceChar != pending.fenceChar)
                    {
                        // Treat as normal line inside code fence
                        processed = false;
                    }
                    else
                    {
                        long closeFenceStart = absoluteLineStart + static_cast<long>(trimmedPos);
                        long closeFenceEnd = absoluteLineEnd;

                        AddHiddenSpanBlock(result, pending.fenceStart, pending.fenceEnd);
                        AddHiddenSpanBlock(result, closeFenceStart, closeFenceEnd);
                        AddProtectedRange(protectedRanges, pending.fenceStart, pending.fenceEnd);
                        AddProtectedRange(protectedRanges, closeFenceStart, closeFenceEnd);

                        if (pending.contentStart < closeFenceStart)
                        {
                        // Apply a small left padding to mimic code block inset.
                        result.blockSpans.push_back({ pending.contentStart, closeFenceStart, MarkdownBlockType::CodeBlock, 240, 0 });
                            result.codeRanges.push_back({ pending.contentStart, closeFenceStart });
                            AddProtectedRange(protectedRanges, pending.contentStart, closeFenceStart);

                            if (pending.isMermaid)
                            {
                                result.mermaidBlocks.push_back({ pending.contentStart, closeFenceStart });
                            }
                        }

                        pending = PendingCodeBlock{};
                        processed = true;
                    }
                }
            }
        }

        if (!processed && pending.active)
        {
            // Inside code block, skip other parsing
            pos = lineEndPos + newlineLength;
            TickProgress(pos, false);
            continue;
        }

        if (!processed && !line.empty())
        {
            // Tables (GFM style)
            // Detect header row when the next line is a separator like: | --- | :---: | ---: |
            // Render by hiding the separator line, and apply zebra background to rows.
            // Column alignment is handled in the UI layer by replacing inner '|' with '\t'
            // and applying paragraph tab stops.
            auto isTableRowCandidate = [&](const std::wstring& row) {
                size_t pipeCount = 0;
                for (wchar_t ch : row)
                {
                    if (ch == L'|')
                        ++pipeCount;
                }
                // 支持无外侧竖线的两列表写法：a | b
                return pipeCount >= 1;
            };

            auto countTableColumns = [&](const std::wstring& row) -> int {
                // Trim whitespace.
                size_t s = 0;
                while (s < row.size() && IsWhitespace(row[s]))
                    ++s;
                size_t e = row.size();
                while (e > s && IsWhitespace(row[e - 1]))
                    --e;
                if (e <= s)
                    return 0;

                // Strip outer pipes if present.
                if (row[s] == L'|')
                    ++s;
                if (e > s && row[e - 1] == L'|')
                    --e;
                if (e <= s)
                    return 0;

                int pipeCount = 0;
                for (size_t i = s; i < e; ++i)
                {
                    if (row[i] == L'|')
                        ++pipeCount;
                }
                return pipeCount + 1;
            };

            auto isTableSeparatorLine = [&](const std::wstring& sep) {
                // ignore leading/trailing whitespace
                size_t s = 0;
                while (s < sep.size() && IsWhitespace(sep[s]))
                    ++s;
                size_t e = sep.size();
                while (e > s && IsWhitespace(sep[e - 1]))
                    --e;

                if (e <= s)
                    return false;

                // Must contain at least one '|'
                bool hasPipe = false;
                for (size_t i = s; i < e; ++i)
                {
                    if (sep[i] == L'|')
                    {
                        hasPipe = true;
                        break;
                    }
                }
                if (!hasPipe)
                    return false;

                // Allow only: '|', '-', ':', spaces, tabs
                bool hasDash = false;
                for (size_t i = s; i < e; ++i)
                {
                    const wchar_t ch = sep[i];
                    if (ch == L'-')
                    {
                        hasDash = true;
                        continue;
                    }
                    if (ch == L'|' || ch == L':' || IsWhitespace(ch))
                        continue;
                    return false;
                }
                if (!hasDash)
                    return false;

                // Require at least 3 dashes total to avoid false positives.
                int dashCount = 0;
                for (size_t i = s; i < e; ++i)
                {
                    if (sep[i] == L'-')
                        ++dashCount;
                }
                return dashCount >= 3;
            };

            if (isTableRowCandidate(line))
            {
                // lookahead next line
                size_t nextStart = lineEndPos + newlineLength;
                if (nextStart < length)
                {
                    size_t nextCursor = nextStart;
                    while (nextCursor < length && text[nextCursor] != L'\n' && text[nextCursor] != L'\r')
                        ++nextCursor;

                    size_t nextLineEndPos = nextCursor;
                    size_t nextNewlineLength = 0;
                    if (nextCursor < length)
                    {
                        if (text[nextCursor] == L'\r' && nextCursor + 1 < length && text[nextCursor + 1] == L'\n')
                            nextNewlineLength = 2;
                        else
                            nextNewlineLength = 1;
                    }

                    std::wstring nextLine = text.substr(nextStart, nextLineEndPos - nextStart);
                    if (isTableSeparatorLine(nextLine))
                    {
                        auto addTablePaddingHiddenSpans = [&](long rowStartAbs, const std::wstring& rowLine) {
                            if (rowLine.empty())
                                return;

                            size_t s = 0;
                            while (s < rowLine.size() && IsWhitespace(rowLine[s]))
                                ++s;
                            size_t e = rowLine.size();
                            while (e > s && IsWhitespace(rowLine[e - 1]))
                                --e;
                            if (e <= s)
                                return;

                            for (size_t i = s; i < e; ++i)
                            {
                                if (rowLine[i] != L'|')
                                    continue;

                                // Hide whitespace before the pipe.
                                size_t left = i;
                                while (left > s && IsWhitespace(rowLine[left - 1]) && rowLine[left - 1] != L'|')
                                    --left;
                                if (left < i)
                                    AddHiddenSpanInline(result, rowStartAbs + static_cast<long>(left), rowStartAbs + static_cast<long>(i));

                                // Hide whitespace after the pipe.
                                size_t right = i + 1;
                                while (right < e && IsWhitespace(rowLine[right]) && rowLine[right] != L'|')
                                    ++right;
                                if (right > i + 1)
                                    AddHiddenSpanInline(result, rowStartAbs + static_cast<long>(i + 1), rowStartAbs + static_cast<long>(right));
                            }
                        };

						constexpr int kMaxTabStops = 32;

						auto splitTableCells = [&](const std::wstring& row) {
							std::vector<std::wstring> cells;
							size_t s = 0;
							while (s < row.size() && IsWhitespace(row[s]))
								++s;
							size_t e = row.size();
							while (e > s && IsWhitespace(row[e - 1]))
								--e;
							if (e <= s)
								return cells;

							if (row[s] == L'|')
								++s;
							if (e > s && row[e - 1] == L'|')
								--e;
							if (e <= s)
								return cells;

							auto isEscaped = [&](size_t index) {
								if (index <= s)
									return false;
								size_t backslashes = 0;
								size_t p = index;
								while (p > s)
								{
									--p;
									if (row[p] == L'\\')
										++backslashes;
									else
										break;
								}
								return (backslashes % 2) == 1;
							};

							size_t cellStart = s;
							bool inCodeSpan = false;
							for (size_t i = s; i < e; ++i)
							{
								if (row[i] == L'`' && !isEscaped(i))
								{
									inCodeSpan = !inCodeSpan;
									continue;
								}

								if (row[i] == L'|' && !inCodeSpan && !isEscaped(i))
								{
									size_t cs = cellStart;
									size_t ce = i;
									while (cs < ce && IsWhitespace(row[cs]))
										++cs;
									while (ce > cs && IsWhitespace(row[ce - 1]))
										--ce;
									cells.push_back((ce > cs) ? row.substr(cs, ce - cs) : std::wstring());
									cellStart = i + 1;
								}
							}
							size_t cs = cellStart;
							size_t ce = e;
							while (cs < ce && IsWhitespace(row[cs]))
								++cs;
							while (ce > cs && IsWhitespace(row[ce - 1]))
								--ce;
							cells.push_back((ce > cs) ? row.substr(cs, ce - cs) : std::wstring());
							return cells;
						};

						auto parseTableAlignments = [&](const std::wstring& sepRow) {
							std::vector<int> aligns;
							auto cells = splitTableCells(sepRow);
							aligns.reserve(cells.size());
							for (auto cell : cells)
							{
								size_t s = 0;
								while (s < cell.size() && IsWhitespace(cell[s]))
									++s;
								size_t e = cell.size();
								while (e > s && IsWhitespace(cell[e - 1]))
									--e;

								int align = 0; // left
								if (e > s)
								{
									const bool leftColon = (cell[s] == L':');
									const bool rightColon = (cell[e - 1] == L':');
									if (leftColon && rightColon)
										align = 1; // center
									else if (rightColon)
										align = 2; // right
									else
										align = 0; // left (default, including leftColon-only)
								}
								aligns.push_back(align);
							}
							return aligns;
						};

						auto cellDisplayUnits = [&](const std::wstring& cell) {
							int units = 0;
							for (wchar_t ch : cell)
							{
								if (ch == L'\t')
								{
									units += 4;
									continue;
								}
								if (ch <= 0x7F)
								{
									units += 1;
								}
								else
								{
									// CJK and other wide glyphs.
									units += 2;
								}
							}
							return units;
						};

						struct TableRowInfo
						{
							long start = 0;
							long end = 0;
							std::wstring text;
							bool isHeader = false;
						};

						std::vector<TableRowInfo> tableRows;
						// header row
						if (absoluteLineStart < absoluteLineEnd)
						{
							TableRowInfo info{};
							info.start = absoluteLineStart;
							info.end = absoluteLineEnd;
							info.text = line;
							info.isHeader = true;
							tableRows.push_back(std::move(info));
						}

                        // hide separator line
                        long sepStartAbs = static_cast<long>(nextStart);
                        long sepEndAbs = static_cast<long>(nextLineEndPos + nextNewlineLength);
                        if (sepEndAbs > static_cast<long>(length))
                            sepEndAbs = static_cast<long>(length);
                        if (sepStartAbs < sepEndAbs)
                        {
                            AddHiddenSpanInline(result, sepStartAbs, sepEndAbs);
                            AddProtectedRange(protectedRanges, sepStartAbs, sepEndAbs);
                        }

						// collect following table rows until a non-row line
						size_t tablePos = nextLineEndPos + nextNewlineLength;
						while (tablePos < length)
						{
                            size_t rowStartPos = tablePos;
                            size_t rowCursor = tablePos;
                            while (rowCursor < length && text[rowCursor] != L'\n' && text[rowCursor] != L'\r')
                                ++rowCursor;

                            size_t rowLineEndPos = rowCursor;
                            size_t rowNewlineLength = 0;
                            if (rowCursor < length)
                            {
                                if (text[rowCursor] == L'\r' && rowCursor + 1 < length && text[rowCursor + 1] == L'\n')
                                    rowNewlineLength = 2;
                                else
                                    rowNewlineLength = 1;
                            }

                            std::wstring rowLine = text.substr(rowStartPos, rowLineEndPos - rowStartPos);
                            if (!isTableRowCandidate(rowLine) || isTableSeparatorLine(rowLine))
                                break;

							long absRowStart = static_cast<long>(rowStartPos);
							long absRowEnd = static_cast<long>(rowLineEndPos);
							if (absRowStart < absRowEnd)
							{
								TableRowInfo info{};
								info.start = absRowStart;
								info.end = absRowEnd;
								info.text = rowLine;
								info.isHeader = false;
								tableRows.push_back(std::move(info));
							}
							tablePos = rowLineEndPos + rowNewlineLength;
						}

						const std::vector<int> parsedAlignments = parseTableAlignments(nextLine);
						// 列数优先以分隔线为准（标准 GFM 语义）。
						// 仅当分隔线异常无法给出列数时，才退化为按数据行推断。
						int detectedColumns = static_cast<int>(parsedAlignments.size());
						if (detectedColumns <= 0)
						{
							for (const auto& tr : tableRows)
							{
								const int splitColumns = static_cast<int>(splitTableCells(tr.text).size());
								detectedColumns = (std::max)(detectedColumns, splitColumns);
							}
						}

						int tabStopCount = detectedColumns;
						if (tabStopCount < 1)
							tabStopCount = 1;
						if (tabStopCount > kMaxTabStops)
							tabStopCount = kMaxTabStops;

						std::vector<int> colAlignments((size_t)tabStopCount, 0);
						for (int ci = 0; ci < tabStopCount && ci < static_cast<int>(parsedAlignments.size()); ++ci)
							colAlignments[(size_t)ci] = parsedAlignments[(size_t)ci];

						std::vector<int> colMaxUnits((size_t)tabStopCount, 1);
						for (const auto& tr : tableRows)
						{
							auto cells = splitTableCells(tr.text);
							for (int ci = 0; ci < tabStopCount; ++ci)
							{
								int units = 0;
								if (ci < (int)cells.size())
									units = cellDisplayUnits(cells[(size_t)ci]);
								if (units > colMaxUnits[(size_t)ci])
									colMaxUnits[(size_t)ci] = units;
							}
						}

						constexpr int kTwipsPerUnit = 120;     // ~8px @96dpi
						constexpr int kCellPaddingTwips = 260; // left+right inner padding
						constexpr int kMinColTwips = 420;      // >= 28px
						constexpr int kMaxColTwips = 18000;    // allow very wide cells
						int tabStops[32]{};
						int accum = 0;
						for (int ci = 0; ci < tabStopCount; ++ci)
						{
							int colTwips = colMaxUnits[(size_t)ci] * kTwipsPerUnit + kCellPaddingTwips;
							if (colTwips < kMinColTwips)
								colTwips = kMinColTwips;
							if (colTwips > kMaxColTwips)
								colTwips = kMaxColTwips;
							accum += colTwips;
							tabStops[ci] = accum;
						}

						int dataRowIndex = 0;
						for (const auto& tr : tableRows)
						{
							if (tr.start >= tr.end)
								continue;
							addTablePaddingHiddenSpans(tr.start, tr.text);

							MarkdownBlockSpan rowSpan{};
							rowSpan.start = tr.start;
							rowSpan.end = tr.end;
							if (tr.isHeader)
								rowSpan.type = MarkdownBlockType::TableHeader;
							else
								rowSpan.type = (dataRowIndex % 2 == 0) ? MarkdownBlockType::TableRowEven : MarkdownBlockType::TableRowOdd;
							rowSpan.tabStopCount = tabStopCount;
							for (int i = 0; i < tabStopCount; ++i)
							{
								rowSpan.tabStopsTwips[i] = tabStops[i];
								rowSpan.tabAlignments[i] = colAlignments[(size_t)i];
							}
							result.blockSpans.push_back(std::move(rowSpan));
							if (!tr.isHeader)
								++dataRowIndex;
						}

						pos = tablePos;
						if (pos > length)
                            pos = length;
                        TickProgress(pos, false);
                        continue;
                    }
                }
            }

            // Headings
            size_t hashCount = 0;
            while (trimmedPos + hashCount < line.size() && line[trimmedPos + hashCount] == L'#')
                ++hashCount;

            if (hashCount > 0 && hashCount <= 6)
            {
                size_t afterHashes = trimmedPos + hashCount;
                if (afterHashes == line.size() || IsWhitespace(line[afterHashes]))
                {
                    size_t textStartInLine = afterHashes;
                    while (textStartInLine < line.size() && IsWhitespace(line[textStartInLine]))
                        ++textStartInLine;

                    long markerStart = absoluteLineStart + static_cast<long>(trimmedPos);
                    long markerEnd = absoluteLineStart + static_cast<long>(textStartInLine);
                    long textStart = markerEnd;
                    long textEnd = absoluteLineEnd;

                    // Hide optional closing hashes, e.g. "## Title ##"
                    // Only when there is at least one whitespace before the closing hashes.
                    size_t scanEnd = line.size();
                    while (scanEnd > textStartInLine && IsWhitespace(line[scanEnd - 1]))
                        --scanEnd;

                    size_t hashRunEnd = scanEnd;
                    size_t hashRunStart = hashRunEnd;
                    while (hashRunStart > textStartInLine && line[hashRunStart - 1] == L'#')
                        --hashRunStart;

                    if (hashRunStart < hashRunEnd)
                    {
                        size_t beforeHashes = hashRunStart;
                        while (beforeHashes > textStartInLine && IsWhitespace(line[beforeHashes - 1]))
                            --beforeHashes;

                        if (beforeHashes < hashRunStart)
                        {
                            long trailingStart = absoluteLineStart + static_cast<long>(beforeHashes);
                            if (trailingStart > textStart && trailingStart < textEnd)
                            {
                                AddHiddenSpanBlock(result, trailingStart, textEnd);
                                AddProtectedRange(protectedRanges, trailingStart, textEnd);
                                textEnd = trailingStart;
                            }
                        }
                    }

                    if (markerStart < markerEnd)
                    {
                        AddHiddenSpanBlock(result, markerStart, markerEnd);
                        AddProtectedRange(protectedRanges, markerStart, markerEnd);
                    }

                    if (textStart < textEnd)
                    {
                        MarkdownBlockType blockType = MarkdownBlockType::Heading1;
                        switch (hashCount)
                        {
                        case 1: blockType = MarkdownBlockType::Heading1; break;
                        case 2: blockType = MarkdownBlockType::Heading2; break;
                        case 3: blockType = MarkdownBlockType::Heading3; break;
                        case 4: blockType = MarkdownBlockType::Heading4; break;
                        case 5: blockType = MarkdownBlockType::Heading5; break;
                        case 6: blockType = MarkdownBlockType::Heading6; break;
                        default: break;
                        }
                        result.blockSpans.push_back({ textStart, textEnd, blockType, 0, 0 });
                    }

                    processed = true;
                }
            }
        }

        // Setext headings: 
        //   Title
        //   =====
        // or
        //   Title
        //   -----
        if (!processed && !line.empty())
        {
            size_t nextStart = lineEndPos + newlineLength;
            if (nextStart < length)
            {
                size_t nextCursor = nextStart;
                while (nextCursor < length && text[nextCursor] != L'\n' && text[nextCursor] != L'\r')
                    ++nextCursor;

                size_t nextLineEndPos = nextCursor;
                size_t nextNewlineLength = 0;
                if (nextCursor < length)
                {
                    if (text[nextCursor] == L'\r' && nextCursor + 1 < length && text[nextCursor + 1] == L'\n')
                        nextNewlineLength = 2;
                    else
                        nextNewlineLength = 1;
                }

                std::wstring nextLine = text.substr(nextStart, nextLineEndPos - nextStart);
                size_t nextTrimmedPos = 0;
                while (nextTrimmedPos < nextLine.size() && IsWhitespace(nextLine[nextTrimmedPos]))
                    ++nextTrimmedPos;

                size_t underlineEnd = nextLine.size();
                while (underlineEnd > nextTrimmedPos && IsWhitespace(nextLine[underlineEnd - 1]))
                    --underlineEnd;

                if (underlineEnd > nextTrimmedPos)
                {
                    wchar_t underlineChar = nextLine[nextTrimmedPos];
                    if (underlineChar == L'=' || underlineChar == L'-')
                    {
                        bool allSame = true;
                        for (size_t i = nextTrimmedPos; i < underlineEnd; ++i)
                        {
                            if (nextLine[i] != underlineChar)
                            {
                                allSame = false;
                                break;
                            }
                        }

                        if (allSame)
                        {
                            // Apply heading style to current line's content.
                            long headingStart = absoluteLineStart + static_cast<long>(trimmedPos);
                            long headingEnd = absoluteLineEnd;
                            if (headingStart < headingEnd)
                            {
                                MarkdownBlockType blockType = (underlineChar == L'=')
                                    ? MarkdownBlockType::Heading1
                                    : MarkdownBlockType::Heading2;
                                result.blockSpans.push_back({ headingStart, headingEnd, blockType, 0, 0 });
                            }

                            // Hide underline marker line.
                            long underlineStartAbs = static_cast<long>(nextStart);
                            long underlineEndAbs = static_cast<long>(nextLineEndPos);
                            if (underlineStartAbs < underlineEndAbs)
                            {
                                AddHiddenSpanBlock(result, underlineStartAbs, underlineEndAbs);
                                AddProtectedRange(protectedRanges, underlineStartAbs, underlineEndAbs);
                            }

                            pos = nextLineEndPos + nextNewlineLength;
                            if (pos > length)
                                pos = length;
                            TickProgress(pos, false);
                            continue;
                        }
                    }
                }
            }
        }

        if (!processed && !line.empty())
        {
            // Horizontal rules: --- / *** / ___ / - - - / * * * / _ _ _
            // Do this before other block parsing so it doesn't get mistaken for list markers.
            auto isHorizontalRule = [&](const std::wstring& value) {
                size_t start = 0;
                while (start < value.size() && IsWhitespace(value[start]))
                    ++start;
                size_t end = value.size();
                while (end > start && IsWhitespace(value[end - 1]))
                    --end;
                if (end <= start)
                    return false;

                wchar_t ruleChar = 0;
                int count = 0;
                for (size_t i = start; i < end; ++i)
                {
                    wchar_t ch = value[i];
                    if (IsWhitespace(ch))
                        continue;
                    if (ch == L'-' || ch == L'*' || ch == L'_')
                    {
                        if (ruleChar == 0)
                            ruleChar = ch;
                        if (ch != ruleChar)
                            return false;
                        ++count;
                        continue;
                    }
                    return false;
                }
                return count >= 3;
            };

            if (isHorizontalRule(line))
            {
                long hrStart = absoluteLineStart;
                long hrEnd = absoluteLineEnd;
                if (hrStart < hrEnd)
                {
                    // Mark HR as protected so inline parsing won't apply.
                    AddProtectedRange(protectedRanges, hrStart, hrEnd);
                    result.blockSpans.push_back({ hrStart, hrEnd, MarkdownBlockType::HorizontalRule, 0, 0 });
                }
                pos = lineEndPos + newlineLength;
                if (pos > length)
                    pos = length;
                TickProgress(pos, false);
                continue;
            }

            // Blockquotes
            size_t blockPos = trimmedPos;
            size_t markerCount = 0;
            size_t afterMarkers = blockPos;
            while (afterMarkers < line.size() && line[afterMarkers] == L'>')
            {
                ++markerCount;
                ++afterMarkers;
                if (afterMarkers < line.size() && line[afterMarkers] == L' ')
                    ++afterMarkers;
            }

            if (markerCount > 0)
            {
                // Hide quote markers (and the standard optional single space after each marker).
                long hiddenStart = absoluteLineStart + static_cast<long>(blockPos);
                long hiddenEnd = absoluteLineStart + static_cast<long>(afterMarkers);
                if (hiddenStart < hiddenEnd)
                {
                    AddHiddenSpanBlock(result, hiddenStart, hiddenEnd);
                    AddProtectedRange(protectedRanges, hiddenStart, hiddenEnd);
                }

                // Hide additional indentation whitespace after the marker(s), but keep it as layout indent.
                size_t contentPos = afterMarkers;
                while (contentPos < line.size() && IsWhitespace(line[contentPos]))
                    ++contentPos;

                const size_t extraIndentColumns = indentColumnsBetween(afterMarkers, contentPos);
                if (contentPos > afterMarkers)
                {
                    long indentHiddenStart = absoluteLineStart + static_cast<long>(afterMarkers);
                    long indentHiddenEnd = absoluteLineStart + static_cast<long>(contentPos);
                    if (indentHiddenStart < indentHiddenEnd)
                    {
                        AddHiddenSpanBlock(result, indentHiddenStart, indentHiddenEnd);
                        AddProtectedRange(protectedRanges, indentHiddenStart, indentHiddenEnd);
                    }
                }

                long textStart = absoluteLineStart + static_cast<long>(contentPos);
                long textEnd = absoluteLineEnd;
                if (textStart < textEnd)
                {
                    const int quoteDepth = static_cast<int>(markerCount);
                    const int extraIndentTwips = static_cast<int>(extraIndentColumns / 2) * 240;
                    result.blockSpans.push_back({ textStart, textEnd, MarkdownBlockType::BlockQuote,
                        360 * quoteDepth + extraIndentTwips, 0 });
                }

                trimmedPos = contentPos;
            }

            // Lists (ordered, unordered, task)
            size_t listPos = trimmedPos;
            size_t markerLength = 0;
            size_t markerVisibleLength = 0;
            const size_t listIndentBasePos = (markerCount > 0) ? afterMarkers : 0;
            const size_t listIndentColumns = indentColumnsBetween(listIndentBasePos, listPos);

            if (listPos < line.size())
            {
                wchar_t marker = line[listPos];
                if (marker == L'-' && listPos + 5 < line.size() &&
                    line[listPos + 1] == L' ' && line[listPos + 2] == L'[' &&
                    (line[listPos + 3] == L' ' || line[listPos + 3] == L'x' || line[listPos + 3] == L'X') &&
                    line[listPos + 4] == L']' && line[listPos + 5] == L' ')
                {
                    // Task list item: "- [ ] " or "- [x] "
                    markerLength = 6;
                    markerVisibleLength = 2; // keep rendered checkbox + one visible space
                }
                else if (marker == L'*' && listPos + 5 < line.size() &&
                    line[listPos + 1] == L' ' && line[listPos + 2] == L'[' &&
                    (line[listPos + 3] == L' ' || line[listPos + 3] == L'x' || line[listPos + 3] == L'X') &&
                    line[listPos + 4] == L']' && line[listPos + 5] == L' ')
                {
                    markerLength = 6;
                    markerVisibleLength = 2; // keep rendered checkbox + one visible space
                }
                else if (marker == L'+' && listPos + 5 < line.size() &&
                    line[listPos + 1] == L' ' && line[listPos + 2] == L'[' &&
                    (line[listPos + 3] == L' ' || line[listPos + 3] == L'x' || line[listPos + 3] == L'X') &&
                    line[listPos + 4] == L']' && line[listPos + 5] == L' ')
                {
                    markerLength = 6;
                    markerVisibleLength = 2; // keep rendered checkbox + one visible space
                }
                else if ((marker == L'-' || marker == L'*' || marker == L'+') &&
                    listPos + 1 < line.size() && line[listPos + 1] == L' ')
                {
                    markerLength = 2;
                    markerVisibleLength = 2; // keep marker + one visible space between bullet and text
                }
                else if (std::iswdigit(marker))
                {
                    size_t digitEnd = listPos;
                    while (digitEnd < line.size() && std::iswdigit(line[digitEnd]))
                        ++digitEnd;

                    if (digitEnd > listPos && digitEnd < line.size() &&
                        (line[digitEnd] == L'.' || line[digitEnd] == L')'))
                    {
                        size_t afterDelim = digitEnd + 1;
                        if (afterDelim < line.size() && line[afterDelim] == L' ')
                        {
                            markerLength = afterDelim - listPos + 1;
                        }
                    }
                }
            }

            if (markerLength > 0)
            {
                // Hide indentation whitespace before the marker so nested list indent is visual-only.
                if (listPos > listIndentBasePos)
                {
                    long indentStartAbs = absoluteLineStart + static_cast<long>(listIndentBasePos);
                    long indentEndAbs = absoluteLineStart + static_cast<long>(listPos);
                    if (indentStartAbs < indentEndAbs)
                    {
                        AddHiddenSpanBlock(result, indentStartAbs, indentEndAbs);
                        AddProtectedRange(protectedRanges, indentStartAbs, indentEndAbs);
                    }
                }

                long hiddenStart = absoluteLineStart + static_cast<long>(listPos + markerVisibleLength);
                long hiddenEnd = hiddenStart + static_cast<long>(markerLength);
                if (markerVisibleLength > 0)
                    hiddenEnd = absoluteLineStart + static_cast<long>(listPos + markerLength);
                if (hiddenStart < hiddenEnd)
                {
                    AddHiddenSpanBlock(result, hiddenStart, hiddenEnd);
                    AddProtectedRange(protectedRanges, hiddenStart, hiddenEnd);
                }

                long textStart = hiddenEnd;
                long textEnd = absoluteLineEnd;
                if (textStart < textEnd)
                {
                    const int quoteDepth = static_cast<int>(markerCount);
                    const int baseIndentTwips = 360;
                    const int nestedIndentTwips = (quoteDepth == 0)
                        ? static_cast<int>(listIndentColumns / 2) * 240
                        : static_cast<int>(listIndentColumns / 2) * 240;
                    const int startIndentTwips = quoteDepth * 360 + baseIndentTwips + nestedIndentTwips;

                    result.blockSpans.push_back({ textStart, textEnd, MarkdownBlockType::ListItem,
                        startIndentTwips, 0 });
                }
            }

            // Indented code block (4 spaces or 1 tab). Keep it simple: single-line detection.
            // Typora treats consecutive indented lines as one code block; we approximate by per-line styling.
            // Do this after list detection so list markers win over indent code.
            {
                const size_t rawIndentColumns = indentColumnsBetween(0, trimmedPos);
                if (rawIndentColumns >= 4)
                {
                    // Hide the leading indentation whitespace and render the remaining text as a code block line.
                    if (trimmedPos > 0)
                    {
                        long indentStartAbs = absoluteLineStart;
                        long indentEndAbs = absoluteLineStart + static_cast<long>(trimmedPos);
                        if (indentStartAbs < indentEndAbs)
                        {
                            AddHiddenSpanBlock(result, indentStartAbs, indentEndAbs);
                            AddProtectedRange(protectedRanges, indentStartAbs, indentEndAbs);
                        }
                    }

                    long textStart = absoluteLineStart + static_cast<long>(trimmedPos);
                    long textEnd = absoluteLineEnd;
                    if (textStart < textEnd)
                    {
                        result.blockSpans.push_back({ textStart, textEnd, MarkdownBlockType::CodeBlock, 240, 0 });
                        result.codeRanges.push_back({ textStart, textEnd });
                        AddProtectedRange(protectedRanges, textStart, textEnd);
                    }

                    pos = lineEndPos + newlineLength;
                    if (pos > length)
                        pos = length;
                    TickProgress(pos, false);
                    continue;
                }
            }
        }

        pos = lineEndPos + newlineLength;
        if (pos > length)
            pos = length;
        TickProgress(pos, false);
    }

    if (pending.active)
    {
        const long docEnd = static_cast<long>(length);

        AddHiddenSpanBlock(result, pending.fenceStart, pending.fenceEnd);
        AddProtectedRange(protectedRanges, pending.fenceStart, pending.fenceEnd);

        if (pending.contentStart < docEnd)
        {
            result.blockSpans.push_back({ pending.contentStart, docEnd, MarkdownBlockType::CodeBlock, 240, 0 });
            result.codeRanges.push_back({ pending.contentStart, docEnd });
            AddProtectedRange(protectedRanges, pending.contentStart, docEnd);

            if (pending.isMermaid)
            {
                result.mermaidBlocks.push_back({ pending.contentStart, docEnd });
            }
        }
    }

    TickProgress(length, true);
}

#if 1  // Legacy inline parsing

void MarkdownParser::ParseInline(const std::wstring& text,
                                 MarkdownParseResult& result,
                                 std::vector<Range>& protectedRanges) const
{
    if (text.empty())
        return;

    const int kInlineBase = 60;
    const int kInlineSpan = 39;
    const int kInlinePasses = 8;
    const size_t safeTotalUnits = (std::max)(text.length(), static_cast<size_t>(1));
    auto beginInlinePass = [&](int passIndex) {
        int passStart = kInlineBase + (kInlineSpan * passIndex) / kInlinePasses;
        int passEnd = kInlineBase + (kInlineSpan * (passIndex + 1)) / kInlinePasses;
        int passSpan = (std::max)(1, passEnd - passStart);
        BeginProgressPhase(passStart, passSpan, safeTotalUnits);
    };

    // Protected ranges should be built first, so subsequent parsers don't apply within code.
    beginInlinePass(0);
    ParseInlineCode(text, result, protectedRanges);

    beginInlinePass(1);
    ParseInlineHtmlTags(text, result, protectedRanges);

    beginInlinePass(2);
    ParseBoldItalic(text, result, protectedRanges);
    beginInlinePass(3);
    ParseStrongEmphasis(text, result, protectedRanges);
    beginInlinePass(4);
    ParseStrikethrough(text, result, protectedRanges);
    beginInlinePass(5);
    ParseEmphasis(text, result, protectedRanges);
    beginInlinePass(6);
    ParseLinks(text, result, protectedRanges);
    beginInlinePass(7);
    ParseAutoLinks(text, result, protectedRanges);
}

void MarkdownParser::ParseInlineHtmlTags(const std::wstring& text,
                                         MarkdownParseResult& result,
                                         std::vector<Range>& protectedRanges) const
{
    // Minimal inline HTML tag support (Phase 1):
    // - Hide tags: <strong>/<b>, <em>/<i>, <u>, <del>/<s>
    // - Hide tags: <code>, <mark>
    // - Style inner text like markdown spans
    // - Convert <br> / <br/> into a line break by inserting '\n'
    // Limitations: attributes are ignored; nested tags are supported in a simple stack manner.

    const long length = static_cast<long>(text.length());
    if (length <= 0)
        return;

    struct Frame
    {
        std::wstring name;
        MarkdownInlineType type = MarkdownInlineType::Italic;
        long openStart = 0;
        long openEnd = 0;
        long innerStart = 0;
    };

    auto toLower = [](std::wstring& s) {
        for (auto& ch : s)
            ch = static_cast<wchar_t>(towlower(ch));
    };

    auto mapType = [&](const std::wstring& tag, MarkdownInlineType& outType) -> bool {
        if (tag == L"strong" || tag == L"b") { outType = MarkdownInlineType::Bold; return true; }
        if (tag == L"em" || tag == L"i") { outType = MarkdownInlineType::Italic; return true; }
        if (tag == L"u") { outType = MarkdownInlineType::Underline; return true; }
        if (tag == L"del" || tag == L"s") { outType = MarkdownInlineType::Strikethrough; return true; }
        if (tag == L"code") { outType = MarkdownInlineType::InlineCode; return true; }
        if (tag == L"mark") { outType = MarkdownInlineType::Mark; return true; }
        return false;
    };

    std::vector<Frame> stack;
    stack.reserve(8);

    long pos = 0;
    while (pos < length)
    {
        TickProgress(static_cast<size_t>(pos), false);
        if (RangeContains(protectedRanges, pos))
        {
            ++pos;
            continue;
        }
        if (text[static_cast<size_t>(pos)] != L'<')
        {
            ++pos;
            continue;
        }

        const long tagStart = pos;
        long scan = pos + 1;
        if (scan >= length)
            break;

        bool closing = false;
        if (text[static_cast<size_t>(scan)] == L'/')
        {
            closing = true;
            ++scan;
        }

        const long nameStart = scan;
        while (scan < length)
        {
            wchar_t ch = text[static_cast<size_t>(scan)];
            if ((ch >= L'a' && ch <= L'z') || (ch >= L'A' && ch <= L'Z'))
            {
                ++scan;
                continue;
            }
            break;
        }
        const long nameEnd = scan;
        if (nameEnd <= nameStart)
        {
            ++pos;
            continue;
        }

        std::wstring name = text.substr(static_cast<size_t>(nameStart), static_cast<size_t>(nameEnd - nameStart));
        toLower(name);

        long gt = scan;
        while (gt < length && text[static_cast<size_t>(gt)] != L'>')
            ++gt;
        if (gt >= length)
            break;

        bool selfClosing = false;
        long back = gt - 1;
        while (back > tagStart && (text[static_cast<size_t>(back)] == L' ' || text[static_cast<size_t>(back)] == L'\t' || text[static_cast<size_t>(back)] == L'\r'))
            --back;
        if (back > tagStart && text[static_cast<size_t>(back)] == L'/')
            selfClosing = true;

        const long tagEnd = gt + 1;

        if (name == L"br")
        {
            // Hide the tag text. (We do not modify the document text to keep indices stable.)
            if (!RangeIntersects(protectedRanges, tagStart, tagEnd))
            {
                AddHiddenSpanInline(result, tagStart, tagEnd);
                AddProtectedRange(protectedRanges, tagStart, tagEnd);
            }
            pos = tagEnd;
            continue;
        }

        MarkdownInlineType mappedType = MarkdownInlineType::Italic;
        if (!mapType(name, mappedType))
        {
            pos = tagEnd;
            continue;
        }

        if (!closing && !selfClosing)
        {
            Frame f;
            f.name = name;
            f.type = mappedType;
            f.openStart = tagStart;
            f.openEnd = tagEnd;
            f.innerStart = tagEnd;
            stack.push_back(std::move(f));
            pos = tagEnd;
            continue;
        }

        if (closing)
        {
            for (int i = static_cast<int>(stack.size()) - 1; i >= 0; --i)
            {
                if (stack[(size_t)i].name != name)
                    continue;
                Frame f = stack[(size_t)i];
                stack.resize((size_t)i);

                if (!RangeIntersects(protectedRanges, f.openStart, f.openEnd) && !RangeIntersects(protectedRanges, tagStart, tagEnd))
                {
                    AddHiddenSpanInline(result, f.openStart, f.openEnd);
                    AddHiddenSpanInline(result, tagStart, tagEnd);
                    AddProtectedRange(protectedRanges, f.openStart, f.openEnd);
                    AddProtectedRange(protectedRanges, tagStart, tagEnd);

                    long innerStart = f.innerStart;
                    long innerEnd = tagStart;
                    if (innerStart < innerEnd && !RangeIntersects(protectedRanges, innerStart, innerEnd))
                    {
                        result.inlineSpans.push_back({ innerStart, innerEnd, f.type });
                        // If it's inline code, protect inner text so other inline parsers don't apply.
                        if (f.type == MarkdownInlineType::InlineCode)
                        {
                            result.codeRanges.push_back({ innerStart, innerEnd });
                            AddProtectedRange(protectedRanges, innerStart, innerEnd);
                        }
                    }
                }
                break;
            }
            pos = tagEnd;
            continue;
        }

        pos = tagEnd;
    }
    TickProgress(static_cast<size_t>(length), true);
}

#endif  // Legacy inline parsing

#if 1  // Legacy inline parsing

void MarkdownParser::ParseBoldItalic(const std::wstring& text,
                                     MarkdownParseResult& result,
                                     std::vector<Range>& protectedRanges) const
{
    // Handle "***text***" and "___text___" as bold + italic combined.
    // Also handle "****text****" / "____text____" (Typora-compatible shorthand seen in test cases)
    // as BoldItalic as well.
    // Do this before strong/emphasis parsing to avoid conflicting ranges.
    const long length = static_cast<long>(text.length());
    long pos = 0;

    while (pos + 5 < length)
    {
        TickProgress(static_cast<size_t>(pos), false);
        if (RangeContains(protectedRanges, pos))
        {
            ++pos;
            continue;
        }

        const bool isTriple =
            (text[pos] == L'*' && text[pos + 1] == L'*' && text[pos + 2] == L'*') ||
            (text[pos] == L'_' && text[pos + 1] == L'_' && text[pos + 2] == L'_');
        const bool isQuad =
            (pos + 3 < length &&
             ((text[pos] == L'*' && text[pos + 1] == L'*' && text[pos + 2] == L'*' && text[pos + 3] == L'*') ||
              (text[pos] == L'_' && text[pos + 1] == L'_' && text[pos + 2] == L'_' && text[pos + 3] == L'_')));

        if (isQuad || isTriple)
        {
            const wchar_t marker = text[pos];
            const int markerCount = isQuad ? 4 : 3;
            const wchar_t needle3[] = { marker, marker, marker, 0 };
            const wchar_t needle4[] = { marker, marker, marker, marker, 0 };
            const wchar_t* needle = isQuad ? needle4 : needle3;
            const long lineEnd = FindLineEnd(text, pos);

            size_t found = text.find(needle, static_cast<size_t>(pos + markerCount));
            if (found != std::wstring::npos)
            {
                long end = static_cast<long>(found);
                if (end > pos + markerCount && end < lineEnd &&
                    !RangeIntersects(protectedRanges, pos, pos + markerCount) &&
                    !RangeIntersects(protectedRanges, end, end + markerCount))
                {
                    AddHiddenSpanInline(result, pos, pos + markerCount);
                    AddHiddenSpanInline(result, end, end + markerCount);
                    AddProtectedRange(protectedRanges, pos, pos + markerCount);
                    AddProtectedRange(protectedRanges, end, end + markerCount);

                    long textStart = pos + markerCount;
                    long textEnd = end;
                    if (textStart < textEnd)
                    {
                        result.inlineSpans.push_back({ textStart, textEnd, MarkdownInlineType::BoldItalic });
                    }

                    pos = end + markerCount;
                    continue;
                }
            }
        }

        ++pos;
    }
    TickProgress(static_cast<size_t>(length), true);
}

#endif  // Legacy inline parsing

#if 1  // Legacy inline parsing

void MarkdownParser::ParseInlineCode(const std::wstring& text,
                                     MarkdownParseResult& result,
                                     std::vector<Range>& protectedRanges) const
{
    const long length = static_cast<long>(text.length());
    long pos = 0;

    while (pos < length)
    {
        TickProgress(static_cast<size_t>(pos), false);
        if (RangeContains(protectedRanges, pos))
        {
            ++pos;
            continue;
        }

        if (text[pos] == L'`')
        {
            const long lineEnd = FindLineEnd(text, pos);
            long end = pos + 1;
            while (end < lineEnd && text[end] != L'`')
                ++end;

            if (end < lineEnd && end > pos + 1 && !RangeIntersects(protectedRanges, pos, end + 1))
            {
                AddHiddenSpanInline(result, pos, pos + 1);
                AddHiddenSpanInline(result, end, end + 1);
                AddProtectedRange(protectedRanges, pos, pos + 1);
                AddProtectedRange(protectedRanges, end, end + 1);

                long codeStart = pos + 1;
                long codeEnd = end;
                if (codeStart < codeEnd)
                {
                    result.inlineSpans.push_back({ codeStart, codeEnd, MarkdownInlineType::InlineCode });
                    result.codeRanges.push_back({ codeStart, codeEnd });
                    AddProtectedRange(protectedRanges, codeStart, codeEnd);
                }

                pos = end + 1;
                continue;
            }
        }
        ++pos;
    }
    TickProgress(static_cast<size_t>(length), true);
}

#endif  // Legacy inline parsing

#if 1  // Legacy inline parsing

void MarkdownParser::ParseStrongEmphasis(const std::wstring& text,
                                         MarkdownParseResult& result,
                                         std::vector<Range>& protectedRanges) const
{
    const long length = static_cast<long>(text.length());
    long pos = 0;

    while (pos + 1 < length)
    {
        TickProgress(static_cast<size_t>(pos), false);
        if (RangeContains(protectedRanges, pos))
        {
            ++pos;
            continue;
        }

        if ((text[pos] == L'*' && text[pos + 1] == L'*') ||
            (text[pos] == L'_' && text[pos + 1] == L'_'))
        {
            const wchar_t marker = text[pos];
            const wchar_t needle[] = { marker, marker, 0 };
            const long lineEnd = FindLineEnd(text, pos);
            size_t found = text.find(needle, static_cast<size_t>(pos + 2));
            if (found != std::wstring::npos)
            {
                long end = static_cast<long>(found);
                // Guard: avoid treating "****text****" as an empty bold span "** **".
                // In that case, the closing marker is adjacent to the opening marker.
                if (end == pos + 2)
                {
                    ++pos;
                    continue;
                }
                if (end > pos + 2 && end < lineEnd &&
                    !RangeIntersects(protectedRanges, pos, pos + 2) &&
                    !RangeIntersects(protectedRanges, end, end + 2))
                {
                    AddHiddenSpanInline(result, pos, pos + 2);
                    AddHiddenSpanInline(result, end, end + 2);
                    AddProtectedRange(protectedRanges, pos, pos + 2);
                    AddProtectedRange(protectedRanges, end, end + 2);

                    long textStart = pos + 2;
                    long textEnd = end;
                    if (textStart < textEnd)
                    {
                        result.inlineSpans.push_back({ textStart, textEnd, MarkdownInlineType::Bold });
                    }

                    pos = end + 2;
                    continue;
                }
            }
        }
        ++pos;
    }
    TickProgress(static_cast<size_t>(length), true);
}

#endif  // Legacy inline parsing

#if 1  // Legacy inline parsing

void MarkdownParser::ParseEmphasis(const std::wstring& text,
                                   MarkdownParseResult& result,
                                   std::vector<Range>& protectedRanges) const
{
    const long length = static_cast<long>(text.length());
    auto isAsciiEnglish = [](wchar_t ch) -> bool {
        return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z');
    };
    auto isUnderscoreWordConnector = [&](long index) -> bool {
        if (index <= 0 || index + 1 >= length)
            return false;
        return isAsciiEnglish(text[static_cast<size_t>(index - 1)]) &&
            isAsciiEnglish(text[static_cast<size_t>(index + 1)]);
    };
    long pos = 0;

    while (pos < length)
    {
        TickProgress(static_cast<size_t>(pos), false);
        if (RangeContains(protectedRanges, pos))
        {
            ++pos;
            continue;
        }

        if (text[pos] == L'*' || text[pos] == L'_')
        {
            const wchar_t marker = text[pos];
            // Treat underscore between English letters as a word connector,
            // not an emphasis delimiter (e.g. foo_bar / user_name).
            if (marker == L'_' && isUnderscoreWordConnector(pos))
            {
                ++pos;
                continue;
            }
            if ((pos > 0 && text[pos - 1] == marker) ||
                (pos + 1 < length && text[pos + 1] == marker))
            {
                ++pos;
                continue;
            }

            const long lineEnd = FindLineEnd(text, pos);

            long end = pos + 1;
            while (end < lineEnd)
            {
                if (text[end] != marker)
                {
                    ++end;
                    continue;
                }
                // Skip connector underscores when searching for closing delimiter.
                if (marker == L'_' && isUnderscoreWordConnector(end))
                {
                    ++end;
                    continue;
                }
                break;
            }

            if (end < lineEnd && end > pos + 1 &&
                !RangeIntersects(protectedRanges, pos, pos + 1) &&
                !RangeIntersects(protectedRanges, end, end + 1))
            {
                AddHiddenSpanInline(result, pos, pos + 1);
                AddHiddenSpanInline(result, end, end + 1);
                AddProtectedRange(protectedRanges, pos, pos + 1);
                AddProtectedRange(protectedRanges, end, end + 1);

                long textStart = pos + 1;
                long textEnd = end;
                if (textStart < textEnd)
                {
                    result.inlineSpans.push_back({ textStart, textEnd, MarkdownInlineType::Italic });
                }

                pos = end + 1;
                continue;
            }
        }
        ++pos;
    }
    TickProgress(static_cast<size_t>(length), true);
}

#endif  // Legacy inline parsing

#if 1  // Legacy inline parsing

void MarkdownParser::ParseStrikethrough(const std::wstring& text,
                                        MarkdownParseResult& result,
                                        std::vector<Range>& protectedRanges) const
{
    const long length = static_cast<long>(text.length());
    long pos = 0;

    while (pos + 1 < length)
    {
        TickProgress(static_cast<size_t>(pos), false);
        if (RangeContains(protectedRanges, pos))
        {
            ++pos;
            continue;
        }

        if (text[pos] == L'~' && text[pos + 1] == L'~')
        {
            const long lineEnd = FindLineEnd(text, pos);
            size_t found = text.find(L"~~", static_cast<size_t>(pos + 2));
            if (found != std::wstring::npos)
            {
                long end = static_cast<long>(found);
                if (end > pos + 2 && end < lineEnd &&
                    !RangeIntersects(protectedRanges, pos, pos + 2) &&
                    !RangeIntersects(protectedRanges, end, end + 2))
                {
                    AddHiddenSpanInline(result, pos, pos + 2);
                    AddHiddenSpanInline(result, end, end + 2);
                    AddProtectedRange(protectedRanges, pos, pos + 2);
                    AddProtectedRange(protectedRanges, end, end + 2);

                    long textStart = pos + 2;
                    long textEnd = end;
                    if (textStart < textEnd)
                    {
                        result.inlineSpans.push_back({ textStart, textEnd, MarkdownInlineType::Strikethrough });
                    }

                    pos = end + 2;
                    continue;
                }
            }
        }
        ++pos;
    }
    TickProgress(static_cast<size_t>(length), true);
}

#endif  // Legacy inline parsing

#if 1  // Legacy inline parsing

void MarkdownParser::ParseLinks(const std::wstring& text,
                                MarkdownParseResult& result,
                                std::vector<Range>& protectedRanges) const
{
    const long length = static_cast<long>(text.length());
    long pos = 0;
    auto findMatchingParen = [&](size_t openParenPos, long lineEnd) -> size_t {
        int depth = 0;
        const size_t textLength = text.length();
        for (size_t i = openParenPos; i < textLength; ++i)
        {
            if (static_cast<long>(i) >= lineEnd)
                break;

            wchar_t ch = text[i];
            if (ch == L'\\')
            {
                if (i + 1 < textLength)
                    ++i;
                continue;
            }

            if (ch == L'(')
            {
                ++depth;
                continue;
            }
            if (ch == L')')
            {
                --depth;
                if (depth == 0)
                    return i;
                if (depth < 0)
                    break;
            }
        }
        return std::wstring::npos;
    };

    while (pos < length)
    {
        TickProgress(static_cast<size_t>(pos), false);
        if (text[pos] == L'[' && !RangeContains(protectedRanges, pos))
        {
            const long lineEnd = FindLineEnd(text, pos);
            size_t closeBracketPos = text.find(L']', static_cast<size_t>(pos + 1));
            if (closeBracketPos != std::wstring::npos)
            {
                long closeBracket = static_cast<long>(closeBracketPos);
                if (closeBracket < lineEnd)
                {
                    size_t openParenPos = closeBracketPos + 1;
                    while (openParenPos < text.length() && text[openParenPos] == L' ')
                        ++openParenPos;

                    if (openParenPos < text.length() && static_cast<long>(openParenPos) < lineEnd && text[openParenPos] == L'(')
                    {
                        size_t endParenPos = findMatchingParen(openParenPos, lineEnd);
                        if (endParenPos != std::wstring::npos)
                        {
                            long endParen = static_cast<long>(endParenPos);
                            if (endParen < lineEnd &&
                                !RangeIntersects(protectedRanges, pos, pos + 1) &&
                                !RangeIntersects(protectedRanges, closeBracket, closeBracket + 1) &&
                                !RangeIntersects(protectedRanges, static_cast<long>(openParenPos), static_cast<long>(openParenPos + 1)) &&
                                !RangeIntersects(protectedRanges, endParen, endParen + 1))
                            {
                                long linkStart = pos + 1;
                                long linkEnd = closeBracket;
                                if (linkStart < linkEnd)
                                {
                                    std::wstring rawHref = text.substr(openParenPos + 1, endParenPos - (openParenPos + 1));
                                    // Support optional title: (url "title") or (url 'title')
                                    // Keep destination only.
                                    size_t hrefEnd = rawHref.find_first_of(L" \t");
                                    std::wstring href = (hrefEnd == std::wstring::npos)
                                        ? rawHref
                                        : rawHref.substr(0, hrefEnd);
                                    result.inlineSpans.push_back({ linkStart, linkEnd, MarkdownInlineType::Link, href });
                                }

                                AddHiddenSpanInline(result, pos, pos + 1);
                                AddHiddenSpanInline(result, closeBracket, endParen + 1);
                                AddProtectedRange(protectedRanges, pos, pos + 1);
                                AddProtectedRange(protectedRanges, closeBracket, endParen + 1);

                                pos = endParen + 1;
                                continue;
                            }
                        }
                    }
                }
            }
        }
        ++pos;
    }
    TickProgress(static_cast<size_t>(length), true);
}

#endif  // Legacy inline parsing

#if 1  // Legacy inline parsing

void MarkdownParser::ParseAutoLinks(const std::wstring& text,
                                    MarkdownParseResult& result,
                                    std::vector<Range>& protectedRanges) const
{
    const long length = static_cast<long>(text.length());
    long pos = 0;

    const std::wstring prefixes[] = { L"http://", L"https://", L"ftp://", L"mailto:", L"www." };

    auto isBoundaryBefore = [&](long index) {
        if (index <= 0)
            return true;
        wchar_t ch = text[index - 1];
        return IsWhitespace(ch) || ch == L'(' || ch == L'<' || ch == L'[';
    };

    while (pos < length)
    {
        TickProgress(static_cast<size_t>(pos), false);
        if (RangeContains(protectedRanges, pos))
        {
            ++pos;
            continue;
        }

        const bool hasAngleWrapper = (pos > 0 && text[pos - 1] == L'<' && !RangeContains(protectedRanges, pos - 1));

        size_t matchedPrefixLen = 0;
        for (const auto& prefix : prefixes)
        {
            size_t prefixLen = prefix.length();
            if (pos + static_cast<long>(prefixLen) <= length &&
                _wcsnicmp(text.c_str() + pos, prefix.c_str(), prefixLen) == 0 &&
                isBoundaryBefore(pos))
            {
                matchedPrefixLen = prefixLen;
                break;
            }
        }

        if (matchedPrefixLen == 0)
        {
            ++pos;
            continue;
        }

        if (hasAngleWrapper)
        {
            long scan = pos + static_cast<long>(matchedPrefixLen);
            while (scan < length && text[scan] != L'>' && !IsWhitespace(text[scan]))
                ++scan;

            if (scan < length && text[scan] == L'>')
            {
                long linkStart = pos;
                long linkEnd = scan;
                if (linkEnd > linkStart && !RangeIntersects(protectedRanges, linkStart, linkEnd))
                {
                    std::wstring href = text.substr(static_cast<size_t>(linkStart), static_cast<size_t>(linkEnd - linkStart));
                    result.inlineSpans.push_back({ linkStart, linkEnd, MarkdownInlineType::AutoLink, href });
                    AddProtectedRange(protectedRanges, linkStart, linkEnd);

                    AddHiddenSpanInline(result, pos - 1, pos);
                    AddHiddenSpanInline(result, scan, scan + 1);
                    AddProtectedRange(protectedRanges, pos - 1, pos);
                    AddProtectedRange(protectedRanges, scan, scan + 1);

                    pos = scan + 1;
                    continue;
                }
            }
        }

        long end = pos + static_cast<long>(matchedPrefixLen);
        while (end < length && !IsWhitespace(text[end]) && !IsLineBreak(text[end]))
            ++end;

        long trimmedEnd = end;
        while (trimmedEnd > pos && wcschr(kTrailingPunct, text[trimmedEnd - 1]) != nullptr)
            --trimmedEnd;

        if (trimmedEnd > pos && !RangeIntersects(protectedRanges, pos, trimmedEnd))
        {
            std::wstring href = text.substr(static_cast<size_t>(pos), static_cast<size_t>(trimmedEnd - pos));
            result.inlineSpans.push_back({ pos, trimmedEnd, MarkdownInlineType::AutoLink, href });
            AddProtectedRange(protectedRanges, pos, trimmedEnd);
        }

        pos = end;
    }

    // Angle-wrapped email autolinks: <user@example.com>
    pos = 0;
    while (pos < length)
    {
        TickProgress(static_cast<size_t>(pos), false);
        if (RangeContains(protectedRanges, pos))
        {
            ++pos;
            continue;
        }

        if (text[pos] != L'<')
        {
            ++pos;
            continue;
        }

        const long lineEnd = FindLineEnd(text, pos);
        long scan = pos + 1;
        bool hasAt = false;
        while (scan < lineEnd && text[scan] != L'>')
        {
            if (IsWhitespace(text[scan]))
                break;
            if (text[scan] == L'@')
                hasAt = true;
            ++scan;
        }

        if (scan < lineEnd && text[scan] == L'>' && hasAt)
        {
            const long emailStart = pos + 1;
            const long emailEnd = scan;
            if (emailEnd > emailStart &&
                !RangeIntersects(protectedRanges, pos, pos + 1) &&
                !RangeIntersects(protectedRanges, scan, scan + 1))
            {
                std::wstring email = text.substr(static_cast<size_t>(emailStart), static_cast<size_t>(emailEnd - emailStart));
                std::wstring href = L"mailto:" + email;
                result.inlineSpans.push_back({ emailStart, emailEnd, MarkdownInlineType::AutoLink, href });
                AddProtectedRange(protectedRanges, emailStart, emailEnd);

                AddHiddenSpanInline(result, pos, pos + 1);
                AddHiddenSpanInline(result, scan, scan + 1);
                AddProtectedRange(protectedRanges, pos, pos + 1);
                AddProtectedRange(protectedRanges, scan, scan + 1);

                pos = scan + 1;
                continue;
            }
        }

        ++pos;
    }
    TickProgress(static_cast<size_t>(length), true);
}

#endif  // Legacy inline parsing
