#include "pch.h"
#include "RenderText.h"

namespace
{
    inline std::wstring TrimWForDisplay(const std::wstring& s)
    {
        size_t a = 0;
        while (a < s.size() && (s[a] == L' ' || s[a] == L'\t' || s[a] == L'\r' || s[a] == L'\n'))
            ++a;
        size_t b = s.size();
        while (b > a && (s[b - 1] == L' ' || s[b - 1] == L'\t' || s[b - 1] == L'\r' || s[b - 1] == L'\n'))
            --b;
        return (b > a) ? s.substr(a, b - a) : std::wstring();
    }
}

	inline bool IsTableWhitespace(wchar_t ch)
	{
		return ch == L' ' || ch == L'\t' || ch == L'\r';
	}

	inline bool IsTableSeparatorLine(const std::wstring& line)
	{
		size_t s = 0;
		while (s < line.size() && IsTableWhitespace(line[s]))
			++s;
		size_t e = line.size();
		while (e > s && IsTableWhitespace(line[e - 1]))
			--e;
		if (e <= s)
			return false;

		bool hasPipe = false;
		int dashCount = 0;
		for (size_t i = s; i < e; ++i)
		{
			wchar_t ch = line[i];
			if (ch == L'|')
				hasPipe = true;
			else if (ch == L'-')
				++dashCount;
			else if (ch == L':' || IsTableWhitespace(ch))
				;
			else
				return false;
		}

		return hasPipe && dashCount >= 3;
	}

	inline bool IsTableRowCandidate(const std::wstring& line)
	{
		int pipeCount = 0;
		for (wchar_t ch : line)
		{
			if (ch == L'|')
				++pipeCount;
		}
		// 支持无外侧竖线的两列表写法：a | b
		return pipeCount >= 1;
	}

	inline void TransformTableLineToTabsInPlace(std::wstring& text, size_t lineStart, size_t lineEnd)
	{
		// lineEnd excludes '\n'. May include trailing '\r' already stripped by caller.
		size_t s = lineStart;
		while (s < lineEnd && (text[s] == L' ' || text[s] == L'\t'))
			++s;
		size_t e = lineEnd;
		while (e > s && (text[e - 1] == L' ' || text[e - 1] == L'\t'))
			--e;
		if (e <= s)
			return;

		auto isEscaped = [&](size_t index) -> bool {
			if (index <= lineStart)
				return false;
			size_t backslashes = 0;
			size_t p = index;
			while (p > lineStart)
			{
				--p;
				if (text[p] == L'\\')
					++backslashes;
				else
					break;
			}
			return (backslashes % 2) == 1;
		};

		const size_t firstNon = s;
		const size_t lastNon = e - 1;
		bool inCodeSpan = false;
		for (size_t i = s; i < e; ++i)
		{
			const wchar_t ch = text[i];
			if (ch == L'`' && !isEscaped(i))
			{
				inCodeSpan = !inCodeSpan;
				continue;
			}
			if (ch != L'|')
				continue;
			if (i == firstNon)
				continue;
			if (inCodeSpan)
				continue;
			if (isEscaped(i))
				continue;

			if (i == lastNon)
			{
				if (text[i] == L'|')
					text[i] = L'\t';
				continue;
			}

			text[i] = L'\t';
		}
	}

	inline void ClipTableCellsInLineInPlace(std::wstring& text, size_t lineStart, size_t lineEnd)
	{
		// 列宽由解析阶段动态计算；这里不截断单元格内容。
		UNREFERENCED_PARAMETER(text);
		UNREFERENCED_PARAMETER(lineStart);
		UNREFERENCED_PARAMETER(lineEnd);
	}

std::wstring RenderText::Transform(const std::wstring& rawText)
	{
		// IMPORTANT: Must preserve text length 1:1 so parser indices remain valid.
		std::wstring out = rawText;
		if (out.empty())
			return out;

		auto isNewlineChar = [](wchar_t ch) { return ch == L'\r' || ch == L'\n'; };
		auto blankRangePreserveNewlines = [&](size_t start, size_t end) {
			if (end > out.size())
				end = out.size();
			for (size_t i = start; i < end; ++i)
			{
				if (!isNewlineChar(out[i]))
					out[i] = L' ';
			}
		};
		auto fillMermaidContentPlaceholder = [&](size_t start, size_t end) {
			// Preserve line structure (newlines) so outline line numbers remain stable,
			// but make the (hidden) mermaid source lines wide enough to allow horizontal
			// scrolling when the rendered diagram overflows.
			if (end > out.size())
				end = out.size();
			for (size_t i = start; i < end; ++i)
			{
				if (isNewlineChar(out[i]))
					continue;
				out[i] = L'\t';
			}
		};

		auto trimLeft = [](const std::wstring& s) -> size_t {
			size_t p = 0;
			while (p < s.size() && (s[p] == L' ' || s[p] == L'\t'))
				++p;
			return p;
		};
		auto startsWithNoCase = [](const std::wstring& s, const std::wstring& prefix) {
			if (s.size() < prefix.size())
				return false;
			for (size_t i = 0; i < prefix.size(); ++i)
			{
				wchar_t a = static_cast<wchar_t>(towlower(s[i]));
				wchar_t b = static_cast<wchar_t>(towlower(prefix[i]));
				if (a != b)
					return false;
			}
			return true;
		};

		auto detectSupportedMermaid = [&](size_t contentStart, size_t contentEnd) -> bool {
			if (contentEnd <= contentStart || contentStart >= out.size())
				return false;
			contentEnd = min(contentEnd, out.size());
			std::wstring block = out.substr(contentStart, contentEnd - contentStart);
			bool hasGraph = false;
			bool hasFlowchart = false;
			bool hasDirVertical = false;
			bool hasDirHorizontal = false;
			bool hasEdge = false;
			bool hasNode = false;
			bool hasSequence = false;
			bool hasSeqSignal = false;
			size_t p = 0;
			while (p < block.size())
			{
				size_t e = block.find(L'\n', p);
				if (e == std::wstring::npos)
					e = block.size();
				size_t ce = e;
				if (ce > p && block[ce - 1] == L'\r')
					--ce;
				std::wstring line = block.substr(p, ce - p);
				line = TrimWForDisplay(line);
				if (!line.empty())
				{
					// sequenceDiagram (minimal support)
					if (startsWithNoCase(line, L"sequenceDiagram"))
						hasSequence = true;
					// Sequence signals: message/participant/note/control keywords.
					if (line.find(L"->") != std::wstring::npos)
						hasSeqSignal = true;
					if (startsWithNoCase(line, L"participant ") || startsWithNoCase(line, L"actor "))
						hasSeqSignal = true;
					if (startsWithNoCase(line, L"note "))
						hasSeqSignal = true;
					if (startsWithNoCase(line, L"alt ") || startsWithNoCase(line, L"else") ||
						startsWithNoCase(line, L"opt ") || startsWithNoCase(line, L"loop ") ||
						startsWithNoCase(line, L"par ") || startsWithNoCase(line, L"and ") ||
						startsWithNoCase(line, L"end"))
						hasSeqSignal = true;
					if (startsWithNoCase(line, L"activate ") || startsWithNoCase(line, L"deactivate ") ||
						startsWithNoCase(line, L"autonumber"))
						hasSeqSignal = true;

					if (startsWithNoCase(line, L"graph "))
					{
						hasGraph = true;
						std::wstring lower = line;
						for (auto& ch : lower) ch = static_cast<wchar_t>(towlower(ch));
						if (lower.find(L" td") != std::wstring::npos || lower.find(L" td;") != std::wstring::npos) hasDirVertical = true;
						if (lower.find(L" tb") != std::wstring::npos || lower.find(L" tb;") != std::wstring::npos) hasDirVertical = true;
						if (lower.find(L" bt") != std::wstring::npos || lower.find(L" bt;") != std::wstring::npos) hasDirVertical = true;
						if (lower.find(L" lr") != std::wstring::npos || lower.find(L" lr;") != std::wstring::npos) hasDirHorizontal = true;
						if (lower.find(L" rl") != std::wstring::npos || lower.find(L" rl;") != std::wstring::npos) hasDirHorizontal = true;
					}
					if (startsWithNoCase(line, L"flowchart "))
					{
						hasFlowchart = true;
						std::wstring lower = line;
						for (auto& ch : lower) ch = static_cast<wchar_t>(towlower(ch));
						if (lower.find(L" td") != std::wstring::npos || lower.find(L" td;") != std::wstring::npos) hasDirVertical = true;
						if (lower.find(L" tb") != std::wstring::npos || lower.find(L" tb;") != std::wstring::npos) hasDirVertical = true;
						if (lower.find(L" bt") != std::wstring::npos || lower.find(L" bt;") != std::wstring::npos) hasDirVertical = true;
						if (lower.find(L" lr") != std::wstring::npos || lower.find(L" lr;") != std::wstring::npos) hasDirHorizontal = true;
						if (lower.find(L" rl") != std::wstring::npos || lower.find(L" rl;") != std::wstring::npos) hasDirHorizontal = true;
					}
					if (line.find(L"-->") != std::wstring::npos)
						hasEdge = true;
					if ((line.find(L"[") != std::wstring::npos && line.find(L"]") != std::wstring::npos) ||
						(line.find(L"{") != std::wstring::npos && line.find(L"}") != std::wstring::npos) ||
						(line.find(L"(") != std::wstring::npos && line.find(L")") != std::wstring::npos))
						hasNode = true;
				}
				p = (e < block.size()) ? (e + 1) : block.size();
			}
			if (hasSequence)
				return hasSeqSignal;
			return (hasGraph || hasFlowchart) && (hasDirVertical || hasDirHorizontal) && hasEdge && hasNode;
		};

		size_t pos = 0;
		while (pos <= out.size())
		{
			size_t lineEnd = out.find(L'\n', pos);
			if (lineEnd == std::wstring::npos)
				lineEnd = out.size();
			size_t contentEnd = lineEnd;
			if (contentEnd > pos && out[contentEnd - 1] == L'\r')
				--contentEnd;
			std::wstring line = out.substr(pos, contentEnd - pos);

			// lookahead next line
			size_t nextPos = (lineEnd < out.size()) ? (lineEnd + 1) : out.size();
			size_t nextLineEnd = (nextPos < out.size()) ? out.find(L'\n', nextPos) : std::wstring::npos;
			if (nextLineEnd == std::wstring::npos)
				nextLineEnd = out.size();
			size_t nextContentEnd = nextLineEnd;
			if (nextContentEnd > nextPos && out[nextContentEnd - 1] == L'\r')
				--nextContentEnd;
			std::wstring nextLine;
			if (nextPos < out.size())
				nextLine = out.substr(nextPos, nextContentEnd - nextPos);

			// Mermaid fenced code blocks: hide fences always; hide content only when supported.
			{
				size_t trimmed = trimLeft(line);
				if (trimmed < line.size() && (line[trimmed] == L'`' || line[trimmed] == L'~'))
				{
					wchar_t fenceChar = line[trimmed];
					size_t fenceCount = 0;
					while (trimmed + fenceCount < line.size() && line[trimmed + fenceCount] == fenceChar)
						++fenceCount;
					if (fenceCount >= 3)
					{
						std::wstring afterFence = line.substr(trimmed + fenceCount);
						afterFence = TrimWForDisplay(afterFence);
						if (startsWithNoCase(afterFence, L"mermaid"))
						{
							// Find close fence.
							size_t openLineStart = pos;
							size_t openLineEnd = contentEnd;
							size_t contentStart = (lineEnd < out.size()) ? (lineEnd + 1) : out.size();
							size_t scanPos = contentStart;
							size_t closeLineStart = std::wstring::npos;
							size_t closeLineEnd = std::wstring::npos;
							while (scanPos <= out.size())
							{
								size_t le = out.find(L'\n', scanPos);
								if (le == std::wstring::npos)
									le = out.size();
								size_t ce = le;
								if (ce > scanPos && out[ce - 1] == L'\r')
									--ce;
								std::wstring l = out.substr(scanPos, ce - scanPos);
								size_t tl = trimLeft(l);
								if (tl < l.size())
								{
									wchar_t ch = l[tl];
									if (ch == fenceChar)
									{
										size_t c = 0;
										while (tl + c < l.size() && l[tl + c] == fenceChar)
											++c;
										if (c >= fenceCount)
										{
											closeLineStart = scanPos;
											closeLineEnd = ce;
											break;
										}
									}
								}
								if (le >= out.size())
									break;
								scanPos = le + 1;
							}

							if (closeLineStart != std::wstring::npos)
							{
								blankRangePreserveNewlines(openLineStart, openLineEnd);
								blankRangePreserveNewlines(closeLineStart, closeLineEnd);
								// Content: hide only if it looks supported by our minimal renderer.
								if (detectSupportedMermaid(contentStart, closeLineStart))
									fillMermaidContentPlaceholder(contentStart, closeLineStart);

								// Continue scanning after the close fence line.
								size_t closeLf = out.find(L'\n', closeLineEnd);
								pos = (closeLf == std::wstring::npos) ? out.size() : (closeLf + 1);
								continue;
							}
						}
					}
				}
			}

			if (IsTableRowCandidate(line) && IsTableSeparatorLine(nextLine))
			{
				// header line
				TransformTableLineToTabsInPlace(out, pos, contentEnd);
				ClipTableCellsInLineInPlace(out, pos, contentEnd);
				// Keep table outer borders stable: set leading and trailing pipes to spaces
				// (we draw borders via overlay). Must preserve length.
				{
					size_t s = pos;
					while (s < contentEnd && (out[s] == L' ' || out[s] == L'\t'))
						++s;
					size_t e = contentEnd;
					while (e > s && (out[e - 1] == L' ' || out[e - 1] == L'\t'))
						--e;
					if (e > s)
					{
						if (out[s] == L'|')
							out[s] = L' ';
						if (e > s + 1 && out[e - 1] == L'|')
							out[e - 1] = L' ';
					}
				}

				// transform subsequent data rows (separator line itself will be hidden by spans)
				size_t tablePos = (nextLineEnd < out.size()) ? (nextLineEnd + 1) : out.size();
				while (tablePos < out.size())
				{
					size_t rowEnd = out.find(L'\n', tablePos);
					if (rowEnd == std::wstring::npos)
						rowEnd = out.size();
					size_t rowContentEnd = rowEnd;
					if (rowContentEnd > tablePos && out[rowContentEnd - 1] == L'\r')
						--rowContentEnd;
					std::wstring rowLine = out.substr(tablePos, rowContentEnd - tablePos);
					if (!IsTableRowCandidate(rowLine) || IsTableSeparatorLine(rowLine))
						break;

					TransformTableLineToTabsInPlace(out, tablePos, rowContentEnd);
					ClipTableCellsInLineInPlace(out, tablePos, rowContentEnd);
					{
						size_t s = tablePos;
						while (s < rowContentEnd && (out[s] == L' ' || out[s] == L'\t'))
							++s;
						size_t e = rowContentEnd;
						while (e > s && (out[e - 1] == L' ' || out[e - 1] == L'\t'))
							--e;
						if (e > s)
						{
							if (out[s] == L'|')
								out[s] = L' ';
							if (e > s + 1 && out[e - 1] == L'|')
								out[e - 1] = L' ';
						}
					}
					if (rowEnd >= out.size())
					{
						tablePos = out.size();
						break;
					}
					tablePos = rowEnd + 1;
				}

				pos = tablePos;
				continue;
			}

			// Typora-like marker rendering:
			// - task list: render marker as checkbox glyph
			// - unordered list: level 1 uses solid bullet, nested levels use hollow bullet
			{
				size_t blockPos = trimLeft(line);
				size_t afterMarkers = blockPos;
				size_t markerCount = 0;
				while (afterMarkers < line.size() && line[afterMarkers] == L'>')
				{
					++markerCount;
					++afterMarkers;
					if (afterMarkers < line.size() && line[afterMarkers] == L' ')
						++afterMarkers;
				}

				const size_t listIndentBasePos = (markerCount > 0) ? afterMarkers : 0;
				size_t markerPos = (markerCount > 0) ? afterMarkers : trimLeft(line);
				while (markerPos < line.size() && (line[markerPos] == L' ' || line[markerPos] == L'\t'))
					++markerPos;

				if (markerPos + 1 < line.size())
				{
					const wchar_t marker = line[markerPos];
					const bool unordered = (marker == L'-' || marker == L'*' || marker == L'+');
					const bool taskList =
						unordered &&
						(markerPos + 5 < line.size()) &&
						line[markerPos + 1] == L' ' &&
						line[markerPos + 2] == L'[' &&
						(line[markerPos + 3] == L' ' || line[markerPos + 3] == L'x' || line[markerPos + 3] == L'X') &&
						line[markerPos + 4] == L']' &&
						line[markerPos + 5] == L' ';

					if (unordered && line[markerPos + 1] == L' ')
					{
						const size_t absPos = pos + markerPos;
						if (absPos < out.size())
						{
							if (taskList)
							{
								const bool checked = (line[markerPos + 3] == L'x' || line[markerPos + 3] == L'X');
								out[absPos] = checked ? L'\x2611' : L'\x2610'; // ☑ / ☐
							}
							else
							{
								size_t indentColumns = 0;
								for (size_t i = listIndentBasePos; i < markerPos; ++i)
								{
									if (line[i] == L'\t')
									{
										const size_t remainder = indentColumns % 4;
										indentColumns += (remainder == 0) ? 4 : (4 - remainder);
									}
									else
									{
										++indentColumns;
									}
								}
								const int listDepth = static_cast<int>(indentColumns / 2);
								out[absPos] = (listDepth <= 0) ? L'\x2022' : L'\x25E6'; // • / ◦
							}
						}
					}
				}
			}

			if (lineEnd >= out.size())
				break;
			pos = lineEnd + 1;
		}

		return out;
	}

RenderText::TransformResult RenderText::TransformWithMapping(const std::wstring& rawText)
{
	TransformResult result;
	result.displayText = Transform(rawText);

	const size_t sourceLen = rawText.size();
	const size_t displayLen = result.displayText.size();

	result.mapping.displayToSource.resize(displayLen + 1);
	result.mapping.sourceToDisplay.resize(sourceLen + 1);

	for (size_t i = 0; i <= displayLen; ++i)
	{
		result.mapping.displayToSource[i] =
			static_cast<long>((i <= sourceLen) ? i : sourceLen);
	}

	for (size_t i = 0; i <= sourceLen; ++i)
	{
		result.mapping.sourceToDisplay[i] =
			static_cast<long>((i <= displayLen) ? i : displayLen);
	}

	return result;
}


