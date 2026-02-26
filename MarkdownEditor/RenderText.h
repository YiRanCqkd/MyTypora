#pragma once

#include <string>
#include <vector>

class RenderText
{
public:
    struct IndexMapping
    {
        // display index -> source index
        std::vector<long> displayToSource;
        // source index -> display index
        std::vector<long> sourceToDisplay;
    };

    struct TransformResult
    {
        std::wstring displayText;
        IndexMapping mapping;
    };

    static std::wstring Transform(const std::wstring& rawText);
    static TransformResult TransformWithMapping(const std::wstring& rawText);
};


