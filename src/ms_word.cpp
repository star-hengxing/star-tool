#include <windows.h>
#include <atlbase.h>
#include <comutil.h>

#include "word.hpp"
#include "macro.hpp"

// https://www.zhihu.com/question/522794038
constexpr auto RVPtr(auto&& value) noexcept
{
    return std::addressof(value);
}

static bool convert(ATL::CComDispatchDriver& words, const wchar_t* src_filename, const wchar_t* dst_filename) noexcept
{
    ATL::CComVariant result;
    if (FAILED(words.Invoke1(L"Open", RVPtr(ATL::CComVariant{src_filename}), &result)))
        return false;

    // https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat
    auto constexpr wdFormatPDF = 17;
    // Document object
    auto word = ATL::CComDispatchDriver{result.pdispVal};

    auto const save_result = SUCCEEDED(word.Invoke2(
        L"SaveAs2",
        RVPtr(ATL::CComVariant{dst_filename}),
        RVPtr(ATL::CComVariant{wdFormatPDF})
    ));

    word.Invoke0(L"Close");
    word.Release();
    return save_result;
}

bool replace_extension_pdf(const std::wstring_view src, wchar_t* dst_filename) noexcept
{
    auto pos = src.find_last_of('.');
    if (pos == std::wstring_view::npos)
        return false;

    pos += 1;
    std::memcpy(dst_filename, src.data(), sizeof(WCHAR) * pos);
    dst_filename[pos] = L'p';
    dst_filename[pos + 1] = L'd';
    dst_filename[pos + 2] = L'f';
    dst_filename[pos + 3] = L'\0';
    return true;
}

bool word_to_pdf(const wchar_t* src_filename, const wchar_t* dst_filename) noexcept
{
    if (!src_filename || !dst_filename)
        return true;

    CoInitialize(nullptr);
    defer { CoUninitialize(); };

    ATL::CComPtr<IDispatch> app;
    if (FAILED(app.CoCreateInstance(L"Word.Application")))
        return false;

    if (FAILED(app.PutPropertyByName(L"Visible", RVPtr(ATL::CComVariant{false}))))
    {
        app.Invoke0(L"Quit");
        app.Release();
        return false;
    }

    ATL::CComVariant result;
    if (FAILED(app.GetPropertyByName(L"Documents", &result)))
    {
        app.Invoke0(L"Quit");
        app.Release();
        return false;
    }

    // Documents object
    auto words = ATL::CComDispatchDriver{result.pdispVal};
    bool ret = convert(words, src_filename, dst_filename);

    app.Invoke0(L"Quit");
    words.Release();
    app.Release();

    return ret;
}

bool word_to_pdf(const wchar_t** filenames, std::size_t size) noexcept
{
    if (!size || !filenames)
        return true;

    CoInitialize(nullptr);
    defer { CoUninitialize(); };

    ATL::CComPtr<IDispatch> app;
    if (FAILED(app.CoCreateInstance(L"Word.Application")))
        return false;

    if (FAILED(app.PutPropertyByName(L"Visible", RVPtr(ATL::CComVariant{false}))))
    {
        app.Invoke0(L"Quit");
        app.Release();
        return false;
    }

    ATL::CComVariant result;
    if (FAILED(app.GetPropertyByName(L"Documents", &result)))
    {
        app.Invoke0(L"Quit");
        app.Release();
        return false;
    }

    bool ret = true;
    WCHAR dst_filename[MAX_PATH];
    // Documents object
    auto words = ATL::CComDispatchDriver{result.pdispVal};
    for (auto i = 0; i < size; i += 1)
    {
        auto const src = std::wstring_view{filenames[i]};
        replace_extension_pdf(src, dst_filename);
        if (!convert(words, src.data(), dst_filename))
            ret = false;
    }

    app.Invoke0(L"Quit");
    words.Release();
    app.Release();

    return ret;
}
