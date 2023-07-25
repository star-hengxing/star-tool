#include <string_view>
#include <memory>

#include <Windows.h>
#define NFD_THROWS_EXCEPTIONS
#define NFD_DIFFERENT_NATIVE_FUNCTIONS
#include <nfd.hpp>

#include "word.hpp"

static std::wstring string_to_wstring(const std::string_view str) noexcept
{
    if (str.empty())
        return {};

    auto size = ::MultiByteToWideChar(
        CP_UTF8, 0, reinterpret_cast<LPCCH>(str.data()),
        static_cast<int>(str.size()), nullptr, 0);

    if (size == 0)
    {
        return {};
    }

    auto ret = std::wstring(size, L'\0');
    size = ::MultiByteToWideChar(
        CP_UTF8, 0, reinterpret_cast<LPCCH>(str.data()),
        static_cast<int>(str.size()), ret.data(), size);

    if (size == 0)
    {
        return {};
    }
    return ret;
}

int main()
{
    NFD::Guard nfd_guard{};

    constexpr nfdu8filteritem_t filiter[] = {{"Word Document", "doc,docx"}};

    NFD::UniquePathSet paths;
    nfdresult_t result = NFD::OpenDialogMultiple(paths, filiter, std::size(filiter));

    if (result == NFD_OKAY)
    {
        nfdpathsetsize_t size;
        result = NFD::PathSet::Count(paths, size);
        if (result == NFD_ERROR)
            return -1;

        if (size == 0)
            return 0;

        NFD::UniquePathSetPathU8 path;
        if (size == 1)
        {
            WCHAR dst_filename[MAX_PATH];
            NFD::PathSet::GetPath(paths, 0, path);

            auto const src = string_to_wstring(path.get());
            replace_extension_pdf(src, dst_filename);
            word_to_pdf(src.data(), dst_filename);
        }
        else
        {
            auto files_w = std::make_unique<std::wstring[]>(size);
            auto arr = std::make_unique<const wchar_t*[]>(size);
            for (nfdpathsetsize_t i = 0; i < size; i += 1)
            {
                NFD::PathSet::GetPath(paths, i, path);
                files_w[i] = string_to_wstring(path.get());
                arr[i] = files_w[i].data();
            }
            word_to_pdf(arr.get(), size);
        }
    }
}
