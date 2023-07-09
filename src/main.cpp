#include <string_view>
#include <memory>

#include <portable-file-dialogs.h>

#include "word.hpp"

static std::wstring string_to_wstring(const std::string_view str) noexcept
{
    if (str.empty())
        return {};

    auto const size = ::MultiByteToWideChar(
        CP_ACP, 0, reinterpret_cast<LPCCH>(str.data()),
        static_cast<int>(str.size()), nullptr, 0);
    if (size == 0)
    {
        return {};
    }

    auto ret = std::wstring(size, 0);
    if (::MultiByteToWideChar(
            CP_ACP, 0, reinterpret_cast<LPCCH>(str.data()),
            static_cast<int>(str.size()), ret.data(), size)
        == 0)
    {
        return {};
    }
    return ret;
}

std::vector<std::string> read_files_from_dialogs() noexcept
{
    auto const filters = std::vector<std::string>{
        "Word Document",
        "*.doc *.docx",
    };

    return pfd::open_file{"Select a file", {}, filters}.result();
}

int main()
{
    auto const files = read_files_from_dialogs();
    auto const size = files.size();
    if (size == 1)
    {
        WCHAR dst_filename[MAX_PATH];
        auto const src = string_to_wstring(files.front());
        replace_extension_pdf(src, dst_filename);
        word_to_pdf(src.data(), dst_filename);
    }
    else
    {
        auto files_w = std::make_unique<std::wstring[]>(size);
        auto arr = std::make_unique<const wchar_t*[]>(size);
        for (auto i = 0; i < size; i += 1)
        {
            files_w[i] = string_to_wstring(files[i]);
            arr[i] = files_w[i].data();
        }
        word_to_pdf(arr.get(), size);
    }
}
