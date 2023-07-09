#pragma once

#include <string_view>

bool replace_extension_pdf(const std::wstring_view src, wchar_t* dst_filename) noexcept;

bool word_to_pdf(const wchar_t* src_filename, const wchar_t* dst_filename) noexcept;
bool word_to_pdf(const wchar_t** filenames, std::size_t size) noexcept;
