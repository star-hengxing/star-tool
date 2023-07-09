Now support platform:

- Windows

# Feature

## word to pdf

converter backend:

- [x] microsoft word
- [ ] libreoffice

# Build

## Prerequisites

- [xmake](https://xmake.io/#/guide/installation)
- Requires C++20 compiler.

## Setup

- [Visual Studio](https://visualstudio.microsoft.com)(If you just want to build without developing, download the [Microsoft C++ Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools))

- Recommend use [scoop](https://scoop.sh) as package manager on Windows.

```sh
scoop install xmake
```

## Build

Clone repo, then try

```sh
xmake -y
```

# Credits

- [portable-file-dialogs](https://github.com/samhocevar/portable-file-dialogs)
- [VC-LTL5](https://github.com/Chuyu-Team/VC-LTL5)
