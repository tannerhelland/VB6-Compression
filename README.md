# VB6 Compression API wrapper(s)

This project provides convenient VB6 wrappers for multiple compression APIs, including...
* zLib -- https://github.com/madler/zlib
* zLib-ng -- https://github.com/Dead2/zlib-ng
* zstd -- https://github.com/facebook/zstd
* lz4 -- https://github.com/lz4/lz4
* ZipArchive -- https://github.com/wqweto/ZipArchive
* Microsoft Compression APIs (Win 8+ only) -- https://msdn.microsoft.com/en-us/library/windows/desktop/hh920921(v=vs.85).aspx

A small sample project allows you to compare compression time, decompression time, and compression ratio across all libraries.  Drag+drop a file onto the text box to test it.

The bulk of this project is adopted from [the PhotoDemon project](https://github.com/tannerhelland/PhotoDemon) which is BSD-licensed.  The 3rd-party compression libraries used in this project have their own licenses; please refer to LICENSE.md for full details.
