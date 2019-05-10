# VB6 compression library wrapper(s)

This project provides convenient VB6 wrappers for a variety of compression libraries, including...
* zLib -- https://github.com/madler/zlib
* zLib-ng -- https://github.com/Dead2/zlib-ng
* libdeflate -- https://github.com/ebiggers/libdeflate
* zstd -- https://github.com/facebook/zstd
* lz4 -- https://github.com/lz4/lz4
* brotli -- https://github.com/google/brotli
* ZipArchive (zero-dependency VB6 lib derived from PuTTY) -- https://github.com/wqweto/ZipArchive
* Microsoft Compression APIs (Win 8+ only) -- https://docs.microsoft.com/en-us/windows/desktop/cmpapi/-compression-portal

A small sample project allows you to compare compression time, decompression time, and compression ratio across all libraries.  Drag+drop a file onto the text box to test it.

The bulk of this project is adopted from [the PhotoDemon project](https://github.com/tannerhelland/PhotoDemon) which is BSD-licensed.  The 3rd-party compression libraries used in this project have their own licenses; please refer to LICENSE.md for full details.
