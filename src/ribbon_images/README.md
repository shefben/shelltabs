# Ribbon button image placeholders

This directory hosts the PNG assets that back the ribbon commands referenced in
`src/CustomRibbonTab.xml` and registered in `src/resource.rc`.

Provide ten PNG files (five small glyphs at 16x16 and five large glyphs at
32x32) using the following naming convention so the resource compiler can pick
them up automatically:

```
button1_small.png    button1_large.png
button2_small.png    button2_large.png
button3_small.png    button3_large.png
button4_small.png    button4_large.png
button5_small.png    button5_large.png
```

The repository intentionally keeps these assets out of source control. Drop
your project-specific images into this folder locally before running the ribbon
compiler or invoking MSBuild. The UICC step resolves the `<Image
Source="ribbon_images/...">` entries in `CustomRibbonTab.xml`, and `rc.exe`
consumes the exact same files when emitting the `IMAGE_BUTTON*_SMALL` and
`IMAGE_BUTTON*_LARGE` resources, so maintaining these filenames ensures both
toolchains stay in sync.

> **Build integration note:** The CMake project watches these exact file names
> as explicit dependencies of `src/resource.rc`. When you copy your PNGs into
> this directory, `cmake --build` will automatically rebuild the resource file
> so that UICC and `rc.exe` pick up your glyph updates without any manual
> project fiddling.
