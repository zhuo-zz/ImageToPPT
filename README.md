# ImageToPPT

ImageToPPT 是一个基于 Qt Widgets 和 C++17 开发的图片转 PPT 半自动重建工具。它不是简单地把整张图片塞进幻灯片，而是让用户手动框选图片区域和文字区域，再导出为可继续编辑的 PPTX 文件。

## 主要功能

- 打开 PNG、JPG、JPEG、BMP 图片。
- 在画布上拖拽框选区域。
- 将区域设置为图片区域或文字区域。
- 图片区域会作为独立图片对象导出到 PPT。
- 文字区域支持 OCR 识别，并可编辑文字内容。
- 自动估计文字颜色、字号、粗细，并支持手动选择字体。
- 支持自动贴合边框，尽量裁掉背景边缘。
- 支持导出前预览。
- 支持背景重建和复杂背景取样填充。
- 支持撤销区域编辑操作。
- 导出标准 `.pptx` 文件。

## 开发环境

- 语言：C++17
- GUI 框架：Qt 6
- 推荐 IDE：Qt Creator
- 构建方式：
  - qmake：`ImageToPPT.pro`
  - CMake：`CMakeLists.txt`
- 可选 OCR 依赖：Tesseract OCR

## 使用方式

1. 启动程序。
2. 点击“打开图片”，选择要重建的图片。
3. 在画布上拖拽框选一个区域。
4. 在弹出的编辑窗口中选择区域类型：
   - 图片区域：用于裁取原图中的图片、图标、装饰块等。
   - 文字区域：用于识别并重建为 PPT 文本框。
5. 必要时点击“自动贴合边框”，让框选范围更贴近实际内容。
6. 对文字区域检查 OCR 结果，并调整字体、字号、颜色和加粗。
7. 点击“导出预览”检查整体效果。
8. 点击“导出 PPTX”生成文件。

## 在 Qt Creator 中运行

推荐使用 qmake 项目运行：

1. 打开 Qt Creator。
2. 选择“打开文件或项目”。
3. 打开 `ImageToPPT.pro`。
4. Kit 选择 Qt MinGW 桌面套件，例如 `Desktop Qt 6.10.2 MinGW 64-bit`。
5. 点击构建并运行。

如果使用 CMake，请确保旧的 build 目录没有缓存其他项目路径。遇到运行旧程序的情况，可以删除旧 build 目录后重新配置。

## 命令行构建

qmake 示例：

```powershell
mkdir build\codex-qmake
cd build\codex-qmake
qmake ..\..\ImageToPPT.pro
mingw32-make
```

CMake 示例：

```powershell
cmake -S . -B build\cmake -G Ninja
cmake --build build\cmake
```

具体命令取决于本机 Qt、CMake、MinGW 和 Ninja 是否已经加入 `PATH`。

## OCR 说明

OCR 功能依赖 Tesseract。没有安装 Tesseract 时，程序仍可正常使用，但文字需要手动输入。

为了提高中文识别效果，建议安装：

- `tesseract.exe`
- 简体中文语言包 `chi_sim.traineddata`
- 英文语言包 `eng.traineddata`

安装后需要确保 `tesseract.exe` 已加入系统 `PATH`，程序才能自动调用。

## 项目结构

```text
.
├── CMakeLists.txt
├── ImageToPPT.pro
├── include
│   ├── ImageCanvas.h
│   ├── MainWindow.h
│   ├── PptxExporter.h
│   ├── Region.h
│   └── RegionDialog.h
└── source
    ├── ImageCanvas.cpp
    ├── MainWindow.cpp
    ├── PptxExporter.cpp
    ├── RegionDialog.cpp
    └── main.cpp
```

## 核心模块

### ImageCanvas

负责图片显示、鼠标拖拽框选、区域绘制、命中测试和选中状态管理。

### RegionDialog

负责区域编辑，包括区域类型、文字内容、OCR、自动贴合、字体、字号、颜色和加粗设置。

### PptxExporter

负责导出逻辑，包括：

- 重建背景图。
- 裁剪图片区域。
- 绘制导出预览。
- 生成 PPTX 所需的 XML。
- 打包为 `.pptx` 文件。

### MainWindow

负责主界面流程，包括打开图片、区域列表、预览、导出、撤销和复杂背景取样提示。

## 已知限制

- OCR 准确率受图片清晰度、字体、背景复杂度和 Tesseract 语言包影响。
- 字体族无法从截图像素中精确反推，目前只做保守估计，并提供手动选择。
- 复杂纹理背景的自动修补可能不如纯色背景稳定。
- 目前主要面向单张图片生成单页 PPT。

