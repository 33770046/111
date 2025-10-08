QT       += core gui widgets axcontainer network
CONFIG   += c++17
TARGET    = GroupingSystem
TEMPLATE  = app

win32: LIBS += -lOle32 -lOleAut32

# 添加设置支持
QT += settings

# 设置编码
win32 {
    QMAKE_CXXFLAGS += /utf-8
    DEFINES += _WIN32_WINNT=0x0601
    DEFINES += WIN32_LEAN_AND_MEAN
    # 设置应用程序图标（Windows 专用）
    RC_FILE = app.rc  # 指定 .rc 文件
}
macx {
    QMAKE_MACOSX_DEPLOYMENT_TARGET = 10.14
}

# 添加资源文件（用于运行时）
RESOURCES += resources.qrc

# 源文件
SOURCES += \
    main.cpp \
    mainwindow_ui.cpp \
    mainwindow_core.cpp \
    mainwindow_algorithm.cpp \
    mainwindow_updates.cpp

# 头文件
HEADERS += \
    mainwindow.h

# 启用调试信息
CONFIG += debug