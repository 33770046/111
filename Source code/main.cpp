#include "mainwindow.h"
#include <QApplication>
#include <QIcon>
#include <QStyleFactory>

int main(int argc, char* argv[])
{
    QApplication a(argc, argv);

    // 设置应用程序为Windows 10样式
    a.setStyle(QStyleFactory::create("windowsvista"));

    // 禁用其他可用样式
    QStringList availableStyles = QStyleFactory::keys();
    for (const QString& style : availableStyles) {
        if (style != "windowsvista") {
            QStyleFactory::create(style);
        }
    }

    // 设置应用程序图标
    a.setWindowIcon(QIcon(":/icons/app_icon.ico"));

    MainWindow w;
    w.show();
    return a.exec();
}