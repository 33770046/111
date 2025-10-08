#include "mainwindow.h"
#include <QNetworkReply>
#include <QJsonDocument>
#include <QJsonObject>
#include <QMessageBox>
#include <QDesktopServices>
#include <QApplication>
#include <QVersionNumber>

void MainWindow::checkForUpdates(bool silent)
{
    if (!silent) {
        updateStatus("正在检查更新...");
    }

    // 混合模式：先尝试自动检查，失败则手动
    QUrl url("https://api.github.com/repos/33770046/111/releases/latest");
    QNetworkRequest request(url);
    request.setHeader(QNetworkRequest::UserAgentHeader, "SmartGroupingSystem/1.0");
    request.setAttribute(QNetworkRequest::Http2AllowedAttribute, false);

    QNetworkReply* reply = networkManager->get(request);

    // 设置超时，如果自动检查失败就转手动
    QTimer* timeoutTimer = new QTimer(this);
    timeoutTimer->setSingleShot(true);
    timeoutTimer->start(8000); // 8秒超时

    connect(timeoutTimer, &QTimer::timeout, this, [this, silent, reply, timeoutTimer]() {
        timeoutTimer->deleteLater();
        if (reply && reply->isRunning()) {
            reply->abort();
        }
        // 自动检查超时，转为手动检查
        manualUpdateCheck(silent);
        });

    connect(reply, &QNetworkReply::finished, this, [this, silent, reply, timeoutTimer]() {
        timeoutTimer->stop();
        timeoutTimer->deleteLater();

        if (reply->error() == QNetworkReply::NoError) {
            // 自动检查成功
            onUpdateCheckFinished(silent, reply);
        }
        else {
            // 自动检查失败，转为手动
            manualUpdateCheck(silent);
            reply->deleteLater();
        }
        });
}

//这里！当前版本:XXX 3377004_6
//这里！当前版本:XXX 3377004_6
//这里！当前版本:XXX 3377004_6
//这里！当前版本:XXX 3377004_6
//这里！当前版本:XXX 3377004_6
void MainWindow::manualUpdateCheck(bool silent)
{
    if (!silent) {
        int result = QMessageBox::information(this, "检查更新",
            "当前版本: 251008.1730\n\n"
            "是否要打开项目页面手动检查更新？",
            QMessageBox::Yes | QMessageBox::No);

        if (result == QMessageBox::Yes) {
            QDesktopServices::openUrl(QUrl("https://github.com/33770046/111/releases"));
            updateStatus("已打开更新页面");
        }
        else {
            updateStatus("取消检查更新");
        }
    }
}

//这里！当前版本:XXX 3377004_6
//这里！当前版本:XXX 3377004_6
//这里！当前版本:XXX 3377004_6
//这里！当前版本:XXX 3377004_6
//这里！当前版本:XXX 3377004_6
void MainWindow::manualUpdateCheck()
{
    int result = QMessageBox::information(this, "检查更新",
        "当前版本: 251008.1730\n\n"
        "自动检查更新失败，是否要手动打开项目页面检查？",
        QMessageBox::Yes | QMessageBox::No);

    if (result == QMessageBox::Yes) {
        QDesktopServices::openUrl(QUrl("https://github.com/33770046/111/releases"));
        updateStatus("已打开更新页面");
    }
    else {
        updateStatus("取消检查更新");
    }
}

void MainWindow::onUpdateCheckFinished(bool silent, QNetworkReply* reply)
{
    isCheckingUpdates = false;

    if (!reply) {
        return;
    }

    bool success = false;
    QString latestVersion;
    QString downloadUrl;
    QString releaseNotes;

    if (reply->error() == QNetworkReply::NoError) {
        QByteArray response = reply->readAll();

        // 调试输出响应内容
        qDebug() << "GitHub API Response:" << response;

        QJsonDocument doc = QJsonDocument::fromJson(response);

        if (!doc.isNull()) {
            QJsonObject json = doc.object();
            latestVersion = json["tag_name"].toString();
            downloadUrl = json["html_url"].toString();
            releaseNotes = json["body"].toString();

            if (!latestVersion.isEmpty()) {
                success = true;

                // 移除版本号前的 'v' 字符（如果有）
                if (latestVersion.startsWith('v')) {
                    latestVersion = latestVersion.mid(1);
                }

                //这里！当前版本:XXX 3377004_6
                //这里！当前版本:XXX 3377004_6
                //这里！当前版本:XXX 3377004_6
                //这里！当前版本:XXX 3377004_6
                //这里！当前版本:XXX 3377004_6
                // 比较版本
                QVersionNumber currentVersion = QVersionNumber::fromString("251008.1730");
                QVersionNumber latest = QVersionNumber::fromString(latestVersion);

                if (latest > currentVersion) {
                    // 有新版本
                    showUpdateDialog(latestVersion, downloadUrl, releaseNotes);
                }
                else {
                    updateStatus("当前已是最新版本");
                    if (!silent) {
                        QMessageBox::information(this, "检查更新", "当前已是最新版本！");
                    }
                }
            }
        }
    }

    if (!success) {
        QString errorMsg;
        if (reply->error() != QNetworkReply::NoError) {
            errorMsg = QString("网络请求失败: %1\n错误代码: %2")
                .arg(reply->errorString())
                .arg(reply->error());
        }
        else {
            errorMsg = "无法解析更新信息，可能是API限制或仓库不存在";
        }

        updateStatus("检查更新失败");
        if (!silent) {
            QMessageBox::warning(this, "检查更新", errorMsg);
        }

        // 调试信息
        qDebug() << "Update check failed:" << errorMsg;
        qDebug() << "Response:" << reply->readAll();
    }

    reply->deleteLater();
}

void MainWindow::showUpdateDialog(const QString& latestVersion, const QString& downloadUrl, const QString& releaseNotes)
{
    QDialog dialog(this);
    dialog.setWindowTitle("发现新版本");
    dialog.setMinimumWidth(500);

    QVBoxLayout* layout = new QVBoxLayout(&dialog);

    //这里！当前版本:XXX 3377004_6
    //这里！当前版本:XXX 3377004_6
    //这里！当前版本:XXX 3377004_6
    //这里！当前版本:XXX 3377004_6
    //这里！当前版本:XXX 3377004_6
    // 版本信息
    QLabel* versionLabel = new QLabel(
        QString("发现新版本: <b>%1</b><br>当前版本: 251008.1730").arg(latestVersion)
    );
    layout->addWidget(versionLabel);

    // 更新说明
    if (!releaseNotes.isEmpty()) {
        QLabel* notesLabel = new QLabel("更新说明:");
        layout->addWidget(notesLabel);

        QTextEdit* notesEdit = new QTextEdit();
        notesEdit->setPlainText(releaseNotes);
        notesEdit->setReadOnly(true);
        notesEdit->setMaximumHeight(150);
        layout->addWidget(notesEdit);
    }

    // 按钮区域
    QHBoxLayout* buttonLayout = new QHBoxLayout();

    QPushButton* downloadBtn = new QPushButton("前往下载");
    QPushButton* laterBtn = new QPushButton("稍后提醒");
    QPushButton* ignoreBtn = new QPushButton("忽略此版本");

    buttonLayout->addWidget(downloadBtn);
    buttonLayout->addWidget(laterBtn);
    buttonLayout->addWidget(ignoreBtn);

    layout->addLayout(buttonLayout);

    // 连接信号
    connect(downloadBtn, &QPushButton::clicked, [&]() {
        QDesktopServices::openUrl(QUrl(downloadUrl));
        dialog.accept();
        });

    connect(laterBtn, &QPushButton::clicked, &dialog, &QDialog::accept);
    connect(ignoreBtn, &QPushButton::clicked, &dialog, &QDialog::reject);

    dialog.exec();
}