#include "mainwindow.h"
#include <QHeaderView>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QFile>
#include <QTextStream>
#include <QDebug>
#include <QDesktopServices>
#include <QDir>
#include <QMenuBar>
#include <QSettings>
#include <QDialog>
#include <QTextEdit>
#include <QMessageBox>
#include <QProgressDialog>
#include <QDateTime>
#include <objbase.h>
#include <comdef.h>
#include <qspinbox.h>
#include <QRegularExpression>
#include <QRegularExpressionMatch>

#ifdef Q_OS_WIN
#include <windows.h>
#include <QAxObject>
#include <QAxBase>
#endif
#include <QTextBrowser>

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent), rng(static_cast<unsigned int>(time(nullptr))),
    fileMenu(nullptr), mainLayout(nullptr), networkManager(nullptr), isCheckingUpdates(false)
{
    setWindowIcon(QIcon(":/icons/app_icon.ico"));
    setWindowTitle("智能分组系统");
    setMinimumSize(1004, 753);

    // 主窗口的中心部件
    QWidget* centralWidget = new QWidget(this);
    setCentralWidget(centralWidget);

    // 主布局 - 垂直布局
    mainLayout = new QVBoxLayout(centralWidget);

    // 分组结果 + 系统日志
    QHBoxLayout* topLayout = new QHBoxLayout();

    // 分组表格
    QGroupBox* groupBox = new QGroupBox("分组结果", this);
    QVBoxLayout* groupLayout = new QVBoxLayout(groupBox);
    groupTable = new QTableWidget(this);
    groupTable->setColumnCount(7);
    groupTable->setHorizontalHeaderLabels(QStringList{ "组号", "成员1", "成员2", "成员3", "成员4", "成员5", "成员6" });
    groupTable->verticalHeader()->setVisible(false);
    groupTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    groupTable->setSelectionBehavior(QAbstractItemView::SelectRows);

    // 设置所有列等宽
    groupTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);

    groupLayout->addWidget(groupTable);
    topLayout->addWidget(groupBox, 2);

    // 系统日志
    QGroupBox* logBox = new QGroupBox("系统日志", this);
    QVBoxLayout* logLayout = new QVBoxLayout(logBox);
    logOutput = new QTextEdit(this);
    logOutput->setReadOnly(true);
    logLayout->addWidget(logOutput);
    topLayout->addWidget(logBox, 1);

    mainLayout->addLayout(topLayout);

    // 按钮区域
    QHBoxLayout* buttonLayout = new QHBoxLayout();
    generateBtn = new QPushButton("生成分组", this);
    personnelCheckBtn = new QPushButton("人员检查", this);
    checkBtn = new QPushButton("要求检查", this);
    resetBtn = new QPushButton("重置", this);

    buttonLayout->addWidget(generateBtn);
    buttonLayout->addWidget(personnelCheckBtn);
    buttonLayout->addWidget(checkBtn);
    buttonLayout->addWidget(resetBtn);
    mainLayout->addLayout(buttonLayout);

    // 要求条件管理
    QGroupBox* constraintBox = new QGroupBox("要求条件管理", this);
    QGridLayout* constraintLayout = new QGridLayout(constraintBox);

    constraintLayout->addWidget(new QLabel("要求类型:"), 0, 0);
    constraintTypeCombo = new QComboBox(this);
    constraintTypeCombo->addItems(QStringList{ "必须同组", "不能同组" });
    constraintLayout->addWidget(constraintTypeCombo, 0, 1);

    constraintLayout->addWidget(new QLabel("姓名(多个用空格分隔):"), 1, 0);
    nameInput = new QLineEdit(this);
    constraintLayout->addWidget(nameInput, 1, 1);

    addConstraintBtn = new QPushButton("添加要求", this);
    removeConstraintBtn = new QPushButton("移除选中", this);
    constraintLayout->addWidget(addConstraintBtn, 0, 2);
    constraintLayout->addWidget(removeConstraintBtn, 1, 2);

    constraintTable = new QTableWidget(this);
    constraintTable->setColumnCount(3);
    constraintTable->setHorizontalHeaderLabels(QStringList{ "类型", "人员列表", "状态" });
    constraintTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    constraintTable->setSelectionBehavior(QAbstractItemView::SelectRows);

    // 设置要求表格的列宽调整
    constraintTable->horizontalHeader()->setSectionResizeMode(0, QHeaderView::ResizeToContents);
    constraintTable->horizontalHeader()->setSectionResizeMode(1, QHeaderView::Stretch);
    constraintTable->horizontalHeader()->setSectionResizeMode(2, QHeaderView::ResizeToContents);

    constraintLayout->addWidget(constraintTable, 2, 0, 1, 3);
    mainLayout->addWidget(constraintBox);

    // 状态栏
    statusLabel = new QLabel("就绪", this);
    statusBar()->addWidget(statusLabel);

    // 现在再初始化名称映射和分组配置
    initNameMapping();  // 这会加载名单、组长和外宿生设置
    initGroupConfigs(); // 这会初始化分组配置
    loadSettings();     // 这会从配置文件加载分组配置

    // 更新要求条件表格
    updateConstraintTable();

    // 连接信号槽
    connect(addConstraintBtn, &QPushButton::clicked, this, &MainWindow::addConstraint);
    connect(removeConstraintBtn, &QPushButton::clicked, this, &MainWindow::removeConstraint);
    connect(generateBtn, &QPushButton::clicked, this, &MainWindow::generateGroups);
    connect(personnelCheckBtn, &QPushButton::clicked, this, &MainWindow::checkPersonnel);
    connect(checkBtn, &QPushButton::clicked, this, &MainWindow::checkConstraints);
    connect(resetBtn, &QPushButton::clicked, this, &MainWindow::resetAll);
    connect(groupTable, &QTableWidget::cellClicked, this, &MainWindow::showGroupDetails);

    // 创建菜单栏
    createMenuBar();

    // 网络管理器初始化
    networkManager = new QNetworkAccessManager(this);
    isCheckingUpdates = false;

    // 启动时静默检查更新（每周一次）
    QSettings settings(configPath, QSettings::IniFormat);
    QDate lastUpdateCheck = settings.value("Update/LastCheck").toDate();
    if (!lastUpdateCheck.isValid() || lastUpdateCheck.daysTo(QDate::currentDate()) >= 7) {
        QTimer::singleShot(3000, this, [this]() {
            checkForUpdates(true);
            });
        settings.setValue("Update/LastCheck", QDate::currentDate());
    }

    // 启动时静默检查更新（每周一次）
    QSettings settings(configPath, QSettings::IniFormat);
    QDate lastUpdateCheck = settings.value("Update/LastCheck").toDate();
    if (!lastUpdateCheck.isValid() || lastUpdateCheck.daysTo(QDate::currentDate()) >= 7) {
        QTimer::singleShot(5000, this, [this]() {
            checkForUpdates(true); // 静默检查
            });
        settings.setValue("Update/LastCheck", QDate::currentDate());
    }

    // 初始化日志
    if (logOutput) {
        logOutput->append("系统初始化完成");
    }
}

void MainWindow::createMenuBar()
{
    QMenuBar* menuBar = new QMenuBar(this);

    // 添加文件菜单
    fileMenu = menuBar->addMenu("文件");

    QAction* newFileAction = new QAction("新建座位表", this);
    QAction* saveAction = new QAction("保存", this);
    QAction* saveAsAction = new QAction("另存为", this);
    QAction* exitAction = new QAction("退出", this);

    fileMenu->addAction(newFileAction);
    fileMenu->addAction(saveAction);
    fileMenu->addAction(saveAsAction);
    fileMenu->addSeparator();
    fileMenu->addAction(exitAction);

    // 添加设置菜单项
    QAction* settingsAction = new QAction("设置", this);
    menuBar->addAction(settingsAction);

    // 添加帮助菜单
    QMenu* helpMenu = menuBar->addMenu("帮助");

    QAction* checkUpdateAction = new QAction("检查更新", this);
    QAction* aboutAction = new QAction("关于", this);

    helpMenu->addAction(checkUpdateAction);
    helpMenu->addSeparator();
    helpMenu->addAction(aboutAction);

    // 连接文件菜单的信号槽
    connect(newFileAction, &QAction::triggered, this, &MainWindow::newFile);
    connect(saveAction, &QAction::triggered, this, &MainWindow::saveFile);
    connect(saveAsAction, &QAction::triggered, this, &MainWindow::saveAs);
    connect(exitAction, &QAction::triggered, this, &QMainWindow::close);

    // 连接其他菜单的信号槽
    connect(settingsAction, &QAction::triggered, this, &MainWindow::showSettingsDialog);
    connect(checkUpdateAction, &QAction::triggered, this, [this]() {
        checkForUpdates(false);
        });
    connect(aboutAction, &QAction::triggered, this, &MainWindow::showAboutDialog);

    this->setMenuBar(menuBar);
}

void MainWindow::showSettingsDialog()
{
    QDialog dialog(this);
    dialog.setWindowTitle("设置");
    dialog.resize(600, 500);

    // 使用选项卡布局
    QTabWidget* tabWidget = new QTabWidget(&dialog);
    QVBoxLayout* mainLayout = new QVBoxLayout(&dialog);
    mainLayout->addWidget(tabWidget);

    // ====================== 人员名单选项卡 ======================
    QWidget* personnelListTab = new QWidget;
    QVBoxLayout* personnelListLayout = new QVBoxLayout(personnelListTab);

    // 男生名单编辑
    QGroupBox maleGroup("男生名单 (每行一个名字)");
    QVBoxLayout maleLayout(&maleGroup);
    QTextEdit maleEdit;
    maleEdit.setPlainText(male_names.join("\n"));
    maleLayout.addWidget(&maleEdit);

    // 女生名单编辑
    QGroupBox femaleGroup("女生名单 (每行一个名字)");
    QVBoxLayout femaleLayout(&femaleGroup);
    QTextEdit femaleEdit;
    femaleEdit.setPlainText(female_names.join("\n"));
    femaleLayout.addWidget(&femaleEdit);

    personnelListLayout->addWidget(&maleGroup);
    personnelListLayout->addWidget(&femaleGroup);
    personnelListLayout->addStretch(1);

    // ====================== 人员设置选项卡 ======================
    QWidget* personnelTab = new QWidget;
    QVBoxLayout* personnelLayout = new QVBoxLayout(personnelTab);

    // 组长设置
    QGroupBox leadersGroup("组长设置");
    QVBoxLayout leadersLayout(&leadersGroup);

    QLabel leadersLabel("选择组长(每行一个名字):");
    QTextEdit* leadersEdit = new QTextEdit();
    leadersEdit->setPlainText(QStringList(leaders.values()).join("\n"));

    leadersLayout.addWidget(&leadersLabel);
    leadersLayout.addWidget(leadersEdit);
    personnelLayout->addWidget(&leadersGroup);

    // 外宿生设置
    QGroupBox boardersGroup("外宿生设置");
    QVBoxLayout boardersLayout(&boardersGroup);

    QLabel boardersLabel("选择外宿生(每行一个名字):");
    QTextEdit* boardersEdit = new QTextEdit();
    boardersEdit->setPlainText(QStringList(boarders.values()).join("\n"));

    boardersLayout.addWidget(&boardersLabel);
    boardersLayout.addWidget(boardersEdit);
    personnelLayout->addWidget(&boardersGroup);

    personnelLayout->addStretch(1);

    // ====================== 分组设置选项卡 ======================
    QWidget* groupConfigTab = new QWidget;
    QGridLayout* groupConfigLayout = new QGridLayout(groupConfigTab);

    // 添加表头
    groupConfigLayout->addWidget(new QLabel("组号"), 0, 0);
    groupConfigLayout->addWidget(new QLabel("启用"), 0, 1);
    groupConfigLayout->addWidget(new QLabel("总人数"), 0, 2);
    groupConfigLayout->addWidget(new QLabel("男生"), 0, 3);
    groupConfigLayout->addWidget(new QLabel("女生"), 0, 4);

    // 创建分组配置控件
    QVector<QComboBox*> enableCombos;
    QVector<QSpinBox*> totalSpins;
    QVector<QSpinBox*> maleSpins;
    QVector<QSpinBox*> femaleSpins;

    for (int i = 0; i < 10; i++) {
        int row = i + 1;

        // 组号标签
        groupConfigLayout->addWidget(new QLabel(QString("组 %1").arg(i + 1)), row, 0);

        // 启用下拉框
        QComboBox* enableCombo = new QComboBox;
        enableCombo->addItem("启用", true);
        enableCombo->addItem("禁用", false);
        enableCombo->setCurrentIndex(groupConfigs[i].enabled ? 0 : 1);
        groupConfigLayout->addWidget(enableCombo, row, 1);
        enableCombos.append(enableCombo);

        // 总人数
        QSpinBox* totalSpin = new QSpinBox;
        totalSpin->setRange(0, 10);
        totalSpin->setValue(groupConfigs[i].total);
        groupConfigLayout->addWidget(totalSpin, row, 2);
        totalSpins.append(totalSpin);

        // 男生人数
        QSpinBox* maleSpin = new QSpinBox;
        maleSpin->setRange(0, 10);
        maleSpin->setValue(groupConfigs[i].males);
        groupConfigLayout->addWidget(maleSpin, row, 3);
        maleSpins.append(maleSpin);

        // 女生人数
        QSpinBox* femaleSpin = new QSpinBox;
        femaleSpin->setRange(0, 10);
        femaleSpin->setValue(groupConfigs[i].females);
        femaleSpin->setEnabled(false);
        groupConfigLayout->addWidget(femaleSpin, row, 4);
        femaleSpins.append(femaleSpin);

        connect(totalSpin, QOverload<int>::of(&QSpinBox::valueChanged), [=](int value) {
            maleSpin->setMaximum(value);
            femaleSpin->setValue(value - maleSpin->value());
            });

        connect(maleSpin, QOverload<int>::of(&QSpinBox::valueChanged), [=](int value) {
            femaleSpin->setValue(totalSpin->value() - value);
            });

        // 初始时根据启用状态设置控件状态
        auto updateGroupState = [=](bool enabled) {
            totalSpin->setEnabled(enabled);
            maleSpin->setEnabled(enabled);
            femaleSpin->setEnabled(false);
            };

        updateGroupState(enableCombo->currentData().toBool());

        // 当启用状态改变时更新控件状态
        connect(enableCombo, QOverload<int>::of(&QComboBox::currentIndexChanged), [=]() {
            bool enabled = enableCombo->currentData().toBool();
            updateGroupState(enabled);
            });
    }

    // ====================== 样式设置选项卡 ======================
    QWidget* styleTab = new QWidget;
    QVBoxLayout* styleLayout = new QVBoxLayout(styleTab);

    QGroupBox styleGroup("座位表样式设置");
    QVBoxLayout styleGroupLayout(&styleGroup);

    QLabel styleLabel("选择座位表样式:");
    styleGroupLayout.addWidget(&styleLabel);

    QComboBox* styleCombo = new QComboBox;
    styleCombo->addItem("样式1", "样式1");
    styleCombo->addItem("样式2", "样式2");
    styleCombo->addItem("自定义", "自定义");

    // 设置当前选择的样式
    int currentIndex = styleCombo->findData(currentStyle);
    if (currentIndex >= 0) {
        styleCombo->setCurrentIndex(currentIndex);
    }

    styleGroupLayout.addWidget(styleCombo);

    // 自定义样式文件选择
    QHBoxLayout* customStyleLayout = new QHBoxLayout;
    QLabel customStyleLabel("自定义样式文件:");
    QLineEdit* customStyleEdit = new QLineEdit;
    QPushButton* browseButton = new QPushButton("浏览...");

    // 如果当前是自定义样式，显示自定义文件路径
    if (currentStyle == "自定义") {
        QSettings settings(configPath, QSettings::IniFormat);
        QString customStylePath = settings.value("Style/CustomPath", "").toString();
        customStyleEdit->setText(customStylePath);
    }

    customStyleLayout->addWidget(&customStyleLabel);
    customStyleLayout->addWidget(customStyleEdit);
    customStyleLayout->addWidget(browseButton);

    // 初始时根据当前选择显示/隐藏自定义样式选项
    bool showCustom = (currentStyle == "自定义");
    customStyleLabel.setVisible(showCustom);
    customStyleEdit->setVisible(showCustom);
    browseButton->setVisible(showCustom);

    styleGroupLayout.addLayout(customStyleLayout);
    styleLayout->addWidget(&styleGroup);

    // 连接样式选择改变信号
    connect(styleCombo, QOverload<int>::of(&QComboBox::currentIndexChanged), [&](int index) {
        QString selectedStyle = styleCombo->itemData(index).toString();
        bool showCustom = (selectedStyle == "自定义");
        customStyleLabel.setVisible(showCustom);
        customStyleEdit->setVisible(showCustom);
        browseButton->setVisible(showCustom);
        });

    // 连接浏览按钮
    connect(browseButton, &QPushButton::clicked, [&]() {
        QString fileName = QFileDialog::getOpenFileName(this, "选择自定义样式文件",
            QCoreApplication::applicationDirPath(),
            "Excel文件 (*.xlsx)");
        if (!fileName.isEmpty()) {
            customStyleEdit->setText(fileName);
        }
        });

    // ====================== 固定位置设置选项卡 ======================
    QWidget* fixedPositionTab = new QWidget;
    QVBoxLayout* fixedPositionLayout = new QVBoxLayout(fixedPositionTab);

    // 姓名输入
    QHBoxLayout* nameLayout = new QHBoxLayout();
    QLabel* nameLabel = new QLabel("姓名:");
    QLineEdit* nameInputLocal = new QLineEdit(&dialog);
    nameLayout->addWidget(nameLabel);
    nameLayout->addWidget(nameInputLocal);

    // 组号和位置输入
    QHBoxLayout* positionLayout = new QHBoxLayout();
    QLabel* groupLabel = new QLabel("组号:");
    QSpinBox* groupSpin = new QSpinBox(&dialog);
    groupSpin->setRange(1, 10);
    QLabel* posLabel = new QLabel("座位位置:");
    QSpinBox* posSpin = new QSpinBox(&dialog);
    posSpin->setRange(1, 6);

    positionLayout->addWidget(groupLabel);
    positionLayout->addWidget(groupSpin);
    positionLayout->addWidget(posLabel);
    positionLayout->addWidget(posSpin);

    // 添加和删除按钮
    QPushButton* addFixedPositionBtn = new QPushButton("添加", &dialog);
    QPushButton* removeFixedPositionBtn = new QPushButton("删除选中", &dialog);
    QPushButton* clearFixedPositionsBtn = new QPushButton("清空所有", &dialog);

    QHBoxLayout* buttonLayoutLocal = new QHBoxLayout();
    buttonLayoutLocal->addWidget(addFixedPositionBtn);
    buttonLayoutLocal->addWidget(removeFixedPositionBtn);
    buttonLayoutLocal->addWidget(clearFixedPositionsBtn);

    // 固定位置表格
    QTableWidget* fixedPositionTableLocal = new QTableWidget(&dialog);
    fixedPositionTableLocal->setColumnCount(3);
    fixedPositionTableLocal->setHorizontalHeaderLabels(QStringList{ "姓名", "组号", "座位位置" });
    fixedPositionTableLocal->setEditTriggers(QAbstractItemView::NoEditTriggers);
    fixedPositionTableLocal->setSelectionBehavior(QAbstractItemView::SelectRows);
    fixedPositionTableLocal->setSelectionMode(QAbstractItemView::MultiSelection);

    // 添加到布局
    fixedPositionLayout->addLayout(nameLayout);
    fixedPositionLayout->addLayout(positionLayout);
    fixedPositionLayout->addLayout(buttonLayoutLocal);
    fixedPositionLayout->addWidget(fixedPositionTableLocal);

    // 更新表格显示
    updateFixedPositionTable(fixedPositionTableLocal);

    // 添加选项卡
    tabWidget->addTab(personnelListTab, "人员名单");
    tabWidget->addTab(personnelTab, "人员设置");
    tabWidget->addTab(groupConfigTab, "分组设置");
    tabWidget->addTab(styleTab, "样式设置");
    tabWidget->addTab(fixedPositionTab, "固定位置");

    // ====================== 按钮区域 ======================
    QHBoxLayout* buttonLayout = new QHBoxLayout();
    QPushButton* saveButton = new QPushButton("保存");
    QPushButton* cancelButton = new QPushButton("取消");
    buttonLayout->addWidget(saveButton);
    buttonLayout->addWidget(cancelButton);
    mainLayout->addLayout(buttonLayout);

    // ====================== 连接固定位置按钮信号槽 ======================
    connect(addFixedPositionBtn, &QPushButton::clicked, [&]() {
        QString name = nameInputLocal->text().trimmed();
        int group = groupSpin->value();
        int seat = posSpin->value();

        if (name.isEmpty()) {
            QMessageBox::warning(&dialog, "错误", "请输入姓名");
            return;
        }

        // 验证姓名
        if (!name_to_id.contains(name)) {
            QMessageBox::warning(&dialog, "错误", QString("姓名 '%1' 不在名单中").arg(name));
            return;
        }

        int personId = name_to_id[name];

        // 验证座位位置
        if (seat < 1 || seat > 6) {
            QMessageBox::warning(&dialog, "错误", "座位位置必须在1-6之间");
            return;
        }

        // 检查是否已存在该人员的固定位置
        if (fixedPositions.contains(personId)) {
            FixedPosition existing = fixedPositions[personId];
            if (QMessageBox::question(&dialog, "确认",
                QString("%1 已被固定到组%2座位%3，是否替换为组%4座位%5?")
                .arg(name).arg(existing.groupNumber).arg(existing.seatPosition)
                .arg(group).arg(seat)) != QMessageBox::Yes) {
                return;
            }
        }

        // 检查目标座位是否已被占用
        bool seatOccupied = false;
        for (auto it = fixedPositions.begin(); it != fixedPositions.end(); ++it) {
            if (it.key() != personId && it.value().groupNumber == group && it.value().seatPosition == seat) {
                QString occupiedName = id_to_name[it.key()];
                QMessageBox::warning(&dialog, "错误",
                    QString("组%1座位%2已被 %3 占用").arg(group).arg(seat).arg(occupiedName));
                seatOccupied = true;
                break;
            }
        }

        if (seatOccupied) {
            return;
        }

        // 添加到固定位置映射
        fixedPositions[personId] = FixedPosition{ group, seat };

        // 更新表格
        updateFixedPositionTable(fixedPositionTableLocal);
        nameInputLocal->clear();
        groupSpin->setValue(1);
        posSpin->setValue(1);
        });

    // 在设置对话框中修改固定位置的删除逻辑
    connect(removeFixedPositionBtn, &QPushButton::clicked, [&]() {
        QModelIndexList selected = fixedPositionTableLocal->selectionModel()->selectedRows();
        if (selected.isEmpty()) {
            QMessageBox::information(&dialog, "提示", "请先选择要删除的固定位置");
            return;
        }

        // 使用集合记录要删除的人员ID
        QList<int> idsToRemove;

        for (const QModelIndex& index : selected) {
            int row = index.row();
            QTableWidgetItem* nameItem = fixedPositionTableLocal->item(row, 0);
            if (nameItem) {
                QString name = nameItem->text();
                if (name_to_id.contains(name)) {
                    idsToRemove.append(name_to_id[name]);
                }
            }
        }

        for (int id : idsToRemove) {
            fixedPositions.remove(id);
        }

        updateFixedPositionTable(fixedPositionTableLocal);
        });

    connect(clearFixedPositionsBtn, &QPushButton::clicked, [&]() {
        if (QMessageBox::question(&dialog, "确认", "确定要清空所有固定位置吗?") == QMessageBox::Yes) {
            fixedPositions.clear();
            updateFixedPositionTable(fixedPositionTableLocal);
        }
        });

    // ====================== 保存按钮逻辑 ======================
    connect(saveButton, &QPushButton::clicked, [&]() {
        // 保存人员名单
        male_names = maleEdit.toPlainText().split("\n", Qt::SkipEmptyParts);
        female_names = femaleEdit.toPlainText().split("\n", Qt::SkipEmptyParts);

        // 保存人员设置
        QStringList leadersList = leadersEdit->toPlainText().split("\n", Qt::SkipEmptyParts);
        QStringList boardersList = boardersEdit->toPlainText().split("\n", Qt::SkipEmptyParts);

        leaders = QSet<QString>(leadersList.begin(), leadersList.end());
        boarders = QSet<QString>(boardersList.begin(), boardersList.end());

        // 保存分组配置
        for (int i = 0; i < 10; i++) {
            groupConfigs[i].enabled = enableCombos[i]->currentData().toBool();
            groupConfigs[i].total = totalSpins[i]->value();
            groupConfigs[i].males = maleSpins[i]->value();
            groupConfigs[i].females = femaleSpins[i]->value();
        }

        // 计算实际启用的组数
        actualGroupCount = 0;
        for (int i = 0; i < 10; i++) {
            if (groupConfigs[i].enabled) actualGroupCount++;
        }

        // 保存样式设置
        QString selectedStyle = styleCombo->currentData().toString();
        currentStyle = selectedStyle;

        QSettings settings(configPath, QSettings::IniFormat);

        // 保存所有设置到配置文件
        settings.setValue("Male/Names", male_names);
        settings.setValue("Female/Names", female_names);
        settings.setValue("Leaders/Names", QStringList(leaders.values()));
        settings.setValue("Boarders/Names", QStringList(boarders.values()));
        settings.setValue("Style/Current", selectedStyle);

        if (selectedStyle == "自定义") {
            settings.setValue("Style/CustomPath", customStyleEdit->text());
        }

        // 保存分组配置和固定位置
        saveSettings();

        // 更新姓名映射
        updateNameMapping();

        dialog.accept();

        // 重置分组数据
        resetAll();

        // 仅在分组已生成时更新显示
        if (!groups.isEmpty()) {
            printGroups();
        }
        });

    // 取消按钮逻辑
    connect(cancelButton, &QPushButton::clicked, &dialog, &QDialog::reject);

    dialog.exec();
}

void MainWindow::showAboutDialog()
{
    QDialog aboutDialog(this);
    aboutDialog.setWindowTitle("关于");
    aboutDialog.resize(500, 400);

    QVBoxLayout* layout = new QVBoxLayout(&aboutDialog);

    // 创建水平布局来放置图片和文本
    QHBoxLayout* contentLayout = new QHBoxLayout();

    // 创建并旋转图片
    QLabel* imageLabel = new QLabel();
    QPixmap pixmap(":/images/about_image.png");
    if (!pixmap.isNull()) {
        QPixmap rotatedPixmap = pixmap.transformed(QTransform().rotate(270));
        rotatedPixmap = rotatedPixmap.scaled(120, 300, Qt::KeepAspectRatio);
        imageLabel->setPixmap(rotatedPixmap);
    }
    contentLayout->addWidget(imageLabel);

    //这里！当前版本:XXX 3377004_6
    //这里！当前版本:XXX 3377004_6
    //这里！当前版本:XXX 3377004_6
    //这里！当前版本:XXX 3377004_6
    //这里！当前版本:XXX 3377004_6
    // 文本内容
    QTextBrowser* textBrowser = new QTextBrowser();
    textBrowser->setOpenExternalLinks(true); // 启用外部链接
    textBrowser->setHtml(
        "<h2>智能分组系统</h2>"
        "<p><b>版本:</b> 251008.1730</p>"
        "<p><b>作者:</b> 3377004_6</p>"
        "<p><b>版权所有 © 2025-2026</b></p>"
        "<p>3377004_6 保留所有权利</p>"
        "<hr>"
        "<p><b>相关链接:</b></p>"
        "<p>• <a href='https://visualstudio.microsoft.com/zh-hans/vs'>Microsoft Visual Studio 2022</a></p>"
        "<p>• <a href='https://www.qt.io/zh-cn'>Qt</a></p>"
        "<p>• <a href='https://github.com/33770046/'>作者主页</a></p>"
        "<p>• <a href='https://www.bilibili.com'>哔哩哔哩</a></p>"
        "<p>• <a href='https://www.douyin.com/'>抖音</a></p>"
        "<p>• <a href='https://www.doubao.com/chat/'>豆包</a></p>"
        "<p>• <a href='https://chat.deepseek.com'>DeepSeek</a></p>"
        "<p>• <a href='https://github.com/'>Github</a></p>"
        "<p>• <a href='https://warthunder.com/zh'>BVVDの征兵网</a></p>"
    );
    contentLayout->addWidget(textBrowser);

    layout->addLayout(contentLayout);

    // 按钮区域
    QHBoxLayout* buttonLayout = new QHBoxLayout();

    QPushButton* updateBtn = new QPushButton("检查更新");
    QPushButton* closeBtn = new QPushButton("关闭");

    buttonLayout->addWidget(updateBtn);
    buttonLayout->addStretch();
    buttonLayout->addWidget(closeBtn);

    layout->addLayout(buttonLayout);

    // 连接信号
    connect(updateBtn, &QPushButton::clicked, [&]() {
        aboutDialog.accept();
        checkForUpdates(false);
        });

    connect(closeBtn, &QPushButton::clicked, &aboutDialog, &QDialog::accept);

    aboutDialog.exec();
}

void MainWindow::printGroups()
{
    // 确定启用的组
    QVector<int> enabledGroups;
    for (int i = 0; i < 10; i++) {
        if (groupConfigs[i].enabled) {
            enabledGroups.append(i);
        }
    }

    if (enabledGroups.isEmpty()) {
        groupTable->setRowCount(0);
        return;
    }

    // 计算最大列数
    int maxColumns = 11; // 组号列
    for (int group_idx : enabledGroups) {
        if (groupConfigs[group_idx].total + 1 > maxColumns) {
            maxColumns = groupConfigs[group_idx].total + 1;
        }
    }

    // 设置表格结构
    groupTable->setRowCount(enabledGroups.size());
    groupTable->setColumnCount(maxColumns);

    // 设置表头
    QStringList headers;
    headers << "组号";
    for (int i = 1; i < maxColumns; i++) {
        headers << QString("成员%1").arg(i);
    }
    groupTable->setHorizontalHeaderLabels(headers);

    // 填充表格
    for (int i = 0; i < enabledGroups.size(); i++) {
        int group_idx = enabledGroups[i];
        GroupConfig config = groupConfigs[group_idx];

        // 完整组号显示
        QTableWidgetItem* groupItem = new QTableWidgetItem(QString("组 %1").arg(group_idx + 1));
        groupTable->setItem(i, 0, groupItem);

        // 成员单元格
        for (int j = 0; j < groups[i].size() && j < maxColumns - 1; j++) {
            QString name = id_to_name[groups[i][j]];
            QTableWidgetItem* item = new QTableWidgetItem(name);

            // 设置性别背景色
            if (getGender(groups[i][j]) == 'M') {
                item->setBackground(QColor(200, 230, 255));
            }
            else {
                item->setBackground(QColor(255, 230, 255));
            }

            // 标记组长
            if (leaders.contains(name)) {
                QFont font = item->font();
                font.setBold(true);
                item->setFont(font);
            }

            // 标记外宿生
            if (boarders.contains(name)) {
                item->setBackground(Qt::yellow);
            }

            groupTable->setItem(i, j + 1, item);
        }

        // 填充空单元格
        for (int j = groups[i].size(); j < maxColumns - 1; j++) {
            groupTable->setItem(i, j + 1, new QTableWidgetItem(""));
        }
    }

    // 设置列宽自适应
    groupTable->resizeColumnsToContents();
    groupTable->horizontalHeader()->setSectionResizeMode(0, QHeaderView::Stretch);
    for (int i = 1; i < maxColumns; i++) {
        groupTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    }
}

void MainWindow::showGroupDetails(int row, int column)
{
    if (row >= GROUP_COUNT || row < 0) return;

    QString details = QString("<b>组 %1 详情:</b><br>").arg(row + 1);

    int maleCount = 0, femaleCount = 0;
    for (int id : groups[row]) {
        QString name = id_to_name[id];
        details += name + " ";
        if (getGender(id) == 'M') maleCount++;
        else femaleCount++;
    }

    // 确定启用的组
    QVector<int> enabledGroups;
    for (int i = 0; i < 10; i++) {
        if (groupConfigs[i].enabled) {
            enabledGroups.append(i);
        }
    }

    if (row >= enabledGroups.size() || row < 0) return;

    // 检查要求满足情况
    QMap<int, int> person_to_group = buildPersonToGroup();

    // 检查必须同组要求
    for (const auto& group : must_together_groups) {
        QSet<int> group_set;
        for (int person : group) {
            group_set.insert(person_to_group[person]);
        }
        if (group_set.size() > 1) {
            details += "<br><font color='red'>警告: 必须同组要求未满足</font>";
        }
    }

    // 检查不能同组要求
    for (const auto& group : must_separate_groups) {
        QMap<int, int> group_count;
        for (int person : group) {
            int g = person_to_group[person];
            group_count[g]++;
            if (group_count[g] > 1) {
                details += "<br><font color='red'>警告: 不能同组要求未满足</font>";
                break;
            }
        }
    }

    logOutput->append(details);
}

void MainWindow::updateFixedPositionTable(QTableWidget* table)
{
    if (!table) {
        // 如果没有传入表格，使用成员变量
        if (!fixedPositionTable) return;
        table = fixedPositionTable;
    }

    table->setRowCount(fixedPositions.size());
    int row = 0;

    for (auto it = fixedPositions.begin(); it != fixedPositions.end(); ++it) {
        int personId = it.key();
        FixedPosition pos = it.value();
        QString name = id_to_name[personId];

        table->setItem(row, 0, new QTableWidgetItem(name));
        table->setItem(row, 1, new QTableWidgetItem(QString::number(pos.groupNumber)));
        table->setItem(row, 2, new QTableWidgetItem(QString::number(pos.seatPosition)));
        row++;
    }

    // 设置表头
    table->setHorizontalHeaderLabels(QStringList{ "姓名", "组号", "座位位置" });
}