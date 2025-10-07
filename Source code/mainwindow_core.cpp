#include "mainwindow.h"
#include <algorithm>
#include <iterator>
#include <random>
#include <climits>

#ifdef Q_OS_WIN
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <comdef.h>
#include <objbase.h>
#endif

void MainWindow::initNameMapping()
{
    // 确定配置文件路径
#ifdef Q_OS_WIN
    // 在Windows上，使用应用程序目录
    configPath = QCoreApplication::applicationDirPath() + "/config.ini";
#else
    // 在其他系统上，使用用户配置目录
    configPath = QStandardPaths::writableLocation(QStandardPaths::AppConfigLocation) + "/config.ini";
    // 确保目录存在
    QDir configDir(QStandardPaths::writableLocation(QStandardPaths::AppConfigLocation));
    if (!configDir.exists()) {
        configDir.mkpath(".");
    }
#endif

    QSettings settings(configPath, QSettings::IniFormat);

    // 从配置文件读取名单
    male_names = settings.value("Male/Names").toStringList();
    female_names = settings.value("Female/Names").toStringList();

    // 从配置文件读取组长和外宿生设置
    QStringList leadersList = settings.value("Leaders/Names").toStringList();
    QStringList boardersList = settings.value("Boarders/Names").toStringList();

    leaders = QSet<QString>(leadersList.begin(), leadersList.end());
    boarders = QSet<QString>(boardersList.begin(), boardersList.end());

    // 从配置文件读取样式设置
    currentStyle = settings.value("Style/Current", "样式1").toString();

    // 初始化映射
    name_to_id.clear();
    id_to_name.clear();

    // 男生映射
    for (int i = 0; i < male_names.size(); i++) {
        name_to_id[male_names[i]] = i + 1;
        id_to_name[i + 1] = male_names[i];
    }

    // 女生映射
    for (int i = 0; i < female_names.size(); i++) {
        name_to_id[female_names[i]] = i + male_names.size() + 1;
        id_to_name[i + male_names.size() + 1] = female_names[i];
    }

    // 记录日志 - 添加空指针检查
    if (logOutput) {
        logOutput->append(QString("加载了 %1 名男生和 %2 名女生").arg(male_names.size()).arg(female_names.size()));
        logOutput->append(QString("加载了 %1 名组长和 %2 名外宿生").arg(leaders.size()).arg(boarders.size()));
        logOutput->append(QString("当前样式: %1").arg(currentStyle));
    }
}

QChar MainWindow::getGender(int person)
{
    return (person >= 1 && person <= male_names.size()) ? 'M' : 'F';
}

void MainWindow::initGroupConfigs()
{
    // 支持最多10个组
    groupConfigs.resize(10);

    // 默认配置 - 所有组使用相同的配置，但不强制人数
    for (int i = 0; i < 10; i++) {
        groupConfigs[i].enabled = (i < 9); // 默认启用前9组
        groupConfigs[i].total = (i < 9) ? 6 : 0; // 默认6人，但不再强制
        groupConfigs[i].males = -1; // -1表示不强制男生数量
        groupConfigs[i].females = -1; // -1表示不强制女生数量
    }
}

void MainWindow::updateNameMapping()
{
    name_to_id.clear();
    id_to_name.clear();

    // 男生映射
    for (int i = 0; i < male_names.size(); i++) {
        name_to_id[male_names[i]] = i + 1;
        id_to_name[i + 1] = male_names[i];
    }

    // 女生映射
    for (int i = 0; i < female_names.size(); i++) {
        name_to_id[female_names[i]] = i + male_names.size() + 1;
        id_to_name[i + male_names.size() + 1] = female_names[i];
    }

    // 记录日志
    logOutput->append(QString("更新了姓名映射: %1 名男生, %2 名女生").arg(male_names.size()).arg(female_names.size()));
}

void MainWindow::loadSettings()
{
    QSettings settings(configPath, QSettings::IniFormat);

    // 加载分组配置
    for (int i = 0; i < 10; i++) {
        settings.beginGroup(QString("GroupConfig_%1").arg(i));
        groupConfigs[i].enabled = settings.value("enabled", i < 9).toBool();
        groupConfigs[i].total = settings.value("total", i < 9 ? 6 : 0).toInt();
        groupConfigs[i].males = settings.value("males", 2).toInt();
        groupConfigs[i].females = settings.value("females", 4).toInt();
        settings.endGroup();
    }

    // 计算实际启用的组数
    actualGroupCount = 0;
    for (int i = 0; i < 10; i++) {
        if (groupConfigs[i].enabled) actualGroupCount++;
    }

    // 加载固定位置 - 修改为新的格式
    fixedPositions.clear();
    QStringList fixedPositionsList = settings.value("FixedPositions/List").toStringList();
    for (const QString& entry : fixedPositionsList) {
        QStringList parts = entry.split(',');
        if (parts.size() == 3) {
            int personId = parts[0].toInt();
            int group = parts[1].toInt();
            int seat = parts[2].toInt();
            fixedPositions[personId] = FixedPosition{ group, seat };
        }
    }

    // 加载要求条件
    must_together_groups.clear();
    QStringList mustTogetherList = settings.value("Constraints/MustTogether").toStringList();
    for (const QString& entry : mustTogetherList) {
        QStringList idsStr = entry.split(',');
        QVector<int> group;
        for (const QString& idStr : idsStr) {
            group.append(idStr.toInt());
        }
        if (!group.isEmpty()) {
            must_together_groups.append(group);
        }
    }

    must_separate_groups.clear();
    QStringList mustSeparateList = settings.value("Constraints/MustSeparate").toStringList();
    for (const QString& entry : mustSeparateList) {
        QStringList idsStr = entry.split(',');
        QVector<int> group;
        for (const QString& idStr : idsStr) {
            group.append(idStr.toInt());
        }
        if (!group.isEmpty()) {
            must_separate_groups.append(group);
        }
    }
    updateConstraintTable();
}

void MainWindow::updateConstraintTable()
{
    constraintTable->setRowCount(0);

    // 添加必须同组要求
    for (const auto& group : must_together_groups) {
        int row = constraintTable->rowCount();
        constraintTable->insertRow(row);

        QStringList names;
        for (int id : group) {
            names.append(id_to_name[id]);
        }

        constraintTable->setItem(row, 0, new QTableWidgetItem("必须同组"));
        constraintTable->setItem(row, 1, new QTableWidgetItem(names.join(", ")));
        constraintTable->setItem(row, 2, new QTableWidgetItem("待验证"));
    }

    // 添加不能同组要求
    for (const auto& group : must_separate_groups) {
        int row = constraintTable->rowCount();
        constraintTable->insertRow(row);

        QStringList names;
        for (int id : group) {
            names.append(id_to_name[id]);
        }

        constraintTable->setItem(row, 0, new QTableWidgetItem("不能同组"));
        constraintTable->setItem(row, 1, new QTableWidgetItem(names.join(", ")));
        constraintTable->setItem(row, 2, new QTableWidgetItem("待验证"));
    }
}

void MainWindow::saveSettings()
{
    QSettings settings(configPath, QSettings::IniFormat);

    // 保存分组配置
    for (int i = 0; i < 10; i++) {
        settings.beginGroup(QString("GroupConfig_%1").arg(i));
        settings.setValue("enabled", groupConfigs[i].enabled);
        settings.setValue("total", groupConfigs[i].total);
        settings.setValue("males", groupConfigs[i].males);
        settings.setValue("females", groupConfigs[i].females);
        settings.endGroup();
    }

    // 保存固定位置
    QStringList fixedPositionsList;
    for (auto it = fixedPositions.begin(); it != fixedPositions.end(); ++it) {
        FixedPosition pos = it.value();
        QString entry = QString("%1,%2,%3").arg(it.key()).arg(pos.groupNumber).arg(pos.seatPosition);
        fixedPositionsList.append(entry);
    }
    settings.setValue("FixedPositions/List", fixedPositionsList);

    // 保存要求条件
    QStringList mustTogetherList;
    for (const auto& group : must_together_groups) {
        QStringList ids;
        for (int id : group) {
            ids.append(QString::number(id));
        }
        mustTogetherList.append(ids.join(","));
    }
    settings.setValue("Constraints/MustTogether", mustTogetherList);

    QStringList mustSeparateList;
    for (const auto& group : must_separate_groups) {
        QStringList ids;
        for (int id : group) {
            ids.append(QString::number(id));
        }
        mustSeparateList.append(ids.join(","));
    }
    settings.setValue("Constraints/MustSeparate", mustSeparateList);

    settings.sync();
}

void MainWindow::addConstraint()
{
    // 添加输入验证
    QString names = nameInput->text().trimmed();
    if (names.contains(",") || names.contains(";")) {
        QMessageBox::warning(this, "输入错误", "姓名不能包含逗号或分号");
        return;
    }
    QString typeStr = constraintTypeCombo->currentText().contains("必须") ? "T" : "S";

    if (names.isEmpty()) {
        QMessageBox::warning(this, "输入错误", "请输入至少一个姓名");
        return;
    }

    QStringList nameList = names.split(" ", Qt::SkipEmptyParts);
    QVector<int> ids;
    QStringList invalidNames;

    for (const QString& name : nameList) {
        if (name_to_id.contains(name)) {
            ids.append(name_to_id[name]);
        }
        else {
            invalidNames.append(name);
        }
    }

    if (!invalidNames.isEmpty()) {
        QString errorMsg = "以下姓名无效: " + invalidNames.join(", ");
        QMessageBox::warning(this, "输入错误", errorMsg);
        return;
    }

    if (ids.size() < 2) {
        QMessageBox::warning(this, "输入错误", "要求需要至少两个有效姓名");
        return;
    }

    // 添加到要求列表
    if (typeStr == "T") {
        must_together_groups.append(ids);
    }
    else {
        must_separate_groups.append(ids);
    }

    // 更新要求表格
    int row = constraintTable->rowCount();
    constraintTable->insertRow(row);

    constraintTable->setItem(row, 0, new QTableWidgetItem(typeStr == "T" ? "必须同组" : "不能同组"));
    QStringList displayNames;
    for (int id : ids) {
        displayNames.append(id_to_name[id]);
    }
    constraintTable->setItem(row, 1, new QTableWidgetItem(displayNames.join(", ")));
    constraintTable->setItem(row, 2, new QTableWidgetItem("待验证"));

    nameInput->clear();
    updateStatus("要求添加成功");
}

void MainWindow::removeConstraint()
{
    QModelIndexList selected = constraintTable->selectionModel()->selectedRows();
    if (selected.isEmpty()) {
        QMessageBox::information(this, "提示", "请先选择要移除的要求");
        return;
    }

    std::sort(selected.begin(), selected.end(), [](const QModelIndex& a, const QModelIndex& b) {
        return a.row() > b.row();
        });

    for (const QModelIndex& index : selected) {
        int row = index.row();

        // 从对应的要求列表中移除
        QString type = constraintTable->item(row, 0)->text();
        if (type == "必须同组") {
            if (row < must_together_groups.size()) {
                must_together_groups.remove(row);
            }
        }
        else {
            if (row < must_separate_groups.size()) {
                must_separate_groups.remove(row);
            }
        }
        constraintTable->removeRow(row);
    }
    updateStatus("要求已移除");
}

void MainWindow::generateGroups()
{
    try {
        // 显示进度对话框
        QProgressDialog progress("正在生成分组...", "取消", 0, 0, this);
        progress.setWindowModality(Qt::WindowModal);
        progress.setCancelButton(nullptr); // 禁用取消按钮
        progress.show();

        // 处理事件，确保UI更新
        QCoreApplication::processEvents();

        // 调用分组算法
        auto result = initializeGroupsWithConstraints();
        groups = result.first;

        // 更新分组表格
        printGroups();

        // 关闭进度对话框
        progress.close();

        updateStatus("分组生成完成");
    }
    catch (const std::exception& e) {
        QMessageBox::critical(this, "错误", QString("生成分组时发生错误: %1").arg(e.what()));
    }
    catch (...) {
        QMessageBox::critical(this, "错误", "生成分组时发生未知错误");
    }
}

void MainWindow::checkPersonnel()
{
    if (groups.isEmpty()) {
        QMessageBox::warning(this, "错误", "请先生成分组结果");
        return;
    }

    // 构建检查报告
    QString report = "<h2>人员检查报告</h2>";
    report += "<table border='1' cellpadding='4'>";
    report += "<tr><th>组号</th><th>人数</th><th>男生</th><th>女生</th><th>组长</th><th>外宿生</th></tr>";

    int totalPersons = 0;
    int totalMales = 0;
    int totalFemales = 0;
    int totalLeaders = 0;
    int totalBoarders = 0;

    for (int i = 0; i < groups.size(); i++) {
        int maleCount = 0;
        int femaleCount = 0;
        int leaderCount = 0;
        int boarderCount = 0;

        for (int person : groups[i]) {
            QString name = id_to_name[person];
            if (getGender(person) == 'M') maleCount++;
            else femaleCount++;

            if (leaders.contains(name)) leaderCount++;
            if (boarders.contains(name)) boarderCount++;
        }

        report += QString("<tr><td>%1</td><td>%2</td><td>%3</td><td>%4</td><td>%5</td><td>%6</td></tr>")
            .arg(i + 1)
            .arg(maleCount + femaleCount)
            .arg(maleCount)
            .arg(femaleCount)
            .arg(leaderCount)
            .arg(boarderCount);

        totalPersons += maleCount + femaleCount;
        totalMales += maleCount;
        totalFemales += femaleCount;
        totalLeaders += leaderCount;
        totalBoarders += boarderCount;
    }

    report += QString("<tr style='background-color:#E6E6E6; font-weight:bold;'><td>总计</td><td>%1</td><td>%2</td><td>%3</td><td>%4</td><td>%5</td></tr>")
        .arg(totalPersons)
        .arg(totalMales)
        .arg(totalFemales)
        .arg(totalLeaders)
        .arg(totalBoarders);

    report += "</table>";

    // 添加统计信息，但不进行固定人数检查
    QString info;
    info += QString("<p>总人数: %1 (男生: %2, 女生: %3)</p>").arg(totalPersons).arg(totalMales).arg(totalFemales);
    info += QString("<p>组长: %1, 外宿生: %2</p>").arg(totalLeaders).arg(totalBoarders);

    // 显示报告对话框
    QDialog reportDialog(this);
    reportDialog.setWindowTitle("人员检查报告");
    reportDialog.resize(350, 400);

    QVBoxLayout* layout = new QVBoxLayout(&reportDialog);

    QTextEdit* reportView = new QTextEdit();
    reportView->setReadOnly(true);
    reportView->setHtml(report + info);
    layout->addWidget(reportView);

    QPushButton* closeBtn = new QPushButton("关闭");
    connect(closeBtn, &QPushButton::clicked, &reportDialog, &QDialog::accept);
    layout->addWidget(closeBtn, 0, Qt::AlignRight);

    reportDialog.exec();
}

void MainWindow::checkConstraints()
{
    showConstraintViolations();
    updateStatus("要求检查完成");
}

void MainWindow::resetAll()
{
    groups.clear();
    must_together_groups.clear();
    must_separate_groups.clear();
    constraintTable->setRowCount(0);
    groupTable->setRowCount(0);
    logOutput->clear();
    updateStatus("系统已重置");
}

void MainWindow::updateStatus(const QString& message)
{
    if (!statusLabel || !logOutput) {
        return;
    }

    QString timestamp = QDateTime::currentDateTime().toString("hh:mm:ss");
    statusLabel->setText(timestamp + " - " + message);
    logOutput->append(timestamp + " - " + message);
}

void MainWindow::showConstraintViolations()
{
    logOutput->append("开始检查要求...");

    // 构建人-组映射
    QMap<int, int> person_to_group = buildPersonToGroup();

    bool constraints_satisfied = true;

    // 检查必须同组要求
    for (const auto& group : must_together_groups) {
        QSet<int> group_set;
        for (int person : group) {
            group_set.insert(person_to_group[person]);
        }
        if (group_set.size() != 1) {
            QString msg = "必须同组要求未满足: ";
            for (int person : group) {
                msg += id_to_name[person] + " ";
            }
            logOutput->append(msg);
            constraints_satisfied = false;
        }
    }

    // 检查不同要求组是否在不同组
    for (int i = 0; i < must_together_groups.size(); i++) {
        int actual_group = person_to_group[must_together_groups[i][0]];
        for (int j = i + 1; j < must_together_groups.size(); j++) {
            if (person_to_group[must_together_groups[j][0]] == actual_group) {
                QString msg = QString("不同要求组在同一组: 要求组%1 和 要求组%2 都在组 %3")
                    .arg(i + 1).arg(j + 1).arg(actual_group + 1);
                logOutput->append(msg);
                constraints_satisfied = false;
            }
        }
    }

    // 检查不能同组要求
    for (const auto& group : must_separate_groups) {
        QMap<int, int> group_count;
        for (int person : group) {
            int g = person_to_group[person];
            group_count[g]++;
            if (group_count[g] > 1) {
                QString msg = "不能同组要求未满足: ";
                for (int person : group) {
                    msg += id_to_name[person] + " ";
                }
                logOutput->append(msg);
                constraints_satisfied = false;
                break;
            }
        }
    }

    if (constraints_satisfied) {
        logOutput->append("所有要求已满足");
    }
    else {
        logOutput->append("部分要求未满足");
    }
}

QMap<int, int> MainWindow::buildPersonToGroup()
{
    QMap<int, int> person_to_group;
    for (int i = 0; i < groups.size(); i++) {
        for (int person : groups[i]) {
            person_to_group[person] = i;
        }
    }
    return person_to_group;
}

void MainWindow::swapPersons(int a, int b)
{
    // 这个函数目前没有实现具体逻辑
    // 保留空实现
}

// 现有的函数
int MainWindow::findSameGenderInGroup(int person, int target_group_idx, const QSet<int>& fixed_people)
{
    QChar gender = getGender(person);
    QVector<int> candidates;

    // 检查目标组索引是否有效
    if (target_group_idx < 0 || target_group_idx >= groups.size()) {
        return -1;
    }

    for (int p : groups[target_group_idx]) {
        // 跳过空位和固定位置人员
        if (p == 0 || fixed_people.contains(p)) continue;

        if (getGender(p) == gender) {
            candidates.append(p);
        }
    }

    if (candidates.isEmpty()) return -1;

    // 使用随机数生成器选择候选者
    std::random_device rd;
    std::mt19937 g(rd());
    std::uniform_int_distribution<int> dist(0, candidates.size() - 1);
    return candidates[dist(g)];
}

// 添加缺少的函数（两个参数版本）
int MainWindow::findSameGenderInGroup(int person, int target_group_idx)
{
    // 调用三个参数的版本，传入空的 fixed_people 集合
    return findSameGenderInGroup(person, target_group_idx, QSet<int>());
}

// 添加固定位置
void MainWindow::addFixedPosition() {
    // 这个函数已整合到设置对话框中，保留空实现
    QMessageBox::information(this, "提示", "固定位置功能已整合到设置对话框中");
}

// 清空固定位置
void MainWindow::clearFixedPositions() {
    if (QMessageBox::question(this, "确认", "确定要清空所有固定位置吗?") == QMessageBox::Yes) {
        fixedPositions.clear();
        updateFixedPositionTable();
        updateStatus("已清空所有固定位置");
    }
}

// 整数溢出警告
void MainWindow::someFunction() {
    int a = 1000000;
    int b = 1000000;
    long long result = static_cast<long long>(a) * b;
}

MainWindow::~MainWindow()
{
    saveSettings();

#ifdef Q_OS_WIN
    try {
        ::CoUninitialize();
    }
    catch (...) {
    }
#endif
}

GroupConfig MainWindow::getGroupConfigForIndex(int index)
{
    int enabledCount = -1;
    for (int i = 0; i < 10; i++) {
        if (groupConfigs[i].enabled) {
            enabledCount++;
            if (enabledCount == index) {
                return groupConfigs[i];
            }
        }
    }
    return GroupConfig{ false, 0, 0, 0 };
}

void MainWindow::newFile()
{
    // 清空分组数据但保留要求条件
    groups.clear();

    // 初始化空分组
    groups.resize(GROUP_COUNT);

    // 清空分组表格
    groupTable->setRowCount(0);

    // 初始化空分组表格
    groupTable->setRowCount(GROUP_COUNT);
    for (int i = 0; i < GROUP_COUNT; i++) {
        groupTable->setItem(i, 0, new QTableWidgetItem(QString("组 %1").arg(i + 1)));
        for (int j = 1; j <= MAX_GROUP_SIZE; j++) {
            groupTable->setItem(i, j, new QTableWidgetItem(""));
        }
    }

    // 询问用户保存位置
    QString defaultFileName = QDateTime::currentDateTime().toString("yyyyMMdd_hhmmss") + "_空座位表.xlsx";
    QString fileName = QFileDialog::getSaveFileName(this, "保存空座位表", defaultFileName, "Excel文件 (*.xlsx)");

    if (fileName.isEmpty()) {
        updateStatus("已创建新的空白座位表（未保存）");
        return;
    }

    // 确保文件扩展名
    if (!fileName.endsWith(".xlsx", Qt::CaseInsensitive)) {
        fileName += ".xlsx";
    }

    // 导出空座位表
    doExportSeatingPlan(fileName);
    updateStatus("已创建新的空白座位表并保存到: " + fileName);
}

void MainWindow::saveFile()
{
    if (groups.isEmpty()) {
        QMessageBox::warning(this, "错误", "没有分组数据可保存");
        return;
    }

    // 在exe同目录下生成文件
    QString fileName = QCoreApplication::applicationDirPath() + "/" +
        QDateTime::currentDateTime().toString("yyyyMMdd_hhmmss") +
        "_座位表.xlsx";

    // 保存文件
    doExportSeatingPlan(fileName);
    updateStatus("座位表已保存到: " + fileName);
}

void MainWindow::saveAs()
{
    if (groups.isEmpty()) {
        QMessageBox::warning(this, "错误", "没有分组数据可保存");
        return;
    }

    // 询问用户保存位置
    QString defaultFileName = QDateTime::currentDateTime().toString("yyyyMMdd_hhmmss") + "_座位表.xlsx";
    QString fileName = QFileDialog::getSaveFileName(this, "另存为", defaultFileName, "Excel文件 (*.xlsx)");

    if (fileName.isEmpty()) return;

    // 确保文件扩展名
    if (!fileName.endsWith(".xlsx", Qt::CaseInsensitive)) {
        fileName += ".xlsx";
    }

    // 保存文件
    doExportSeatingPlan(fileName);
    updateStatus("座位表已另存为: " + fileName);
}

void MainWindow::exportSeatingPlan()
{
    if (groups.isEmpty()) {
        QMessageBox::warning(this, "错误", "请先生成分组结果");
        return;
    }

    QString fileName = QFileDialog::getSaveFileName(this, "导出座位表", "", "Excel文件 (*.xlsx)");
    if (fileName.isEmpty()) return;

    doExportSeatingPlan(fileName);
    updateStatus("座位表已导出到: " + fileName);
}