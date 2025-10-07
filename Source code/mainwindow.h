#pragma once

#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QTableWidget>
#include <QGroupBox>
#include <QPushButton>
#include <QTextEdit>
#include <QComboBox>
#include <QLineEdit>
#include <QLabel>
#include <QStatusBar>
#include <QVector>
#include <QMap>
#include <QSet>
#include <QList>
#include <QString>
#include <random>
#include <ctime>
#include <unordered_map>
#include <set>
#include <QProgressDialog>
#include <QFileDialog>
#include <QMessageBox>
#include <QDateTime>
#include <QCoreApplication>
#include <QStringConverter>
#include <QDesktopServices>
#include <QAxObject>
#include <QMenuBar>
#include <QSettings>
#include <QDialog>
#include <QMenu>
#include <QAction>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QStandardPaths>
#include <QDir>
#include <QRegularExpression>
#include <QSpinBox>
#include <QTabWidget>
#include <QTextEdit>
#include <QGroupBox>
#include <queue>
#include <utility>

// 分组配置结构体
struct GroupConfig {
    bool enabled;
    int total;
    int males;
    int females;
};

// 座位信息结构体
struct SeatInfo {
    QString address;
    int row;
    int col;
    int regionIdx;
    int seatIdx;
};

// 固定位置结构体
struct FixedPosition {
    int groupNumber;
    int seatPosition;

    FixedPosition() : groupNumber(1), seatPosition(1) {}
    FixedPosition(int group, int seat) : groupNumber(group), seatPosition(seat) {}
};

// 位置结构体
struct Position {
    int group;
    int seat;
    Position(int g = -1, int s = -1) : group(g), seat(s) {}
    bool operator==(const Position& other) const {
        return group == other.group && seat == other.seat;
    }
};

// 交换状态结构体
struct SwapState {
    int group;
    int seat;
    QVector<QPair<Position, Position>> path;

    SwapState(int g, int s, const QVector<QPair<Position, Position>>& p = {})
        : group(g), seat(s), path(p) {
    }
};

class MainWindow : public QMainWindow {
    Q_OBJECT

public:
    MainWindow(QWidget* parent = nullptr);
    ~MainWindow();
    QStringList getMaleNames() const { return male_names; }
    QStringList getFemaleNames() const { return female_names; }

private slots:
    void addConstraint();
    void removeConstraint();
    void generateGroups();
    void exportSeatingPlan();
    void checkPersonnel();
    void checkConstraints();
    void resetAll();
    void showGroupDetails(int row, int column);
    void showSettingsDialog();
    void showAboutDialog();
    void newFile();
    void saveFile();
    void saveAs();
    void addFixedPosition();
    void clearFixedPositions();
    void updateFixedPositionTable(QTableWidget* table = nullptr);

private:
    // 初始化函数
    void initNameMapping();
    void initGroupConfigs();
    void createMenuBar();
    void updateNameMapping();
    void loadSettings();
    void updateConstraintTable();
    void saveSettings();
    void updateStatus(const QString& message);
    QMenu* groupTableContextMenu;
    QAction* moveToSeatAction;

    // 分组算法相关函数
    QPair<QVector<QVector<int>>, QSet<int>> initializeGroupsWithConstraints();
    QMap<int, int> buildPersonToGroup();
    void swapPersons(int a, int b);
    int findSameGenderInGroup(int person, int target_group_idx);
    int selectedPersonId;
    int selectedGroup;
    int selectedSeat;
    void showConstraintViolations();
    GroupConfig getGroupConfigForIndex(int index);

    // 导出相关函数
    void doExportSeatingPlan(const QString& fileName);

    // 辅助函数
    QChar getGender(int person);
    void printGroups();
    void someFunction();

    // 位置交换相关函数
    bool swapPersonsInGroup(int groupIndex, int seatA, int seatB);
    bool swapPersonsBetweenGroups(int groupA, int seatA, int groupB, int seatB);
    bool movePersonToPosition(int personId, int targetGroup, int targetSeat);
    QVector<QPair<Position, Position>> findSwapPath(int startGroup, int startSeat, int targetGroup, int targetSeat);
    void optimizeFixedPositions();

    // UI组件
    QTableWidget* groupTable;
    QTableWidget* constraintTable;
    QTableWidget* fixedPositionTable;
    QTextEdit* logOutput;
    QComboBox* constraintTypeCombo;
    QLineEdit* nameInput;
    QLineEdit* fixedPositionInput;
    QPushButton* addConstraintBtn;
    QPushButton* removeConstraintBtn;
    QPushButton* generateBtn;
    QPushButton* personnelCheckBtn;
    QPushButton* checkBtn;
    QPushButton* resetBtn;
    QLabel* statusLabel;
    QMenu* fileMenu;

    // 数据成员
    QStringList male_names;
    QStringList female_names;
    QMap<QString, int> name_to_id;
    QMap<int, QString> id_to_name;
    QVector<QVector<int>> groups;
    QSet<int> special_groups;
    QVector<QVector<int>> must_together_groups;
    QVector<QVector<int>> must_separate_groups;
    QSet<QString> leaders;
    QSet<QString> boarders;
    QVector<GroupConfig> groupConfigs;
    int actualGroupCount;
    int findSameGenderInGroup(int person, int target_group_idx, const QSet<int>& fixed_people);

    // 固定位置相关
    QMap<int, FixedPosition> fixedPositions;

    // 样式和模板相关
    QString currentTemplatePath;
    QString currentStyle;

    // 布局相关
    QVBoxLayout* mainLayout;

    // 随机数生成器
    std::mt19937 rng;

    // 配置文件路径
    QString configPath;

    // 常量
    static const int GROUP_COUNT = 9;
    static const int SPECIAL_GROUP_COUNT = 2;
    static const int MAX_GROUP_SIZE = 6;
};

#endif // MAINWINDOW_H