#include "mainwindow.h"
#include <algorithm>
#include <iterator>
#include <random>
#include <climits>
#include <queue>
#include <utility>

#ifdef Q_OS_WIN
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <comdef.h>
#include <objbase.h>
#endif

QPair<QVector<QVector<int>>, QSet<int>> MainWindow::initializeGroupsWithConstraints()
{
    // 确定启用的组
    QVector<int> enabledGroups;
    for (int i = 0; i < 10; i++) {
        if (groupConfigs[i].enabled) {
            enabledGroups.append(i);
        }
    }

    if (enabledGroups.isEmpty()) {
        QMessageBox::warning(this, "错误", "没有启用的分组");
        return { QVector<QVector<int>>(), QSet<int>() };
    }

    // 初始化分组数据结构
    QVector<QVector<int>> local_groups(enabledGroups.size());
    for (int i = 0; i < local_groups.size(); i++) {
        int group_idx = enabledGroups[i];
        local_groups[i] = QVector<int>(groupConfigs[group_idx].total, 0);
    }

    QSet<int> local_special_groups;

    // 创建人员集合
    QSet<int> all_people;
    for (int i = 1; i <= male_names.size() + female_names.size(); i++) {
        all_people.insert(i);
    }

    // 创建随机数生成器
    std::random_device rd;
    std::mt19937 g(rd());

    // 第一步：最高优先级
    QSet<int> fixed_people;

    // 按组号排序处理固定位置，确保前面的组先处理
    QList<int> fixedPersonIds = fixedPositions.keys();
    std::sort(fixedPersonIds.begin(), fixedPersonIds.end(), [&](int a, int b) {
        return fixedPositions[a].groupNumber < fixedPositions[b].groupNumber;
        });

    for (int personId : fixedPersonIds) {
        FixedPosition fixedPos = fixedPositions[personId];
        int groupNumber = fixedPos.groupNumber;
        int seatPosition = fixedPos.seatPosition;
        int groupIdx = groupNumber - 1;

        logOutput->append(QString("处理固定位置: %1 到组%2座位%3")
            .arg(id_to_name[personId]).arg(groupNumber).arg(seatPosition));

        // 检查组是否启用
        if (groupIdx >= 0 && groupIdx < 10 && groupConfigs[groupIdx].enabled) {
            int localIdx = enabledGroups.indexOf(groupIdx);
            if (localIdx != -1) {
                // 检查座位位置是否有效
                if (seatPosition >= 1 && seatPosition <= local_groups[localIdx].size()) {
                    int seatIndex = seatPosition - 1;

                    // 检查目标座位是否已被占用
                    if (local_groups[localIdx][seatIndex] != 0) {
                        int occupiedPerson = local_groups[localIdx][seatIndex];

                        if (fixed_people.contains(occupiedPerson)) {
                            logOutput->append(QString("错误: 座位%1已被固定位置人员 %2 占用，无法放置 %3")
                                .arg(seatPosition).arg(id_to_name[occupiedPerson]).arg(id_to_name[personId]));
                            continue;
                        }

                        // 移动占用者到其他空位
                        bool moved = false;
                        QString movedLog;

                        // 首先尝试同组其他空位
                        for (int j = 0; j < local_groups[localIdx].size() && !moved; j++) {
                            if (local_groups[localIdx][j] == 0 && j != seatIndex) {
                                local_groups[localIdx][j] = occupiedPerson;
                                moved = true;
                                movedLog = QString("同组移动到座位%1").arg(j + 1);
                            }
                        }

                        // 如果同组没有空位，尝试其他组
                        if (!moved) {
                            for (int otherGroup = 0; otherGroup < local_groups.size() && !moved; otherGroup++) {
                                if (otherGroup == localIdx) continue;

                                for (int j = 0; j < local_groups[otherGroup].size() && !moved; j++) {
                                    if (local_groups[otherGroup][j] == 0) {
                                        local_groups[otherGroup][j] = occupiedPerson;
                                        moved = true;
                                        movedLog = QString("移动到组%1座位%2").arg(otherGroup + 1).arg(j + 1);
                                    }
                                }
                            }
                        }

                        if (moved) {
                            logOutput->append(QString("移动 %1 %2").arg(id_to_name[occupiedPerson]).arg(movedLog));
                        }
                        else {
                            logOutput->append(QString("错误: 无法为 %1 找到空座位，固定位置分配失败")
                                .arg(id_to_name[occupiedPerson]));
                            continue;
                        }
                    }

                    // 分配固定位置
                    local_groups[localIdx][seatIndex] = personId;
                    fixed_people.insert(personId);
                    all_people.remove(personId);

                    logOutput->append(QString("成功分配固定位置: %1 到组%2座位%3")
                        .arg(id_to_name[personId]).arg(groupNumber).arg(seatPosition));
                }
                else {
                    logOutput->append(QString("错误: 座位位置%1超出范围(1-%2)")
                        .arg(seatPosition).arg(local_groups[localIdx].size()));
                }
            }
            else {
                logOutput->append(QString("错误: 组%1未启用").arg(groupNumber));
            }
        }
        else {
            logOutput->append(QString("错误: 组%1未启用或索引无效").arg(groupNumber));
        }
    }

    // 第二步：处理必须同组的要求
    QSet<int> constrained_people;

    // 为要求组预分配组
    QVector<int> available_groups = enabledGroups;
    std::shuffle(available_groups.begin(), available_groups.end(), g);

    QSet<int> constrained_groups;

    for (int i = 0; i < must_together_groups.size(); i++) {
        if (available_groups.isEmpty()) {
            logOutput->append("错误：没有足够的组来分配所有要求组");
            break;
        }

        int group_index = available_groups.takeFirst();
        constrained_groups.insert(group_index);
        int local_idx = enabledGroups.indexOf(group_index);

        // 为要求组分配位置，从第一个座位开始顺序分配
        int currentSeat = 0;
        for (int person : must_together_groups[i]) {
            // 如果人员已经被固定位置占用，跳过
            if (fixed_people.contains(person)) {
                logOutput->append(QString("人员 %1 已被固定位置占用，跳过要求分配").arg(id_to_name[person]));
                continue;
            }

            // 找到下一个空位
            while (currentSeat < local_groups[local_idx].size() &&
                local_groups[local_idx][currentSeat] != 0) {
                currentSeat++;
            }

            if (currentSeat < local_groups[local_idx].size()) {
                local_groups[local_idx][currentSeat] = person;
                constrained_people.insert(person);
                all_people.remove(person);
                currentSeat++;
            }
            else {
                logOutput->append(QString("错误: 组%1没有足够的空位分配给要求组").arg(group_index + 1));
                break;
            }
        }
    }

    // 分离男女生
    QVector<int> free_males;
    QVector<int> free_females;

    for (int person : all_people) {
        if (person <= male_names.size()) {
            free_males.append(person);
        }
        else {
            free_females.append(person);
        }
    }

    // 随机打乱
    std::shuffle(free_males.begin(), free_males.end(), g);
    std::shuffle(free_females.begin(), free_females.end(), g);

    // 第一步：分配女生
    for (int i = 0; i < enabledGroups.size() && !free_females.isEmpty(); i++) {
        int group_idx = enabledGroups[i];
        int local_idx = i;

        if (constrained_groups.contains(group_idx)) {
            continue;
        }

        GroupConfig config = groupConfigs[group_idx];
        int current_females = 0;
        for (int person : local_groups[local_idx]) {
            if (person != 0 && getGender(person) == 'F') current_females++;
        }

        int need = config.females - current_females;
        if (need > 0) {
            for (int j = 0; j < need && !free_females.isEmpty(); j++) {
                // 找到第一个空位
                bool assigned = false;
                for (int k = 0; k < local_groups[local_idx].size() && !assigned; k++) {
                    if (local_groups[local_idx][k] == 0) {
                        local_groups[local_idx][k] = free_females.takeFirst();
                        assigned = true;
                        logOutput->append(QString("分配女生 %1 到组%2座位%3")
                            .arg(id_to_name[local_groups[local_idx][k]]).arg(group_idx + 1).arg(k + 1));
                    }
                }
                if (!assigned) {
                    break;
                }
            }
        }
    }

    // 第二步：分配男生
    for (int i = 0; i < enabledGroups.size() && !free_males.isEmpty(); i++) {
        int group_idx = enabledGroups[i];
        int local_idx = i;

        if (constrained_groups.contains(group_idx)) {
            continue;
        }

        GroupConfig config = groupConfigs[group_idx];
        int current_males = 0;
        for (int person : local_groups[local_idx]) {
            if (person != 0 && getGender(person) == 'M') current_males++;
        }

        int need = config.males - current_males;
        if (need > 0) {
            for (int j = 0; j < need && !free_males.isEmpty(); j++) {
                // 找到第一个空位
                bool assigned = false;
                for (int k = 0; k < local_groups[local_idx].size() && !assigned; k++) {
                    if (local_groups[local_idx][k] == 0) {
                        local_groups[local_idx][k] = free_males.takeFirst();
                        assigned = true;
                        logOutput->append(QString("分配男生 %1 到组%2座位%3")
                            .arg(id_to_name[local_groups[local_idx][k]]).arg(group_idx + 1).arg(k + 1));
                    }
                }
                if (!assigned) {
                    break; // 该组没有空位了
                }
            }
        }
    }

    // 第三步：分配剩余女生到任何空位
    while (!free_females.isEmpty()) {
        bool assigned = false;
        for (int i = 0; i < enabledGroups.size(); i++) {
            int local_idx = i;

            // 检查是否有空位
            for (int k = 0; k < local_groups[local_idx].size(); k++) {
                if (local_groups[local_idx][k] == 0) {
                    local_groups[local_idx][k] = free_females.takeFirst();
                    assigned = true;
                    logOutput->append(QString("分配剩余女生 %1 到组%2座位%3")
                        .arg(id_to_name[local_groups[local_idx][k]]).arg(i + 1).arg(k + 1));
                    break;
                }
            }

            if (free_females.isEmpty()) {
                break;
            }
        }

        if (!assigned && !free_females.isEmpty()) {
            logOutput->append("警告: 无法分配所有女生，可能座位不足");
            break;
        }
    }

    // 第四步：分配剩余男生到任何空位
    while (!free_males.isEmpty()) {
        bool assigned = false;
        for (int i = 0; i < enabledGroups.size(); i++) {
            int local_idx = i;

            // 检查是否有空位
            for (int k = 0; k < local_groups[local_idx].size(); k++) {
                if (local_groups[local_idx][k] == 0) {
                    local_groups[local_idx][k] = free_males.takeFirst();
                    assigned = true;
                    logOutput->append(QString("分配剩余男生 %1 到组%2座位%3")
                        .arg(id_to_name[local_groups[local_idx][k]]).arg(i + 1).arg(k + 1));
                    break;
                }
            }

            if (free_males.isEmpty()) {
                break;
            }
        }

        if (!assigned && !free_males.isEmpty()) {
            logOutput->append("警告: 无法分配所有男生，可能座位不足");
            break;
        }
    }

    // 第五步：重新设计组长分配逻辑，确保每组最多一个组长
    QVector<int> availableLeaders;
    for (const QString& leaderName : leaders) {
        if (name_to_id.contains(leaderName)) {
            int leaderId = name_to_id[leaderName];
            // 检查该组长是否已被分配
            bool alreadyAssigned = false;
            for (int i = 0; i < local_groups.size() && !alreadyAssigned; i++) {
                for (int person : local_groups[i]) {
                    if (person == leaderId) {
                        alreadyAssigned = true;
                        break;
                    }
                }
            }
            if (!alreadyAssigned) {
                availableLeaders.append(leaderId);
            }
            else {
                logOutput->append(QString("组长 %1 已被分配").arg(leaderName));
            }
        }
    }

    std::shuffle(availableLeaders.begin(), availableLeaders.end(), g);

    // 首先标记哪些组已经有组长
    QVector<bool> groupHasLeader(local_groups.size(), false);
    for (int i = 0; i < local_groups.size(); i++) {
        for (int person : local_groups[i]) {
            if (person != 0 && leaders.contains(id_to_name[person])) {
                groupHasLeader[i] = true;
                logOutput->append(QString("组%1 已有组长: %2").arg(i + 1).arg(id_to_name[person]));
                break;
            }
        }
    }

    // 为没有组长的组分配组长
    for (int i = 0; i < local_groups.size() && !availableLeaders.isEmpty(); i++) {
        if (!groupHasLeader[i]) {
            // 在组内找一个空位放置组长
            bool leaderAssigned = false;
            for (int j = 0; j < local_groups[i].size() && !leaderAssigned; j++) {
                if (local_groups[i][j] == 0) {
                    int leaderId = availableLeaders.takeFirst();
                    local_groups[i][j] = leaderId;
                    leaderAssigned = true;
                    groupHasLeader[i] = true;
                    logOutput->append(QString("分配组长 %1 到组%2座位%3")
                        .arg(id_to_name[leaderId]).arg(i + 1).arg(j + 1));
                }
            }

            if (!leaderAssigned) {
                logOutput->append(QString("警告: 无法为组%1 分配组长，没有空位").arg(i + 1));
            }
        }
    }

    // 如果还有剩余的组长，记录警告
    if (!availableLeaders.isEmpty()) {
        logOutput->append(QString("警告: 有 %1 个组长未能分配").arg(availableLeaders.size()));
        for (int leaderId : availableLeaders) {
            logOutput->append(QString("未分配组长: %1").arg(id_to_name[leaderId]));
        }
    }

    // 第六步：彻底重写组长冲突解决逻辑
    bool leaderConflictResolved = false;
    int maxLeaderConflictIterations = 50;
    int leaderConflictIteration = 0;

    while (!leaderConflictResolved && leaderConflictIteration < maxLeaderConflictIterations) {
        leaderConflictResolved = true;
        leaderConflictIteration++;

        // 首先统计每个组的组长数量
        QVector<int> leaderCounts(local_groups.size(), 0);
        QVector<QVector<int>> leaderPositions(local_groups.size());

        for (int i = 0; i < local_groups.size(); i++) {
            for (int j = 0; j < local_groups[i].size(); j++) {
                if (local_groups[i][j] != 0 && leaders.contains(id_to_name[local_groups[i][j]])) {
                    leaderCounts[i]++;
                    leaderPositions[i].append(j);
                }
            }
        }

        // 找出组长过多和过少的组
        QVector<int> groupsWithExtraLeaders;
        QVector<int> groupsWithoutLeaders;

        for (int i = 0; i < local_groups.size(); i++) {
            if (leaderCounts[i] > 1) {
                groupsWithExtraLeaders.append(i);
            }
            else if (leaderCounts[i] == 0) {
                groupsWithoutLeaders.append(i);
            }
        }

        // 如果没有冲突，退出循环
        if (groupsWithExtraLeaders.isEmpty() || groupsWithoutLeaders.isEmpty()) {
            leaderConflictResolved = true;
            break;
        }

        // 尝试移动多余的组长到没有组长的组
        bool movedInThisIteration = false;

        for (int sourceGroup : groupsWithExtraLeaders) {
            if (leaderCounts[sourceGroup] <= 1) continue;

            // 从该组找出可以移动的组长
            for (int leaderPos : leaderPositions[sourceGroup]) {
                int leaderId = local_groups[sourceGroup][leaderPos];

                // 跳过固定位置的组长
                if (fixed_people.contains(leaderId)) {
                    continue;
                }

                // 尝试移动到没有组长的组
                for (int targetGroup : groupsWithoutLeaders) {
                    // 在目标组找空位
                    for (int k = 0; k < local_groups[targetGroup].size(); k++) {
                        if (local_groups[targetGroup][k] == 0) {
                            // 移动组长
                            local_groups[targetGroup][k] = leaderId;
                            local_groups[sourceGroup][leaderPos] = 0;

                            logOutput->append(QString("解决组长冲突: 将组长 %1 从组%2 移动到组%3")
                                .arg(id_to_name[leaderId]).arg(sourceGroup + 1).arg(targetGroup + 1));

                            movedInThisIteration = true;
                            leaderConflictResolved = false;
                            break;
                        }
                    }

                    if (movedInThisIteration) break;
                }

                if (movedInThisIteration) break;
            }

            if (movedInThisIteration) break;
        }

        // 如果没有移动成功，尝试交换
        if (!movedInThisIteration) {
            for (int sourceGroup : groupsWithExtraLeaders) {
                if (leaderCounts[sourceGroup] <= 1) continue;

                for (int leaderPos : leaderPositions[sourceGroup]) {
                    int leaderId = local_groups[sourceGroup][leaderPos];

                    // 跳过固定位置的组长
                    if (fixed_people.contains(leaderId)) {
                        continue;
                    }

                    // 尝试与没有组长的组中的非组长交换
                    for (int targetGroup : groupsWithoutLeaders) {
                        for (int k = 0; k < local_groups[targetGroup].size(); k++) {
                            int targetPerson = local_groups[targetGroup][k];
                            if (targetPerson != 0 && !leaders.contains(id_to_name[targetPerson]) &&
                                !fixed_people.contains(targetPerson) &&
                                getGender(leaderId) == getGender(targetPerson)) {

                                // 交换人员
                                std::swap(local_groups[sourceGroup][leaderPos], local_groups[targetGroup][k]);

                                logOutput->append(QString("解决组长冲突(交换): 将组长 %1 从组%2 与组%3的 %4 交换")
                                    .arg(id_to_name[leaderId]).arg(sourceGroup + 1).arg(targetGroup + 1).arg(id_to_name[targetPerson]));

                                movedInThisIteration = true;
                                leaderConflictResolved = false;
                                break;
                            }
                        }

                        if (movedInThisIteration) break;
                    }

                    if (movedInThisIteration) break;
                }

                if (movedInThisIteration) break;
            }
        }

        if (!movedInThisIteration) {
            // 如果这一轮没有任何移动，退出循环避免无限循环
            break;
        }
    }

    // 最终检查组长分配
    QMap<int, int> finalLeaderCounts;
    for (int i = 0; i < local_groups.size(); i++) {
        int count = 0;
        for (int person : local_groups[i]) {
            if (person != 0 && leaders.contains(id_to_name[person])) {
                count++;
            }
        }
        finalLeaderCounts[i] = count;

        if (count > 1) {
            logOutput->append(QString("警告: 经过冲突解决后，组%1 仍有 %2 个组长")
                .arg(i + 1).arg(count));
        }
        else if (count == 0) {
            logOutput->append(QString("警告: 经过冲突解决后，组%1 没有组长")
                .arg(i + 1));
        }
    }

    // 创建不可移动人员集合
    QSet<int> immovable_people = fixed_people;
    for (int i = 0; i < local_groups.size(); i++) {
        for (int person : local_groups[i]) {
            if (person != 0 && leaders.contains(id_to_name[person])) {
                immovable_people.insert(person);
            }
        }
    }

    // 男生连续在前，女生连续在后
    logOutput->append("进行最终性别分组排列...");

    for (int i = 0; i < local_groups.size(); i++) {
        // 收集固定位置信息
        QMap<int, int> fixed_positions_in_group;
        for (int j = 0; j < local_groups[i].size(); j++) {
            int person = local_groups[i][j];
            if (person != 0 && fixed_people.contains(person)) {
                fixed_positions_in_group[j] = person;
            }
        }

        // 收集非固定位置的男生和女生
        QVector<int> non_fixed_males;
        QVector<int> non_fixed_females;

        for (int j = 0; j < local_groups[i].size(); j++) {
            int person = local_groups[i][j];
            if (person != 0 && !fixed_positions_in_group.contains(j)) {
                if (getGender(person) == 'M') {
                    non_fixed_males.append(person);
                }
                else {
                    non_fixed_females.append(person);
                }
            }
        }

        // 随机打乱非固定位置的男生和女生
        std::shuffle(non_fixed_males.begin(), non_fixed_males.end(), g);
        std::shuffle(non_fixed_females.begin(), non_fixed_females.end(), g);

        // 重新组合，保留固定位置，男生在前连续排列，女生在后连续排列
        QVector<int> new_arrangement(local_groups[i].size(), 0);

        // 首先放置固定位置
        for (int pos : fixed_positions_in_group.keys()) {
            new_arrangement[pos] = fixed_positions_in_group[pos];
        }

        // 然后按顺序先放置非固定男生（连续），再放置非固定女生（连续）
        int current_seat = 0;

        // 放置非固定男生
        for (int j = 0; j < non_fixed_males.size(); j++) {
            // 找到第一个空位
            while (current_seat < new_arrangement.size() && new_arrangement[current_seat] != 0) {
                current_seat++;
            }
            if (current_seat < new_arrangement.size()) {
                new_arrangement[current_seat] = non_fixed_males[j];
                current_seat++;
            }
        }

        // 放置非固定女生
        for (int j = 0; j < non_fixed_females.size(); j++) {
            // 找到第一个空位
            while (current_seat < new_arrangement.size() && new_arrangement[current_seat] != 0) {
                current_seat++;
            }
            if (current_seat < new_arrangement.size()) {
                new_arrangement[current_seat] = non_fixed_females[j];
                current_seat++;
            }
        }

        local_groups[i] = new_arrangement;

        // 记录性别分布用于调试
        QString distribution;
        int male_count = 0;
        int female_count = 0;
        bool in_male_section = true;

        for (int person : local_groups[i]) {
            if (person != 0) {
                if (getGender(person) == 'M') {
                    distribution += "M";
                    male_count++;
                    if (!in_male_section) {
                        logOutput->append(QString("警告: 组%1 中男生出现在女生区域").arg(i + 1));
                    }
                }
                else {
                    distribution += "F";
                    female_count++;
                    in_male_section = false;
                }
                distribution += " ";
            }
        }

        logOutput->append(QString("组%1 最终性别分布: %2 (男生:%3, 女生:%4)")
            .arg(i + 1).arg(distribution).arg(male_count).arg(female_count));

        // 验证男生是否连续且在前
        bool found_female = false;
        bool male_after_female = false;
        for (int person : local_groups[i]) {
            if (person != 0) {
                if (getGender(person) == 'F') {
                    found_female = true;
                }
                else if (getGender(person) == 'M' && found_female) {
                    male_after_female = true;
                    break;
                }
            }
        }

        if (male_after_female) {
            logOutput->append(QString("错误: 组%1 中男生出现在女生后面").arg(i + 1));
        }
    }

    // 处理不能同组要求
    logOutput->append("处理不能同组要求...");
    QMap<int, int> person_to_group;
    for (int i = 0; i < local_groups.size(); i++) {
        for (int person : local_groups[i]) {
            if (person != 0) { // 跳过空位
                person_to_group[person] = i;
            }
        }
    }

    bool changed = false;
    do {
        changed = false;
        for (const auto& group : must_separate_groups) {
            QMap<int, QVector<int>> group_members;
            for (int person : group) {
                int g = person_to_group[person];
                group_members[g].append(person);
            }

            // 检查每个组，如果同一个组内有超过1个该要求组的人，则需要移动
            for (auto it = group_members.begin(); it != group_members.end(); ++it) {
                if (it.value().size() <= 1)
                    continue;

                // 对于每个多余的人（从第二个开始），尝试移动到其他组
                for (int i = 1; i < it.value().size(); i++) {
                    int person = it.value()[i];
                    int current_group = person_to_group[person];

                    // 如果这个人在不可移动集合中，不能移动
                    if (immovable_people.contains(person)) {
                        logOutput->append(QString("警告: 人员 %1 在不可移动位置，无法移动以满足不能同组要求").arg(id_to_name[person]));
                        continue;
                    }

                    // 尝试移动到其他组
                    bool moved = false;
                    QVector<int> candidate_groups;
                    for (int g = 0; g < local_groups.size(); g++) {
                        if (g != current_group) {
                            candidate_groups.append(g);
                        }
                    }
                    std::shuffle(candidate_groups.begin(), candidate_groups.end(), g);

                    for (int target_group : candidate_groups) {
                        // 检查目标组中是否有当前要求组的其他人
                        bool conflict = false;
                        for (int p : group) {
                            if (p != person && person_to_group[p] == target_group) {
                                conflict = true;
                                break;
                            }
                        }
                        if (conflict) continue;

                        // 在目标组找一个同性别的人交换
                        int x = findSameGenderInGroup(person, target_group, immovable_people);
                        if (x != -1) {
                            // 确保x不在当前要求组中且不是不可移动人员
                            if (group.contains(x) || immovable_people.contains(x))
                                continue;

                            // 交换
                            int person_index = local_groups[current_group].indexOf(person);
                            int x_index = local_groups[target_group].indexOf(x);

                            if (person_index != -1 && x_index != -1) {
                                local_groups[current_group][person_index] = x;
                                local_groups[target_group][x_index] = person;

                                // 更新映射
                                person_to_group[person] = target_group;
                                person_to_group[x] = current_group;

                                changed = true;
                                moved = true;
                                break;
                            }
                        }
                    }
                    if (moved) break;
                }
            }
        }
    } while (changed);

    // 平衡外宿生分布
    logOutput->append("平衡外宿生分布（不移动组长）...");
    QMap<int, int> boarderCount;
    for (int group_idx = 0; group_idx < local_groups.size(); group_idx++) {
        int count = 0;
        for (int person : local_groups[group_idx]) {
            if (person != 0 && boarders.contains(id_to_name[person])) {
                count++;
            }
        }
        boarderCount[group_idx] = count;
    }

    // 平衡外宿生分布，确保各组外宿生数量差异不超过1
    bool balanced = false;
    int maxIterations = 100; // 防止无限循环
    int iteration = 0;

    while (!balanced && iteration < maxIterations) {
        balanced = true;
        iteration++;

        // 找到外宿生最多和最少的组
        int maxBoarders = -1, minBoarders = INT_MAX;
        int maxGroup = -1, minGroup = -1;

        for (int i = 0; i < local_groups.size(); i++) {
            int count = boarderCount[i];
            if (count > maxBoarders) {
                maxBoarders = count;
                maxGroup = i;
            }
            if (count < minBoarders) {
                minBoarders = count;
                minGroup = i;
            }
        }

        // 如果差异大于1，尝试平衡
        if (maxBoarders - minBoarders > 1) {
            // 尝试从外宿生多的组移动一个外宿生到外宿生少的组
            bool moved = false;
            for (int i = 0; i < local_groups[maxGroup].size() && !moved; i++) {
                int person = local_groups[maxGroup][i];
                if (person != 0 && boarders.contains(id_to_name[person]) && !immovable_people.contains(person)) {
                    // 在外宿生少的组找一个非外宿生交换
                    for (int j = 0; j < local_groups[minGroup].size() && !moved; j++) {
                        int otherPerson = local_groups[minGroup][j];
                        if (otherPerson != 0 && !boarders.contains(id_to_name[otherPerson]) &&
                            !immovable_people.contains(otherPerson) &&
                            getGender(person) == getGender(otherPerson)) {
                            // 交换
                            std::swap(local_groups[maxGroup][i], local_groups[minGroup][j]);
                            boarderCount[maxGroup]--;
                            boarderCount[minGroup]++;
                            moved = true;
                            balanced = false;
                            logOutput->append(QString("平衡外宿生: 将 %1 从组%2 移动到组%3")
                                .arg(id_to_name[person]).arg(maxGroup + 1).arg(minGroup + 1));
                            break;
                        }
                    }
                }
            }
        }
        else {
            balanced = true;
        }
    }

    // 使用交换算法优化固定位置
    optimizeFixedPositions();

    // 最终验证：检查固定位置和组长分配
    logOutput->append("开始最终验证...");

    // 验证固定位置
    for (auto it = fixedPositions.begin(); it != fixedPositions.end(); ++it) {
        int personId = it.key();
        FixedPosition fixedPos = it.value();
        int groupNumber = fixedPos.groupNumber;
        int seatPosition = fixedPos.seatPosition;
        int groupIdx = groupNumber - 1;

        int localIdx = enabledGroups.indexOf(groupIdx);
        if (localIdx != -1) {
            int seatIndex = seatPosition - 1;
            if (local_groups[localIdx][seatIndex] == personId) {
                logOutput->append(QString("✓ 固定位置验证通过: %1 在组%2座位%3")
                    .arg(id_to_name[personId]).arg(groupNumber).arg(seatPosition));
            }
            else {
                logOutput->append(QString("✗ 固定位置验证失败: %1 应该在组%2座位%3，实际在组%4")
                    .arg(id_to_name[personId]).arg(groupNumber).arg(seatPosition)
                    .arg(localIdx + 1));
            }
        }
    }

    // 验证组长分配
    QMap<int, int> leaderGroupCount;
    for (int i = 0; i < local_groups.size(); i++) {
        int leaderCount = 0;
        for (int person : local_groups[i]) {
            if (person != 0 && leaders.contains(id_to_name[person])) {
                leaderCount++;
            }
        }
        leaderGroupCount[i] = leaderCount;

        if (leaderCount == 1) {
            logOutput->append(QString("✓ 组长分配验证通过: 组%1 有1个组长").arg(i + 1));
        }
        else if (leaderCount == 0) {
            logOutput->append(QString("⚠ 组长分配警告: 组%1 没有组长").arg(i + 1));
        }
        else {
            logOutput->append(QString("✗ 组长分配失败: 组%1 有%2个组长").arg(i + 1).arg(leaderCount));
        }
    }

    // 移除所有空位
    for (int i = 0; i < local_groups.size(); i++) {
        QVector<int> nonZero;
        for (int person : local_groups[i]) {
            if (person != 0) {
                nonZero.append(person);
            }
        }
        local_groups[i] = nonZero;
    }

    logOutput->append("分组完成");
    return { local_groups, local_special_groups };
}

// 组内交换
bool MainWindow::swapPersonsInGroup(int groupIndex, int seatA, int seatB)
{
    if (groupIndex < 0 || groupIndex >= groups.size()) return false;
    if (seatA < 0 || seatA >= groups[groupIndex].size()) return false;
    if (seatB < 0 || seatB >= groups[groupIndex].size()) return false;

    std::swap(groups[groupIndex][seatA], groups[groupIndex][seatB]);
    logOutput->append(QString("组内交换: 位置%1 ↔ 位置%2").arg(seatA + 1).arg(seatB + 1));
    return true;
}

// 跨组交换
bool MainWindow::swapPersonsBetweenGroups(int groupA, int seatA, int groupB, int seatB)
{
    if (groupA < 0 || groupA >= groups.size()) return false;
    if (groupB < 0 || groupB >= groups.size()) return false;
    if (seatA < 0 || seatA >= groups[groupA].size()) return false;
    if (seatB < 0 || seatB >= groups[groupB].size()) return false;

    // 检查性别要求
    int personA = groups[groupA][seatA];
    int personB = groups[groupB][seatB];

    if (getGender(personA) != getGender(personB)) {
        logOutput->append("错误: 跨组交换需要同性别人员");
        return false;
    }

    std::swap(groups[groupA][seatA], groups[groupB][seatB]);
    logOutput->append(QString("跨组交换: 组%1位置%2 ↔ 组%3位置%4")
        .arg(groupA + 1).arg(seatA + 1).arg(groupB + 1).arg(seatB + 1));
    return true;
}

// 移动人员到指定位置
bool MainWindow::movePersonToPosition(int personId, int targetGroup, int targetSeat)
{
    // 找到人员当前位置
    int currentGroup = -1, currentSeat = -1;
    for (int i = 0; i < groups.size(); i++) {
        for (int j = 0; j < groups[i].size(); j++) {
            if (groups[i][j] == personId) {
                currentGroup = i;
                currentSeat = j;
                break;
            }
        }
        if (currentGroup != -1) break;
    }

    if (currentGroup == -1) {
        logOutput->append(QString("错误: 未找到人员 %1").arg(id_to_name[personId]));
        return false;
    }

    // 如果已经在目标位置，直接返回成功
    if (currentGroup == targetGroup && currentSeat == targetSeat) {
        return true;
    }

    // 查找交换路径
    auto swapPath = findSwapPath(currentGroup, currentSeat, targetGroup, targetSeat);

    if (swapPath.isEmpty()) {
        logOutput->append(QString("错误: 无法找到从组%1位置%2到组%3位置%4的交换路径")
            .arg(currentGroup + 1).arg(currentSeat + 1)
            .arg(targetGroup + 1).arg(targetSeat + 1));
        return false;
    }

    // 执行交换链
    for (const auto& swap : swapPath) {
        if (swap.first.group == swap.second.group) {
            // 组内交换
            swapPersonsInGroup(swap.first.group, swap.first.seat, swap.second.seat);
        }
        else {
            // 跨组交换
            swapPersonsBetweenGroups(swap.first.group, swap.first.seat,
                swap.second.group, swap.second.seat);
        }
    }

    logOutput->append(QString("成功移动 %1 到组%2位置%3")
        .arg(id_to_name[personId]).arg(targetGroup + 1).arg(targetSeat + 1));
    return true;
}

// 查找交换路径
QVector<QPair<Position, Position>> MainWindow::findSwapPath(int startGroup, int startSeat, int targetGroup, int targetSeat)
{
    std::queue<SwapState> queue;
    QSet<QString> visited;

    queue.push(SwapState(startGroup, startSeat, {}));
    visited.insert(QString("%1-%2").arg(startGroup).arg(startSeat));

    while (!queue.empty()) {
        SwapState current = queue.front();
        queue.pop();

        // 如果到达目标位置
        if (current.group == targetGroup && current.seat == targetSeat) {
            return current.path;
        }

        // 尝试所有可能的交换

        // 1. 同组交换
        for (int otherSeat = 0; otherSeat < groups[current.group].size(); otherSeat++) {
            if (otherSeat != current.seat) {
                QString newState = QString("%1-%2").arg(current.group).arg(otherSeat);
                if (!visited.contains(newState)) {
                    QVector<QPair<Position, Position>> newPath = current.path;
                    newPath.append(qMakePair(Position(current.group, current.seat), Position(current.group, otherSeat)));
                    queue.push(SwapState(current.group, otherSeat, newPath));
                    visited.insert(newState);
                }
            }
        }

        // 2. 跨组交换（只考虑同性别）
        int currentPerson = groups[current.group][current.seat];
        for (int otherGroup = 0; otherGroup < groups.size(); otherGroup++) {
            if (otherGroup == current.group) continue;

            for (int otherSeat = 0; otherSeat < groups[otherGroup].size(); otherSeat++) {
                int otherPerson = groups[otherGroup][otherSeat];
                if (getGender(currentPerson) == getGender(otherPerson)) {
                    QString newState = QString("%1-%2").arg(otherGroup).arg(otherSeat);
                    if (!visited.contains(newState)) {
                        QVector<QPair<Position, Position>> newPath = current.path;
                        newPath.append(qMakePair(Position(current.group, current.seat), Position(otherGroup, otherSeat)));
                        queue.push(SwapState(otherGroup, otherSeat, newPath));
                        visited.insert(newState);
                    }
                }
            }
        }
    }

    return {};
}

// 优化固定位置分配
void MainWindow::optimizeFixedPositions()
{
    logOutput->append("开始优化固定位置分配...");

    int successCount = 0;
    int failCount = 0;

    // 构建当前分组的人员位置映射
    QMap<int, QPair<int, int>> currentPositions;
    for (int i = 0; i < groups.size(); i++) {
        for (int j = 0; j < groups[i].size(); j++) {
            if (groups[i][j] != 0) {
                currentPositions[groups[i][j]] = qMakePair(i, j);
            }
        }
    }

    for (auto it = fixedPositions.begin(); it != fixedPositions.end(); ++it) {
        int personId = it.key();
        FixedPosition fixedPos = it.value();
        int targetGroup = fixedPos.groupNumber - 1;
        int targetSeat = fixedPos.seatPosition - 1;

        // 检查人员是否在当前分组中
        if (!currentPositions.contains(personId)) {
            logOutput->append(QString("错误: 人员 %1 不在当前分组中").arg(id_to_name[personId]));
            failCount++;
            continue;
        }

        // 获取当前位置
        QPair<int, int> currentPos = currentPositions[personId];
        int currentGroup = currentPos.first;
        int currentSeat = currentPos.second;

        // 如果已经在目标位置，直接返回成功
        if (currentGroup == targetGroup && currentSeat == targetSeat) {
            successCount++;
            continue;
        }

        // 尝试移动人员到固定位置
        if (movePersonToPosition(personId, targetGroup, targetSeat)) {
            successCount++;
        }
        else {
            failCount++;
            logOutput->append(QString("固定位置分配失败: %1 到组%2位置%3")
                .arg(id_to_name[personId]).arg(targetGroup + 1).arg(targetSeat + 1));
        }
    }

    logOutput->append(QString("固定位置优化完成: 成功 %1, 失败 %2").arg(successCount).arg(failCount));
}

void MainWindow::doExportSeatingPlan(const QString& fileName)
{
    // 参数校验
    if (fileName.isEmpty()) {
        QMessageBox::warning(this, "错误", "文件名不能为空");
        return;
    }

#ifdef Q_OS_WIN
    // COM初始化
    HRESULT hr = ::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        _com_error err(hr);
        QMessageBox::warning(this, "错误", "COM初始化失败: " + QString::fromWCharArray(err.ErrorMessage()));
        return;
    }
#else
    QMessageBox::information(this, "提示", "Excel导出功能仅在Windows平台可用");
    return;
#endif

    // 确保文件扩展名
    QString filePath = fileName;
    if (!filePath.endsWith(".xlsx", Qt::CaseInsensitive)) {
        filePath += ".xlsx";
    }

    // 根据当前选择的样式确定模板文件
    QString templateDir = QCoreApplication::applicationDirPath();
    QString template1 = templateDir + "/座位样式1.xlsx";
    QString template2 = templateDir + "/座位样式2.xlsx";

    QString selectedTemplate;

    if (currentStyle == "样式1") {
        selectedTemplate = template1;
    }
    else if (currentStyle == "样式2") {
        selectedTemplate = template2;
    }
    else if (currentStyle == "自定义") {
        QSettings settings(configPath, QSettings::IniFormat);
        selectedTemplate = settings.value("Style/CustomPath", "").toString();

        if (selectedTemplate.isEmpty() || !QFile::exists(selectedTemplate)) {
            QMessageBox::warning(this, "错误", "自定义样式文件不存在，请重新选择");
#ifdef Q_OS_WIN
            ::CoUninitialize();
#endif
            return;
        }
    }
    else {
        // 默认使用样式1
        selectedTemplate = template1;
    }

    // 检查模板文件是否存在
    if (!QFile::exists(selectedTemplate)) {
        // 尝试使用备用模板
        if (QFile::exists(template1)) {
            selectedTemplate = template1;
        }
        else if (QFile::exists(template2)) {
            selectedTemplate = template2;
        }
        else {
            QMessageBox::warning(this, "错误", "未找到座位模板文件，请确保同目录下有'座位样式1.xlsx'或'座位样式2.xlsx'");
#ifdef Q_OS_WIN
            ::CoUninitialize();
#endif
            return;
        }
    }

    currentTemplatePath = selectedTemplate;

    QScopedPointer<QAxObject> excel;
    bool isWPS = false;

    // 先尝试创建 Excel 对象
    excel.reset(new QAxObject("Excel.Application"));
    if (excel->isNull()) {
        // 尝试WPS表格
        excel.reset(new QAxObject("KWPS.Application"));
        if (!excel->isNull()) {
            isWPS = true;
            logOutput->append("检测到WPS Office (Kingsoft)");
        }
        else {
            excel.reset(new QAxObject("KET.Application"));
            if (!excel->isNull()) {
                isWPS = true;
                logOutput->append("检测到WPS表格 (KET)");
            }
            else {
                excel.reset(new QAxObject("ET.Application"));
                if (!excel->isNull()) {
                    isWPS = true;
                    logOutput->append("检测到WPS表格 (ET)");
                }
                else {
                    QMessageBox::warning(this, "错误", "无法创建 Excel 或 WPS 对象，请确保已安装 Microsoft Excel 或 WPS Office");
#ifdef Q_OS_WIN
                    ::CoUninitialize();
#endif
                    return;
                }
            }
        }
    }
    else {
        logOutput->append("检测到Microsoft Excel");
    }

    QScopedPointer<QAxObject> workbooks;
    QScopedPointer<QAxObject> workbook;
    QScopedPointer<QAxObject> sheets;
    QScopedPointer<QAxObject> sheet;

    try {
        // 配置Excel应用
        excel->setProperty("Visible", false);
        excel->setProperty("DisplayAlerts", false);
        excel->setProperty("ScreenUpdating", false);

        // 打开模板文件
        workbooks.reset(excel->querySubObject("Workbooks"));
        if (workbooks.isNull()) throw std::runtime_error("无法获取Workbooks对象");

        QString winTemplatePath = QDir::toNativeSeparators(currentTemplatePath);
        workbook.reset(workbooks->querySubObject("Open(const QString&)", winTemplatePath));
        if (workbook.isNull()) throw std::runtime_error("无法打开模板文件");

        sheets.reset(workbook->querySubObject("Sheets"));
        if (sheets.isNull()) throw std::runtime_error("无法获取Sheets集合");

        sheet.reset(sheets->querySubObject("Item(int)", 1));
        if (sheet.isNull()) throw std::runtime_error("无法获取工作表");

        // 获取使用的单元格范围
        QScopedPointer<QAxObject> usedRange(sheet->querySubObject("UsedRange"));
        QScopedPointer<QAxObject> rows(usedRange->querySubObject("Rows"));
        QScopedPointer<QAxObject> columns(usedRange->querySubObject("Columns"));

        int rowCount = rows->property("Count").toInt();
        int colCount = columns->property("Count").toInt();

        // 遍历所有单元格，查找需要替换的内容
        for (int row = 1; row <= rowCount; row++) {
            for (int col = 1; col <= colCount; col++) {
                QScopedPointer<QAxObject> cell(sheet->querySubObject("Cells(int,int)", row, col));
                QVariant value = cell->property("Value");

                if (value.isValid() && !value.isNull()) {
                    QString cellValue = value.toString();

                    // 替换日期
                    if (cellValue.contains("XXXX.XX.XX")) {
                        QString currentDate = QDateTime::currentDateTime().toString("yyyy.MM.dd");
                        cell->setProperty("Value", currentDate);
                    }
                    // 处理组-人员编码
                    else if (cellValue.contains(QRegularExpression("^\\d+-\\d+$"))) {
                        QStringList parts = cellValue.split('-');
                        if (parts.size() == 2) {
                            bool ok1, ok2;
                            int groupNum = parts[0].toInt(&ok1);
                            int personNum = parts[1].toInt(&ok2);

                            if (ok1 && ok2 && groupNum >= 1 && groupNum <= groups.size() &&
                                personNum >= 1 && personNum <= groups[groupNum - 1].size()) {
                                // 获取对应的学生ID和姓名
                                int studentId = groups[groupNum - 1][personNum - 1];
                                QString studentName = id_to_name[studentId];

                                // 设置学生姓名
                                cell->setProperty("Value", studentName);

                                // 设置格式
                                cell->setProperty("HorizontalAlignment", -4108); // 水平居中
                                cell->setProperty("VerticalAlignment", -4108);   // 垂直居中
                                cell->setProperty("WrapText", true);             // 自动换行

                                // 设置字体
                                QScopedPointer<QAxObject> font(cell->querySubObject("Font"));
                                font->setProperty("Size", 16);

                                // 标记组长
                                if (leaders.contains(studentName)) {
                                    font->setProperty("Bold", true);
                                }

                                // 标记外宿生
                                if (boarders.contains(studentName)) {
                                    QScopedPointer<QAxObject> interior(cell->querySubObject("Interior"));
                                    interior->setProperty("Color", QColor(Qt::yellow).rgb());
                                }
                            }
                            else {
                                // 如果位置超出范围，留空
                                cell->setProperty("Value", "");
                            }
                        }
                    }
                }
            }
        }

        // 保存文件
        QString winPath = QDir::toNativeSeparators(filePath);

        // WPS特殊处理
        if (isWPS) {
            try {
                // WPS特有保存方式
                workbook->dynamicCall("SaveAs(const QString&)", winPath);
                logOutput->append("使用WPS特有保存方式");
            }
            catch (...) {
                try {
                    // 备用保存方式
                    QScopedPointer<QAxObject> activeWorkbook(excel->querySubObject("ActiveWorkbook"));
                    if (!activeWorkbook.isNull()) {
                        activeWorkbook->dynamicCall("SaveAs(QVariant)", winPath);
                        logOutput->append("使用WPS备用保存方式");
                    }
                }
                catch (...) {
                    // 最后尝试
                    workbook->dynamicCall("SaveAs(QString)", winPath);
                    logOutput->append("使用WPS最后尝试保存方式");
                }
            }
        }
        else {
            // Microsoft Excel标准保存方式
            QVariantList args;
            args << winPath;
            args << 51;
            workbook->dynamicCall("SaveAs(const QString&, int)", args);
        }

        // 验证保存结果
        QFileInfo savedFile(winPath);
        if (!savedFile.exists() || savedFile.size() == 0) {
            throw std::runtime_error("文件保存失败，请检查路径和权限");
        }

        // 完成操作
        excel->setProperty("ScreenUpdating", true);
        excel->setProperty("DisplayAlerts", true);
        QDesktopServices::openUrl(QUrl::fromLocalFile(winPath));
        updateStatus("座位表已导出到: " + winPath);
    }
    catch (const std::exception& e) {
        QMessageBox::critical(this, "错误", QString("导出失败: %1").arg(e.what()));
    }
    catch (...) {
        QMessageBox::critical(this, "错误", "导出失败: 未知错误");
    }

    // 清理资源
    if (!workbook.isNull()) {
        try {
            workbook->dynamicCall("Close(Boolean)", false);
        }
        catch (...) {
            // 忽略异常
        }
    }
    if (!excel.isNull()) {
        try {
            excel->dynamicCall("Quit()");
        }
        catch (...) {
            // 忽略异常
        }
    }

#ifdef Q_OS_WIN
    // 释放COM
    ::CoUninitialize();
#endif
}