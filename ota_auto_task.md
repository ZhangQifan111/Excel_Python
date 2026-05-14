# OTA 优化自动任务计划

## 背景
OTA 代码重构已进行到第三轮。#if defined 从 36 减到 13。
项目路径：`/Daily Work/2026.5.9 NewOta/aio/Debug_V2/`
源码目录：`src/OTA/src/` 和 `src/OTA/inc/`
ProductConfig 目录：`src/OTA/src/ProductConfig/`

## 今晚自动执行的 3 个任务

### 任务 1：check_bms_ext1/ext2_upgrade_status 搬到 the_st_config.h

**当前状态：**
- Ota_Config.c 中有 `#if defined(THE_ST)` 包裹的 `check_bms_ext1_upgrade_status` 和 `check_bms_ext2_upgrade_status` 函数定义
- 这两个函数被 the_st_config.h 的 `THE_ST_BMSCheckResult` 调用

**操作步骤：**
1. 读 Ota_Config.c，找到 `#if defined(THE_ST)` 包裹的 check_bms_ext1/ext2_upgrade_status 函数
2. 把这两个函数搬到 the_st_config.h，放在 BMSCheckResult 之前，改为 static 函数
3. 在 the_st_config.h 中更新 THE_ST_BMSCheckResult 的调用（如果函数名不变则不需要）
4. 从 Ota_Config.c 删除这两个函数的定义和 #if defined 块
5. 从 Ota_Master.h 删除对应的 extern 声明（如果有）

**注意：**
- 函数签名：`bool check_bms_ext1_upgrade_status(void)` 和 `bool check_bms_ext2_upgrade_status(void)`
- 函数内部引用了全局变量（bMcu_UpFlag, PcsRegData 等），这些在 Ota_Mast.c 的上下文中都可见
- 改名规则：按惯例加 THE_ST_ 前缀（如 THE_ST_CheckBmsExt1UpgradeStatus），但要同步更新调用处

### 任务 2：日志字符串统一

**当前状态：**
- Ota_Mast.c 中有 4 个 #if defined 块用于日志字符串（约在 705、2075、2231、2293 行）
- 日志格式因产品不同而略有差异（如 Send 方向名称不同）

**操作步骤：**
1. 读 Ota_Mast.c，找到所有日志相关的 #if defined 块
2. 分析日志差异：是 Send 方向名称不同，还是其他内容不同
3. 如果只是 Send 方向名称不同：
   - 在 ProductOps 中添加 `const char* (*GetSendDirDesc)(uint8_t dir)` 函数指针
   - 在三个 ProductConfig 文件中实现该函数
   - 替换日志中的 #if defined 为 `g_ProductOps.GetSendDirDesc(dir)`
4. 如果差异更复杂，保留 #if defined 并记录原因

**注意：**
- 日志中可能有 `__FUNCTION__` 或硬编码函数名，这些不需要统一
- 只统一 Send 方向描述字符串
- 如果改动太大（>50行），跳过这个任务，记录原因

### 任务 3：BMS/DCDC wrapper 函数简化

**当前状态：**
- Ota_Config.c 中有 BMS_CheckRequireFunc、DCDC_CheckRequireFunc、BMS_CheckResultFunc、DCDC_CheckResultFunc 等 wrapper 函数
- 它们只是调用 g_ProductOps 的对应函数指针
- Ota_Mcu_Info 表中填的是 wrapper 函数名

**操作步骤：**
1. 读 Ota_Config.c，列出所有 wrapper 函数
2. 读 Ota_Mcu_Info 表（三个 ProductConfig 文件），看表中填的是什么
3. 如果表中可以直接填 g_ProductOps 的函数指针：
   - 修改 Ota_Mcu_Info 表，直接填 g_ProductOps 的函数指针
   - 删除 wrapper 函数
4. 如果函数签名不匹配（wrapper 有额外参数），保留 wrapper 并记录原因

**注意：**
- Ota_Mcu_Info 表的函数指针签名是 `(uint8_t firmwareIndex, uint8_t MCU)`
- g_ProductOps 的函数指针签名也是 `(uint8_t, uint8_t)` 或 `(uint8_t)`
- 签名不匹配的不能直接替换
- 如果改动太大，跳过这个任务

## 完成后
1. 打包更新 `ota_refactor_v3_final.7z`
2. 更新 README 和 CHANGELOG
3. 通过微信通知奥特曼结果

## 参考文件
- Ota_Master.h: ProductOps 结构体定义（19 个函数指针）
- Ota_Mast.c: 主状态机（2792 行）
- Ota_Config.c: BMS/DCDC check 函数
- Ota_Common.c: 公共函数（#if defined 已清零）
- the_st_config.h: THE_ST 产品配置
- she_st_config.h: SHE_ST 产品配置
- default_config.h: 默认产品配置
