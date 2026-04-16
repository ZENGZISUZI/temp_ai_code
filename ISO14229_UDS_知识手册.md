# ISO 14229 - UDS 统一诊断服务 知识手册

> 全称：ISO 14229 - Road vehicles — Unified diagnostic services  
> 适用领域：车载诊断、ECU通信、刷写、故障码读取等

---

## 一、概述

### 1.1 什么是 UDS？

UDS（Unified Diagnostic Services）是 ISO 14229 定义的**统一诊断服务协议**，用于诊断测试设备（Tester）与车辆 ECU 之间的通信。

**核心思想**：标准化诊断交互，不同厂商的 ECU 使用统一的诊断协议。

### 1.2 UDS 与 OBD 的区别

| 对比项 | OBD（ISO 15031） | UDS（ISO 14229） |
|--------|-------------------|-------------------|
| 全称 | On-Board Diagnostics | Unified Diagnostic Services |
| 目的 | 排放相关法规诊断 | 全功能诊断（开发/生产/售后） |
| 服务范围 | 仅排放相关 | 全部ECU功能 |
| 强制性 | 法规强制 | 非强制（行业约定） |
| 服务数量 | 有限（约10个） | 丰富（6大类30+服务） |
| 应用场景 | 年检、维修站 | 开发调试、产线刷写、深度诊断 |

### 1.3 诊断通信模型

```
┌─────────────┐                    ┌─────────────┐
│   Tester    │ ◄── 诊断请求 ──►  │     ECU     │
│ (诊断设备)  │ ── 诊断响应 ──►   │  (控制器)   │
└─────────────┘                    └─────────────┘

通信方式：
  - 请求/响应模式（Client-Server）
  - Tester = Client（发起请求）
  - ECU = Server（返回响应）
```

---

## 二、标准结构

ISO 14229 共 7 个部分：

| 部分 | 标准编号 | 名称 | 说明 |
|------|----------|------|------|
| 1 | ISO 14229-1 | 应用层 | **核心部分**，定义所有诊断服务 |
| 2 | ISO 14229-2 | 会话层服务 | 会话管理 |
| 3 | ISO 14229-3 | 基于CAN的实现 | 数据链路层（ISO 15765-2） |
| 4 | ISO 14229-4 | 基于FlexRay的实现 | FlexRay总线 |
| 5 | ISO 14229-5 | 基于IP的实现 | DoIP（ISO 13400） |
| 6 | ISO 14229-6 | 基于K线的实现 | K-Line总线 |
| 7 | ISO 14229-7 | 基于LIN的实现 | LIN总线 |

**最常用**：Part 1（应用层）+ Part 3（CAN实现）+ Part 5（DoIP实现）

---

## 三、协议栈架构

```
┌─────────────────────────────────────┐
│          应用层 (Application)        │  ISO 14229-1 (UDS)
├─────────────────────────────────────┤
│          会话层 (Session)            │  ISO 14229-2
├─────────────────────────────────────┤
│          传输层 (Transport)          │  ISO 15765-2 (ISO-TP)
├─────────────────────────────────────┤
│          网络层 (Network)            │  ISO 15765-2
├─────────────────────────────────────┤
│          数据链路层 (Data Link)      │  ISO 11898 (CAN) / ISO 13400 (DoIP)
├─────────────────────────────────────┤
│          物理层 (Physical)           │  CAN-H/L, Ethernet, K-Line, LIN
└─────────────────────────────────────┘
```

---

## 四、诊断服务详解

### 4.1 服务分类总览

| 功能单元 | SID范围 | 服务组 | 说明 |
|----------|---------|--------|------|
| 0x10-0x3E | 诊断管理 | 会话/安全/通信控制 | 基础管理 |
| 0x40-0x7F | 数据传输 | 读/写数据 | 数据操作 |
| 0x80-0xBE | 已存储数据 | DTC操作 | 故障码管理 |
| 0xC0-0xFE | 输入输出控制 | IO/例程/刷写 | 执行控制 |

### 4.2 核心服务详解

---

#### 🔧 0x10 - DiagnosticSessionControl（诊断会话控制）

**功能**：切换ECU的诊断会话模式

| 子功能 | 名称 | 说明 |
|--------|------|------|
| 0x01 | defaultSession | 默认会话（上电默认） |
| 0x02 | programmingSession | 编程会话（刷写用） |
| 0x03 | extendedDiagnosticSession | 扩展诊断会话（高级诊断） |
| 0x04 | safetySystemDiagnosticSession | 安全系统诊断会话 |

**示例**：
```
请求：10 03        → 进入扩展诊断会话
正响应：50 03      → 成功进入
负响应：7F 10 12   → 子功能不支持
```

**重要规则**：
- 默认会话下，很多服务被限制（如安全访问、刷写等）
- 切换到非默认会话后，需要周期性发送 0x3E（TesterPresent）保活
- S3定时器超时（通常5秒），ECU自动回到默认会话

---

#### 🔧 0x11 - ECUReset（ECU复位）

**功能**：复位ECU

| 子功能 | 名称 | 说明 |
|--------|------|------|
| 0x01 | hardReset | 硬复位（完全重启） |
| 0x02 | keyResetOn | 软复位（按键复位） |
| 0x03 | softReset | 软复位 |
| 0x80 | enableRapidPowerShutDown | 启用快速断电 |
| 0x81 | disableRapidPowerShutDown | 禁用快速断电 |

**示例**：
```
请求：11 01        → 硬复位
正响应：51 01      → 复位执行中
```

---

#### 🔧 0x14 - ClearDTC（清除故障码）

**功能**：清除ECU中存储的所有DTC信息

**请求格式**：
```
14 FF FF FF    → 清除所有DTC
14 XX XX XX    → 清除指定DTC组
```

**注意**：
- 通常需要先进入扩展会话 + 安全解锁
- 清除后DTC计数归零，快照数据也被清除

---

#### 🔧 0x19 - ReadDTCInformation（读取DTC信息）

**功能**：读取ECU中存储的故障码信息

| 子功能 | 名称 | 说明 |
|--------|------|------|
| 0x01 | reportNumberOfDTCByStatus | 按状态查DTC数量 |
| 0x02 | reportDTCByStatus | 按状态查DTC列表 |
| 0x04 | reportDTCSnapshotIdentification | 查DTC快照ID |
| 0x05 | reportDTCSnapshotRecord | 查DTC快照数据 |
| 0x06 | reportDTCExtendedDataRecord | 查DTC扩展数据 |
| 0x0A | reportSupportedDTCs | 查支持的DTC列表 |

**DTC状态位定义**：
```
Bit 0: testFailed              - 当前测试失败
Bit 1: testFailedThisOperationCycle - 本次操作周期失败
Bit 2: pendingDTC              - 待确认DTC
Bit 3: confirmedDTC            - 已确认DTC
Bit 4: testNotCompletedSinceLastClear - 上次清除后测试未完成
Bit 5: testFailedSinceLastClear - 上次清除后测试失败过
Bit 6: testNotCompletedThisOperationCycle - 本次周期测试未完成
Bit 7: warningIndicatorRequested - 请求点亮警告灯
```

**示例**：
```
请求：19 01 01        → 查确认的DTC数量
正响应：59 01 01 00 03  → 有3个确认的DTC

请求：19 02 09        → 查所有DTC列表（状态掩码0x09=confirmed+pending）
正响应：59 02 09 [DTC1高] [DTC1中] [DTC1低] [状态1] [DTC2高] ...
```

---

#### 🔧 0x22 - ReadDataByIdentifier（按标识符读数据）

**功能**：通过DID（Data Identifier）读取ECU数据

**常用DID**：

| DID | 名称 | 说明 |
|-----|------|------|
| 0xF180 | VehicleManufacturerSparePartNumber | 零件号 |
| 0xF181 | VehicleManufacturerECUSoftwareNumber | 软件版本号 |
| 0xF182 | VehicleManufacturerECUSoftwareVersionNumber | 软件版本 |
| 0xF183 | SystemSupplierIdentifier | 供应商ID |
| 0xF184 | ECUManufacturingDate | 制造日期 |
| 0xF185 | ECUSerialNumber | 序列号 |
| 0xF186 | SupportedFunctionalUnits | 支持的功能单元 |
| 0xF187 | VehicleManufacturerECUHardwareNumber | 硬件版本号 |
| 0xF188 | SystemNameOrEngineType | 系统名称 |
| 0xF189 | RepairShopCode | 维修站代码 |
| 0xF18A | DiagnosticIdentifier | 诊断标识 |
| 0xF18B | DiagnosticVariantCode | 诊断变体码 |
| 0xF190 | VIN | 车辆识别码 |
| 0xF191 | VehicleManufacturerECUHardwareVersionNumber | 硬件版本 |
| 0xF193 | SystemIdentification | 系统标识 |
| 0xF195 | BootSoftwareIdentification | 引导程序标识 |
| 0xF197 | ProgrammingDate | 编程日期 |
| 0xF198 | CalibrationRepairShopCode | 标定维修站代码 |

**示例**：
```
请求：22 F1 90        → 读取VIN码
正响应：62 F1 90 LSVAU2189N2123456  → VIN = LSVAU2189N2123456

请求：22 F1 81        → 读取软件版本号
正响应：62 F1 81 SoftV2.1.0        → 软件版本 = SoftV2.1.0
```

---

#### 🔧 0x23 - ReadMemoryByAddress（按地址读内存）

**功能**：读取ECU指定内存地址的数据

**请求格式**：
```
23 [地址长度参数] [地址] [数据长度]
```

**注意**：
- 通常需要安全解锁后才能使用
- 地址和长度格式由addressAndLengthFormatIdentifier定义

---

#### 🔧 0x27 - SecurityAccess（安全访问）

**功能**：解锁ECU受保护的服务（刷写、写数据等）

**流程**：
```
Tester                          ECU
  │                               │
  │── 27 [Seed SID] ────────────►│  请求种子
  │                               │
  │◄── 67 [Seed SID] [Seed] ─────│  返回种子
  │                               │
  │  [计算Key]                    │
  │                               │
  │── 27 [Key SID] [Key] ───────►│  发送密钥
  │                               │
  │◄── 67 [Key SID] ─────────────│  解锁成功
  │                               │
```

**安全等级配对**：

| Seed SID | Key SID | 安全等级 |
|----------|---------|----------|
| 0x01 | 0x02 | Level 1 |
| 0x03 | 0x04 | Level 2 |
| 0x05 | 0x06 | Level 3 |
| ... | ... | ... |
| 0x41 | 0x42 | Level 33 |
| 0x43 | 0x44 | Level 34 |

**常见安全算法**：
- 简单异或/移位算法
- 基于种子的AES/DES加密
- 车企自定义算法（保密）

**防暴力破解**：
- 连续失败N次后延迟（指数退避）
- 超过最大次数后锁定（需硬复位）

---

#### 🔧 0x28 - CommunicationControl（通信控制）

**功能**：控制ECU的通信行为

| 子功能 | 名称 | 说明 |
|--------|------|------|
| 0x00 | enableRxAndTx | 启用收发 |
| 0x01 | enableRxAndDisableTx | 启用接收，禁用发送 |
| 0x02 | disableRxAndEnableTx | 禁用接收，启用发送 |
| 0x03 | disableRxAndTx | 禁用收发 |

**通信类型**：
| 控制类型 | 说明 |
|----------|------|
| 0x01 | 正常应用通信 |
| 0x02 | 网络管理通信 |
| 0x03 | 应用+网络管理通信 |

**典型场景**：刷写前禁用正常通信，避免干扰

---

#### 🔧 0x2E - WriteDataByIdentifier（按标识符写数据）

**功能**：通过DID向ECU写入数据

**示例**：
```
请求：2E F1 90 [17字节VIN]   → 写入VIN码
正响应：6E F1 90              → 写入成功
```

**注意**：
- 通常需要安全解锁
- 部分DID只读，写入会返回NRC 0x31

---

#### 🔧 0x2F - InputOutputControlByIdentifier（输入输出控制）

**功能**：控制ECU的输入输出信号

**控制模式**：
| 模式 | 说明 |
|------|------|
| 0x00 | returnControlToECU | 归还控制权 |
| 0x01 | resetToDefault | 恢复默认值 |
| 0x02 | freezeCurrentState | 冻结当前状态 |
| 0x03 | shortTermAdjustment | 短期调整（手动控制） |

**典型场景**：
- 手动控制风扇转速
- 强制点亮/熄灭指示灯
- 模拟传感器信号

---

#### 🔧 0x31 - RoutineControl（例程控制）

**功能**：启动/停止/查询ECU中的例程

| 子功能 | 名称 | 说明 |
|--------|------|------|
| 0x01 | startRoutine | 启动例程 |
| 0x02 | stopRoutine | 停止例程 |
| 0x03 | requestRoutineResults | 查询例程结果 |

**常用例程ID**：

| Routine ID | 名称 | 说明 |
|------------|------|------|
| 0x0203 | CheckProgrammingDependencies | 检查刷写依赖 |
| 0x0501 | EraseMemory | 擦除内存 |
| 0x0201 | CheckProgrammingPreConditions | 检查刷写前置条件 |
| 0xFF00 | EraseFlash | 擦除Flash |
| 0x0301 | CheckMemory | 校验内存 |

**示例**：
```
请求：31 01 FF 00        → 启动擦除Flash例程
正响应：71 01 FF 00      → 例程已启动

请求：31 03 FF 00        → 查询擦除结果
正响应：71 03 FF 00 00   → 擦除成功
```

---

#### 🔧 0x34 - RequestDownload（请求下载）

**功能**：请求ECU准备接收数据（刷写第一步）

**请求格式**：
```
34 [压缩/加密方法] [地址和长度格式] [起始地址] [数据长度]
```

**示例**：
```
请求：34 00 44 00 00 00 00 00 00 10 00   → 请求下载0x0000-0x1000
正响应：74 20 [maxBlockLength]            → 允许，最大块长度=0x20
```

---

#### 🔧 0x36 - TransferData（数据传输）

**功能**：传输刷写数据块

**格式**：
```
请求：36 [序号] [数据...]     → 传输数据块
正响应：76 [序号]             → 接收成功
```

**序号规则**：
- 从0x01开始
- 每次递增1，到0xFF后回绕到0x00
- 接收方校验序号连续性

---

#### 🔧 0x37 - RequestTransferExit（结束传输）

**功能**：通知ECU数据传输完成

```
请求：37              → 结束传输
正响应：77            → 传输完成确认
```

---

#### 🔧 0x3E - TesterPresent（心跳保活）

**功能**：保持当前诊断会话不超时

| 子功能 | 说明 |
|--------|------|
| 0x00 | 需要响应 |
| 0x80 | 不需要响应（suppressPosRspMsgIndicationBit） |

**示例**：
```
请求：3E 80        → 保活（不需要响应）
请求：3E 00        → 保活（需要响应）
正响应：7E         → 保活确认
```

**S3定时器**：通常5秒，超时ECU回到默认会话

---

#### 🔧 0x85 - ControlDTCSetting（控制DTC记录）

**功能**：启用/禁用DTC记录

| 子功能 | 名称 | 说明 |
|--------|------|------|
| 0x01 | on | 启用DTC记录 |
| 0x02 | off | 禁用DTC记录 |

**典型场景**：刷写或调试时禁用DTC记录，避免产生误报故障码

---

### 4.3 负响应码（NRC）

所有诊断服务失败时返回负响应，格式：`7F [SID] [NRC]`

| NRC | 名称 | 说明 |
|-----|------|------|
| 0x10 | generalReject | 通用拒绝 |
| 0x11 | serviceNotSupported | 服务不支持 |
| 0x12 | subFunctionNotSupported | 子功能不支持 |
| 0x13 | incorrectMessageLengthOrInvalidFormat | 消息长度错误或格式无效 |
| 0x14 | responseTooLong | 响应过长 |
| 0x21 | busyRepeatRequest | 忙碌，请重试 |
| 0x22 | conditionsNotCorrect | 条件不满足 |
| 0x24 | requestSequenceError | 请求序列错误 |
| 0x25 | noResponseFromSubnetComponent | 子网组件无响应 |
| 0x26 | failurePreventsExecutionOfRequestedAction | 故障阻止执行 |
| 0x31 | requestOutOfRange | 请求超出范围 |
| 0x33 | securityAccessDenied | 安全访问被拒绝 |
| 0x35 | invalidKey | 密钥无效 |
| 0x36 | exceededNumberOfAttempts | 超过尝试次数 |
| 0x37 | requiredTimeDelayNotExpired | 时间延迟未到期 |
| 0x70 | uploadDownloadNotAccepted | 上传/下载不被接受 |
| 0x71 | transferDataSuspended | 数据传输挂起 |
| 0x72 | generalProgrammingFailure | 通用编程失败 |
| 0x73 | wrongBlockSequenceCounter | 块序号错误 |
| 0x78 | requestCorrectlyReceived-ResponsePending | 请求已收到，响应稍后（**重要！**） |
| 0x7E | subFunctionNotSupportedInActiveSession | 当前会话不支持此子功能 |
| 0x7F | serviceNotSupportedInActiveSession | 当前会话不支持此服务 |

**⭐ NRC 0x78（ResponsePending）特别说明**：
- ECU收到请求但需要较长时间处理时，先返回0x78
- 作用：重置S3定时器，防止超时
- Tester收到0x78后应继续等待，不要超时报错

---

## 五、典型工作流程

### 5.1 诊断会话建立流程

```
Tester                          ECU
  │                               │
  │── 10 03 ────────────────────►│  进入扩展诊断会话
  │◄── 50 03 ───────────────────│  成功
  │                               │
  │── 3E 80 ────────────────────►│  保活（周期性，如2秒一次）
  │  （无需响应）                  │
  │                               │
  │── 22 F1 90 ─────────────────►│  读取VIN
  │◄── 62 F1 90 [VIN] ──────────│  返回VIN
  │                               │
  │── 19 02 09 ─────────────────►│  读取DTC
  │◄── 59 02 09 [DTC列表] ──────│  返回DTC
  │                               │
```

### 5.2 安全访问流程

```
Tester                          ECU
  │                               │
  │── 10 03 ────────────────────►│  进入扩展会话
  │◄── 50 03 ───────────────────│
  │                               │
  │── 27 01 ────────────────────►│  请求Seed（Level 1）
  │◄── 67 01 [Seed] ────────────│  返回Seed
  │                               │
  │  [用Seed计算Key]              │
  │                               │
  │── 27 02 [Key] ──────────────►│  发送Key
  │◄── 67 02 ───────────────────│  解锁成功！
  │                               │
  │── 2E F1 90 [VIN] ───────────►│  写入VIN（需要安全解锁）
  │◄── 6E F1 90 ────────────────│  写入成功
  │                               │
```

### 5.3 ECU刷写流程（完整）

```
Tester                          ECU
  │                               │
  │── 10 02 ────────────────────►│  ① 进入编程会话
  │◄── 50 02 ───────────────────│
  │                               │
  │── 27 05 ────────────────────►│  ② 安全访问（Level 3）
  │◄── 67 05 [Seed] ────────────│
  │── 27 06 [Key] ──────────────►│
  │◄── 67 06 ───────────────────│  解锁成功
  │                               │
  │── 28 03 01 ─────────────────►│  ③ 禁用正常通信
  │◄── 68 03 01 ────────────────│
  │                               │
  │── 85 02 ────────────────────►│  ④ 禁用DTC记录
  │◄── C5 02 ───────────────────│
  │                               │
  │── 31 01 02 01 ──────────────►│  ⑤ 检查刷写前置条件
  │◄── 71 01 02 01 ─────────────│  条件满足
  │                               │
  │── 31 01 FF 00 ──────────────►│  ⑥ 擦除Flash
  │◄── 71 01 FF 00 ─────────────│  擦除中...
  │── 31 03 FF 00 ──────────────►│  查询擦除结果
  │◄── 71 03 FF 00 [结果] ──────│  擦除完成
  │                               │
  │── 34 00 44 [地址] [长度] ───►│  ⑦ 请求下载
  │◄── 74 [maxBlockLen] ────────│  允许下载
  │                               │
  │── 36 01 [数据块1] ──────────►│  ⑧ 传输数据
  │◄── 76 01 ───────────────────│
  │── 36 02 [数据块2] ──────────►│
  │◄── 76 02 ───────────────────│
  │  ...                          │
  │── 36 NN [数据块N] ──────────►│
  │◄── 76 NN ───────────────────│
  │                               │
  │── 37 ───────────────────────►│  ⑨ 结束传输
  │◄── 77 ──────────────────────│
  │                               │
  │── 31 01 02 03 ──────────────►│  ⑩ 校验数据完整性
  │◄── 71 01 02 03 ─────────────│
  │── 31 03 02 03 ──────────────►│
  │◄── 71 03 02 03 [校验结果] ──│  校验通过
  │                               │
  │── 11 01 ────────────────────►│  ⑪ 硬复位ECU
  │◄── 51 01 ───────────────────│
  │                               │
  │── 10 03 ────────────────────►│  ⑫ 进入扩展会话
  │◄── 50 03 ───────────────────│
  │                               │
  │── 28 00 01 ─────────────────►│  ⑬ 恢复正常通信
  │◄── 68 00 01 ────────────────│
  │                               │
  │── 85 01 ────────────────────►│  ⑭ 启用DTC记录
  │◄── C5 01 ───────────────────│
  │                               │
  │── 22 F1 81 ─────────────────►│  ⑮ 验证软件版本
  │◄── 62 F1 81 [新版本号] ─────│  刷写成功！
  │                               │
```

---

## 六、传输层（ISO 15765-2 / ISO-TP）

### 6.1 为什么需要传输层？

CAN单帧最大8字节（CAN FD最大64字节），但诊断服务数据可能超过这个限制。ISO-TP负责**数据分包和重组**。

### 6.2 帧类型

| 类型 | PCI字节 | 说明 |
|------|---------|------|
| 单帧 (SF) | 0x0N | 数据≤7字节，一帧搞定 |
| 首帧 (FF) | 0x1N NN | 数据>7字节，第一帧，含总长度 |
| 连续帧 (CF) | 0x2N | 后续数据帧，带序号 |
| 流控帧 (FC) | 0x30 | 控制发送节奏 |

### 6.3 多帧传输示例

```
发送方（Tester）              接收方（ECU）
  │                             │
  │── [FF] 1A 22 F1 90 4C 53 56 ►│  首帧：总长26字节
  │                             │
  │◄── [FC] 30 00 14 00 ────────│  流控：连续发，间隔20ms
  │                             │
  │── [CF] 21 41 55 32 31 38 39 ►│  连续帧 #1
  │── [CF] 22 4E 32 31 32 33 34 ►│  连续帧 #2
  │── [CF] 23 35 36 ────────────►│  连续帧 #3（最后）
  │                             │
```

### 6.4 流控参数

| 参数 | 名称 | 说明 |
|------|------|------|
| BS (Block Size) | 块大小 | 每发BS个CF后等待FC，0=不限制 |
| STmin | 最小间隔时间 | 连续帧之间的最小间隔（ms） |

---

## 七、DoIP（ISO 13400）

### 7.1 概述

DoIP = Diagnostics over Internet Protocol，基于以太网的诊断，用于高速刷写和大数据量诊断。

### 7.2 优势

| 对比项 | CAN诊断 | DoIP诊断 |
|--------|---------|----------|
| 传输速率 | ~500 kbps | ~100 Mbps |
| 刷写时间 | 30-60分钟 | 3-5分钟 |
| 数据量 | 小 | 大 |
| 成本 | 低 | 较高 |

### 7.3 DoIP报文格式

```
┌──────────┬──────────┬──────────────┬──────────┬──────────┐
│ Protocol │ Protocol │  Payload     │ Source   │ Target   │
│ Version  │ Type     │  Length      │ Node ID  │ Node ID  │
│ (1 byte) │ (2 byte) │  (4 byte)    │ (2 byte) │ (2 byte) │
└──────────┴──────────┴──────────────┴──────────┴──────────┘
```

### 7.4 常用Payload Type

| Type | 名称 | 说明 |
|------|------|------|
| 0x0000 | Generic DoIP header nack | 头部否定响应 |
| 0x0001 | Vehicle identification request | 车辆识别请求 |
| 0x0004 | Vehicle announcement | 车辆公告 |
| 0x0005 | Routing activation request | 路由激活请求 |
| 0x0006 | Routing activation response | 路由激活响应 |
| 0x8001 | Diagnostic message | 诊断消息 |
| 0x8002 | Diagnostic message ack | 诊断消息确认 |

---

## 八、诊断ID配置

### 8.1 CAN诊断ID

| 类型 | ID | 说明 |
|------|-----|------|
| 请求ID（物理寻址） | 0x7E0 | Tester → ECU |
| 响应ID（物理寻址） | 0x7E8 | ECU → Tester |
| 请求ID（功能寻址） | 0x7DF | Tester → 所有ECU |

**物理寻址 vs 功能寻址**：
- 物理寻址：一对一通信，ECU必须响应
- 功能寻址：一对多广播，ECU不一定要响应

### 8.2 常见ECU诊断ID

| ECU | 请求ID | 响应ID |
|-----|--------|--------|
| 发动机 | 0x7E0 | 0x7E8 |
| 变速器 | 0x7E1 | 0x7E9 |
| ABS | 0x7E2 | 0x7EA |
| 安全气囊 | 0x7E3 | 0x7EB |
| 仪表 | 0x7E4 | 0x7EC |
| 网关 | 0x7E5 | 0x7ED |

---

## 九、实战技巧

### 9.1 诊断开发常用工具

| 工具 | 类型 | 说明 |
|------|------|------|
| CANoe | 软件 | Vector出品，行业标杆 |
| CANape | 软件 | 标定+诊断 |
| INCA | 软件 | ETAS出品，标定工具 |
| PCAN | 硬件 | Peak出品，CAN接口 |
| CANable | 硬件 | 开源CAN适配器 |
| VSPY | 软件 | Vehicle Spy，诊断+网络 |
| OBD2扫描仪 | 硬件 | 基础OBD诊断 |

### 9.2 诊断数据库文件

| 格式 | 说明 |
|------|------|
| .cdd | CANdelaStudio诊断数据库 |
| .odx | 开放诊断数据交换格式 |
| .pdx | 打包的ODX文件 |
| .arxml | AUTOSAR诊断描述 |

### 9.3 常见问题排查

| 问题 | 可能原因 | 解决方案 |
|------|----------|----------|
| 无响应 | ID错误/ECU未唤醒 | 检查ID，先唤醒网络 |
| 7F xx 12 | 子功能不支持 | 检查当前会话是否支持 |
| 7F xx 13 | 消息长度错误 | 检查请求报文长度 |
| 7F xx 22 | 条件不满足 | 检查前置条件（如车速=0） |
| 7F xx 24 | 序列错误 | 检查服务调用顺序 |
| 7F xx 33 | 安全锁定 | 先执行安全访问 |
| 7F xx 35 | Key错误 | 检查安全算法 |
| 7F xx 36 | 尝试次数过多 | 等待延迟或硬复位 |
| 7F xx 31 | 超出范围 | 检查DID/RID是否正确 |
| 超时无响应 | ECU处理中 | 等待0x78响应 |

### 9.4 suppressPosRspMsgIndicationBit

**子功能字节的最高位（Bit 7）**：
- 0：需要正响应
- 1：抑制正响应（不需要回复正响应）

```
10 03 → 进入扩展会话，需要响应
10 83 → 进入扩展会话，不需要响应（0x03 | 0x80 = 0x83）
3E 80 → 保活，不需要响应（0x00 | 0x80 = 0x80）
```

**注意**：负响应永远不能被抑制！

---

## 十、与其他标准的关系图

```
                    ┌──────────────────────┐
                    │   ISO 14229 (UDS)    │  应用层
                    │   诊断服务定义        │
                    └──────────┬───────────┘
                               │
              ┌────────────────┼────────────────┐
              │                │                │
    ┌─────────▼──────┐ ┌──────▼───────┐ ┌──────▼───────┐
    │ ISO 15765-2    │ │ ISO 13400    │ │ ISO 14230    │
    │ (ISO-TP/CAN)   │ │ (DoIP/Eth)   │ │ (K-Line)     │
    └─────────┬──────┘ └──────┬───────┘ └──────┬───────┘
              │                │                │
    ┌─────────▼──────┐ ┌──────▼───────┐ ┌──────▼───────┐
    │ ISO 11898      │ │ Ethernet     │ │ K-Line       │
    │ (CAN/CAN-FD)   │ │ (100Mbps+)   │ │ (10.4kbps)   │
    └────────────────┘ └──────────────┘ └──────────────┘

    ┌──────────────────────────────────────────────────────┐
    │              ISO 15031 (OBD)                         │
    │   排放法规诊断，基于UDS子集                            │
    │   使用固定ID（0x7DF/0x7E0-0x7E7）                    │
    └──────────────────────────────────────────────────────┘
```

---

## 附录A：SID速查表

| SID | 服务名称 | 正响应SID | 功能 |
|-----|----------|-----------|------|
| 0x10 | DiagnosticSessionControl | 0x50 | 会话控制 |
| 0x11 | ECUReset | 0x51 | ECU复位 |
| 0x14 | ClearDTC | 0x54 | 清除DTC |
| 0x19 | ReadDTCInformation | 0x59 | 读取DTC |
| 0x22 | ReadDataByIdentifier | 0x62 | 按ID读数据 |
| 0x23 | ReadMemoryByAddress | 0x63 | 按地址读内存 |
| 0x24 | ReadScalingDataByIdentifier | 0x64 | 读缩放数据 |
| 0x27 | SecurityAccess | 0x67 | 安全访问 |
| 0x28 | CommunicationControl | 0x68 | 通信控制 |
| 0x2C | DynamicallyDefineDataIdentifier | 0x6C | 动态定义DID |
| 0x2E | WriteDataByIdentifier | 0x6E | 按ID写数据 |
| 0x2F | InputOutputControlByIdentifier | 0x6F | IO控制 |
| 0x31 | RoutineControl | 0x71 | 例程控制 |
| 0x34 | RequestDownload | 0x74 | 请求下载 |
| 0x35 | RequestUpload | 0x75 | 请求上传 |
| 0x36 | TransferData | 0x76 | 数据传输 |
| 0x37 | RequestTransferExit | 0x77 | 结束传输 |
| 0x38 | RequestFileTransfer | 0x78 | 文件传输 |
| 0x3D | WriteMemoryByAddress | 0x7D | 按地址写内存 |
| 0x3E | TesterPresent | 0x7E | 心跳保活 |
| 0x85 | ControlDTCSetting | 0xC5 | 控制DTC记录 |
| 0x86 | ResponseOnEvent | 0xC6 | 事件响应 |

---

## 附录B：NRC速查表

| NRC | 名称 | 常见原因 |
|-----|------|----------|
| 0x10 | generalReject | 未知错误 |
| 0x11 | serviceNotSupported | SID错误或ECU不支持 |
| 0x12 | subFunctionNotSupported | 子功能值错误 |
| 0x13 | incorrectMessageLengthOrInvalidFormat | 请求长度不对 |
| 0x21 | busyRepeatRequest | ECU忙，稍后重试 |
| 0x22 | conditionsNotCorrect | 前置条件不满足 |
| 0x24 | requestSequenceError | 服务调用顺序错误 |
| 0x31 | requestOutOfRange | DID/RID/地址不存在 |
| 0x33 | securityAccessDenied | 未解锁就访问受保护服务 |
| 0x35 | invalidKey | Key计算错误 |
| 0x36 | exceededNumberOfAttempts | 安全解锁失败太多次 |
| 0x37 | requiredTimeDelayNotExpired | 延迟时间未到 |
| 0x72 | generalProgrammingFailure | 刷写数据校验失败 |
| 0x73 | wrongBlockSequenceCounter | TransferData序号错误 |
| 0x78 | responsePending | 请求已收到，处理中 |
| 0x7E | subFunctionNotSupportedInActiveSession | 当前会话不支持 |
| 0x7F | serviceNotSupportedInActiveSession | 当前会话不支持 |

---

> 📝 整理时间：2026-04-15  
> 📖 参考标准：ISO 14229-1:2020, ISO 15765-2, ISO 13400
