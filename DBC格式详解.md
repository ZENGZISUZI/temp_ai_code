# DBC 文件格式详解

DBC (Database Container) 是 Vector 公司定义的 CAN 总线数据库文件格式，用于描述 CAN 网络中的报文、信号、节点等信息。本文档对 DBC 文件的每个结构进行详细剖析。

---

## 目录

1. [文件概述](#文件概述)
2. [版本信息](#版本信息)
3. [符号定义](#符号定义)
4. [消息定义](#消息定义)
5. [信号定义](#信号定义)
6. [值表定义](#值表定义)
7. [节点定义](#节点定义)
8. [注释定义](#注释定义)
9. [属性定义](#属性定义)
10. [信号组定义](#信号组定义)
11. [信号扩展类型](#信号扩展类型)
12. [完整示例](#完整示例)

---

## 文件概述

DBC 文件是纯文本格式，采用特定的语法结构。每个部分以关键字开头，以分号 `;` 结束。

**基本特点：**
- 文本格式，可读性强
- 层级结构清晰
- 支持扩展属性
- 广泛应用于汽车电子开发

---

## 版本信息

### 语法

```
VERSION "版本字符串";
```

### 说明

- 位于文件开头
- 标识 DBC 文件的版本
- 通常为空字符串或工具生成的版本号

### 示例

```
VERSION "1.0";
```

---

## 符号定义

### 语法

```
NS_:
    [符号列表]
;
```

### 说明

- 定义 DBC 文件中使用的符号关键字
- 不同工具可能支持不同的符号集
- 通常由工具自动生成

### 常见符号

| 符号 | 说明 |
|------|------|
| NS_ | 符号定义 |
| BS_ | 位时序定义（已废弃） |
| BU_ | 节点定义 |
| BO_ | 消息定义 |
| SG_ | 信号定义 |
| EV_ | 环境变量定义 |
| CM_ | 注释定义 |
| BA_DEF_ | 属性定义 |
| BA_ | 属性值 |
| VAL_ | 值表定义 |
| SIG_GROUP_ | 信号组定义 |
| SIG_VALTYPE_ | 信号值类型 |

### 示例

```
NS_:
    NS_DESC_
    CM_
    BA_DEF_
    BA_
    VAL_
    BO_
    SG_
    BU_
;
```

---

## 消息定义

### 语法

```
BO_ 消息ID 消息名称: 消息长度 发送节点
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| 消息ID | 整数 | CAN ID，十进制表示，标准帧 0-2047，扩展帧可达 29 位 |
| 消息名称 | 字符串 | 消息标识符，建议命名规范 |
| 消息长度 | 整数 | 数据场字节数，1-8 字节 |
| 发送节点 | 字符串 | 发送该消息的节点名称，无发送者用 `Vector__XXX` |

### 示例

```
BO_ 100 Engine_Status: 8 ECU_Engine
BO_ 200 Vehicle_Speed: 8 ECU_ABS
BO_ 300 Light_Control: 4 Vector__XXX
```

### 消息ID说明

- **标准帧**: ID 范围 0x000 - 0x7FF (0-2047)
- **扩展帧**: ID 范围 0x00000000 - 0x1FFFFFFF
- 扩展帧 ID 通常通过属性 `VFrameFormat` 标识

---

## 信号定义

### 语法

```
SG_ 信号名称: 起始位|位长度@字节序 (系数,偏移量) [最小值|最大值] "单位" 接收节点
```

### 参数详解

| 参数 | 类型 | 说明 |
|------|------|------|
| 信号名称 | 字符串 | 信号标识符 |
| 起始位 | 整数 | 信号在数据场中的起始位 (0-63) |
| 位长度 | 整数 | 信号占用的位数 (1-64) |
| 字节序 | 0/1 | 0=Motorola(大端)，1=Intel(小端) |
| 系数 | 浮点数 | 物理值转换系数 (scale) |
| 偏移量 | 浮点数 | 物理值转换偏移 (offset) |
| 最小值 | 浮点数 | 物理值最小范围 |
| 最大值 | 浮点数 | 物理值最大范围 |
| 单位 | 字符串 | 物理单位，如 "km/h", "rpm" |
| 接收节点 | 字符串 | 接收该信号的节点，多个用逗号分隔 |

### 字节序详解

**Motorola 格式 (字节序=0):**
- 大端序，高位在前
- 起始位为最高有效位 (MSB)
- 汽车行业常用格式

**Intel 格式 (字节序=1):**
- 小端序，低位在前
- 起始位为最低有效位 (LSB)
- 符合 x86 处理器习惯

### 位编码示例

**Motorola 格式 (起始位 7，长度 12):**

```
字节:  [0]     [1]     [2]     [3]
位:    7-0    15-8    23-16   31-24
       |----信号----|
       MSB         LSB
```

**Intel 格式 (起始位 0，长度 12):**

```
字节:  [0]     [1]     [2]     [3]
位:    7-0    15-8    23-16   31-24
       |----信号----|
       LSB         MSB
```

### 物理值转换

```
物理值 = 原始值 × 系数 + 偏移量
原始值 = (物理值 - 偏移量) / 系数
```

### 示例

```
BO_ 100 Engine_Status: 8 ECU_Engine
 SG_ Engine_RPM : 0|16@1+ (0.25,0) [0|16383.75] "rpm" ECU_Cluster
 SG_ Engine_Temp : 16|8@1+ (1,-40) [-40|215] "°C" ECU_Cluster
 SG_ Engine_Status : 24|2@0+ (1,0) [0|3] "" ECU_Cluster
```

### 信号符号类型

| 符号 | 说明 |
|------|------|
| + | 无符号整数 |
| - | 有符号整数 |

---

## 值表定义

### 语法

```
VAL_ 消息ID 信号名称 值1 "描述1" 值2 "描述2" ... ;
```

### 说明

- 定义信号值与描述的映射关系
- 用于枚举类型信号
- 值为整数，描述为字符串

### 示例

```
VAL_ 100 Engine_Status 0 "Off" 1 "Running" 2 "Error" 3 "Reserved";
VAL_ 200 Gear_Position 0 "Park" 1 "Reverse" 2 "Neutral" 3 "Drive";
```

### 应用场景

- 状态信号：开/关、模式选择
- 档位信号：P/R/N/D
- 错误码：故障类型描述

---

## 节点定义

### 语法

```
BU_: 节点1 节点2 节点3 ... ;
```

### 说明

- 定义 CAN 网络中的 ECU 节点
- 节点名称需唯一
- 用于标识消息的发送者和接收者

### 示例

```
BU_: ECU_Engine ECU_ABS ECU_Cluster ECU_Body ECU_Airbag;
```

### 节点属性

节点可以关联属性，如：
- 节点地址
- 通信速率
- 厂商信息

---

## 注释定义

### 语法

```
CM_ 对象类型 对象标识 "注释内容";
```

### 对象类型

| 类型 | 说明 | 对象标识格式 |
|------|------|-------------|
| BU_ | 节点注释 | 节点名称 |
| BO_ | 消息注释 | 消息ID |
| SG_ | 信号注释 | 消息ID 信号名称 |
| EV_ | 环境变量注释 | 环境变量名称 |

### 示例

```
CM_ BU_ ECU_Engine "发动机控制单元";
CM_ BO_ 100 "发动机状态报文，周期 100ms";
CM_ SG_ 100 Engine_RPM "发动机转速，精度 0.25 rpm";
```

### 注释规范

- 使用双引号包裹
- 支持多行注释（使用 `\n` 换行）
- 建议使用英文或 UTF-8 编码

---

## 属性定义

### 语法

```
BA_DEF_ 对象类型 "属性名称" 属性类型 [值范围];
BA_DEF_DEF_ "属性名称" 默认值;
BA_ "属性名称" 对象类型 对象标识 属性值;
```

### 对象类型

| 类型 | 说明 |
|------|------|
| (空) | 全局属性 |
| BU_ | 节点属性 |
| BO_ | 消息属性 |
| SG_ | 信号属性 |
| EV_ | 环境变量属性 |

### 属性类型

| 类型 | 语法 | 说明 |
|------|------|------|
| INT | INT 最小值 最大值 | 整数类型 |
| HEX | HEX 最小值 最大值 | 十六进制整数 |
| FLOAT | FLOAT 最小值 最大值 | 浮点数类型 |
| STRING | STRING | 字符串类型 |
| ENUM | ENUM "值1","值2",... | 枚举类型 |

### 常用预定义属性

| 属性名称 | 适用对象 | 类型 | 说明 |
|----------|----------|------|------|
| GenMsgCycleTime | BO_ | INT | 消息周期时间 (ms) |
| GenMsgSendType | BO_ | ENUM | 发送类型 (周期/事件) |
| GenSigStartValue | SG_ | FLOAT | 信号初始值 |
| VFrameFormat | BO_ | ENUM | 帧格式 (标准/扩展) |
| NmStationAddress | BU_ | HEX | 节点网络管理地址 |

### 示例

```
BA_DEF_ BO_ "GenMsgCycleTime" INT 0 65535;
BA_DEF_DEF_ "GenMsgCycleTime" 0;
BA_ "GenMsgCycleTime" BO_ 100 100;

BA_DEF_ SG_ "GenSigStartValue" FLOAT -1e6 1e6;
BA_DEF_DEF_ "GenSigStartValue" 0;
BA_ "GenSigStartValue" SG_ 100 Engine_RPM 0;
```

---

## 信号组定义

### 语法

```
SIG_GROUP_ 消息ID 组名称 重复次数 信号1 信号2 ... ;
```

### 说明

- 将多个信号组合成一组
- 用于信号复用或条件传输
- 重复次数通常为 1

### 示例

```
SIG_GROUP_ 100 Engine_Data 1 Engine_RPM Engine_Temp Engine_Status;
```

---

## 信号扩展类型

### 语法

```
SIG_VALTYPE_ 消息ID 信号名称 类型值;
```

### 类型值

| 值 | 类型 | 说明 |
|----|------|------|
| 0 | Signed/Unsigned | 整数类型（默认） |
| 1 | Float | IEEE 754 单精度浮点 |
| 2 | Double | IEEE 754 双精度浮点 |

### 示例

```
SIG_VALTYPE_ 100 Temperature_Float 1;
SIG_VALTYPE_ 100 Pressure_Double 2;
```

### 浮点信号说明

- 单精度浮点：32 位
- 双精度浮点：64 位
- 起始位和位长度需匹配浮点类型位数
- 系数通常为 1，偏移量为 0

---

## 环境变量定义

### 语法

```
EV_ 变量名称: 变量类型 [最小值|最大值] 初始值 "单位" 访问类型 节点列表;
```

### 变量类型

| 类型 | 值 | 说明 |
|------|-----|------|
| 整数 | 0 | 整数类型 |
| 浮点 | 1 | 浮点类型 |
| 字符串 | 2 | 字符串类型 |

### 访问类型

| 类型 | 说明 |
|------|------|
| DUMMY_NODE_VECTOR0 | 只读 |
| DUMMY_NODE_VECTOR1 | 读写 |
| DUMMY_NODE_VECTOR2 | 写入 |

### 示例

```
EV_ Ambient_Temp: 1 [-40|80] 25 "°C" DUMMY_NODE_VECTOR1 ECU_Cluster;
```

---

## 完整示例

以下是一个完整的 DBC 文件示例：

```
VERSION "1.0";

NS_:
    NS_DESC_
    CM_
    BA_DEF_
    BA_
    VAL_
    BO_
    SG_
    BU_
;

BS_:

BU_: ECU_Engine ECU_ABS ECU_Cluster ECU_Body;

BO_ 100 Engine_Status: 8 ECU_Engine
 SG_ Engine_RPM : 0|16@1+ (0.25,0) [0|16383.75] "rpm" ECU_Cluster
 SG_ Engine_Temp : 16|8@1+ (1,-40) [-40|215] "°C" ECU_Cluster
 SG_ Engine_Status : 24|2@0+ (1,0) [0|3] "" ECU_Cluster
 SG_ Fuel_Level : 32|8@1+ (0.4,0) [0|100] "%" ECU_Cluster
;

BO_ 200 Vehicle_Speed: 8 ECU_ABS
 SG_ Vehicle_Speed : 0|16@1+ (0.01,0) [0|655.35] "km/h" ECU_Cluster,ECU_Engine
 SG_ Wheel_Speed_FL : 16|16@1+ (0.01,0) [0|655.35] "km/h" ECU_ABS
 SG_ Wheel_Speed_FR : 32|16@1+ (0.01,0) [0|655.35] "km/h" ECU_ABS
 SG_ Wheel_Speed_RL : 48|16@1+ (0.01,0) [0|655.35] "km/h" ECU_ABS
;

BO_ 300 Light_Control: 4 ECU_Body
 SG_ Head_Light : 0|2@0+ (1,0) [0|3] "" ECU_Body
 SG_ Turn_Signal : 2|2@0+ (1,0) [0|3] "" ECU_Body
 SG_ Brake_Light : 4|1@0+ (1,0) [0|1] "" ECU_Body
;

CM_ BU_ ECU_Engine "发动机控制单元";
CM_ BU_ ECU_ABS "防抱死制动系统";
CM_ BU_ ECU_Cluster "仪表盘";
CM_ BU_ ECU_Body "车身控制模块";

CM_ BO_ 100 "发动机状态报文，周期 100ms";
CM_ BO_ 200 "车速报文，周期 50ms";
CM_ BO_ 300 "灯光控制报文，事件触发";

CM_ SG_ 100 Engine_RPM "发动机转速，分辨率 0.25 rpm";
CM_ SG_ 100 Engine_Temp "发动机冷却液温度";

BA_DEF_ BO_ "GenMsgCycleTime" INT 0 65535;
BA_DEF_DEF_ "GenMsgCycleTime" 0;
BA_ "GenMsgCycleTime" BO_ 100 100;
BA_ "GenMsgCycleTime" BO_ 200 50;

BA_DEF_ SG_ "GenSigStartValue" FLOAT -1e6 1e6;
BA_DEF_DEF_ "GenSigStartValue" 0;
BA_ "GenSigStartValue" SG_ 100 Engine_RPM 0;
BA_ "GenSigStartValue" SG_ 100 Engine_Temp 25;

VAL_ 100 Engine_Status 0 "Off" 1 "Running" 2 "Error" 3 "Reserved";
VAL_ 300 Head_Light 0 "Off" 1 "Low" 2 "High" 3 "Auto";
VAL_ 300 Turn_Signal 0 "Off" 1 "Left" 2 "Right" 3 "Hazard";
VAL_ 300 Brake_Light 0 "Off" 1 "On";
```

---

## 附录：DBC 关键字速查表

| 关键字 | 说明 |
|--------|------|
| VERSION | 版本信息 |
| NS_ | 符号定义 |
| BS_ | 位时序（已废弃） |
| BU_ | 节点定义 |
| BO_ | 消息定义 |
| SG_ | 信号定义 |
| EV_ | 环境变量定义 |
| CM_ | 注释定义 |
| BA_DEF_ | 属性定义 |
| BA_DEF_DEF_ | 属性默认值定义 |
| BA_ | 属性值赋值 |
| VAL_ | 值表定义 |
| SIG_GROUP_ | 信号组定义 |
| SIG_VALTYPE_ | 信号扩展类型定义 |

---

## 参考资料

- Vector CANdb++ Documentation
- AUTOSAR Specification of DBC File Format
- ISO 15765 (CAN Protocol)

---

*文档生成时间: 2026-03-26*
