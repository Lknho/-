# 财务管理系统

> 桌面端财务管理系统，集成 RapidOCR 票据智能识别，支持发票/PDF 自动录入、凭证管理、账簿报表生成，SQLite 本地存储，单文件 exe 离线运行。

---

## 目录

- [功能特性](#功能特性)
- [技术栈](#技术栈)
- [快速开始](#快速开始)
- [项目结构](#项目结构)
- [开发指南](#开发指南)
- [测试](#测试)
- [打包发布](#打包发布)
- [贡献指南](#贡献指南)
- [许可证](#许可证)

---

## 功能特性

### 票据智能识别
- 支持图片（JPG/PNG/BMP）和 PDF 格式票据导入
- 基于 RapidOCR + ONNX Runtime 的中文 OCR 引擎，离线运行
- 自动识别发票代码、发票号码、开票日期、金额、税额、价税合计等关键字段
- OpenCV 图像预处理：去噪、倾斜矫正、二值化、透视变换
- Shapely + pyclipper 版面分析，自动定位表格与字段区域
- 识别结果人工校正界面，低置信度字段高亮提示

### 账务管理
- 记账凭证填制、审核、查询、打印
- 借贷平衡自动校验，审核后凭证冻结
- 会计科目多级管理，支持自定义科目体系
- 往来单位、部门、职员基础档案管理

### 账簿报表
- 总账、明细账、日记账、科目余额表自动生成
- 资产负债表、利润表、现金流量表
- 期末结账与反结账，结账后当期凭证锁定

### 数据安全
- SQLite 本地单文件数据库，数据完全存储在本机
- 所有金额使用 decimal 高精度十进制运算，杜绝浮点误差
- 多级自动备份策略（结账备份、退出备份、手动备份）
- 操作日志全程记录，制单/审核/记账/结账权限分离
- OCR 推理完全离线，财务数据不上传任何服务器

---

## 技术栈

本项目基于 Python 3.13 开发，主要技术组件如下：

| 类别 | 技术 | 版本 | 用途 |
|---|---|---|---|
| 语言 | Python | 3.13 | 主开发语言 |
| GUI | Tkinter / ttk | 8.6 | 桌面界面 |
| 数据库 | SQLite | 3 | 本地数据持久化 |
| OCR | RapidOCR (ONNX Runtime) | 1.2.3 | 中文票据文字检测与识别 |
| 推理引擎 | ONNX Runtime | — | CPU 端模型推理 |
| 图像处理 | OpenCV | — | 图像预处理、矫正、二值化 |
| 图像处理 | Pillow | — | 图像格式转换 |
| 数值计算 | NumPy | 2.5.2 | 矩阵运算 |
| PDF 处理 | PyMuPDF | — | PDF 页面渲染 |
| 几何计算 | Shapely | — | 版面区域几何运算 |
| 多边形裁剪 | pyclipper | — | 票据区域裁剪 |
| XML 解析 | lxml | — | 结构化文档解析 |
| 配置 | PyYAML | — | OCR 模型配置 |
| 高精度计算 | decimal | — | 金额精确计算 |
| 系统集成 | pywin32 / WMI | — | Windows 系统信息 |
| 并发 | multiprocessing | — | 批量 OCR 并行处理 |
| 打包 | PyInstaller | — | 单文件 exe 打包 |

---

## 快速开始

### 方式一：直接运行（推荐）

1. 从 [Releases](../../releases) 下载最新版 `财务管理系统.exe`
2. 双击运行，无需安装 Python 或任何依赖
3. 首次启动会自动创建数据库与目录结构

> 系统要求：Windows 10 / 11（64 位），内存 4 GB 以上

### 方式二：从源码运行

```bash
# 克隆仓库
git clone https://github.com/你的用户名/财务管理系统.git
cd 财务管理系统

# 安装依赖
pip install -r requirements.txt

# 运行主程序
python finance_app.py
```

---

## 项目结构

```
财务管理系统/
├── finance_app.py              # 主程序入口
├── requirements.txt            # Python 依赖清单
├── AGENTS.md                   # 项目协作规范
├── 技术方案.md                 # 技术方案文档
├── README.md                   # 项目说明（本文件）
├── src/                        # 源代码目录
│   ├── gui/                    # Tkinter 界面模块
│   ├── ocr/                    # OCR 识别模块
│   ├── image/                  # 图像处理模块
│   ├── finance/                # 财务业务逻辑
│   ├── db/                     # 数据库访问层
│   └── utils/                  # 工具函数
├── models/                     # OCR 模型文件
│   ├── ch_PP-OCRv3_det_infer.onnx
│   ├── ch_PP-OCRv3_rec_infer.onnx
│   └── ch_ppocr_mobile_v2.0_cls_infer.onnx
├── data/                       # 运行时数据（数据库、票据图像）
├── tests/                      # 测试用例
└── docs/                       # 文档
```

---

## 开发指南

### 环境搭建

```bash
# 创建虚拟环境
python -m venv venv
venv\Scripts\activate

# 安装开发依赖
pip install -r requirements.txt
pip install pytest pyinstaller
```

### 编码规范

- 所有金额运算必须使用 `decimal.Decimal`，禁止使用 `float`
- 数据库操作使用参数化查询，防止 SQL 注入
- 凭证保存前必须校验借贷平衡
- 每次改动后必须编写或更新相关测试
- 每次改动完成后必须创建对应的 Git commit

### OCR 模型配置

模型配置文件位于 `models/config.yaml`，可调整以下参数：

```yaml
det:
  model_path: models/ch_PP-OCRv3_det_infer.onnx
  threshold: 0.3
rec:
  model_path: models/ch_PP-OCRv3_rec_infer.onnx
  threshold: 0.5
cls:
  model_path: models/ch_ppocr_mobile_v2.0_cls_infer.onnx
```

---

## 测试

```bash
# 运行全部测试
pytest tests/ -v

# 运行文档完整性验证
python test_技术方案.py
```

测试覆盖：
- 金额计算精度（decimal）
- 凭证借贷平衡校验
- 数据库 CRUD 与事务
- OCR 样本集回归测试
- 备份与恢复
- 文档完整性

---

## 打包发布

使用 PyInstaller 打包为单文件 exe：

```bash
pyinstaller --onefile --windowed \
  --name "财务管理系统" \
  --add-data "models;models" \
  --add-data "src;src" \
  --hidden-import=PIL._tkinter_finder \
  finance_app.py
```

打包产物位于 `dist/财务管理系统.exe`。

---

## 贡献指南

欢迎提交 Issue 和 Pull Request。提交代码前请确保：

1. 所有测试通过
2. 代码符合项目编码规范
3. 已创建对应的 Git commit
4. PR 描述清晰说明改动内容与原因

---

## 许可证

本项目采用 [MIT License](LICENSE) 开源协议。

---

## 免责声明

本软件仅供学习与个人财务管理使用，不构成任何专业财务或税务建议。正式财务核算请以专业财务软件和会计师意见为准。
