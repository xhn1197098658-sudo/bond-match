# 债券买家匹配 (Bond Buyer Match)

基于 **iFinD（同花顺）** 数据源的债券信息查询与买家匹配工具。

---

## 一、环境要求

- **Python 3.8+**（推荐使用 Anaconda）
- **依赖**：PyQt5、pandas、openpyxl（见 `requirements.txt`）
- **iFinD**：同花顺数据接口，需单独安装并配置账号

---

## 二、安装与配置

### 1. 安装依赖

**方式 A（推荐）**：双击运行  
- `install_requirements.bat`（自动检测 Anaconda 并安装依赖）

**方式 B**：命令行（已激活 conda 或 Python 在 PATH 时）  
```powershell
cd 项目目录
pip install -r requirements.txt
```

### 2. 安装 iFinD 接口

- 双击运行 `setup_ifind.bat`，或执行：`pip install iFinDAPI`
- 账号申请与文档：[同花顺数据接口](https://quantapi.51ifind.com/)

### 3. 配置 iFinD 账号

- 启动程序后，打开 **帮助 → iFinD 设置**
- 填写同花顺数据接口的**账号**和**密码**并确定  
- 命令行使用时也可设置环境变量：`IFIND_USER`、`IFIND_PASSWORD`

### 4. 环境自检（可选）

```powershell
python check_env.py
```

可检查 Python、依赖、数据库、iFinD 是否正常。

---

## 三、运行程序

- **图形界面**：双击 `run_app.bat`，或执行 `python app.py`
- **命令行查询**：`python bond_lookup.py 010107.SH`

### 打包为单 exe

如需分发给无 Python 环境的电脑，可打包为单个 exe：

1. 安装 PyInstaller：`pip install pyinstaller`
2. 双击运行 `build_exe.bat`，或执行：`pyinstaller --clean bond_match.spec`
3. 生成文件位于 `dist/债券买家匹配.exe`，数据库 `bond_buyer_match.db` 将保存在 exe 同目录

---

## 四、数据与「潜在买家」

- **潜在买家**来自本地数据库的「可买名单」和「持仓」，需先导入 Excel 才有数据。
- **数据流**：在 **data/** 目录放置 Excel → 程序内 **帮助 → 更新数据** 选择 data 所在文件夹 → 导入到 **database/** 下的 SQLite 数据库。

### Excel 要求（放在 `data/` 或任意文件夹，导入时选择）

| 文件       | 说明     | 主要列（示例） |
|------------|----------|----------------|
| 持仓       | 机构/基金持仓 | company_name, fund_name, issuer_name, bond_code, bond_name, amount |
| 可买名单   | 机构可买发行人 | company_name, issuer_name |
| 联系人     | 机构联系人   | company_name, name, position, email, phone |

- 文件名需包含 `holdings`、`can_buy_lists`、`contacts`（或使用 `sample_holdings.xlsx` 等）。
- 若 **data/** 为空，可先生成示例数据：  
  `python create_sample_data.py`  
  再在程序中「更新数据」选择 `data` 文件夹，导入后搜索 **010107.SH** 可看到示例潜在买家。

### 查看当前数据库内容

```powershell
python list_db_contents.py
```

可查看各表条数，确认是否已导入可买名单、持仓等。

---

## 五、项目结构

```
bond-match/
├── app.py                  # 主程序（PyQt5 界面）
├── bond_lookup.py          # 命令行债券查询
├── check_env.py            # 环境自检
├── create_sample_data.py   # 生成 data/ 下示例 Excel
├── list_db_contents.py     # 查看数据库各表条数
├── data_provider.py        # 数据 API 入口（iFinD）
├── ifind_api.py            # iFinD 数据接口
├── import_data.py          # 命令行 Excel 导入
├── test_ifind_login.py     # 命令行测试 iFinD 登录
├── database/
│   ├── db_manager.py       # 数据库管理
│   ├── schema.sql          # 表结构
│   └── *.db                # SQLite（导入后生成，不提交）
├── data/                   # 放置待导入的 Excel（持仓/可买名单/联系人）
├── run_app.bat             # 一键启动主程序
├── install_requirements.bat # 一键安装依赖
├── setup_ifind.bat         # 安装 iFinD 接口
├── requirements.txt
└── README.md
```

---

## 六、常见问题

- **iFinD 未连接**：在「帮助 → iFinD 设置」填写账号密码；账号须为 quantapi.51ifind.com 的**数据接口**账号。
- **登录返回值 -201**：表示重复登录，已按成功处理，可正常使用。
- **无法找到债券信息**：确认 iFinD 已连接；若提示 "params are invalid"，程序会尝试多组参数；银行间债券(.IB) 可能需在 iFinD 中用不同代码或接口。
- **潜在买家为空**：先通过「帮助 → 更新数据」导入持仓、可买名单 Excel；或运行 `create_sample_data.py` 生成示例后导入。
- **conda 在 Cursor 终端不可用**：在终端执行 `conda init powershell` 后重启 Cursor；或使用「Anaconda Prompt」运行命令。

---

## 七、开发与贡献

- 原仓库：Fork 后可在 GitHub 提 Pull Request。
- 同步上游：`git fetch upstream` → `git merge upstream/main` → `git push`。
