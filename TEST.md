# 测试说明

## 一、环境准备

1. **安装依赖**（在项目目录下执行）：
   ```bash
   pip install -r requirements.txt
   ```

2. **iFinD 账号**（用于债券查询）  
   - 若尚未申请：打开 [同花顺数据接口](https://quantapi.51ifind.com/) 申请试用/正式账号。  
   - 若已安装过 iFinD 接口：可执行 `setup_ifind.bat` 或 `pip install iFinDAPI` 确保本机能导入 iFinD。

---

## 二、不依赖 iFinD 的测试（先做）

### 1. 检查模块是否能正常导入

```bash
python -c "from database.db_manager import DatabaseManager; from data_provider import get_data_api; print('导入成功')"
```

应输出：`导入成功`。

### 2. 测试数据导入（可选，需自备 Excel）

在 `data/` 目录下放置三个 Excel 文件（或任意路径，运行时输入）：

- 持仓：文件名包含 `holdings`（列需含：company_name, bond_code, issuer_name 等）
- 可买名单：包含 `can_buy_lists`（列需含：company_name, issuer_name）
- 联系人：包含 `contacts`（列需含：company_name, name 等）

然后执行：

```bash
python import_data.py
```

按提示选择或输入文件路径，确认导入。若成功会提示“导入完成”。

### 3. 启动图形界面（不查债券也可看界面）

```bash
python app.py
```

- 应弹出“债券买家匹配”窗口。  
- 若未配置 iFinD，会提示“iFinD 未连接”，属正常。  
- 可点击 **帮助 → iFinD 设置** 检查/填写账号密码（填好后点确定会重新连接）。

---

## 三、依赖 iFinD 的测试（需有效账号）

### 1. 配置 iFinD 账号

**方式 A：在图形界面里配置（推荐）**

1. 运行：`python app.py`  
2. 菜单 **帮助 → iFinD 设置**  
3. 填写同花顺数据接口的**账号**和**密码**，点确定。  
4. 若连接成功，后续搜索债券时会使用该账号。

**方式 B：命令行用环境变量**

Windows PowerShell：

```powershell
$env:IFIND_USER = "你的账号"
$env:IFIND_PASSWORD = "你的密码"
```

Windows CMD：

```cmd
set IFIND_USER=你的账号
set IFIND_PASSWORD=你的密码
```

（仅当前终端窗口有效。）

### 2. 命令行查询债券

```bash
python bond_lookup.py 132001.SH
```

或先运行 `python bond_lookup.py`，再按提示输入债券代码。

- 若连接成功且代码有效：会打印债券名称、发行人、行业、评级等信息。  
- 若提示“无法连接 iFinD”：检查账号密码、网络，以及是否已安装 `iFinDAPI`。  
- 若提示“未找到债券”：可能是代码格式或接口权限问题，可换一个已知存在的债券代码再试。

### 3. 在图形界面里查询债券

1. 运行：`python app.py`  
2. 确保已在 **帮助 → iFinD 设置** 中填写账号密码并确定。  
3. 在“债券代码”输入框输入代码（如 `132001.SH`），点击 **搜索** 或按回车。  
4. 若正常：会显示债券名称、发行人、潜在买家列表等；点击左侧公司可看联系人详情。

---

## 四、快速自检脚本

在项目目录下执行：

```bash
python check_env.py
```

会检查：Python 版本、依赖包、database/ifind_api 导入、iFinD 是否可用等，并在控制台输出结果（见下方脚本说明）。

---

## 五、常见问题

| 现象 | 可能原因 | 建议 |
|------|----------|------|
| 提示“无法连接 iFinD” | 未配置账号/密码、或账号未开通/过期 | 在「iFinD 设置」中填写正确账号密码，或设置环境变量 |
| 提示“No module named 'iFinDPy'” | 未安装 iFinD 接口 | 执行 `pip install iFinDAPI` 或运行 `setup_ifind.bat` |
| 搜索债券无结果 | 债券代码错误或接口无权限 | 核对代码格式（如 132001.SH），换其他债券试，或联系同花顺确认权限 |
| 导入 Excel 报错 | 文件格式或列名不符 | 检查 Excel 是否包含所需列（如 company_name, issuer_name 等） |
