<div align=center>

  <h1>Bilibili_follow_export</h1>
  <h2>B站关注列表导出工具</h2>

  <img src="icon.ico" width=200></img>

  <a href="https://t.me/cnbigjackson"><img src="https://img.shields.io/badge/Telegram-2CA5E0?style=for-the-badge&logo=telegram&logoColor=white"></img></a>

  <a href="https://www.bigjackson.vip"><img src="https://trendshift.io/api/badge/trendshift/repositories/11432/yearly?language=Dart"></img></a>
  <a href="https://www.bigjackson.vip" target="_blank"><img src="https://abroad.hellogithub.com/v1/widgets/recommend.svg?rid=68d824ea55ee4b07aba6fe1dd61ac939&claim_uid=J9Qu6aDd8LT1nU0"/></img></a>
  
  <p>一个带 GUI 界面的 Python 工具，输入 B站 Cookie 即可获取当前账号的全部关注列表，并导出为格式化的 Excel 表格。</p>
</div>

## 功能

- 支持在程序内弹出独立 B 站登录窗口，完成登录后自动获取 `SESSDATA`
- 可直接获取当前账号的全部关注列表
- 自动分页拉取全部关注列表
- 导出为 `.xlsx` 文件，包含：序号、UID、用户名、个人空间链接
- 导出的 Excel 带有表格样式（冻结首行、超链接、蓝底白字表头）

## 使用方法

### 1. 安装依赖

在项目目录执行：

```bash
pip install requests openpyxl playwright
playwright install chromium
```

如果系统提示未安装浏览器内核，请重新执行最后一条命令。

### 2. 运行程序

```bash
python bilibili_follow_export.py
```

### 3. 使用 EXE 版本

项目已提供 Windows 单文件程序：`dist/BilibiliFollowExport.exe`。

也可以自行重新打包：

```bash
pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name BilibiliFollowExport --icon icon.ico bilibili_follow_export.py
```

打包完成后，EXE 文件位于 `dist/BilibiliFollowExport.exe`。

EXE 中不包含 Playwright 的 Chromium 浏览器内核。首次在新电脑上使用「网页登录获取」功能时，请先安装 Python、Playwright 及浏览器内核：

```bash
pip install playwright
playwright install chromium
```

### 4. 获取 Cookie

点击右侧的「网页登录获取」按钮，程序会打开一个独立的 B 站登录窗口。用户在窗口中登录后，程序会自动读取 `SESSDATA`，填入输入框，并关闭该窗口。

如果不想用登录窗口，也可以直接把 `SESSDATA` 手动粘贴到输入框中。

### 5. 操作流程

1. 点击「网页登录获取」或手工填写 `SESSDATA`
2. 点击「获取关注列表」
3. 获取完成后点击「导出 Excel」

### 6. 手动获取 Cookie（备用方案）

如果登录窗口方式不适用，可按以下步骤手动获取：

1. 浏览器打开 [bilibili.com](https://www.bilibili.com) 并登录
2. 按 `F12` 打开开发者工具
3. 进入 **Application** → **Cookies** → `bilibili.com`
4. 复制 `SESSDATA` 的值

## 技术栈

- Python 3
- tkinter（GUI）
- requests（HTTP 请求）
- openpyxl（Excel 导出）
- playwright（B站登录窗口与 Cookie 获取）

## 许可证

[MIT](LICENSE)
