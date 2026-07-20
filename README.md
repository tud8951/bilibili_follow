# B站关注列表导出工具

一个带 GUI 界面的 Python 工具，输入 B站 Cookie 即可获取当前账号的全部关注列表，并导出为格式化的 Excel 表格。

## 功能

- 输入 `SESSDATA` Cookie 即可登录
- 自动分页拉取全部关注列表
- 导出为 `.xlsx` 文件，包含：序号、UID、用户名、个人空间链接
- 导出的 Excel 带有表格样式（冻结首行、超链接、蓝底白字表头）

## 使用方法

### 1. 安装依赖

```bash
pip install requests openpyxl
```

### 2. 获取 Cookie

1. 浏览器打开 [bilibili.com](https://www.bilibili.com) 并登录
2. 按 `F12` 打开开发者工具
3. 进入 **Application** → **Cookies** → `bilibili.com`
4. 复制 `SESSDATA` 的值

### 3. 运行程序

```bash
python bilibili_follow_export.py
```

### 4. 操作

1. 将 `SESSDATA` 粘贴到输入框
2. 点击「获取关注列表」
3. 获取完成后点击「导出 Excel」

## 技术栈

- Python 3
- tkinter（GUI）
- requests（HTTP 请求）
- openpyxl（Excel 导出）

## 许可证

[MIT](LICENSE)
