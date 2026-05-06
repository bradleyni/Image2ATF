# GitHub 上传步骤（给新手）

这个目录已经按 GitHub 仓库整理好了。你只需要把它上传成一个新仓库。

## 方法 A：网页上传（最简单）

1. 登录 GitHub。
2. 点击右上角 `+` → `New repository`。
3. Repository name 填：`Image2ATF`。
4. Visibility 选 `Private` 或 `Public`。
   - 不确定就选 `Private`。
5. 创建仓库后，点击 `uploading an existing file`。
6. 把本目录里的所有文件拖进去。
7. Commit message 填：`Initial commit`。
8. 点击 `Commit changes`。

## 方法 B：命令行上传

先在 GitHub 建一个空仓库，然后在本目录打开终端，执行：

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_NAME/Image2ATF.git
git push -u origin main
```

把 `YOUR_NAME` 换成你的 GitHub 用户名。

## 本地运行

Windows 双击：

```text
install_windows.bat
run_windows.bat
```

macOS / Linux：

```bash
chmod +x install_macos_linux.sh run_macos_linux.sh
./install_macos_linux.sh
./run_macos_linux.sh
```

## 重要提醒

- 不要把 API Key 写进代码，也不要提交 `.env`、`.venv`。
- 这个仓库里的 `.gitignore` 已经帮你忽略虚拟环境、日志和常见输出文件。
- 这份整理尽量不改动 `Image2ATF_V2.8.py` 的核心逻辑，只补齐 GitHub 项目文件。
