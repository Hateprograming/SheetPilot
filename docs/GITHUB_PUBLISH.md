# 创建 SheetPilot 仓库与首次推送

本指南发布的是应用目录本身。GitHub 仓库名为 **SheetPilot**；本地目录仍为 `SheetPilotApp`，不需要重命名。研究工作区中的 `DataAnalysis/`、`Test/` 和 `Guidance/` 不包含在应用仓库里。

## 哪些文件应该上传

上传 `src/`、`web/src/`、`web/public/`、测试与评估源码、文档、Docker 配置、`pyproject.toml`，以及 `web/package.json` 和 `web/package-lock.json`。

点开头不代表不能上传：`.gitignore`、`.dockerignore` 和 `.env.example` 都应上传。真实 `.env`、`.venv/`、`.pytest*/`、`.phase1/`、`.phase2/`、运行数据、截图产物和网页副本由忽略规则排除。`.git/` 是 Git 自身元数据，不需要手工上传。

项目根目录的空 `package.json` 与 `package-lock.json` 已忽略；真正的前端依赖清单位于 `web/`。忽略文件不会删除本地内容。

## 方案一：GitHub 网页创建，再用 Git 推送

适用于已经安装 Git、没有安装 GitHub CLI 的情况。

1. 登录 GitHub，点击右上角 **+ → New repository**。
2. Repository name 填 **SheetPilot**。可先选 Private，展示时再决定是否公开；网站能公开访问与源码是否公开是两件事。
3. **不要初始化 README、.gitignore 或 License**，本地已有相应配置；本项目尚未选定开源许可证。
4. 点击 Create repository，复制页面提供的 HTTPS 仓库地址。

在 PowerShell 中逐段执行，某段报错时先解决，再继续：

```powershell
Set-Location "D:\Files\实习\SheetPilot\SheetPilotApp"
git init -b main
```

如果尚未设置 Git 提交身份，使用自己的名称和邮箱，仅对此仓库设置。邮箱可以使用 GitHub 邮箱设置中的 noreply 地址：

```powershell
git config user.name "你的名字或 GitHub 用户名"
git config user.email "你的 GitHub 邮箱或 noreply 地址"
```

暂存后检查将上传的文件名：

```powershell
git add .
git status --short
git diff --cached --stat
git ls-files -- .env .env.example web/package.json web/package-lock.json
```

最后一条应只列出 `.env.example` 和 `web/` 下两份依赖清单，不应出现 `.env`。暂存列表也不应包含 `.venv/`、`data/`、`data_web/`、`node_modules/` 或 `web_副本/`。

确认后提交并推送。将下面示例中的 `YOUR_GITHUB_USERNAME` 替换为实际仓库所有者；也可以直接使用网页复制的完整地址：

```powershell
git commit -m "Initial commit: SheetPilot web application"
git remote add origin https://github.com/YOUR_GITHUB_USERNAME/SheetPilot.git
git push -u origin main
```

首次推送可能要求通过浏览器登录 GitHub。GitHub HTTPS 推送不使用账户密码；按 Git 凭据管理器的登录提示操作即可。

## 方案二：GitHub CLI 创建并推送

此方案替代方案一，不要重复创建同名仓库。需要已安装 GitHub CLI（`gh`）。

先按方案一初始化、设置提交身份、暂存并检查文件，然后执行：

```powershell
git commit -m "Initial commit: SheetPilot web application"
gh auth login
gh auth setup-git
gh repo create SheetPilot --private --source=. --remote=origin --push
```

这条 `gh repo create` 命令会创建名为 **SheetPilot** 的私有仓库、添加远端并推送当前提交。若明确希望公开源码，可将 `--private` 改为 `--public`。

## 后续更新

```powershell
Set-Location "D:\Files\实习\SheetPilot\SheetPilotApp"
git add .
git diff --cached --stat
git commit -m "Describe this update"
git push
```

首次推送成功后，云平台部署时仓库根目录就是应用根目录，**不再需要设置 `SheetPilotApp` 子目录**。

## 官方参考

- [GitHub：将本地代码添加到 GitHub](https://docs.github.com/en/migrations/importing-source-code/using-the-command-line-to-import-source-code/adding-locally-hosted-code-to-github)
- [GitHub CLI：gh repo create](https://cli.github.com/manual/gh_repo_create)
