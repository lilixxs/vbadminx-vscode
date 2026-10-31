# CI/CD 说明

本仓库使用 GitHub Actions。工作流文件位于 `.github/workflows/`。

## 1. 工作流一览

| 工作流 | 文件 | 触发条件 | 动作 |
|---|---|---|---|
| CI | `ci.yml` | 每次 push（任意分支）、每个 PR | 安装依赖 → 编译 → 单元测试（117）→ VBA 用例回归 → 性能基准 → 打包 VSIX；上传 VSIX 与验证报告为产物 |
| Release | `release.yml` | push tag `v*`，或手动 `workflow_dispatch` | 编译 → 单元测试 → 打包 → 发布到 VS Code Marketplace 与 Open VSX → 创建 GitHub Release 并附加 VSIX |

CI 的第二个 job（`host`）在 Linux 上用 `xvfb` 运行 VS Code 宿主集成检查（31 项）。该 job 设为 `continue-on-error: true`，因为它依赖图形环境，失败不会阻断 CI。

## 2. 每次提交自动做什么

每次 `git push`：

1. `npm ci` → `npm run compile`
2. `npm test`（单元测试，结果写入 `output/unit-test-check.json`）
3. `npm run cases:vba`（依据 MS-VBAL 的用例回归，写入 `output/vba-cases-report.json`）
4. `npm run benchmark:lsp`（1000 文件 / 10 万行基准，写入 `output/lsp-benchmark.json`）
5. `npm run package`（生成 `output/vbadminx-vscode-<version>.vsix`）
6. 上传产物：VSIX 与 `output/*.json`、`output/*.md`

即：**每次提交都会自动构建、自动验证，并产出可下载的 VSIX 产物**。

## 3. 发布到应用市场

正式发版到市场需要 tag 与密钥：

```bash
git tag v0.1.0
git push origin v0.1.0
```

推送 tag 后 Release 工作流会：

1. 编译、测试、打包；
2. 若配置了 `VSCE_PAT` → `vsce publish --pat` 到 VS Code Marketplace；否则用 `vsce publish --oidc`（Trusted Publishing）；
3. 若配置了 `OVSX_PAT` → `ovsx publish` 到 Open VSX Registry；
4. 用内置的 `GITHUB_TOKEN` 创建 GitHub Release，并附加 VSIX。

Release 正文由工作流自动生成，固定包含：

1. **更新内容**：上一个 tag 到当前 tag 之间的提交标题（无提交时写“本次发布包含维护性更新”）；
2. **SHA-256**：VSIX 的校验值（`sha256sum` 输出，含文件名）；
3. **VSIX 安装方法**：下载附件、`从 VSIX 安装...` 与 `code --install-extension` 命令；
4. **VS Code Marketplace**：<https://marketplace.visualstudio.com/items?itemName=lilixxs.vbadminx-vscode>。

生成逻辑在 `release.yml` 的“生成发布说明”步骤，输出 `release-notes.md` 并以 `body_path` 提交给 GitHub Release。

**发布凭据的两种来源**：有 `VSCE_PAT` 用 PAT；没有则走 OIDC，此时 publisher 必须在 Marketplace 已存在且配好 Trusted Publishing，否则该步骤报 404。市场发布失败**不会**中断 Open VSX 与 GitHub Release（这两个步骤带 `if: ${{ !cancelled() }}`）。

## 4. 需要配置的 Secrets

> **重要时间点**：Azure DevOps 的全局 Personal Access Token（PAT）于 **2026-12-01 退役**。官方推荐改用
> **Trusted Publishing**（GitHub OIDC 联邦身份），不保存长期密钥。工作流已同时支持两种方式：
> 配置了 `VSCE_PAT` 就用 PAT，否则走 `vsce publish --oidc`。

### 方式 A（推荐，无长期密钥）：Trusted Publishing

1. 登录 <https://marketplace.visualstudio.com/manage>，打开 publisher `lilixxs`。
2. 选择 **Trusted Publishing → Create**，填写：
   - Policy name：`Release workflow`
   - Repository owner：`lilixxs`
   - Repository：`vbadminx-vscode`
   - Workflow：`release.yml`
   - Environment：留空（工作流未使用 environment）
3. 保存。之后打 tag 时 `release.yml` 会用 `--oidc` 换取 15 分钟有效的发布凭据（OIDC token 只能交换一次）。

无需在 GitHub 配置任何 Secret。

### 方式 B（过渡方案）：Personal Access Token

| Secret | 用途 | 获取方式 |
|---|---|---|
| `VSCE_PAT` | 发布到 VS Code Marketplace | 见下方步骤；**2026-12-01 后失效** |
| `OVSX_PAT` | 发布到 Open VSX Registry | 在 `https://open-vsx.org` 登录后，于用户设置中创建 access token |

`GITHUB_TOKEN` 由 Actions 自动提供，无需配置。

#### 申请 VSCE_PAT

1. 若没有 Azure DevOps 组织，先创建：<https://learn.microsoft.com/azure/devops/organizations/accounts/create-organization>
2. 打开 <https://dev.azure.com>，选择组织。
3. 右上角头像 → **Personal access tokens** → **New Token**。
4. 填写：
   - Name：任意，例如 `vscode-marketplace-publish`
   - **Organization：必须选 `All accessible organizations`**（选具体组织会导致 401/403）
   - Expiration：按需（到期后发布失败）
   - Scopes：**Custom defined** → 点 **Show all scopes** → 找到 **Marketplace** → 勾选 **Manage**
5. **Create**，立刻复制 token（只显示一次）。

#### 先创建 publisher（发布前必须）

1. <https://marketplace.visualstudio.com/manage> 用**同一个 Microsoft 账号**登录。
2. **Create publisher**，ID 填 `lilixxs`（必须与 `package.json` 的 `publisher` 完全一致，创建后不可修改）。
3. 本地验证：

```powershell
npx --yes @vscode/vsce login lilixxs
# 粘贴 PAT，出现 “The Personal Access Token verification succeeded” 即成功
```

#### 配置到 GitHub

仓库 **Settings → Secrets and variables → Actions → New repository secret**，名称填 `VSCE_PAT`，值粘贴 token。

> 常见错误：`401/403` 几乎都是「组织未选 All accessible organizations」或「scope 不是 Marketplace: Manage」。

## 5. 已知故障与排查

| 现象 | 原因 | 处理 |
|---|---|---|
| `error: unknown option '--oidc'` | 本地 `@vscode/vsce` 低于 4.1.0（3.x 不支持 Trusted Publishing） | 升级到 `^4.1.0`；工作流使用 `./node_modules/.bin/vsce` |
| `Marketplace OIDC token exchange failed with 404: The requested publisher doesn't exist` | Marketplace 上没有与 `package.json` 的 `publisher` 同名的 publisher | 在 <https://marketplace.visualstudio.com/manage> 创建 publisher，ID 必须与 `package.json` 的 `publisher` 完全一致 |
| `401/403`（PAT 模式） | PAT 的 Organization 未选 `All accessible organizations`，或 scope 不是 `Marketplace: Manage` | 重新创建 PAT |
| `apt-get install` 失败（CI host job） | Ubuntu 24.04 把 `libgtk-3-0`/`libasound2` 改名为 `*t64` | 只安装 `xvfb`，不要写死这些包名 |
| 发布成功但 GitHub Release 未创建 | 市场发布步骤失败会中断后续 step | 工作流已用 `if: ${{ !cancelled() }}` 保证 Release 仍创建 |

## 6. 发布前检查清单

- `package.json` 的 `version` 已按 SemVer 更新（与 tag 一致，去掉 `v` 前缀）。
- `package.json` 的 `publisher` 为 `lilixxs`，且该 publisher 在 Marketplace 已创建并配置 Trusted Publishing。
- 建议补充 `repository` 字段（指向本仓库），`vsce` 在缺少时会告警：
  ```json
  "repository": { "type": "git", "url": "https://github.com/<owner>/vbadminx-vscode" }
  ```
- `README.md` 是 Marketplace 详情页内容；`docs/` 与 `output/` 的说明不参与详情页。
- 本地先跑一遍：`npm run compile && npm test && npm run verify:package`。

## 7. 本地等价命令

| CI 步骤 | 本地命令 |
|---|---|
| 编译 | `npm run compile` |
| 单元测试 | `npm test` |
| 用例回归 | `npm run cases:vba` |
| 性能基准 | `npm run benchmark:lsp` |
| 打包 | `npm run package` |
| 交付核对 | `npm run verify:package` |
| 截图（需 Windows + VS Code） | `npm run capture` |

## 7. 已知限制

- `npm run capture`（截图）与 `npm run smoke:go` 依赖本机环境（Windows、VS Code、模型代理），**不在 CI 中运行**。
- `npm run test:host` 在 CI 的 Linux + xvfb 下运行；本机 Windows 下无需 xvfb。
- 发布工作流未做版本号与 tag 一致性校验，需人工保证（见 §5）。
