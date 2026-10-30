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
2. 若配置了 `VSCE_PAT` → `vsce publish` 到 VS Code Marketplace；
3. 若配置了 `OVSX_PAT` → `ovsx publish` 到 Open VSX Registry；
4. 用内置的 `GITHUB_TOKEN` 创建 GitHub Release，并附加 VSIX。

**未配置密钥时**，市场发布步骤只输出警告并跳过（工作流仍成功），GitHub Release 照常创建。这样未配置密钥也不会让流水线变红。

## 4. 需要配置的 Secrets

在仓库 **Settings → Secrets and variables → Actions → New repository secret** 添加：

| Secret | 用途 | 获取方式 |
|---|---|---|
| `VSCE_PAT` | 发布到 VS Code Marketplace | Azure DevOps 个人访问令牌（Publisher `vbadminx` 所属组织），Marketplace 范围设为 **All accessible organizations / Marketplace: Manage**。创建地址：`https://dev.azure.com/<org>/_usersSettings/tokens` |
| `OVSX_PAT` | 发布到 Open VSX Registry | 在 `https://open-vsx.org` 登录后，于用户设置中创建 access token |

`GITHUB_TOKEN` 由 Actions 自动提供，无需配置。

## 5. 发布前检查清单

- `package.json` 的 `version` 已按 SemVer 更新（与 tag 一致，去掉 `v` 前缀）。
- `package.json` 的 `publisher` 为 `vbadminx`，且该 publisher 在 Marketplace 已创建。
- 建议补充 `repository` 字段（指向本仓库），`vsce` 在缺少时会告警：
  ```json
  "repository": { "type": "git", "url": "https://github.com/<owner>/vbadminx-vscode" }
  ```
- `README.md` 是 Marketplace 详情页内容；`docs/` 与 `output/` 的说明不参与详情页。
- 本地先跑一遍：`npm run compile && npm test && npm run verify:package`。

## 6. 本地等价命令

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
