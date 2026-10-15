<!-- 本文件由 `python scripts/docx_scan.py --list-protections` 生成，不要手改。 -->

# 文本级保护区（掩码）家族清单

这些 token 里的点号与标点是语义的一部分，机械层（I/S/P）**既不修改也不报告**。
规则在 `scripts/protect.py`；为什么这么做见 `references/docx-protection.md` 第六节。

| ID | 名称 | 覆盖什么 |
| --- | --- | --- |
| `P-URL` | URL | 带协议的网址 |
| `P-WWW` | www 地址 | 没有协议头的网址 |
| `P-MAIL` | 电子邮箱 | 邮箱地址里的点和 @ 都是语义 |
| `P-DOMAIN` | 域名 / 主机 | 域名、主机名、带端口或路径的地址（也顺带保护 os.path 这类点号名） |
| `P-WIN` | Windows 路径 | 盘符路径：C:\Users\Administrator\.pi\agent\skills\（中文目录名也一起保护） |
| `P-UNC` | UNC 网络路径 | \\server\share 形式的路径 |
| `P-POSIX` | POSIX 绝对路径 | /usr/local/bin、/etc/hosts |
| `P-REL` | 相对 / 家目录路径 | ./x、../y、~/.pi/agent 形式的路径 |
| `P-SLASH` | 含斜杠的 token | tests/、a/b/c、and/or、n/a 这种一段带一段的写法 |
| `P-DOTNAME` | 点号开头的名字 | .github、.gitignore、.pi、os.path、git.exe 这类「点+字母」整体保护 |
| `P-FILEEXT` | 带扩展名的文件名 | run-all.ps1、ci.yml、data.7z（含数字开头的扩展名） |
| `P-CODE` | 反引号代码段 | `…` 之间的内容（含反引号）一个字都不许动 |
| `P-FLAG` | 命令行参数 | --flag、-o；前面必须是分隔符，且不能是连字符，避免误吞 A--B |
| `P-ENV` | 环境变量 | $HOME、${VAR}、%USERPROFILE% |
| `P-IDENT` | 下划线标识符 | snake_case 形式的代码标识符 |
| `P-CAMEL` | 驼峰标识符 | getElementById、parseXML 这类驼峰命名 |
| `P-TASKID` | 编号类 token | JIRA-1234、UTF-8、ISO-9001 |
| `P-KEYCOLON` | 配置行的键 | 段落以 key: 开头时保护那个半角冒号（避免被 P01 转成全角） |
| `P-NUM` | 小数 / 版本号 | 1.5、v0.1.0（也保护全角点，因为没有规则会去修它） |
| `P-THOUSAND` | 千分位数字 | 1,234,567 |
| `P-TIME` | 时间 | 只保护半角冒号；全角 `9：30` 要留给 P08 去修 |
| `P-DATE` | 日期 | 2026-10-13、2026/10/13 |
| `P-UNIT` | 带单位的数值 | 10 kg、100 MB（不含 % 与 ℃ —— 那两个要留给 S06 修） |
| `P-ENTITY` | HTML 实体 | &nbsp;、&#160;、&#xA0; |

## 刻意**不**保护的东西（否则会挡住该做的修复）

| 不保护 | 因为 |
| --- | --- |
| 半角 `%`、`℃`、`°` 前的空格 | `50 %` → `50%` 是 `S06` 要修的目标 |
| 全角 `：`（如 `9：30`） | `9：30` → `9:30` 是 `P08` 要修的目标，所以时间掩码只匹配半角冒号 |
| 全角数字 `１２３` | → `123` 是 `P07` 要修的目标 |
| 中文语境里的半角 `,` `.` | 转全角是 `P01` 要修的目标 |
| 字母间的双连字符 `A--B` | `--` → `——` 是 `P04` 要修的目标，所以命令行掩码要求 `--` 前是分隔符 |
| 全角波浪号 `～`（如 `3～5`） | → `-` 是 `P09` 要修的目标 |
| 紧跟路径之后、且其后没有分隔符的中文串 | 保护它会把那句正文里的问题一起吞掉；实测 `放在C:\temp下` 只保护 `C:\temp` |

> 这些都有反向测试守着（`tests/selftest.py` 的 T8）：谁要是把掩码写宽了，测试立刻变红。
