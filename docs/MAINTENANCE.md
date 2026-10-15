# 维护说明

给改代码的人看的东西。用户向的说明在根目录 [README.md](../README.md)。

## 文件地图

```text
SKILL.md                        边界 · 五条铁律 · 导航表 · 核查规范要点（≤200 行，超了就该下沉）
workflows/
  audit-workflow.md             「审计」5 阶段，只读
  apply-workflow.md             「清理」5 阶段，含校验门与回滚
references/
  taxonomy.md                   【核查规范全文】八层逐条判据 / 反例 / 误伤风险
  rule-index.md                 规则目录（自动生成，勿手改）
  docx-protection.md            保护区：哪些字节 / 哪些字符序列绝对不能碰（OOXML + 掩码两层）
  protection-index.md           文本级保护区（掩码）家族清单（自动生成，勿手改）
  report-format.md              报告契约：report.md / report.json / .applied.json + 退出码
  verification.md               【最后的校验规则】G1–G7、自检 T1–T8、不变量清单
scripts/
  docx_text.py                  共享层：段落/run 字符映射、保护区切分、字节级回写
  protect.py                    文本级保护区（24 个掩码家族）—— 单一真理源
  rules_core.py                 I/S/P 三层规则（可落盘）+ Issue 结构 + 定点应用
  rules_report.py               M/B/F/L/C 五层规则（只出建议）
  docx_scan.py                  审计器（只读）→ report.md + report.json
  docx_apply.py                 机械层替换（默认 dry-run；落盘前自动校验）
  docx_guard.py                 校验器（七项，可独立调用）
assets/lexicon-zh.txt           中文 AI 味词表（`词|类别|建议`，UTF-8，直接加词）
assets/readme/                  README 用的 SVG（hero.svg / gate.svg）
tests/make_sample.py            生成「AI 痕迹大全」脏样本（黄金样本）
tests/selftest.py               回归闸门
tests/word_roundtrip.py         真 Microsoft Word 往返验证（COM）
tests/word_make.ps1 / word_read.ps1  上面那个的 COM 助手（纯 ASCII）
tests/run-all.ps1               一键跑两个测试（已处理控制台编码）
```

零依赖：主链路只用 Python 3 标准库（`zipfile` / `re` / `dataclasses` / `unicodedata` / `xml.etree`）。
`python-docx` 只在 `selftest.py` 里当**第三方交叉验证器**用，找不到就跳过，不影响结论。

## 设计取舍（为什么这么做）

1. **不用 ElementTree 重新序列化，只在 XML 字节流上做手术。**
   代价是要自己处理转义与 `xml:space`；收益是能给出「除 N 处替换外逐字节未变」的**证明**。
   见 [docx-protection.md](../references/docx-protection.md) 第四节。

2. **段落文字必须映射回 run。** Word 会把一句话按格式与编辑历史切成任意多个 `<w:r><w:t>`，
   所以「共 3 个」里的空格可能**独自占一个 run**。不建字符映射，跨 run 的空格永远删不掉。

3. **不可见元素变成哨兵 `\t` / `\n`。** `<w:tab/>`、`<w:br/>` 不是文字但占排版位；
   规则不许跨越哨兵，否则「上一行末尾 + 下一行开头」会被当成一句话误伤。

4. **保护区用 flag 标记，不靠「我记得跳过」。** 域结果、修订插入各自带 flag，
   `Issue.auto` 的计算显式包含它们 —— 想让某类位置被动到，必须改代码，不能靠疏忽。

5. **哪些能自动、哪些不能，是硬编码的分层线（`AUTO_LAYERS = ("I", "S", "P")`）。**
   这不是配置项。加一条会改版式的「自动规则」，等于把整套安全论证作废。

6. **定点迭代。** `P07`（全角数字转半角）会造出新的 `S01` 命中机会，
   所以一次替换不够，要迭代到没有可改的为止（上限 8 轮，检测到状态往复立即中止）。

7. **护栏用两个互补的检查。** G3（逐字节重放）证明「说到的都做到了」，
   G4（token 集合）证明「计划本身没吃掉事实」。任一单独都不够，
   见 [verification.md](../references/verification.md) 第二节。

8. **覆盖率门写进测试。** `selftest.py` 要求每条规则都有样本命中 ——
   加规则不加样本 → 构建失败。这是防「规则悄悄失效」的唯一可靠办法。

9. **规则目录自动生成**（`--rules-out`），避免文档与代码跑偏。

10. **报告里的命中处用 `« »` 而不是 `⟪ ⟫`。** 后者（U+27EA/27EB）在等宽字体里普遍缺字形，
    Consolas 会回退到中日韩字体把它画成全宽括号，读起来像书名号。`«»` 是 Latin-1，任何字体都有。
    看不见的字符（零宽、NBSP、SHY…）一律换成人能读的名字，否则「命中了」但屏幕上什么都没有。

## 加一条保护掩码（比加规则更常见）

发现「某个 token 不该被任何规则碰」时，**加掩码，不要加规则**。

```bash
# 1) 在 scripts/protect.py 的 FAMILIES 里加一项：(id, 名称, 覆盖什么, 正则)
# 2) 必须同时在 tests/selftest.py 的两个列表里各补一个用例：
#    · MUST_MASK      —— 这个 token 必须整体落在掩码里
#    · MUST_NOT_MASK  —— 确认没顺手把该修的也保护进去
# 3) 重生成清单 + 跑自检
python scripts/docx_scan.py --protections-out references/protection-index.md
python tests/selftest.py
# 4) 在 references/docx-protection.md 第五节的家族表里补一行说明
```

**加掩码的两个禁忌**（T8 会当场抓住）：

1. **不许挡住该做的修复**。反例：给 `%` 加掩码 → `50 %` 永远修不了；给全角 `：` 加掩码 → `9：30` 永远修不了。
   所以时间/单位/数字类掩码**只匹配半角字符**。
2. **不许吞掉紧跟路径的中文正文**。掩码里的路径尾部不能以「后面没分隔符的中文串」结尾，
   否则 `放在C:\temp下 共 3 个` 这种句子里，那句正文的 `共 3` 会静默漏报。
   现在的写法只能到 `C:\temp`，`下` 留在掩码外 —— 这是故意的。

掩码基本上只影响机械层（I/S/P）的「改不改」，对 M/B/F/L/C 的「报不报」没影响：
代码段里的 `` `x` `` 依然会被 M05 指出来。**唯一例外是 `M11`（转义符残留）** ——
它是唯一一条看掩码的报告层规则，因为 `C:\Users\x\.pi` 里的 `\.` 不是 Markdown 转义。

## 加一条规则：四步

```bash
# 1) 在 scripts/rules_core.py 的 RULES 里加一项（机械层）或 scripts/rules_report.py（报告层）
#    机械层规则函数签名： fn(text, ctx) -> [(start, end, repl, note)]
#    报告层：加进 M_PATTERNS / L_STRUCT / C_PATTERNS，或写计数函数并登记进 REPORT_RULE_META
# 2) 在 tests/make_sample.py 里加一条能命中它的脏数据（否则 selftest 的覆盖率门会红）
# 3) 重生成目录并跑自检
python scripts/docx_scan.py --rules-out references/rule-index.md
python tests/selftest.py
# 4) 在 references/taxonomy.md 里补它的 判据 / 反例 / 误伤风险（三样缺一不可）
```

**新规则进机械层的准入条件**（四条全满足，缺一条就放报告层）：

1. 是**字符级替换**，不是格式操作；
2. 换完之后**语义 token 集合不变**；
3. 有明确的**反例**，且在反例上不会命中；
4. 万一判断错了，最坏后果只是「多删/多换了一个标点或空格」。

## 自己人踩过的坑

| 坑 | 表现 | 处理 |
| --- | --- | --- |
| Windows 上 `>` 重定向会二次编码 | `python x.py > a.txt` 之后文件里中文变成 **GBK** 而不是 UTF-8（用 UTF-8 读会失败）；输出真 emoji 时 Python 直接报 `UnicodeEncodeError: 'gbk' codec can't encode character` | **别用 shell 重定向存中文输出。** 可靠做法：让脚本自己用 `open(..., encoding='utf-8')` 写文件（本仓库的 `--rules-out PATH` 就是这样）；或用 PowerShell 的 `[IO.File]::WriteAllText($p, $s, [Text.UTF8Encoding]::new($false))`。实测：**先设 `[Console]::OutputEncoding=[Text.Encoding]::UTF8` 并不能修好它**（子进程写重定向句柄时用的是 ANSI 代码页 cp936） |
| PowerShell 控制台默认 GBK | 直接打印中文报 `UnicodeEncodeError` | 所有 CLI 入口都调 `docx_text.u()` 把 stdout/stderr 重设为 UTF-8 |
| PowerShell 里反引号是转义符 | 用 `python -c "...`code`..."` 传含反引号的代码 → 代码被吃掉 | **别用 `python -c` 传含反引号的代码**，写成脚本文件再跑 |
| `.ps1` 里写中文注释 | Windows PowerShell 5.1 按 ANSI 读，报莫名其妙的假错 | 存成 UTF-8 **with BOM**，或用 `pwsh`；本仓库的 `tests/*.ps1` 一律**纯 ASCII** 以绕过这个坑 |
| 段落号 `p0001` 在不同部件里重号 | 报告定位串了（页眉的 p0001 盖掉了正文的 p0001） | 唯一键是 `(part, pid)`，`para_map` 用元组做键 |
| `core.autocrlf=true` 的开发机 | .svg / .md 出现 CRLF 噪声 | 仓库根有 `.gitattributes`，统一 `eol=lf` |

## 跑测试

```powershell
powershell -File tests/run-all.ps1      # 两个测试一起跑（已处理控制台编码）
python tests/selftest.py                # 只要黄金样本回归
python tests/word_roundtrip.py          # 只要真 Word 往返（需装 Word）
```

`selftest.py` 里的 T1 覆盖率门是这套东西能长期活下去的关键：它把「加了一条规则却没人验证它到底会不会命中」
变成一次构建失败。T8 是它的镜像：它保证「加了一条保护却没人验证它到底有没有生效」也变不成静默漏报。`word_roundtrip.py` 则用 Word COM 生成真 `.docx`，清理后再让 Word 读回来比对文本 ——
这是「Word 打得开、没触发修复」的唯一可靠证据，手工拼的最小 docx 证明不了。

## 改 README 视觉层

`assets/readme/*.svg` 手写，遵循 [beautify-github-readme](https://github.com/oil-oil/beautify-github-readme) 的约束：
1200 单位 `viewBox`、系统字体、不引用远程资源、自带背景（明暗两色都要能看）、
关键文字 ≥20 单位（在 900px 渲染宽度下约 15px）。
改完跑一次结构审计，并在 900px 与 360px 两个宽度下都看一眼：

```bash
python <beautify-github-readme>/scripts/audit_readme.py README.md
```
