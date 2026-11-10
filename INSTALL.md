# paper-to-slides — 新项目接入指南

这个 zip 是 paper-to-slides skill 的**完整安装包**：8 个 Python 脚本、5 份 references 规范、
2 份文档，以及**内置默认模板** `assets/_default/template.pptx`（AlphaFold 母板已归一化版本）。

把它装到新项目后，按下方步骤准备依赖即可开始使用。**项目自带模板时可覆盖默认模板**（见第 3 节）。

---

## 1. 解压

把 zip 解到项目的 **skill 安装目录**（建议放 `.workbuddy/skills/` 下）：

```bash
# 工作区根目录
mkdir -p .workbuddy/skills
unzip paper-to-slides.zip -d .workbuddy/skills/paper-to-slides
```

解压后应有如下结构：

```
.workbuddy/skills/paper-to-slides/
├── SKILL.md                                   ← 主入口：四步决策顺序 + 阶段 0–5 + 双重验收
├── README.md                                  ← 功能概览 + 可个性化清单
├── INSTALL.md                                 ← 本文件
├── assets/
│   ├── README.md                              ← 模板放置规范
│   └── _default/
│       └── template.pptx                      ← ★ 内置默认模板（AlphaFold 母板已归一化版）
├── references/
│   ├── claim-evidence-spec.md                 ← ★ 论证审查：证据四级 + 主张—证据矩阵 + 结论边界
│   ├── script-spec.md                         ← 讲稿：逐页五要素 + 答疑清单 + 时长换算
│   ├── figure-and-author.md
│   ├── layout-spec.md
│   ├── outline-spec.md
│   ├── template-audit.md
│   └── typography-spec.md                     ← ★ 内容文字层级规范（4 档，26/22/20/14 pt）
└── scripts/
    ├── crop_figure.py
    ├── extract_pdf_assets.py
    ├── inspect_template.py
    ├── measure_deck.py
    ├── normalize_template.py                  ← ★ 母版字体归一化 CLI（Calibri + 黑体）
    ├── plan_layout.py
    ├── theme_override.py
    └── verify_deck.py
```

---

## 2. 装依赖

skill 用到的 Python 包（**3 个**，无新增依赖）：

| 包 | 用途 | 最低版本 |
|----|------|----------|
| `python-pptx` | 读写 .pptx | 1.0 |
| `lxml` | OOXML XML 树操作 | 5.0 |
| `PyMuPDF` (import 时叫 `fitz`) | PDF 嵌入图提取 + 子图裁剪渲染 | 1.24 |

装在 venv 里即可：

```bash
python -m venv .venv
. .venv/Scripts/activate   # Windows
# 或 source .venv/bin/activate  # macOS / Linux
pip install python-pptx lxml PyMuPDF
```

> 不要 `pip install -g`。WorkBuddy 的 Python 是 managed 版（隔离目录），
> 在 managed 路径外跑命令会污染全局 site-packages。

---

## 3. 模板查找（自动 fallback，零配置开箱即用）

SKILL.md 阶段 0 的查找顺序：

1. `assets/template.pptx`（**项目自有模板**，放这里会优先用）
2. `assets/_default/template.pptx`（**skill 内置默认模板**，随 zip 自带）
3. 都没有 → 向用户索取

**新项目最简用法**：什么都不做，直接进入 SKILL.md 阶段 0，skill 自动用内置默认模板。

**项目自有模板**：把模板放到 `assets/template.pptx` 即可覆盖默认：

```bash
cp /path/to/your/template.pptx .workbuddy/skills/paper-to-slides/assets/template.pptx
```

随后跑一遍自检（**这一步必做**，模板若不合规会在阶段 0 告警）：

```bash
python .workbuddy/skills/paper-to-slides/scripts/inspect_template.py \
    .workbuddy/skills/paper-to-slides/assets/template.pptx \
    -o template_profile.json
cat template_profile.json | python -m json.tool | head -40
```

告警分三档处理（详见 SKILL.md 阶段 0）：

- **致命类**（空壳模板、版式 `custom_design=false`、缺页型）→ 必须先修模板
- **重要类**（配色走 schemeClr、占位符硬编码非主题字体）→ **skill 已自动补偿**：阶段 4 会跑三步归一化链（rename_theme → apply_typography_normalization → apply_theme_override），不阻断生成
- **卫生类**（字体不一致、内嵌源码、硬编码页码）→ 可先确认再修

---

## 4. 端到端冒烟（推荐先跑一遍）

最简验证：用内置默认模板 + 任意一篇 PDF，按 SKILL.md 阶段 0–5 跑一遍。
最快的验证路径是**只跑阶段 0 + 阶段 4 的模板归一化**，确认模板走通：

```python
from pptx import Presentation
from theme_override import (
    rename_theme, apply_typography_normalization, apply_theme_override,
    PAPER_MINIMAL,
)

prs = Presentation("assets/_default/template.pptx")  # 或 assets/template.pptx
rename_theme(prs, "paper-minimalist")
apply_typography_normalization(prs, {
    "title_major_latin": "+mj-lt", "title_major_ea": "+mj-ea",
    "body_minor_latin":  "+mn-lt", "body_minor_ea":  "+mn-ea",
})
apply_theme_override(prs, PAPER_MINIMAL)
prs.save("/tmp/normalized.pptx")
print("OK")
```

然后把 `/tmp/normalized.pptx` 用 PowerPoint 打开，确认：
- 主题名显示为 `paper-minimalist`
- 所有文字字体统一（layout 里没有 `Microsoft YaHei` 等硬编码残留）
- accent1 渲染色为 `#0B0F1A`（近黑）

生成完第一版 deck 后，再跑一次内容合规门禁（**字号档位检查在里面**）：

```bash
python .workbuddy/skills/paper-to-slides/scripts/verify_deck.py output/deck.pptx \
    --profile template_profile.json --pages 9 --exempt 1,2,6,9
```

`[type scale]` 段逐页列出用到的字号与对应级别。正常应形如：

```
   slide3: 14pt, 20pt, 26pt  -> L1+L3+L4
   slide8: 11.5pt, 13pt       -> -  FAIL [below 14pt minimum: 11.5pt, 13pt]
```

出现 FAIL（小于 14pt / 档外字号 / 单页 >3 级 / 层级倒挂）时，
按 `references/typography-spec.md` 把文字重新定级——**不要**靠调字号去凑版面。

同一份报告里的 `[copy style]` 段检查文案符号：出现 `·` 这类装饰性分隔符会 FAIL，
把分隔改成空格即可（默认禁用 `· ・ ∙ •`，用 `--banned-symbols ""` 可关闭）。

---

## 5. 完整流程

跑通冒烟后，按 SKILL.md 阶段 0–5 走完整流程：

1. **阶段 0** — 汇报任务（时长/听众/目的/语言/深度取舍）+ 模板探查（见上文自检命令）
2. **阶段 1** — 精读与登记（Read PDF，产出 analysis.md：论文类型 + Figure Inventory + 原文位置）
3. **阶段 2** — 论证审查（产出 claims.md：结论等级 L1–L4 + 替代解释 + 结论边界）
4. **阶段 3** — 讲述主线与大纲（按 references/outline-spec.md，**必须用户确认后再进阶段 4**）
5. **阶段 4** — 生成 PPTX 与讲稿（三步归一化链 + 子图裁剪 + plan_layout + 按字号阶梯写文字）
6. **阶段 5** — 双重验收（科学质量表 + measure_deck 溢出 + verify_deck 内容/字号门禁）

每阶段的输出位置与判据在 SKILL.md 里都有。

---

## 6. 出问题怎么办

| 现象 | 看哪里 |
|------|--------|
| 不知道版式怎么命名算 cover/content/section-divider | SKILL.md 阶段 0 判定要点 + references/template-audit.md |
| 子图放不下（`ok: false`）/ 图多到排不下 / 图占比不足 | references/layout-spec.md「五、异常处理」 |
| 模板审查告警看不懂 | references/template-audit.md 判据表 + 单位换算 + 判读陷阱 |
| 大纲格式不对 | references/outline-spec.md（含完整结构 + 写法示例 + 常见错误） |
| 字号不知道用多大 / 出现十几种字号 / 组标题比正文小 | references/typography-spec.md（4 档阶梯 + 层级判定流程 + 14pt 下限） |
| 结论说得太满 / 被问"这能说明因果吗" | references/claim-evidence-spec.md（L1–L4 + Confidence boundary） |
| 哪些图该上页、哪些留大纲 | references/figure-and-author.md「三级深度」 |
| 作者简介怎么写 | references/figure-and-author.md「作者调研」节（是否成页见 SKILL.md 3.8） |
| 讲稿怎么写才不像念 PPT | references/script-spec.md §2 逐页五要素 |
| 图下没有图题 / 图题被算进 30 词 / 底部图注带抢走「正文块」导致倒挂漏报 | references/layout-spec.md「子图标签与图题」；图题命名 `cap-<label>`、底部图注带命名 `cap-band`，坐标取 `cap_x/cap_y/cap_w/cap_h` |
| 底部图注带的内容像流水账、抓不到理解重难点 | SKILL.md 3.6「瓶颈三问」：先回答"听众读这页图卡在哪"，从读图钥匙 / 逻辑前提 / 易误读点里选最实质的一条写；每行过删除测试，不卡就删 |
| 想让一组内容看起来是"一个模块" | references/typography-spec.md「模块块」：无边框填充 + 白字，每页 ≤3 个，命名 `Mod-*` |
| 白字压在浅色底上读不清 | 同上「对比度硬门槛」：填充色 `L ≤ 0.175`，否则 `darkened(rgb, 0.35)` |
| 内置默认模板用腻了想换 | 把自己设计的 `your_template.pptx` 拷成 `assets/template.pptx`，SKILL.md 阶段 0 优先用它 |

---

## 7. 版本

> 本节按时间倒序记录。**当前生效的数值以上方最新条目为准**；
> 更早条目里的 20 词 / 28 词是各版本当时的口径，不是现状。

本次打包时间：2026-10-29（**v10.5**）

### v10.5 —— 作者页标题固定为 `Authors`

**背景**：作者页原先只规定"封面不出现角色括号说明、身份说明只放作者页"，
但没规定**这一页自己的标题写什么**，生成端于是自由发挥——写论文题目、写作者姓名、
或写成 "Who are they" 这类疑问句。作者页不是论证，标题一旦写成问题或题目，
就和内容页的标题句式撞车，听众会以为这页要讲科学问题。

**改法**：

- **页标题固定为 `Authors`**。这是全 deck 唯一的固定标签式标题；其余内容页标题仍是
  科学问题 / 结论句式。不写论文题目、姓名、疑问句，**不加副标题**；中文汇报同写 `Authors`，
  与 `Introduction` / `Methods` / `Discussion` 共用一套节标签语汇。位置沿用大纲
  （封面之后、Introduction 之前）。
- **`outline-spec.md`**：页面标题句式那条加"非内容页用固定标签"例外块——封面（写论文题目）、
  分节页、**作者页（`Authors`）**、收尾页的 take-home 句都不受"科学问题句式"与"≤12 词"约束，
  并注明**其中只有作者页是硬性固定措辞**；页面条目模板的 `Type:` 行注明 `author（标题固定写 Authors）`。
- **`figure-and-author.md`**：作者调研「注意事项」首条加同一条规定，并写明"调研到的信息进栏内，
  不进标题"。
- **`verify_deck.py` 新增 `[authors page]` 段**（诊断，不进退出码）：靠页内出现
  `corresponding author` / `co-first author` / `photo placeholder` 识别作者页，
  核对标题是否为 `Authors`，不符报 WARN 并给出改法。实测：标题写论文题目 → WARN；
  写 `Authors` → OK；无作者页的 deck 报"normal for a group meeting"。
- `SKILL.md` 3.8 呈现规则首条、5.2 清单、脚本说明行，`README.md` 速查表各一处同步。

### v10.4 —— 降权只针对补充行；名词注释行恢复常驻

**背景**：v10.3 把整块 `cap-band` 降权，连名词注释行也要求"先试正文措辞 / 图上标注"。
用户澄清口径：**"标记新名词的图注每一页都完全允许；不希望出现的是作为正文补充内容、
没起到关键提示作用、导致冗余的图注行。"** —— 要抑制的是**行**，不是**框**。

**改法**：图注带里的行分成两类，判据彻底分开。

- **A 类 · 名词注释行（恢复常驻）**：`名词: ≤10 词说明`，判据只有"这个词在本页之前建立过没有"。
  首现即写，**每页都可以有，不设频率上限**，**不参加删除测试**（删了必然卡）。
  载体回到页底 `cap-band`——**正文句中的括注仍然禁止**（一框一号：20pt bullet + 14pt 括注
  同框即 `mixed sizes`，或在正文里拆第二个字号）。
- **B 类 · 补充行（瓶颈 / 试剂原理）收紧**：新增一道门槛——"删掉这行，听众会不会
  **把图读错**或**把结论判高**？"只是"少知道一点"不算。并列出**三类典型冗余**：
  ① 复述正文 bullet（同义改写）；② 复述图上已画清的流程/分组/坐标；③ 解释本领域常识。
  三类见到就删。**冗余图注几乎都出在这一类，不在名词注释上。**
- **权重指标改口径**：从"带 `cap-band` 的页 ≤ 1/3"改为"**带补充行（即 2 行以上）的页 ≤ 1/3**"，
  名词注释行**不计入**这个比例——一行名词注释的图注带是常态，不报警。
- **`verify_deck.py` 的 `[caption band]` 段重写**：分开报"带 `cap-band` 的页数"（只陈述）
  与"含补充行（band 内 2 行以上）的页数与占比"（要看的就是这个），提示语改成
  "删了会不会读错图 / 判高结论"的追问 + 三类冗余清单。仍是诊断，不进退出码。
- **撤销 v10.3 的四级载体优先级表**（前文 / 正文措辞 / 图上标注 / 图注带）：
  它把注释赶去正文措辞，与"每页都可以有"相冲突。`caption_band_in` 的 **0 行 0.00** 档保留
  （确实没有首现名词的页仍然不建框），但**1 行 0.42 恢复为常见形态**，不再是"多数页应落在 0"。
- 同步更新：`SKILL.md`（A/B 档表两行 / 原则一 / 3.2 两处 / 3.6 整节重写 / 232 行模板回填 /
  4.4 生成规则 / 5.2 检查清单四条 / 脚本说明）、`outline-spec.md`（`Bottleneck`·`Background`
  字段语义 / 字段说明 / 30 词分配 / 图注两处 / 页面内逻辑 / 错误表 + 两行新反模式）、
  `layout-spec.md`（承载表 / 分工表 / 行数-高度表的典型构成）、`figure-and-author.md`
  （"该上 PPT"三类重排 + 判断口径）、`typography-spec.md`（一框一号第 8 条）、
  `README.md`（B 档表 / 二.4 / 二.8 / 二.9 / 4.7 自定义口径）、`scripts/verify_deck.py`。

> 仍无自动门禁：脚本只能读行数，读不出"这一行是不是在复述正文"。
> 名词注释行有"首现"这个客观判据；补充行的取舍仍靠 3.6 的 B 类门槛 + 三类冗余清单人工判定。

### v10.3 —— 图注带回到"非必须"，理解性信息改按载体优先级选

> **已被 v10.4 修订**：v10.3 把整块 `cap-band` 一起降权，其中包括了本不该降的名词注释行；
> 其中"载体优先级 ①–④""删除测试三行一律适用""带 `cap-band` 的页 ≤ 1/3"三条已作废。
> 保留的只有：`caption_band_in` 的 `0 行 0.00` 档、`plan_layout` 的 `content_area()` 取 0 的写法、
> `verify_deck` 的 `[caption band]` 诊断段（口径已在 v10.4 改写）。

**背景**：v10.2 把新名词背景行定成"首现即必写、不过删除测试"，加上瓶颈行与试剂原理行，
结果是**每一页都有 `cap-band`**。但图注带在原设计里一直是**非必须项**：它是几种载体里最贵的
一种（占版面高度、把视线从图拉到页底、写偏一点就成讲稿）。100% 出现说明规则把"内容要求"
和"载体要求"绑死了。

**改法**：拆开"内容必须"与"载体自由"，再给载体排优先级。

- **载体优先级（3.6 新表，降序，前一级够用就不走下一级）**：
  ① 前文页 / 背景页已建立 → ② 正文自解释措辞（`the stem-cell marker Lgr5`）→
  ③ 图上标注 → ④ 页底 `cap-band`。内容要求不降（新名词首现必须有背景，仍是 A 档），
  降的是"必须落在页底"这一条。**唯一仍禁止的是正文句中的括注**（一框一号）。
- **删除测试恢复为三行一律适用**，取消 v10.2 给新名词背景行的豁免；新名词行多一道前置门槛：
  先确认 ①②③ 都用不上。
- **权重指标（诊断，非门禁）**：带 `cap-band` 的页应明显少于一半，**超过 1/3 就该回查**，
  连续两三页都出现基本等于拿它当讲稿。
- **`caption_band_in` 默认值改为 0**：不写图注带的页把字段设 0 或干脆不写，
  `plan_layout.py` 的 `content_area()` 改为 `p.get("caption_band_in") or 0`，
  且 band = 0 时底边距回到 `margin_in`（不再吃固定 0.1 in 的窄边）。行数对照表新增
  **0 行 0.00** 一档；`1 行 0.42 / 2 行 0.66 / 3 行 0.90` 不变。
- **新增诊断段 `[caption band]`**（`verify_deck.py`）：报"N of M content page 带 `cap-band`
  （x%）"，超 1/3 时给一段提示，**只诊断、不进退出码**；页数 < 3 的 deck 不做比例提示。
- **单页 demo 重出**：`demo/slide-demo-v103.pptx` / `.png`。原来两行（新名词 + 瓶颈），
  按新规则砍掉瓶颈行——那句 `The 17 nt step is a length gate, not a time course` 既复述了
  卡片上的 `17 nt is a switch, not a rate`，又过不了删除测试（横轴写着 `R-loop length (nt)`，
  不会有人读成时间）。结果：`caption_band_in` 0.66 → 0.42，图高 2.338 → 2.381 in，
  `scaffold` 词数 58 → 39。
- 同步更新：`SKILL.md`（A/B 档表 / 原则一 / 3.2 两处 / 232 行模板回填 / 4.4 生成规则 /
  5.2 检查清单四条 / 3.6 整节重写）、`outline-spec.md`（`Bottleneck`·`Background` 字段 /
  字段说明 / 30 词分配 / 图注两处 / 页面内逻辑 / 错误表一行）、`layout-spec.md`
  （承载表 / 固定模块 / 分工表 / 行数-高度表）、`figure-and-author.md`（"哪些上 PPT"三类 +
  判断口径）、`README.md`（二.4 / 二.8 / 二.9 / 4.7 参数表 / 自定义口径）、
  `scripts/plan_layout.py`（schema 默认值与注释 / `content_area()` / 提示文案）、
  `scripts/verify_deck.py`（`[caption band]` 诊断）。

> 仍无自动门禁：`verify_deck.py` 认不出"这句话是否真需要"。
> 图注带的**出现频率**现在有读数（`[caption band]` 段），**内容取舍**仍靠 3.6 的三问 + 删除测试。

---

### v10.2 —— 新名词背景一律进页底图注带（`cap-band`）

**背景**：v10.1 及以前，新名词首现可以写成"正文句中的同位语/括注"（`Lgr5 (intestinal
stem-cell marker)`）或"紧跟名词"。这与 v10 自己立的一框一号直接冲突：正文 bullet 是 20pt，
括注是 14pt，同框即 `mixed sizes in one textbox`；把括注放大到 20pt 又让注释与主句等重，
读不出主次；括注留在正文还要吃掉 30 词预算。

**改法**：注释类文字统一归页面底部 `cap-band`，正文只留名词本身。

- `cap-band` 的构成固定为 **新名词背景行 → 瓶颈行 → 试剂原理行**，最多 3 行。
- **取舍规则不同，分开写**：新名词背景行**首现即必写**（原则一，A 档），判据是"这个词本页
  之前建立过没有"，**不参加删除测试**；瓶颈行（3.6 三问有实质卡点才写）与试剂原理行
  （首现且原理非常识才写）仍然过删除测试。
- **注释 vs 句子成分**：`Lgr5 (…marker)` 是注释 → 进图注带；`Lgr5+ stem cells`、
  `ILC2-derived IL-13` 是名词作限定语 → 不另起一行。同一个名词只在首现页注释一次；
  背景页讲过的后续不再重复。
- **`caption_band_in` 必须按行数预留**：14pt 一行 0.237 in + 0.18 in 上下留白，
  即 **1 行 0.42 / 2 行 0.66 / 3 行 0.90**（两行是最常见的形态，旧默认 0.45 只够一行半，
  第二行会压到图或页脚）。`plan_layout.py` docstring 的默认值由 0.45 改为 0.66 并写明口径。
- **30 词分配顺序改为四档**：① 实验设计 → ② 本页结论 → ③ 必要对照 → ④ 修饰与重复（先删）。
  原 ③"关键试剂原理"与 ⑤"新名词背景括注"移出正文预算（它们现在都在图注带）。
- **页标题不计入 30 词（口径统一）**：`outline-spec.md` 原写"正文区总词数（标题 + bullets）≤ 30"，
  与 `verify_deck.py` 的实际统计（title 占位符单独计，报 `title=` / `body=` 两个数，门禁只判 `body`）
  不一致；现统一为 **bullets ≤ 30 词、标题另受 ≤12 词约束**。同步 `SKILL.md`（3.2 首段 / A·B 档表 /
  验收清单）、`outline-spec.md` §2、`typography-spec.md` §3 第 6 条、`README.md` 三处，
  并在 `verify_deck.py` 的 `--max-body-words` help 里写明"标题、`cap-*` 图题、`y ≥ scaffold_y`
  的文本都不进这个数"。
- 同步更新：`SKILL.md`（总览表 A/B 档 / 原则一 / 3.2 分配与图注两处 / 3.6 三行表 +
  `caption_band_in` 口径 + 删除测试例外 / 4.4 生成规则 / 232 行模板回填 / 检查清单三处）、
  `outline-spec.md`（`Background` 字段 / 页面内逻辑 / 30 词分配 / 图注两处 / 背景预热 /
  常见错误表两行）、`layout-spec.md`（图注带承载表 / 与正文区分工 + 行数-高度表）、
  `figure-and-author.md`（"哪些上 PPT"三类 + 删除测试例外）、`typography-spec.md`
  （一框一号第 8 条 + `mixed sizes` 判据 + 常见错误）、`README.md`（四条）、
  `scripts/plan_layout.py`（schema 注释 + 模块注释里 cap-band 的 20pt 笔误改 14pt）。

> 未加自动门禁：`verify_deck.py` 不认"这是不是一个新名词"，只认字号与位置。
> 图注带文本按 `y ≥ scaffold_y` 天然算进 `scaffold` 词数、不算正文，所以这轮不需要改脚本逻辑。
> **标题 ≤12 词同样没有自动门禁**（只在人工检查清单里）——demo 的标题原为 13 词
> （`Cas9 unwinds from the PAM, and cuts only past a 17 nt R-loop`），本轮改为 12 词
> （`…, cutting only past a 17 nt R-loop`），`build_demo_v102.py` 里有断言守着。

---

### v10.1 —— L3 正文档 18 → 20 pt

四档变为 **26 / 22 / 20 / 14 pt**（L1 页标题 / L2 组标题 / L3 正文要点 / L4 图注脚注）。

- `verify_deck.py` 的 `DEFAULT_TIERS_PT` 改为 `(26.0, 22.0, 20.0, 14.0)`；**18pt 现在是离档值**，
  会被判 `off-ladder size`。命令行不传 `--tiers` 时即用新阶梯。
- **副作用（已在 `typography-spec.md` §1 写明）**：`22 ∶ 20` 只差 2pt，比例 1.10，同页
  L2 与 L3 并存时字号本身分不开——必须靠**字重（SemiBold vs Regular）+ 位置（统领关系、
  缩进、上方留白）**把两级拉开；若某页的 L2 只是为了分段、没有真正的统领关系，
  就用 L3 加重，别硬加一级 L2。
- **行高变化**：L3 一行 20 × 1.22 / 72 = **0.339 in**（原 18pt 为 0.305）。`SKILL.md` 的
  实测版式参数里 bullets 盒高由 0.72 提到 **0.74 in**（2 行 = 0.68 in），行高坑示例改为
  2 行 20pt + `space_after=6pt` ≈ 0.76 in。
- 同步更新：`typography-spec.md`（档位表 / px 换算 / 比例说明 / `TIER_PT` / 硬约束举例 /
  模块块表 / 验收表 / 常见错误）、`layout-spec.md`（正文区承载 + 倒挂示例）、
  `outline-spec.md`（档位表）、`SKILL.md`（档位表 / `TIER_PT` / A 档表 / 检查清单 /
  参考条目 / 实测参数 / 行高坑）、`README.md`（档位表 / `--tiers` 示例）、
  `template-audit.md`、本文件目录树。

---

### v10 —— 在 v9.7 视觉基线上只并回三样东西

> v9.8/v9.9 那轮"三档字号 + 三色制（禁灰）"被判"效果不行"，已整体撤回；
> v10 的基线就是 v9.7（**四档 26/22/18/14、PAPER_MINIMAL 含灰阶**），
> 只把与视觉风格无关的能力并回来。

**一、内容页逐页自适应排版（来自 10-27 的改造）**

- `plan_layout.py` 新增 `body_zone: {pos: top|bottom|left|right, in}`：当页正文区作为
  一等模块参与规划，输出 `body_box`（生成时照抄，内部自由排）；布局名带
  `bodyT|/bodyB|/bodyL|/bodyR|` 前缀。不声明 = 旧行为（正文条贴图注带上沿）。
- `text_col_in` / `text_col_side` 保留为 left/right 别名。
- 新增 `structure: "hierarchy"|"parallel"`（±4 确定性微调）与面板级 `role: "key"`
  （key 图自动进大格，每页最多一张）。
- 候选布局族扩充：2 图 `1+1-big-left/right`；3 图 `wide-top+2` / `2+wide-bottom`；
  4 图 `1+3-big-left/right`；>4 图枚举全部分解网格 + 大格变体。
- 防呆：`left/right` 占内容宽 ≥60%、`top/bottom` 占高 ≥50% 时直接报错退出。
- **图注带 ≠ 正文区**：图注带 = 页底 14pt `cap-band`（永远底部，硬预留）；
  正文区 = 当页关键正文（`body_zone` 每页声明方位）。标题带 / 页码 / 图注带 /
  子图图题是固定模块，先扣再枚举，任何候选都不可侵占。

**二、一框一号（来自 10-27）**

- 一个文字框内只允许一个字号，需要两个层级就拆两个框。
- `verify_deck.py` 的 `[type scale]` 段新增判据：非标题文字框内出现 ≥2 个显式字号
  → `FAIL [mixed sizes in one textbox]`，并逐框列出字号、给"拆框"提示。
- 豁免两处：`Fig. 5a` 那种同字号加粗；模板标题占位符。

**三、字体归一化 + latin 槽 bug（来自 10-29，纯 bug 修复，不改风格）**

- `theme_override.py` 新增 `apply_font_scheme(prs, latin="Calibri", ea="黑体")`：
  改**所有** `theme*.xml` 的 majorFont/minorFont + master/layout 的 `buFont`，
  清掉换族后失效的 panose/pitchFamily/charset。
- 新增 `fix_latin_slots(prs)`：把 master/layout 里 `latin=+mj-ea / +mn-ea`
  （拉丁槽误指东亚字体槽 → 英文按黑体渲染、字距发宽）改回 `+mj-lt / +mn-lt`。
- 新脚本 `normalize_template.py`：一条命令归一化任意 .pptx 的字体层，存后自动读回校验
  （`--dry-run` 只报告）。**不动配色、不动字号档位**。
- 内置 `template.pptx` 已归一化，发现并修掉两处既存缺陷：
  ① `slideLayout2` 标题 + `slideLayout3` 5 处 `latin` 槽指到东亚字体槽；
  ② `theme2/theme3` 一直是旧 Office 默认的 `等线 Light / 等线`（之前只改 theme1，
     所以引用这两个主题的版式实际渲染成等线）。
- `verify_deck.py` 新增 `[fonts]` 段（进退出码，`--no-font-check` 可关）：theme 的
  major/minor 必须都是 Calibri + 黑体且两两一致；slides/layouts/master 上的显式
  typeface 非 Calibri/黑体即 FAIL；`+mj-ea`/`+mn-ea` 指错槽也 FAIL。
- **不保留** v9.9 的 `[colors]` 禁灰门禁——配色回到 `PAPER_MINIMAL`（含次要灰 `#64748B`）。

**同步更新**：SKILL.md（A 档表 2 条 + 4.3 自适应 + 4.4 一框一号/字体 + 清单 2 项 +
脚本清单）、layout-spec（选型三输入 + 图注带≠正文区对照表 + 固定模块 A 档 + 布局目录）、
typography-spec（硬约束第 8 条 + 验收判据 + 常见错误）、README（门禁表 + 2.7 + 字体节 +
目录树）、assets/README、template-audit。

---

### v9.7 —— 图注带字数豁免 + 缩写规则

v9.7 相对 v9.6 的变化 —— **图注带设为字数豁免区 + 新增缩写使用规则**：

- **图注带去掉字数限制**：原先"瓶颈行 ≤12 词 + 试剂原理 ≤12 词"的每行上限全部取消，
  图注带明确声明为**字数豁免区**（不计 30 词预算、不设每行字数上限）。动机：限字数曾把
  生成端逼成"压术语成缩写、压因果成电报句"，比多几个词更伤理解。保留的硬约束：
  删除测试逐行必过、不写结论、不复述图上已画清的流程、L4 字号与版面高度不变。
- **逐图图题（`cap-*`）仍 ≤12 词**——豁免只针对底部图注带，逐图图题口径不变。
- **新增缩写使用规则（SKILL.md 3.6，全文适用）**：缩写代替全称出现，当且仅当——
  ① 该术语在 deck 全文出现 **>3 次**；或
  ② 图注带（或紧跟名词）给出一行展开，格式 **`缩写：全称（翻译）`**，
     如 `ILC2: group 2 innate lymphoid cell（2 型固有淋巴细胞）`。
  两条都不满足 → 写全称；禁止为凑字数把首次出现的术语直接压成缩写。
- **大纲字段**：`Bottleneck:` 去掉 ≤12 词、标注豁免；`Background:` 拆开口径
  （名词括注 ≤10 词计正文预算；原理行放图注带豁免；缩写展开写规定格式）。
- **同步更新**：SKILL.md（总表/原则一/3.2/3.6/清单 5 处）、outline-spec（字段 + 常见错误表
  新增"为凑字数压缩写"一行）、figure-and-author（3 处）、README（2 处）。
- 脚本无需改动：`verify_deck.py` 本就把图注带排除在词数统计外，无每行字数检查。

v9.6 相对 v9.5 的变化 —— **底部图注带的内容标准重写：从"固定公式"改为"瓶颈判定"**：

- **问题**：图注带第一行原来定义为"实验设计一行：体系 + 关键操作 + 对照"，生成端照公式
  填模板，于是每页都出现一行格式正确但抓不住理解重难点的小字。
- **`SKILL.md` 3.6 重写**：图注带第一行 = **本页最大理解瓶颈**（≤12 词）。写之前先回答
  "目标听众读这页图，最可能卡在哪一步？"，从三个候选里选最实质的一条：
  **读图钥匙**（颜色/标记/读出信号对应什么生物学状态）、**逻辑前提**（这个证据要成立必须先
  接受的前提，含决定性对照）、**易误读点**（相关当因果、技术重复当生物学重复、共用数据）。
  素材取自 Figure Inventory 与读图笔记，不许临时编。
- **体系句降级**：`Tamoxifen-induced Cre; littermate controls` 这类公式句不再默认必写，
  只有当"这是什么体系"本身就是听众的最大卡点时才参选。
- **删除测试取代"每页必填"**：图注带每行自问"删掉它听众读图会卡住吗"，不会卡就删；
  三问都查不到卡点的页允许整条省略。
- **大纲字段 `Design:` 改名 `Bottleneck:`**（outline-spec.md）：要求标注三选一的类别前缀，
  写 `—` 必须给出理由。脚本不解析该字段，改名无副作用。
- **同步更新**：`figure-and-author.md`（瓶颈行条目 + 判断口径）、`layout-spec.md`（图注带分工表）、
  `README.md`（三处）、`SKILL.md`（B 档总表、4.4 两处、5.2.2/5.2.4 清单共 5 处）、
  `INSTALL.md` §6 排障表新增一行。
- 字号阶梯、图注带几何（`y ≥ scaffold_y`、`cap-band` 命名）、词数豁免均**不变**——只改内容标准，不改排版。

v9.5 相对 v9.4 的变化 —— **整体复审，清除迭代中弃用但未清除的内容**（无新功能，只做减法与纠错）：

*文档内自相矛盾的残留*

- `SKILL.md` 3.5 首句加粗标记错乱（`**…**最大值**` 套了三层），修正。
- `SKILL.md` 4.3 约束优先级表里"图内文字大小不在本表内"**重复了两遍**（早期合并编辑残留），删一条并修正缩进。
- `figure-and-author.md` 分页前言仍写"③ 才谈排版（图占比…**≥50%**）"，且"①②一旦定下…**不得回头改写**"
  是绝对禁止——与现行的"三类科学理由可回退"冲突。改为参考线 + 补上回退口子。
- `figure-and-author.md`"图题写在这一页的问题层面，子图只配一句话图注"是**旧的单图注模型**，
  与"标题写问题 + 每图自己一行图题"冲突，改写。

*图数配额的残留（"每页 2–4 张"早已降为诊断指标）*

- `layout-spec.md` 回退触发点里仍带条件"证据 **> 4 张**且其中确有可独立成问的子问题"——图数不该进判据，去掉。
- `layout-spec.md` 5.2 表格仍列"**证据 > 4 张** | 脚本提示"一行，改写为"子图很多 = 仅提示"。
- `SKILL.md` B 档总表把"**每页 2–4 张图**"塞在图注规则那一行，删掉（它不属于图注规则）。
- `template-audit.md` 把"满足**每页 2~4 子图**"当成模板合格判据，改为"能容纳本 deck 实际需要的子图"。

*旧门槛与旧模型的表述*

- `layout-spec.md` 规则 3 标题里的 `> 50%`、"规划全绿、verify **挂 5 页**"、5.2 里"图占比 **FAIL**"、
  小节标题"文字大小：只记录，**不上报**"（与 5.1 的复核触发矛盾）——全部按"参考线 / WARN / 只记录但可复核"改写。
- `SKILL.md` 5.2.4 检查项"图注 = **结论** + 一行实验设计"是旧单图注模型，改为"底部图注带 = Design +
  试剂原理，结论不属图注"。
- `SKILL.md` 4.5 引 `script-spec.md 第 7、11 条`——该文档只有 §1–§12，改为 §8、§12。
- `outline-spec.md` 的 `Captions` 模板用 `fig5a — …`（破折号分隔、编号小写），与生成规范的
  `Fig. 5a` + 两个空格不一致，破折号本身也是被禁的分隔符，统一。

*悬空引用*

- `SKILL.md` 4.4"不用的占位符直接删（`drop_empty_placeholders`）"——**任何脚本里都没有这个函数**。
  改为给出可直接抄的实现（保留标题与页码字段，只删空的正文类占位符）。
- `README.md` 的 `layout_role_map` 示例写 `空白内页: "figure-only"`，而 `inspect_template.py`
  实际产出 `blank`（SKILL.md 也用 `blank`）——统一为 `blank`。

*代码与文档相反的三处行为（`plan_layout.py`）*

- 文档说"传 `min_text_px: 0` 后只优化不重叠 + 面积 + 留白"，实际 `target_text_px` 的惩罚仍在跑 →
  改为 `min_text_px <= 0` 时整体关闭字号项。实测两种参数现在给出**完全相同的坐标**。
- "跨子图文字大小差 **> 3 倍**"惩罚仍在运行，而 `layout-spec.md` 已把该判据标注**已作废**
  （它会把第 4 步拉回决策链）→ 删除该惩罚。
- 覆盖率告警仍在建议"**consider splitting the page or moving a panel in**"（按排版改分页），
  `n > 4` 仍打 `[WARN] exceeds the 2-4 per page rule`，摘要行仍写 `[under the 50% target]` →
  分别改为"接受并注明，不要为提升数字删/拆证据"、`[NOTE]`（问"有没有可独立成问的子问题"）、
  `[under the 50% reference line — diagnostic, not a failure]`。

*新增一条命名约定（因上面本轮实验发现的事实）*

- **底部图注带必须命名 `cap-band`**。`verify_deck.py` 的层级检查要在非图注文本框里挑"正文块"，
  而图注带通常词数最多：不按 `cap-` 命名时它会被当成正文块（L4 末档），任何短标签都不可能
  "更低一级"，于是**倒挂检查整体失效、真倒挂被漏报**（实测同一页：不命名不报，命名后正确报出
  "ShortLabel 18pt < GroupHeading 22pt"）。此前文档把这个机制的原因写反成"误报"，一并纠正。
  同步写进 `layout-spec.md` 的图注小节、`SKILL.md` 4.4/5.2.4、`README.md` §2.7、`INSTALL` §6 排障表。

v9.4 相对 v9.3 的变化 —— **正文词预算 20 → 30（核心证据页例外 28 → 40）**：

- **每页正文区词数上限 20 → 30**；核心证据页的具名例外 28 → 40（同一定义：多出的词
  只能用于读图指引或关键数值，大纲标 `Budget: 40`，验收传 `--body-words-override "9=40"`）。
- **约束强度不变**：仍是 B 档默认建议 + 具名例外，不是科学约束。文案规范里新增一句
  显式表述——**讲清科学问题优先**：一页讲不清"做了什么、看到什么、这个证据为什么站得住"时
  接受更长文案；能压的只有修饰与复述，实验设计、关键对照、定量事实、新名词背景不动（A 档）。
  真空不下时的次序是**删修饰 → 拆页**，不是继续删科学内容。
- `verify_deck.py` 的 `--max-body-words` 默认值 20 → 30；`--body-words-override` 帮助文案
  改为 `"3=40,7=40"` / `Budget: 40`。
- 同步范围：`SKILL.md`（总表 B 档、讲演原则一、3.2 整节、3.6、4.4 模块块、5.2、
  5.4 讲稿、交付清单）、`README.md`（§2.2 B 档、§2.5 门禁表、§2.8 模块块、§2.9 图题、
  §4.6 命令行）、`references/outline-spec.md`（§0.1 整节、页面模板 `Budget:` 行、豁免页说明）、
  `typography-spec.md`、`figure-and-author.md`、`layout-spec.md`、`script-spec.md`。
- **图题（`cap-*`）与底部图注带仍不计入这 30 词**，这条没变——它们是图注，不是正文。

v9.3 相对 v9.2 的变化 —— **新增两个页面级视觉元素（模块块 + 每图图题）**：

- **① 模块块：无边框填充几何块 + 白色文字**（`typography-spec.md`「模块块：填充底 + 反白文字」，
  `SKILL.md` 4.4 有可直接抄的 python-pptx 代码，大纲新增 `Modules:` 字段）：
  - 用法有条件：只在**同一模块 ≥2 个并列**或**一个结论块**时用，每页 ≤3 个；
    **不许给每条 bullet 都套框**（全框等于无框，一页七彩色块最像模板）；
  - 形态固定：矩形/圆角矩形、**无描边无阴影**、填充 `accent1`、白字、字号**仍取 4 档阶梯**
    （不放大，加粗可以）、内边距 0.14/0.08 in；
  - **新增对比度硬门槛**：白字要求填充色相对亮度 `L ≤ 0.175`（`1.05/(L+0.05) ≥ 4.5`），
    主题色偏浅时用加深 30–40% 的版本——白字压浅底投影后既看不清字也丢了边界。
    文档里给了 `luminance()` / `white_readable()` / `darkened()` 三个可直接抄的函数；
  - 命名必须 `Mod-<slug>`：`verify_deck` 把 `Rectangle` 前缀的形状算进图占比，撞名会让诊断数字虚高；
  - 框内文字**仍计入正文 20 词**（套框不产生额外文字配额）。
- **② 每张插入的论文子图，正下方加「编号 + 一句话图题」**：
  - 形式 `Fig. 5a  IL-13 suffices to expand Tuft cells`，紧贴该子图下沿（**不是**页面底部一行概括），
    宽 = 该子图宽、左对齐、≤12 词、最多 2 行、L4 = 14pt 次要灰、`Fig. 5a` 加粗（同字号同色）；
  - 内容只写**观察**（这张子图直接测到什么），不写本页结论、不写推断、不复述图上已画清的流程；
    一页图题合计 ≤40 词，超了说明在写第二套正文；
  - **排版侧真的预留了空间**：`plan.json` 新增 `panel_caption_in`（一行 0.30 / 两行 0.53，
    或给每张图 `caption_lines` 让脚本推算）。`plan_layout.py` 把这段从格子底部扣掉再排图、
    图**贴底对齐**（同一行图题基线一致），并给每个 placement 输出
    `cap_x / cap_y / cap_w / cap_h`——生成时照抄，不要再自己算。缺省 0 时行为与 v9.2 完全一致；
  - **词数与层级统计跟着改**：`verify_deck.py` 按名字前缀 `cap-` 认图注——这些行算 scaffold
    （**不计入正文 20 词**），并且**不再参与"正文块"评选**（否则词数最多的图注会抢走"正文块"的位置，使层级检查失去参照、真倒挂被漏报，
    这正是旧版"把 bullet 压到 ≤图注词数"那个土办法的成因）。新增 `[panel captions]` 段，
    逐页报 `caps=N/pictures`，有图无题的页会被点出来（诊断，不进退出码）；
  - 与页面底部图注带的分工写进了 `layout-spec.md`：**逐图**那行说"这张图是什么"，
    **底部**那行说"这批图怎么做"（Design / 关键试剂原理，全页共用一次）。两处都必须有、不重复。
- **`measure_deck.py` 现在计入文本框内边距**（`margin_left/right/top/bottom`）：
  模块块 0.14/0.08 in 的内边距是真实空间，旧版只扣段落 `marL`，会低估行数与高度、漏报溢出。
  `--ignore-insets` 可回到旧的宽松口径。
- 配套文档同步：`outline-spec.md` 页模板把 `Caption:` 拆成逐图 `Captions:` 块 + `Modules:` 行、
  `figure-and-author.md` 把"一句话图注"改成"每图一行观察"、`SKILL.md` 5.2.3/5.2.4 加对应验收项、
  README 新增 §2.8（原 §2.8 两条讲演原则顺延为 §2.9）。

v9.2 相对 v9.1 的变化 —— **建立规则优先级总表，规则归位 + 去重**：

- **新增「规则优先级总表」**（`SKILL.md` 开头，README §2.2 有速览版）。
  规则分三档，冲突时上一档压下一档：
  - **A 不可妥协**（11 条）：科学完整性优先于排版、结论不越级、不杜撰可追溯、
    分页按问题不按图数、回退只因科学理由、用户确认门禁、产物改则新建副本、
    理解性信息不可删、工程硬失败必修、14pt 字号地板、不用装饰性分隔符。
  - **B 默认建议**（8 条）：正文 ≤20 词、图占比取最大值、一页 ≤3 级字号、
    每图一条 `cap-*` 图题 + 底部 Design 行、每页 ≤3 个模块块、
    各节权重、讲稿时长 90–100%、Discussion 四要素、清单核心五列填满。
  - **C 按场景决定**（10 项）：作者页、Introduction 页数、Method 页数、讲述方式、
    take-home、讲稿是否产出、类型检查项、拓展讨论/收尾页、语言、可读性复核是否触发。
- **另列三类纯诊断值**（图占比数值 / 图内文字 `text_display_px` / 每页图数）：
  只汇报，永不驱动决策。
- **消除"最高原则"字样重复**：原先"四步决策顺序""两条讲演原则""证据清单不可改"三处
  各自宣称最高优先级，现统一引用总表；各 reference 开头都加了一行**档位说明**
  （claim-evidence 全 A、typography 阶梯 A、layout 的 0/1 条 A、outline 分档、
  figure-and-author 作者调研 C、script-spec 产出与否 C）。
- **规则去重**：`SKILL.md` 3.5 与 `layout-spec.md` 三重复的 `group` 适用范围、
  高瘦子图配对、近方形图选型三处细节，改为**只在 `layout-spec.md` 维护**
  （此前两处措辞已出现不一致）。
- README 小节顺延：原 §2.3–2.7 → §2.4–2.8，新增 §2.2 总表速览。

v9.1 相对 v9 的变化 —— **图占比改为诊断指标；Figure Inventory 补重复与独立性**：

- **图占比不再阻塞交付**。优先级明确低于"本页叙述充实、证据链完整、关键证据可读"：
  唯一要求是**取到布局允许的最大值**，50% 降为**参考线与诊断指标**。
  取到最大仍 < 50% 时**只查一次**"这一页有没有相关但非必需的图"——有则移走重排
  （凑数本来就在第 2 步被禁止），没有就**接受并在交付说明注明**，不再回头重新分组。
  同步改写：SKILL.md 第 3 步 / 3.5 / 4.x 约束优先级 / 5.2.3 / 5.2.5，
  `outline-spec.md` §0.3 第 3 步与 §0.4，`layout-spec.md` 三.3 与五 5.2。
- **脚本侧同步**（`verify_deck.py`）：图占比从失败项改为 **WARN**，新增 `[figure coverage]`
  汇总段；新增 `--figure-pct-override "12=42"`（声明某页接受值，标 `accepted`，
  即使在 `--figure-blocking` 下也不再判死）与 `--figure-blocking`（opt-in 恢复阻塞）。
  **退出码不再包含图占比。**
- **Figure Inventory 补两处**：`实验设计` 列加"该对照排除了哪种解释"；
  新增 **`重复与独立性`** 列（生物学重复 / 技术重复 / 独立批次数 / 是否与同页其他子图共用
  同一批动物、同一次建库、同一个模型）。它是 5.1「统计与定量：区分生物学重复与技术重复」
  的承载字段，也是判断"几张图算一条证据还是多条"的依据
  ——共用同一批数据的多张图在证据链上只算一条。

v9 相对 v8.3 的变化 —— **从"逐图总结"升级为"论文级论证"**：

- 新增 **阶段 2 论证审查**（`references/claim-evidence-spec.md`）：结论分四级
  （L1 观察 / L2 关联 / L3 因果 / L4 机制），每条核心主张进 Claim–Evidence 矩阵
  （观察 / 推断 / 强度 / 替代解释 / 决定性实验 / 结论边界），产出 `claims.md`。
  分页以它为依据，没有矩阵不写大纲。
- 新增 **阶段 1.2 论文类型识别**（机制 / 组学 / 结构 / 计算 / 临床 / 理论 / 资源），
  按类型取检查项，不再把同一套清单压给每篇论文。
- Figure Inventory 增 **结论等级 / 影响程度（decisive / bounded / cosmetic）/ A–C 深度分级**；
  逐图三问改为 Why this experiment / What did it show / What does it establish。
  **所有子图都要审查，但不是所有子图都要上页**——两件事分开。
- **图数不再是门槛**：单图页合法（独立问题 + 决定性证据），多图证据页合法；
  2–4 张降为诊断指标。
- **回跳规则改为有条件迭代**：禁止为排版回跳；允许因问题定义错、证据链断、
  替代解释未处理而回退，且必须记进大纲「修订记录」。
- **Introduction 改为任务驱动**（四项任务，不设 3 角度 / 2–5 页下限）；
  **作者页改为按需**（普通组会不单独成页）。
- **阶段 5 改为双重验收**：科学质量表（通过 / 需人工复核 / 未通过，含追溯链抽查）
  + 工程门禁；脚本合格不再等于可交付。
- 讲稿从 SKILL.md 附录迁出为 `references/script-spec.md`，新增逐页五要素
  （目的 / 读图顺序 / 关键结果 / 推理桥梁 / 过渡）与核心图答疑清单。
- 词预算新增具名例外：核心证据页 ≤28 词，
  `verify_deck.py --body-words-override "9=28"`。
- 图内文字偏小不再一律"只记录"：**关键证据（剂量/组别/统计标记）不可读**时触发可读性复核，
  按换布局 → 裁白边 → 回第 1–2 步 → 如实注明的顺序，不得直接缩正文。

v8.3 相对 v8.2 的变化 —— **字号阶梯改为 4 档，单位改 pt，14pt 为硬地板**：

- L5 并入 L4：图注 / 引用不再有专用小字档，与子项共用末档，靠颜色区分。
- 四档改为 **26 / 22 / 18 / 14 pt**（原 25.5 / 21 / 18 / 15 / 12.75）。单位基准由 px 改为 pt，
  四档全为整数，OOXML 直接写 `sz="2600"`，不再涉及 ¾ pt 换算。
- **14pt 是硬下界**：任何小于 14pt 的内容文字判 `below minimum tier`。
  该判据此前实际不可达（pasted ladder 的最小值恒等于下限），现已修正为真正生效，
  且优先于 `off-ladder` 报告——12.75pt 会提示"太小"而不是"不在档位"。
- 模板标题 26pt 正好落在 L1；为避免把模板决定的这档算进"单页 3 级"，
  门禁新增"标题占位符字号不计入层级上限"。
- `--tiers` 单位改为 pt；`plan_layout.py` 舒适线由 17px 改为 14pt（≈18.67px）。
- 图注盒高示例 0.46 → 0.50 in（14pt 两行所需）。

v8.2 相对 v8.1 的变化 —— **面向听众的理解成本与时长约束**：

- 新增「两条讲演原则」（SKILL.md 四步之后）：① 听众不会思考——新名词首现给背景、
  实验设计与关键试剂原理写在页上、图上画不出的必须补；② 时长是硬预算——
  阶段 3.0 第三问确认时长，按 90–100% 目标反推页数，不够时删页、富余时拓展。
- 新增 **3.6 理解性信息**（背景 ≤10 词 / 实验设计 ≤12 词 / 试剂原理 ≤12 词，放图注带不计预算）
  与 **3.7 拓展讨论**（争议、文献不足、后续实验、成药可行性；标 `Beyond the paper`）。
- 大纲格式新增 `Design` / `Background` 字段与 `Duration` / `Audience` 头字段；
  `figure-and-author.md` 的「上页/留大纲」默认值改写：三类理解性信息默认上页。
- 修正 outline-spec §0.1 的口径矛盾（原文写"含图注 ≤20"，与 `verify_deck` 图注带不计数的
  实际行为不符）。
- 阶段 5 清单 +4 项（新名词背景、试剂原理、实验设计在页、时长 90–100%）；
  README 新增 §2.7 与 §4.9，§6 排障 +2 行。

v8.1 相对 v8 的变化 —— **新增产物修改规则**：

- **产物要改，一律新建副本（硬性）**：已交付的 PPT 需要任何修改时不在原文件上改，
  复制为 `deck_v2.pptx`（再改再递增）在副本中修改，改完重过阶段 5 门禁后交付新版；
  原文件保持用户上次看到的状态。落地位置：SKILL.md §4.4 首条 + §5.6 交付清单新增核验项。
- README §6 排障表新增「交付后的 PPT 需要修改」一行。

v8 相对 v7 的变化 —— **附录瘦身，实战经验并入正文**：

- **附录 A（生成阶段高频坑）整节并入正文**：页序先定死、占位符几何四值 → §4.4；
  图内字号实测（第 8 百分位）、实测版式参数 → §4.3；词数按正则数 → §3.2；
  盒高含 `space_after`、PIL 预览自查 → §4.5。`min_text_px: 0` 一条与正文重复，直接删。
- **附录 B 只保留通用知识**：非 PDF 输入（页号 ≠ 章节号，务必回问）→ 阶段 1 开头；
  自渲染坐标图（mmCIF / AlphaFold pLDDT 色阶）与动效 GIF 首帧 → §4.2；
  归一化改不到字面 hex → §4.4；方形子图面积上限 → `layout-spec.md` 新增「方形子图的特别处理」。
  案例专属细节删除。
- **附录 C + D 合并为单一附录「限时中文讲稿」**：11 条，任意页数通用；
  音节/字词两套换算口径合并为一条（相差 <5%），超时定点删句与工程坑并入。
- README §2.5 相应改写为「实战经验（并入正文与附录）」，目录结构与排障指引同步更新。

v7 相对 v6 的变化 —— **结构与格式统一，内容无回退**：

- **SKILL.md 阶段 3 子节编号补齐**：`3.0` 讲述方式 → `3.1` 分页 → `3.2` 文案预算
  → `3.3` 文字层级 → `3.4` 外源配图 → `3.5` 图占比规则。
- **SKILL.md 阶段 4 拆成五个子步**：`4.1` 提图 → `4.2` 裁剪闭环 → `4.3` 排版规划
  → `4.4` 落盘生成 → `4.5` 导出后校验。
- **图占比规则去重**：原先在 SKILL.md 阶段 3、附录 A、`layout-spec.md` 三处各写一遍，
  现统一为「SKILL.md 3.5 写要点 + `layout-spec.md` 三写全量」；人工加大图区的三个技巧、
  `group` 同图等比适用范围、高瘦子图处理、两个 `figure_coverage` 分母不一致
  全部归入 `layout-spec.md` 三的对应小节。
- **阶段 5 清单按主题分六组**：内容完整性 / 分页与举证 / 排版与图占比 / 模板与文字 /
  导出与工程校验 / 交付，共 30+ 项，不再是一长条平铺列表。
- **`outline-spec.md` 第 0 节重编**：原 `0.` / `0.0` / `0.1` / `0.2` / `0.3` / `0.4`
  → `0.1` 文案预算 / `0.2` 讲述方式 / `0.3` 四步顺序 / `0.4` 图占比 / `0.5` 文字层级
  / `0.6` 外源配图，并同步修正文档内全部交叉引用。
- **统一章节引用风格**：`layout-spec.md` 一律用中文序号（三 / 五 / 5.2），不再混用"第三节"。
- **README.md 重写**：与 SKILL.md 编号对齐；修正「三项硬性门禁」→「四项」；
  配色表补注与 `accent4` 的区别；新增两个 `figure_coverage` 分母不一致的提示；
  新增"文案里混进了 `·`，`[copy style]` 报 FAIL"的排障入口。
- 修正 `INSTALL.md` 首段的文档计数笔误。

v6 相对 v5.3 的变化：Figure Inventory 增 **实验设计 / Hole** 两列并打通 Holes → Limitations 汇入；
新增 **三档呈现粒度**（`panel` / `caption-only` / `outline-only`），读图笔记只把会致误读的部分上页，
其余进大纲的「未上榜内容备查」节；
阶段 3 **3.0 第一步改为问用户讲述方式**（`sequential` 顺序 / `backward` 倒推 / `hybrid` 混用），
选定倒推后再确认 take-home + killer data；
Method 页补 **为什么选这个策略 / 备选差在哪**，Key Innovations 补 **why high-impact**；
文案禁用 `·` 等装饰性分隔符，并由 `verify_deck.py` 新增 `[copy style]` 段自动校验。
来源工作区：`C:/Users/10595/.workbuddy/skills/paper-to-slides/`
内置默认模板来源：原项目 `AlphaFold_母板.pptx` 经三步归一化（rename_theme → apply_typography_normalization → apply_theme_override PAPER_MINIMAL）后落到 `assets/_default/template.pptx`；
v10 又跑了一遍 `normalize_template.py`（三个 theme 全改 Calibri + 黑体、master/layout 的 `buFont` 统一、6 处 `latin` 槽从 `+mj-ea/+mn-ea` 改回 `+mj-lt/+mn-lt`），配色仍是 `PAPER_MINIMAL`
配套审查报告：原项目留档 `AlphaFold_母板_模板审查报告_v5.md`（描述 5 个 P 编号修复路径）