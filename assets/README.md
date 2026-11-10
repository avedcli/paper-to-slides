# assets/ — 固定模板放这里

把常用的组会/学术汇报模板复制为本目录下的 **`template.pptx`**，skill 每次生成时直接复用。

    cp "你的模板.pptx" assets/template.pptx

## 模板查找顺序（skill 阶段 0 自动 fallback）

1. `assets/template.pptx` — **项目自有模板**（优先用）
2. `assets/_default/template.pptx` — **skill 内置默认模板**（随包自带，归一化后的 AlphaFold 母板）
3. 都没有 → 向用户索取

> 想让 skill 开箱即用？什么都不放，`_default/` 下已有默认。
> 想换风格？把自己的 `your_template.pptx` 拷成 `assets/template.pptx` 即可覆盖。

## 一份能用的模板至少要有什么

1. **4 类版式**：封面、内容页（**必须带标题占位符**）、分节页（section divider）、纯图页。
   - 内容页标题 = 页面所要回答的科学问题，所以没有标题占位符的"空白页"只能当纯图页用。
   - **不需要**预置多子图格位：子图由生成阶段按英寸坐标绝对定位，版式里画几个图片框无关紧要；
     真正决定每页能放几张图的是可用图形区尺寸。
2. **主题里定死配色与字体**，中英文各自指定（`latin` + `ea`）。中文不设 `ea` 会回退成系统字体。
   v10 起统一为**外文 Calibri + 中文黑体**，一份 deck 只留一个拉丁族 + 一个中文字族；
   `Inter` 这类非系统字体在别人机器上会静默回退、换行错位，Arial/等线/宋体混入会被
   `verify_deck.py` 的 `[fonts]` 段点名。
   - **三个 `theme*.xml` 都要改**：只改 theme1 时，凡引用 theme2/theme3 的版式仍按旧字体渲染
     （内置母版就踩过：theme2/theme3 长期停在 `等线 Light / 等线`）。
   - **`latin` 槽不许指向 `+mj-ea` / `+mn-ea`**：那是东亚槽，英文会被渲染成黑体，字距发宽而
     查不出是哪出错（内置母版这处 bug 已在 v10 修掉）。
   - 直接跑 `python scripts/normalize_template.py assets/template.pptx` 一步搞定两者，
     它只动字体层，不碰配色与字号档位。
3. **每个承载文字的占位符都要有自己的字号**。留空会继承母版默认（Office 原生 44pt 标题 / 32pt 正文），
   通常比设计意图大两倍以上，甚至超出占位框高度而与相邻框压叠。
   - 内容页标题建议落在 **L1 档**（本母板为 26pt，正好等于 L1）；正文区块由生成阶段按
     `references/typography-spec.md` 的 4 档阶梯显式设定，**末档 14pt 即下界**。
     版式里若预设小于 14pt 的文字，投影后必然看不清。
4. **页面尺寸确定**（16:9 → 13.333 × 7.5 in）。
5. **内容区留出净空**：装饰色带 / 水印 / 大 logo 别压到内容区——这块面积直接决定每页能放几张子图。
6. **页码用幻灯片编号域**，别写死 `01 / 09`；并确认母版里页眉页脚开关是打开的（`<p:hf sldNum="1"/>`），
   否则页码域存在也不会显示。
7. 交付前删掉示例页、清空 `cNvPr/@descr` 里的导出器残渣。

## 放进来之前先自检

    python scripts/inspect_template.py assets/template.pptx -o template_profile.json

`warnings` 分三档处理：**致命类**（空壳模板、版式 `custom_design=false`、缺页型）必须先修；
**重要类**（占位符无字号、页码开关关闭、颜色走 `schemeClr` 槽位）建议修，或明确改由生成阶段显式设定；
**卫生类**（字体不一致、非系统字体、内嵌源码、硬编码页码、残留示例页）可先确认再修。

判据细节、单位换算与常见误判见 `references/template-audit.md`。
