---
name: training-course-template
description: "将任意源课件套用「培训课件」模板：文字与图表与源一致，背景/版式/装饰继承模板。当用户要求培训课件风格、套用模板-培训课件、或 content-identical restyle 时使用。"
---

# 培训课件模板风格（内容保真 + 模板适配）

模板文件：仓库根目录 `模板-培训课件-20260703.pptx`。

本 skill **面向任意章节/任意源 PPT**（不限于某一章）。规则按「源特征 → 模板版式」泛化，禁止写死章号、节名或页数。

## 用户目标（硬性契约）

| 维度 | 要求 |
|------|------|
| **文字** | 与源课件**语义与用词相同**，不改写、不润色；章号以**源正文**为准 |
| **图表** | 源业务图必须迁入；WMF 宜转高清 PNG；图组在 16:9 内容区等比适配（保持相对布局） |
| **主题模板** | 背景/logo/蓝带/目录左图/标题装饰来自模板 layout，禁止 PptxGenJS 仿色重画 |
| **排版** | 可为 16:9 与可读性重排；目录必须单页；字号紧凑、避免满页留白 |
| **页数** | 默认与源内容页大致对应；目录不额外拆页 |

## 禁止事项

1. 禁止 PptxGenJS / 自绘仿封面  
2. 禁止只复制 theme、丢弃 layouts + 模板 media  
3. 禁止丢弃源业务图  
4. 禁止改写原文；禁止擅自改章号  
5. 禁止 `Expand-Archive` 处理 `.ppt`；禁止 `Compress-Archive` 打包  
6. 禁止用 accent1 `#4874CB` 重绘 chrome（装饰色以 layout 硬编码为准）  
7. 禁止把目录拆成两页；追加目录行必须与模板行视觉一致  

## 模板版式（泛化选用）

| 源页特征 | Layout | 规则 |
|----------|--------|------|
| 封面（课名/作者/单位） | Layout1 | 原文进 ctrTitle / subTitle |
| **目录** | Layout2/3/4 | 见下方「目录单页算法」 |
| 节分隔（仅大标题） | Layout5 | 节名 + 节号 |
| 要点/参数/公式文字 | Layout6 | 标题 + 正文 |
| 左右对比 | Layout7 | 两栏均为原文 |
| 图多文少 | Layout6 | 图放入内容区；正文可短 |
| 源无结束页 | — | 不擅自加「谢谢观看」 |

版式细节：[layouts.md](layouts.md)。

### 目录单页算法（适用于任意源）

1. 从源前几页中识别目录项：匹配 `^\d+\.\d+` 的行（如 `3.1 …`、`10.2 …`），排序去重  
2. **必须单页**  
3. **去重 logo**：Layout2/3/4 模板里「飞行校验中心」logo 常出现两次（同一 `image5.png`）。生成前必须删掉多余 pic，只保留最下方一个。  
4. 条数 `n`：  
   - `n==3` → Layout2；`n==4` → Layout3；`n==5` → Layout4；填占位符即可  
   - `n<3` → Layout2，多余槽位用空白占位  
   - `n>5` → 仍用 Layout4：**前 5 条只填版式占位符（不改、不遮罩、不重绘）**；第 6 条起在同页**同构追加**  
5. **仅追加行须与前五条视觉一致**（禁止改前五条去迁就末条）：  
   - Layout4 真身是「透明底 + `#2576B5` 描边框」，不是实心蓝底白字  
   - 序号框：幻灯片绝对坐标（由版式组合变换映射，勿直接用子坐标系里的 x），数字 `#0070C0`、约 28pt、两位  
   - 标题：同占位符几何，字色 `#2576B5`、约 24pt 加粗微软雅黑；另附同宽描边空心框  
   - 行距 = Layout4 相邻行 `Δy`（约 826135 EMU）；下沿避开底部 logo  
6. 禁止 Layout4 + Layout2「续页」拆目录  

## 字号与留白（泛化）

| 角色 | 建议 |
|------|------|
| 封面标题 / 副标题 | ~30pt / ~16pt |
| 内容标题 | ~22pt |
| 正文 | ~14–15pt，段落间距紧凑 |
| 目录条 | ~15pt |

避免标题过大导致一页只有几行；图组在内容区 scale≈0.96 填满，减少空带。

## 硬性工作流

工具：`.claude/skills/pptx/scripts/`。源可以是任意 `*.ppt` / `*.pptx`。

```bash
# 0) .ppt → .pptx（COM / soffice / WPS）
# 1) 解包模板
python .claude/skills/pptx/scripts/office/unpack.py "模板-培训课件-20260703.pptx" work/
# 2) 迁入源 media（WMF→PNG 优先）；按页建 slide、填占位符、放图
# 3) 打包
python .claude/skills/pptx/scripts/clean.py work/
python .claude/skills/pptx/scripts/office/pack.py work/ "output/{源主名}_培训课件_内容保真.pptx" --original "模板-培训课件-20260703.pptx"
```

验收：[acceptance.md](acceptance.md)。保真细则：[content-fidelity.md](content-fidelity.md)。  
全流程文件清单与数据流：[workflow.md](workflow.md)。

## 输出命名

`output/{源主名}_培训课件_内容保真.pptx`
