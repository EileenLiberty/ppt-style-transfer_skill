## 环境准备

### 0. 安装 Node.js [支持Windows，未安装过的必做]

1. 打开 https://nodejs.org/zh-cn/
2. 下载 **长期支持版 (LTS)** 的 Windows 安装程序 (`.msi`)
3. 双击安装：路径保持默认 `C:\Program Files\nodejs\`，一路 Next（不要改成中文路径）
4. **关掉所有** CMD / PowerShell 窗口，再新开一个终端
5. 验证：

```bash
node -v
npm -v
```

能显示版本号即可。

6. 换国内镜像（避免下载卡住，可以跳过）：

```bash
npm config set registry https://registry.npmmirror.com
```

7. 全局安装 openskills：

```bash
npm install -g openskills
openskills --version
```

### 1. 克隆本仓库

```bash
git clone https://github.com/EileenLiberty/ppt-style-transfer_skill.git
cd ppt-style-transfer_skill
```

### 2. 安装 Anthropic skills（不要手拷官方 .claude）

```bash
openskills install anthropics/skills
```

### 3. 同步 AGENTS.md

```bash
openskills sync
```

同步后会生成 `AGENTS.md`，本仓库的 `training-course-template/`应该保持在`.claude\skills\`目录下。

### 4. Python 依赖

```bash
pip install -r requirements.txt
```

本机还需安装 **Microsoft PowerPoint**，用于 .ppt 转换与 WMF→PNG。

### 5.使用
将源课件放到项目目录后：

```bash
python scripts/build_training_course.py "yourPPT.pptx"
# 或
python scripts/build_training_course.py "yourPPT.ppt"
```
输出默认在 `output/`。
