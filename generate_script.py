#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
短视频文案生成器 v2.0
支持两种模式：
  1. AI模式 (配置Deepseek API后) - 生成完整口播文案
  2. 模板模式 (无API) - 生成框架模板供你填空

首次使用需配置API密钥，支持从 .env 文件读取。
"""

import json
import os
import sys
import urllib.request
from datetime import datetime
from pathlib import Path

# ===== 配置 =====

SCRIPTS_DIR = Path(__file__).parent / "scripts"
ENV_FILE = Path(__file__).parent / ".env"
DEFAULT_API_BASE = "https://api.deepseek.com/v1"
DEFAULT_MODEL = "deepseek-chat"

# ===== 文案框架定义 =====

FRAMEWORKS = [
    {
        "id": 1,
        "name": "钩子-展开-收尾 (通用)",
        "desc": "万能公式：强钩子开头，中间讲干货，结尾总结引导互动",
        "sections": [
            {"name": "钩子", "time": "0-5秒", "hint": "前5秒决定去留，用反常识/好奇/痛点抓住注意力"},
            {"name": "展开", "time": "5-45秒", "hint": "核心内容区，围绕主题展开，讲清你的观点"},
            {"name": "收尾 + CTA", "time": "45-60秒", "hint": "总结升华，引导关注/点赞/评论"},
        ],
    },
    {
        "id": 2,
        "name": "PAS 痛点驱动",
        "desc": "先戳痛点，再给解药：适合解决问题类的干货内容",
        "sections": [
            {"name": "抛出问题", "time": "0-8秒", "hint": "描述观众正在经历的痛点或困惑"},
            {"name": "放大后果", "time": "8-30秒", "hint": "强化问题的负面影响，制造紧迫感"},
            {"name": "给出方案", "time": "30-55秒", "hint": "提供你的方法/观点/产品作为解药"},
            {"name": "行动号召", "time": "55-60秒", "hint": "明确告诉观众下一步该做什么"},
        ],
    },
    {
        "id": 3,
        "name": "3幕故事法",
        "desc": "讲好一个故事：铺垫 -> 冲突 -> 结局，适合个人经历/案例分享",
        "sections": [
            {"name": "铺垫", "time": "0-10秒", "hint": "建立场景和人物，让观众代入"},
            {"name": "冲突/转折", "time": "10-45秒", "hint": "矛盾出现，制造情绪起伏"},
            {"name": "结局 + 感悟", "time": "45-60秒", "hint": "结果揭晓，点出核心主题和领悟"},
        ],
    },
    {
        "id": 4,
        "name": "AIDA 说服法则",
        "desc": "Attention -> Interest -> Desire -> Action，适合种草/推荐",
        "sections": [
            {"name": "吸引注意", "time": "0-5秒", "hint": "用反常识/惊人事实/强反差开场"},
            {"name": "激发兴趣", "time": "5-20秒", "hint": "展开话题，让观众产生继续看的欲望"},
            {"name": "唤起渴望", "time": "20-50秒", "hint": "展示核心价值，让观众觉得'我也想要'"},
            {"name": "促进行动", "time": "50-60秒", "hint": "明确的CTA，告诉观众下一步"},
        ],
    },
]

TONES = {
    "1": {"label": "口语轻松", "desc": "像朋友聊天一样自然"},
    "2": {"label": "激情澎湃", "desc": "情绪饱满，有感染力"},
    "3": {"label": "专业深度", "desc": "理性干货，有说服力"},
    "4": {"label": "幽默风趣", "desc": "轻松搞笑，让人会心一笑"},
    "5": {"label": "温暖治愈", "desc": "温柔走心，给人力量"},
}

# ===== API 配置管理 =====


def load_config():
    """从 .env 文件加载 API 配置"""
    config = {"api_key": "", "api_base": DEFAULT_API_BASE, "model": DEFAULT_MODEL}
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text("utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, val = line.split("=", 1)
                key = key.strip()
                val = val.strip().strip("\"'")
                if key == "DEEPSEEK_API_KEY":
                    config["api_key"] = val
                elif key == "DEEPSEEK_API_BASE":
                    config["api_base"] = val.rstrip("/")
                elif key == "DEEPSEEK_MODEL":
                    config["model"] = val
    return config


def save_config(api_key, api_base, model):
    """保存 API 配置到 .env 文件"""
    content = f"""# Deepseek API 配置
DEEPSEEK_API_KEY={api_key}
DEEPSEEK_API_BASE={api_base}
DEEPSEEK_MODEL={model}
"""
    ENV_FILE.write_text(content, encoding="utf-8")
    print("\n  配置已保存到 .env 文件")


def setup_config():
    """首次配置向导"""
    print("\n" + "-" * 56)
    print("\n首次使用，请配置 Deepseek API：\n")
    print("  (如果你还没有API密钥，可前往 https://platform.deepseek.com 获取)")
    print()

    api_key = input("  API Key: ").strip()
    while not api_key:
        api_key = input("  API Key: ").strip()

    api_base = input(f"  API Base URL (回车默认 {DEFAULT_API_BASE}):\n  > ").strip()
    if not api_base:
        api_base = DEFAULT_API_BASE

    model = input(f"  模型名 (回车默认 {DEFAULT_MODEL}):\n  > ").strip()
    if not model:
        model = DEFAULT_MODEL

    save_config(api_key, api_base, model)
    return {"api_key": api_key, "api_base": api_base, "model": model}


# ===== Deepseek API 调用 =====


def _clean_surrogates(text):
    """移除字符串中的孤儿代理字符 (lone surrogates)"""
    if not text:
        return text
    # Python 的 UTF-8 编码器不支持代理字符, 需要逐个字符清理
    result = []
    for ch in text:
        if '\ud800' <= ch <= '\udfff':
            result.append('?')  # 替换为问号
        else:
            result.append(ch)
    return ''.join(result)


def call_deepseek(system_prompt, user_prompt, config):
    """调用 Deepseek 兼容 API"""
    url = f"{config['api_base']}/chat/completions"

    # 清理可能损坏的字符
    system_prompt = _clean_surrogates(system_prompt)
    user_prompt = _clean_surrogates(user_prompt)

    payload = json.dumps({
        "model": config["model"],
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.8,
        "max_tokens": 4096,
        "stream": False,
    }, ensure_ascii=False).encode("utf-8")

    req = urllib.request.Request(
        url,
        data=payload,
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "Authorization": f"Bearer {config['api_key']}",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return result["choices"][0]["message"]["content"]
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"API 请求失败 (HTTP {e.code}): {error_body}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"网络连接失败: {e.reason}")
    except (KeyError, json.JSONDecodeError) as e:
        raise RuntimeError(f"API 响应解析失败: {e}")


# ===== AI 文案提示词构造 =====


def build_prompts(framework, tone, info):
    """构造 AI 提示词"""
    sections_desc = "\n".join(
        f"  [{s['time']}] {s['name']} - {s['hint']}"
        for s in framework["sections"]
    )

    points_text = "\n".join(f"  - {p}" for p in info["points"]) if info["points"] else "  (无)"

    system_prompt = f"""你是一个专业的短视频文案写手。你的任务是用{info['target']}能听懂的口语，写一段60秒以内的短视频口播文案。

## 核心要求
- 语言必须口语化，适合口播，像在和朋友聊天
- 不要书面语、不要堆砌形容词、不要假大空
- 每句话都要有信息量，节奏要紧凑
- 开头必须有强钩子，让观众舍不得划走
- 严格按照给定的框架和时间分配来写
- 画面建议要具体、可执行，不能只说"展示相关内容"
- 语气：{tone['label']}（{tone['desc']}）

## 输出格式要求
你必须严格按照以下 JSON 格式输出，不要加 markdown 代码块标记，直接输出纯 JSON：

{{
  "title_ideas": ["标题1", "标题2", "标题3", "标题4", "标题5"],
  "punchlines": [
    "金句1 - 一句让人记住的话，适合做视频封面标题",
    "金句2"
  ],
  "hashtags": ["#标签1", "#标签2", "#标签3", "#标签4", "#标签5"],
  "sections": [
    {{
      "script": "这部分的完整口播台词...",
      "visual": "这段对应的画面建议..."
    }}
  ]
}}

sections 数组的长度和顺序必须与下面给出的框架分段完全一致。"""

    user_prompt = f"""## 主题
{info['theme']}

## 目标观众
{info['target']}

## 关键要点（必须覆盖）
{points_text}

## 文案框架
框架：{framework['name']}
分段：
{sections_desc}

## 行动召唤 (CTA)
{info['cta']}"""

    return system_prompt, user_prompt


def parse_ai_output(content, framework):
    """解析 AI 返回的 JSON，提取各段内容"""
    # 清理 AI 返回中的损坏字符
    content = _clean_surrogates(content)
    json_str = content.strip()
    if json_str.startswith("```"):
        # 去掉代码块标记
        lines = json_str.splitlines()
        lines = [l for l in lines if not l.startswith("```")]
        json_str = "\n".join(lines).strip()

    data = json.loads(json_str)
    sections = data.get("sections", [])

    # 校验段数
    expected = len(framework["sections"])
    if len(sections) != expected:
        raise ValueError(
            f"AI 返回了 {len(sections)} 段，但框架需要 {expected} 段"
        )

    # 构建结构化输出
    result = []
    for i, sec in enumerate(framework["sections"]):
        result.append({
            "title": sec["name"],
            "time": sec["time"],
            "hint": sec["hint"],
            "script": sections[i].get("script", ""),
            "visual": sections[i].get("visual", ""),
        })

    return {
        "sections": result,
        "title_ideas": data.get("title_ideas", []),
        "punchlines": data.get("punchlines", []),
        "hashtags": data.get("hashtags", []),
    }


# ===== 界面交互 =====


def _fix_stdin_encoding():
    """修复 Windows 终端中文输入编码问题"""
    if os.name == "nt" and sys.stdin.encoding and sys.stdin.encoding.lower() in ("gbk", "gb2312", "cp936"):
        try:
            sys.stdin.reconfigure(encoding="utf-8")
        except Exception:
            try:
                # 备用方案：用 buffer 包装
                import codecs
                sys.stdin = codecs.getreader("utf-8")(sys.stdin.buffer)
            except Exception:
                pass  # 放弃治疗，不修了


def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")


def print_banner(has_api):
    clear_screen()
    mode = "[AI模式]" if has_api else "[模板模式]"
    print("=" * 56)
    print(f"          短视频文案生成器 v2.0  {mode}")
    print("   输入主题 -> 选择框架 -> 生成结构化文案")
    print("=" * 56)


def choose_framework():
    print("\n" + "-" * 56)
    print("\n请选择文案框架：\n")
    for fw in FRAMEWORKS:
        print(f"  [{fw['id']}] {fw['name']}")
        print(f"       {fw['desc']}")
        print()
    while True:
        try:
            choice = int(input("  输入编号 (1-4): ").strip())
            if 1 <= choice <= len(FRAMEWORKS):
                return FRAMEWORKS[choice - 1]
        except ValueError:
            pass
        print("  无效选择，请重新输入\n")


def choose_tone():
    print("\n请选择文案语气：\n")
    for key, tone in TONES.items():
        print(f"  [{key}] {tone['label']} -- {tone['desc']}")
    print()
    while True:
        choice = input("  输入编号 (1-5): ").strip()
        if choice in TONES:
            return TONES[choice]
        print("  无效选择，请重新输入\n")


def collect_info():
    print("\n" + "-" * 56)
    print("\n请填写以下信息：\n")

    theme = input("  核心主题 (一句话说清这个视频想表达什么):\n  > ").strip()
    while not theme:
        theme = input("  > ").strip()

    target = input("\n  目标观众 (谁会在看这个视频):\n  > ").strip()

    print("\n  关键要点 (每行一个，输入空行结束):")
    points = []
    while True:
        point = input("  * ").strip()
        if not point:
            break
        points.append(point)

    cta = input("\n  希望观众做什么 (关注/点赞/评论/分享等):\n  > ").strip()

    return {
        "theme": theme,
        "target": target or "通用观众",
        "points": points,
        "cta": cta or "关注我，获取更多干货",
    }


# ===== 文案生成 =====


def generate_script_ai(framework, tone, info, config):
    """AI 模式：调用 API 生成完整文案"""
    print("\n  AI 正在为你写文案，请稍候...\n")

    system_prompt, user_prompt = build_prompts(framework, tone, info)
    raw = call_deepseek(system_prompt, user_prompt, config)
    parsed = parse_ai_output(raw, framework)

    return {
        "meta": {
            "title": info["theme"],
            "framework": framework["name"],
            "tone": tone["label"],
            "target": info["target"],
            "cta": info["cta"],
            "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        },
        "sections": parsed["sections"],
        "key_points": info["points"],
        "title_ideas": parsed["title_ideas"],
        "punchlines": parsed["punchlines"],
        "hashtags": parsed["hashtags"],
    }


def generate_script_template(framework, tone, info):
    """模板模式：生成框架模板"""
    sections = []
    for sec in framework["sections"]:
        sections.append({
            "title": sec["name"],
            "time": sec["time"],
            "hint": sec["hint"],
        })
    return {
        "meta": {
            "title": info["theme"],
            "framework": framework["name"],
            "tone": tone["label"],
            "target": info["target"],
            "cta": info["cta"],
            "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        },
        "sections": sections,
        "key_points": info["points"],
        "title_ideas": [],
        "punchlines": [],
        "hashtags": [],
    }


# ===== Markdown 输出 =====


def render_markdown(script, is_template):
    """将文案渲染为 Markdown"""
    # 清理所有数据中的损坏字符
    for sec in script.get("sections", []):
        for key in ("script", "visual", "title", "hint"):
            if key in sec:
                sec[key] = _clean_surrogates(sec[key])
    for key in ("title_ideas", "punchlines", "hashtags"):
        if key in script:
            script[key] = [_clean_surrogates(x) for x in script[key]]
    m = script["meta"]
    lines = []

    lines.append(f"# {m['title']}")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("| 项目 | 内容 |")
    lines.append("| --- | --- |")
    lines.append(f"| 核心主题 | {m['title']} |")
    lines.append(f"| 文案框架 | {m['framework']} |")
    lines.append(f"| 语气风格 | {m['tone']} |")
    lines.append(f"| 目标观众 | {m['target']} |")
    lines.append(f"| 期望时长 | 约 60 秒 |")
    lines.append(f"| 生成时间 | {m['created']} |")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 分段脚本")
    lines.append("")

    for sec in script["sections"]:
        lines.append(f"### [{sec['time']}] {sec['title']}")
        lines.append("")
        lines.append(f"> 提示：{sec['hint']}")
        lines.append("")

        if is_template:
            lines.append("**台词：**")
            lines.append("")
            lines.append("> (在这里填写你的口播文案)")
            lines.append("")
            lines.append("**画面建议：**")
            lines.append("")
            lines.append("- (描述这里用什么画面或素材)")
        else:
            lines.append("**台词：**")
            lines.append("")
            for para in sec.get("script", "").split("\n"):
                para = para.strip()
                if para:
                    lines.append(f"> {para}")
            lines.append("")
            lines.append("**画面建议：**")
            lines.append("")
            for cue in sec.get("visual", "").split("\n"):
                cue = cue.strip()
                if cue:
                    lines.append(f"- {cue}")

        lines.append("")
        lines.append("---")
        lines.append("")

    if script["key_points"]:
        lines.append("## 务必覆盖的要点")
        lines.append("")
        for pt in script["key_points"]:
            lines.append(f"- [ ] {pt}")
        lines.append("")

    if script["title_ideas"]:
        lines.append("## 推荐标题")
        lines.append("")
        for t in script["title_ideas"]:
            lines.append(f"- {t}")
        lines.append("")

    if script["punchlines"]:
        lines.append("## 金句/高光台词")
        lines.append("")
        for p in script["punchlines"]:
            lines.append(f"- 「{p}」")
        lines.append("")

    lines.append("## 行动召唤 (CTA)")
    lines.append("")
    lines.append(f"{m['cta']}")
    lines.append("")

    if script["hashtags"]:
        lines.append("## 推荐话题标签")
        lines.append("")
        for tag in script["hashtags"]:
            lines.append(f"  {tag}")

    return "\n".join(lines)


def save_script(content, theme):
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = "".join(c for c in theme if c.isalnum() or c in " _-").strip() or "unnamed"
    safe_name = safe_name[:20]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{timestamp}_{safe_name}.md"
    filepath = SCRIPTS_DIR / filename
    # 清理任何残留的代理字符
    content = _clean_surrogates(content)
    filepath.write_text(content, encoding="utf-8")
    return filepath


# ===== 输出摘要 =====


def show_summary(filepath, script, is_template):
    m = script["meta"]
    mode = "模板" if is_template else "完整文案"
    print()
    print("=" * 56)
    print(f"  {mode}生成完成！")
    print("=" * 56)
    print(f"\n  主题：{m['title']}")
    print(f"  框架：{m['framework']}")
    print(f"  语气：{m['tone']}")
    print(f"  文件：{filepath}")
    print()

    if is_template:
        print("  接下来：")
        print("    1. 打开 .md 文件填写台词")
        print("    2. 对照画面建议准备素材")
        print("    3. 录制时检查'务必覆盖的要点'")
    else:
        print("  文案已包含完整台词和画面建议，可直接使用")
        print("  当然，你完全可以根据自己的风格再调整")
    print()

    if not is_template and script["punchlines"]:
        print("  金句预览：")
        for p in script["punchlines"][:2]:
            print(f"    「{p}」")
        print()

    print("  再次运行：python generate_script.py")
    print()


# ===== 主流程 =====


def main():
    try:
        _fix_stdin_encoding()
        # 1. 加载配置
        config = load_config()
        has_api = bool(config["api_key"])

        if not has_api:
            print_banner(False)
            print("\n  [!] 未检测到 API 配置，将以模板模式运行")
            print("  (只能生成框架模板，需要你手动填写台词)")
            resp = input("\n  是否现在配置 Deepseek API？(y/n): ").strip().lower()
            if resp == "y":
                config = setup_config()
                has_api = True
            print()
        else:
            print_banner(True)

        # 2. 收集信息
        info = collect_info()

        # 3. 选择框架和语气
        framework = choose_framework()
        tone = choose_tone()

        # 4. 生成文案
        is_template = not has_api
        if has_api:
            try:
                script = generate_script_ai(framework, tone, info, config)
            except Exception as e:
                print(f"\n  [!] AI 生成失败: {e}")
                resp = input("  是否降级为模板模式继续？(y/n): ").strip().lower()
                if resp == "y":
                    script = generate_script_template(framework, tone, info)
                    is_template = True
                else:
                    print("  已取消。")
                    sys.exit(1)
        else:
            script = generate_script_template(framework, tone, info)

        # 5. 渲染和保存
        markdown = render_markdown(script, is_template)
        filepath = save_script(markdown, info["theme"])
        show_summary(filepath, script, is_template)

        # 6. 打开文件
        if os.name == "nt":
            resp = input("  用默认编辑器打开文件？(y/n): ").strip().lower()
            if resp == "y":
                os.startfile(filepath)

    except KeyboardInterrupt:
        print("\n\n  已取消。")
        sys.exit(0)


if __name__ == "__main__":
    main()
