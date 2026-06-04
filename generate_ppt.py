#!/usr/bin/env python3
"""
學術簡報生成器 v2（內建學習功能）
回答問題 → 在桌面產生 .pptx 檔案 → 評分 → 自動學習偏好
"""

import sys
import json
import os
import shutil
from pathlib import Path
from datetime import datetime, date
from collections import Counter

# 安裝依賴
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError:
    print("ERROR: python-pptx 未安裝，正在安裝...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx", "-q"])
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

# ─── 安裝函式 ───────────────────────────────────────────
def install_to_hermes():
    """將本腳本和 SKILL.md 安裝到 ~/.hermes/skills/"""
    src_dir = Path(__file__).parent.resolve()
    target_dir = Path.home() / ".hermes" / "skills" / "productivity" / "academic-ppt"
    force = "--force" in sys.argv or "-y" in sys.argv

    if target_dir.exists() and not force:
        ans = input(f"已安裝於 {target_dir}，是否覆寫？ (y/N): ").strip().lower()
        if ans not in ("y", "yes"):
            print("安裝中止")
            return

    target_dir.mkdir(parents=True, exist_ok=True)
    scripts_dir = target_dir / "scripts"
    scripts_dir.mkdir(exist_ok=True)

    # 安裝 SKILL.md（從 pip 安裝的 package 目錄讀）
    skill_md = src_dir / "SKILL.md"
    if skill_md.exists():
        shutil.copy2(skill_md, target_dir / "SKILL.md")
        print(f"✓ 安裝 SKILL.md")
    else:
        # fallback: 從預設路徑找（开发模式）
        default_skill = Path(__file__).resolve().parents[1] / "SKILL.md"
        if default_skill.exists():
            shutil.copy2(default_skill, target_dir / "SKILL.md")
            print(f"✓ 安裝 SKILL.md")
        else:
            print(f"⚠️  SKILL.md 未找到，跳過（可手動複製）")

    # 安裝本腳本
    shutil.copy2(__file__, scripts_dir / "generate_ppt.py")
    print(f"✓ 安裝 generate_ppt.py")

    # 安裝依賴
    try:
        import importlib
        importlib.import_module("pptx")
        print("✓ python-pptx 已安裝")
    except ImportError:
        print("📦 安裝 python-pptx...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx", "-q"])
        print("✓ python-pptx 安裝完成")

    print(f"\n安裝完成！執行方式：")
    print(f"  python3 {scripts_dir / 'generate_ppt.py'}")


# ─── 配色設定 ───────────────────────────────────────────
COLOR_SCHEMES = {
    "深藍": {
        "primary": RGBColor(0x1F, 0x4E, 0x79),
        "background": RGBColor(0xFF, 0xFF, 0xFF),
        "text": RGBColor(0x1F, 0x4E, 0x79),
        "accent": RGBColor(0x2E, 0x75, 0xB6),
        "light_bg": RGBColor(0xF2, 0xF2, 0xF2),
    },
    "深綠": {
        "primary": RGBColor(0x2E, 0x7D, 0x32),
        "background": RGBColor(0xFF, 0xFF, 0xFF),
        "text": RGBColor(0x2E, 0x7D, 0x32),
        "accent": RGBColor(0x60, 0xAD, 0x4E),
        "light_bg": RGBColor(0xE8, 0xF5, 0xE9),
    },
    "灰黑": {
        "primary": RGBColor(0x33, 0x33, 0x33),
        "background": RGBColor(0xFF, 0xFF, 0xFF),
        "text": RGBColor(0x33, 0x33, 0x33),
        "accent": RGBColor(0x75, 0x75, 0x75),
        "light_bg": RGBColor(0xF5, 0xF5, 0xF5),
    },
}

SKILL_DIR = Path.home() / ".hermes" / "skills" / "productivity" / "academic-ppt"
LEARNING_FILE = SKILL_DIR / "learning.json"
DEFAULT_OCCASIONS = ["研討會", "課堂報告", "論文答辯", "學術會議", "其他"]
DEFAULT_DURATIONS = ["5", "10", "15", "20", "30"]
DEFAULT_COLORS = ["深藍", "深綠", "灰黑"]
DEFAULT_SPECIALS = ["無", "Q&A 頁", "附錄/參考文獻", "圖表", "資料視覺化"]

# ─── 安裝模式 ────────────────────────────────────────────
if "--install" in sys.argv:
    install_to_hermes()
    sys.exit(0)

# ─── 學習系統 ──────────────────────────────────────────

def get_learning_path():
    """確保學習檔案目錄存在並回傳路徑"""
    SKILL_DIR.mkdir(parents=True, exist_ok=True)
    return LEARNING_FILE


def load_learning():
    """載入學習資料，如果不存在則回傳空結構"""
    path = get_learning_path()
    if not path.exists():
        return {"sessions": [], "preferences": {}, "stats": {}}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return {"sessions": [], "preferences": {}, "stats": {}}


def save_learning(data):
    """儲存學習資料"""
    path = get_learning_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def compute_preferences(sessions):
    if not sessions:
        return {}

    now = datetime.now()
    prefs = {}
    scored_sessions = []

    for s in sessions:
        rating = s.get("rating", 0)
        if rating == 0:
            continue
        try:
            ts = datetime.fromisoformat(s["timestamp"])
            days_ago = (now - ts).days
            weight = max(0.3, 1.0 - days_ago * 0.01) if days_ago > 0 else 1.0
            score = rating * weight
            scored_sessions.append((s, score, rating))
        except Exception:
            continue

    if not scored_sessions:
        return {}

    for key in ["occasion", "duration", "slides_count", "color_scheme", "special"]:
        values_with_scores = []
        for s, score, rating in scored_sessions:
            val = s.get(key)
            if val:
                values_with_scores.append((val, score))
        if values_with_scores:
            best = max(values_with_scores, key=lambda x: x[1])
            prefs[key] = best[0]

    return prefs


def get_preferred_defaults(learning):
    prefs = compute_preferences(learning.get("sessions", []))
    return prefs


def record_session(learning, session_data):
    sessions = learning.get("sessions", [])
    sessions.append(session_data)
    sessions = sessions[-100:]
    learning["sessions"] = sessions
    learning["preferences"] = compute_preferences(sessions)
    learning["stats"] = {
        "total": len(sessions),
        "avg_rating": round(
            sum(s.get("rating", 0) for s in sessions if s.get("rating", 0) > 0)
            / max(1, sum(1 for s in sessions if s.get("rating", 0) > 0)),
            2
        ),
    }
    save_learning(learning)
    return learning


# ─── 工具函式 ──────────────────────────────────────────

def ask(prompt, default=None, options=None, hint=None):
    if options:
        print(f"\n{prompt}")
        for i, opt in enumerate(options, 1):
            marker = ""
            if opt == default:
                marker = "（預設）"
            print(f"  {i}. {opt} {marker}")
        if hint and hint.get("suggested") and hint["suggested"] != default:
            print(f"  ★ 使用記錄建議: {hint['suggested']}")
        while True:
            try:
                choice = input("請選擇或直接輸入: ").strip()
                if not choice:
                    if hint and hint.get("suggested"):
                        return hint["suggested"]
                    return default
                idx = int(choice) - 1
                if 0 <= idx < len(options):
                    return options[idx]
                print("無效選項，請重試")
            except ValueError:
                if choice and choice in options:
                    return choice
                print("無效選項，請重試")
    else:
        while True:
            prompt_str = f"\n{prompt}"
            if default:
                prompt_str += f"（預設: {default}）"
            val = input(f"{prompt_str}: ").strip()
            if not val and default:
                return default
            if val:
                return val
            print("此題為必填，請輸入內容")


def get_desktop():
    home = Path.home()
    candidates = [
        home / "Desktop",
        home / "desktop",
        home / "OneDrive" / "Desktop",
        home,
    ]
    for p in candidates:
        if p.exists() and p.is_dir():
            return p
    return home / "Desktop"


# ─── PPT 產生函式 ──────────────────────────────────────

def add_title_slide(prs, title, subtitle, color_scheme):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color_scheme["primary"]

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.33), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.alignment = PP_ALIGN.CENTER

    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(12.33), Inches(1))
    tf2 = sub_box.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = subtitle
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    p2.alignment = PP_ALIGN.CENTER


def add_outline_slide(prs, outline_items, color_scheme):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _fill_bg(slide, color_scheme["background"])
    _add_header(slide, "大綱", color_scheme)

    content = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5))
    tf = content.text_frame
    tf.word_wrap = True
    for i, item in enumerate(outline_items):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"{i+1}. {item}"
        p.font.size = Pt(22)
        p.font.color.rgb = color_scheme["text"]
        p.space_before = Pt(12)
        p.space_after = Pt(6)


def add_content_slide(prs, title, body_lines, color_scheme):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _fill_bg(slide, color_scheme["background"])
    _add_header(slide, title, color_scheme)

    content = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5.5))
    _fill_text_box(content, "\n".join(body_lines), color_scheme)


def add_qna_slide(prs, color_scheme):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _fill_bg(slide, color_scheme["primary"])

    box = slide.shapes.add_textbox(Inches(0), Inches(2.8), Inches(13.33), Inches(1.5))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = "Q & A"
    p.font.size = Pt(52)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.alignment = PP_ALIGN.CENTER

    sub = slide.shapes.add_textbox(Inches(0), Inches(4.5), Inches(13.33), Inches(0.8))
    tf2 = sub.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = "謝謝聆聽"
    p2.font.size = Pt(24)
    p2.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    p2.alignment = PP_ALIGN.CENTER


def add_reference_slide(prs, refs, color_scheme):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _fill_bg(slide, color_scheme["light_bg"])
    _add_header(slide, "參考文獻", color_scheme)

    content = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5.5))
    tf = content.text_frame
    tf.word_wrap = True
    for i, ref in enumerate(refs):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"[{i+1}] {ref}"
        p.font.size = Pt(16)
        p.font.color.rgb = color_scheme["text"]
        p.space_before = Pt(8)


# ─── 輔助函式 ──────────────────────────────────────────

def _fill_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_header(slide, title, color_scheme):
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Inches(0.9))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color_scheme["primary"]
    bar.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(12), Inches(0.7))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def _fill_text_box(shape, text, color_scheme):
    tf = shape.text_frame
    tf.word_wrap = True
    lines = text.strip().split("\n")
    for i, line in enumerate(lines):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = line.strip()
        p.font.size = Pt(20)
        p.font.color.rgb = color_scheme["text"]
        p.space_before = Pt(4)
        p.space_after = Pt(4)


def generate_outline(topic, slides_count, occasion):
    base = [
        f"{topic} 研究背景",
        f"{topic} 研究動機與問題意識",
        f"{topic} 研究方法",
        f"{topic} 研究結果",
        f"{topic} 討論",
        f"{topic} 結論",
    ]
    if slides_count >= 12:
        base.insert(3, f"{topic} 文獻回顧")
        base.insert(5, f"{topic} 資料分析")
    if slides_count >= 15:
        base.insert(2, f"{topic} 理論框架")
    return base[:min(len(base), slides_count - 2)]


def save_prs(prs, topic, desktop):
    today = date.today().isoformat()
    safe_topic = "".join(c if c.isalnum() or c in (" ", "-", "_") else "" for c in topic)
    safe_topic = safe_topic.strip()[:30]
    filename = f"{today}_{safe_topic}.pptx"
    path = desktop / filename
    prs.save(str(path))
    return path


def get_body_for_section(item):
    if "背景" in item or "文獻" in item:
        return [
            "■ 研究背景與文獻探討",
            "",
            "● 現有研究的發現",
            "● 研究缺口（Research Gap）",
            "● 本研究與既有文獻的關聯",
        ]
    elif "動機" in item or "問題" in item:
        return [
            "■ 研究動機與問題意識",
            "",
            "● 為何此問題重要？",
            "● 現有方法的不足之處",
            "● 本研究欲回答的問題",
        ]
    elif "方法" in item or "理論" in item or "框架" in item:
        return [
            "■ 研究方法",
            "",
            "● 研究設計與架構",
            "● 資料來源與分析方法",
            "● 研究限制與假設",
        ]
    elif "結果" in item or "分析" in item or "資料" in item:
        return [
            "■ 研究結果",
            "",
            "● 主要發現摘要",
            "● 數據與證據",
            "● 重要圖表或圖示",
        ]
    elif "討論" in item:
        return [
            "■ 討論",
            "",
            "● 研究結果的詮釋",
            "● 與假說或預期的對照",
            "● 對後續研究的啟示",
        ]
    elif "結論" in item:
        return [
            "■ 結論",
            "",
            "● 本研究的主要貢獻",
            "● 實務或學術意涵",
            "● 未來研究方向",
        ]
    else:
        return [
            f"■ {item}",
            "",
            "● 重點說明 1",
            "● 重點說明 2",
            "● 重點說明 3",
        ]


def ask_rating():
    print("\n─── 滿意度回饋 ───")
    print("本次簡報滿意嗎？")
    print("  1. 很差（需要大幅修改）")
    print("  2. 不太好（有些問題）")
    print("  3. 普通（基本可用）")
    print("  4. 不錯（符合預期）")
    print("  5. 很好（超出預期）")
    while True:
        try:
            r = input("請選擇（1-5）: ").strip()
            if r in ["1", "2", "3", "4", "5"]:
                return int(r)
            print("請輸入 1-5 的數字")
        except (ValueError, EOFError):
            return 0


def ask_feedback():
    print("(直接按 Enter 跳過，或輸入回饋意見)")
    try:
        fb = input("回饋（可选）: ").strip()
        return fb
    except EOFError:
        return ""


# ─── 主程式 ────────────────────────────────────────────

def main():
    print("=" * 50)
    print("  學術簡報生成器 v2")
    print("=" * 50)

    learning = load_learning()
    prefs = get_preferred_defaults(learning)
    stats = learning.get("stats", {})
    if stats.get("total", 0) > 0:
        print(f"\n📊 已學習 {stats['total']} 次，平均滿意度 {stats.get('avg_rating', '?')}")
        if prefs:
            print("   偵測到你的偏好設定，自動套用")

    topic = ask("簡報主題（一句話描述）")

    suggested_occasion = prefs.get("occasion", "研討會")
    hint = {"suggested": suggested_occasion, "used": False}
    occasion = ask("場合", default="研討會", options=DEFAULT_OCCASIONS, hint=hint)

    suggested_duration = prefs.get("duration", "15")
    hint = {"suggested": suggested_duration, "used": False}
    duration = ask("總時長（分鐘）", default="15", options=DEFAULT_DURATIONS, hint=hint)

    suggested_slides = prefs.get("slides_count", 10)
    suggested_slides_str = str(suggested_slides)
    hint = {"suggested": suggested_slides_str, "used": False}
    slides_str = ask("投影片數量（6-30）", default="10", hint=hint)
    try:
        slides_count = max(6, min(30, int(slides_str)))
    except ValueError:
        slides_count = 10

    print("\n大綱：直接按 Enter 由 AI 自動規劃，或輸入自訂章節（用換行分隔）")
    custom = input("（每行一章節）: ").strip()
    if custom:
        outline_items = [line.strip() for line in custom.split("\n") if line.strip()]
    else:
        outline_items = generate_outline(topic, slides_count, occasion)

    print("\n生成大綱：")
    for i, item in enumerate(outline_items, 1):
        print(f"  {i}. {item}")

    confirm = input("\n大綱確認？ (Y/n): ").strip().lower()
    if confirm in ("n", "no"):
        custom = input("輸入自訂大綱（每行一章節）: ").strip()
        outline_items = [line.strip() for line in custom.split("\n") if line.strip()]

    color_scheme = ask("配色", default="深藍", options=DEFAULT_COLORS)
    scheme = COLOR_SCHEMES.get(color_scheme, COLOR_SCHEMES["深藍"])

    special = ask("特殊需求", default="無", options=DEFAULT_SPECIALS)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    add_title_slide(prs, topic, occasion, scheme)
    add_outline_slide(prs, outline_items, scheme)

    for item in outline_items:
        add_content_slide(prs, item, get_body_for_section(item), scheme)

    if "Q&A" in special or "Q&A 頁" in special:
        add_qna_slide(prs, scheme)

    if "附錄" in special or "參考文獻" in special:
        add_reference_slide(prs, ["（請自行填入參考文獻）"], scheme)

    desktop = get_desktop()
    output_path = save_prs(prs, topic, desktop)
    print(f"\n✓ 簡報已儲存：{output_path}")

    rating = ask_rating()
    feedback = ask_feedback()
    session_data = {
        "timestamp": datetime.now().isoformat(),
        "topic": topic,
        "occasion": occasion,
        "duration": duration,
        "slides_count": slides_count,
        "color_scheme": color_scheme,
        "special": special,
        "rating": rating,
        "feedback": feedback,
    }
    learning = record_session(learning, session_data)

    print("\n完成！")
    print(f"📍 {output_path}")


# ─── CLI 入口 ────────────────────────────────────────────
def main_cli():
    """pip 安裝後的入口：無參數進入引導式問答，--install 安裝到 Hermes"""
    if len(sys.argv) == 1:
        main()
    elif sys.argv[1] == "--install":
        install_to_hermes()
    elif sys.argv[1] == "--help":
        print("學術簡報生成器")
        print("  academic-ppt           啟動簡報生成")
        print("  academic-ppt --install  安裝到 ~/.hermes/skills/")
        print("  academic-ppt --help      顯示說明")
    else:
        main()


if __name__ == "__main__":
    main_cli()