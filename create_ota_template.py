#!/usr/bin/env python3
"""大田区用PowerPointテンプレート生成スクリプト"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# 大田区カラー
OTA_NAVY    = RGBColor(0x00, 0x35, 0x6B)   # メインカラー（濃紺）
OTA_BLUE    = RGBColor(0x00, 0x72, 0xBC)   # アクセントカラー（青）
OTA_LIGHT   = RGBColor(0xE8, 0xF4, 0xFF)   # 背景薄青
OTA_GOLD    = RGBColor(0xC8, 0xA0, 0x00)   # アクセント（金）
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_TEXT   = RGBColor(0x44, 0x44, 0x44)
LIGHT_GRAY  = RGBColor(0xF2, 0xF2, 0xF2)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def set_fill_solid(shape, color):
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text_box(slide, text, left, top, width, height,
                 font_size=18, bold=False, color=GRAY_TEXT,
                 align=PP_ALIGN.LEFT, font_name="メイリオ"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    return txBox


def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    set_fill_solid(shape, color)
    shape.line.fill.background()
    return shape


# ─────────────────────────────────────────
# スライド1: タイトルスライド
# ─────────────────────────────────────────
def create_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # 背景
    bg = slide.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    set_fill_solid(bg, OTA_NAVY)
    bg.line.fill.background()

    # 左帯（装飾）
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, OTA_GOLD)

    # 中央白ボックス
    box_top = Inches(1.8)
    box_h   = Inches(3.8)
    box = add_rect(slide, Inches(1.2), box_top, Inches(10.9), box_h, WHITE)

    # 上ライン（青）
    add_rect(slide, Inches(1.2), box_top, Inches(10.9), Inches(0.08), OTA_BLUE)

    # 下ライン（金）
    add_rect(slide, Inches(1.2), box_top + box_h - Inches(0.08),
             Inches(10.9), Inches(0.08), OTA_GOLD)

    # タイトル
    add_text_box(slide, "プレゼンテーションタイトル",
                 Inches(1.5), Inches(2.2), Inches(10.3), Inches(1.2),
                 font_size=36, bold=True, color=OTA_NAVY, align=PP_ALIGN.CENTER)

    # サブタイトル
    add_text_box(slide, "サブタイトル・説明文をここに入力",
                 Inches(1.5), Inches(3.4), Inches(10.3), Inches(0.7),
                 font_size=20, bold=False, color=OTA_BLUE, align=PP_ALIGN.CENTER)

    # 部署名・日付エリア
    add_text_box(slide, "大田区　●●部　●●課",
                 Inches(1.5), Inches(6.1), Inches(7.0), Inches(0.5),
                 font_size=14, color=WHITE, align=PP_ALIGN.LEFT)
    add_text_box(slide, "令和　　年　　月　　日",
                 Inches(9.0), Inches(6.1), Inches(3.0), Inches(0.5),
                 font_size=14, color=WHITE, align=PP_ALIGN.RIGHT)

    # 区名ロゴ文字（右上）
    add_text_box(slide, "大田区",
                 Inches(10.5), Inches(0.3), Inches(2.5), Inches(0.7),
                 font_size=22, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)


# ─────────────────────────────────────────
# スライド2: 目次スライド
# ─────────────────────────────────────────
def create_toc_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 背景
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)

    # ヘッダー帯
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.2), OTA_NAVY)
    add_rect(slide, 0, Inches(1.2), SLIDE_W, Inches(0.06), OTA_GOLD)

    # タイトル
    add_text_box(slide, "目　次",
                 Inches(0.5), Inches(0.2), Inches(12.0), Inches(0.85),
                 font_size=30, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # 目次アイテム
    items = ["１．背景・目的", "２．現状分析", "３．課題の整理",
             "４．施策の方向性", "５．今後のスケジュール"]
    for i, item in enumerate(items):
        y = Inches(1.6 + i * 0.95)
        add_rect(slide, Inches(1.0), y, Inches(0.45), Inches(0.55), OTA_BLUE)
        add_text_box(slide, str(i + 1),
                     Inches(1.0), y - Inches(0.02), Inches(0.45), Inches(0.6),
                     font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(slide, item,
                     Inches(1.65), y, Inches(10.0), Inches(0.6),
                     font_size=18, color=GRAY_TEXT)

    # フッター
    _add_footer(slide, "目次")


# ─────────────────────────────────────────
# スライド3: コンテンツスライド（汎用）
# ─────────────────────────────────────────
def create_content_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.2), OTA_NAVY)
    add_rect(slide, 0, Inches(1.2), SLIDE_W, Inches(0.06), OTA_GOLD)

    add_text_box(slide, "スライドタイトル",
                 Inches(0.5), Inches(0.2), Inches(11.5), Inches(0.85),
                 font_size=28, bold=True, color=WHITE)

    add_text_box(slide, "１．大見出し",
                 Inches(0.5), Inches(1.5), Inches(12.0), Inches(0.6),
                 font_size=18, bold=True, color=OTA_NAVY)
    add_rect(slide, Inches(0.5), Inches(2.1), Inches(12.0), Inches(0.04), OTA_BLUE)

    add_text_box(slide,
                 "・本文テキストをここに入力します。\n"
                 "・箇条書きや説明文を記述してください。\n"
                 "・フォントはメイリオを使用しています。",
                 Inches(0.7), Inches(2.2), Inches(11.5), Inches(1.5),
                 font_size=16, color=GRAY_TEXT)

    add_text_box(slide, "２．大見出し",
                 Inches(0.5), Inches(3.9), Inches(12.0), Inches(0.6),
                 font_size=18, bold=True, color=OTA_NAVY)
    add_rect(slide, Inches(0.5), Inches(4.5), Inches(12.0), Inches(0.04), OTA_BLUE)

    add_text_box(slide,
                 "・本文テキストをここに入力します。\n"
                 "・必要に応じて表・グラフ・図を挿入してください。",
                 Inches(0.7), Inches(4.6), Inches(11.5), Inches(1.2),
                 font_size=16, color=GRAY_TEXT)

    _add_footer(slide, "スライドタイトル")


# ─────────────────────────────────────────
# スライド4: 2カラムスライド
# ─────────────────────────────────────────
def create_two_column_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.2), OTA_NAVY)
    add_rect(slide, 0, Inches(1.2), SLIDE_W, Inches(0.06), OTA_GOLD)

    add_text_box(slide, "スライドタイトル（2カラム）",
                 Inches(0.5), Inches(0.2), Inches(11.5), Inches(0.85),
                 font_size=28, bold=True, color=WHITE)

    for col, (x, label) in enumerate([(Inches(0.4), "左カラム"), (Inches(6.9), "右カラム")]):
        w = Inches(6.0)
        add_rect(slide, x, Inches(1.4), w, Inches(0.55), OTA_BLUE)
        add_text_box(slide, label, x, Inches(1.4), w, Inches(0.55),
                     font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, x, Inches(1.95), w, Inches(4.8), OTA_LIGHT)
        add_text_box(slide, "・内容をここに記述します\n・図表を挿入することも可能です",
                     x + Inches(0.1), Inches(2.05), w - Inches(0.2), Inches(4.5),
                     font_size=15, color=GRAY_TEXT)

    _add_footer(slide, "スライドタイトル（2カラム）")


# ─────────────────────────────────────────
# スライド5: まとめ・終了スライド
# ─────────────────────────────────────────
def create_closing_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg = slide.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    set_fill_solid(bg, OTA_NAVY)
    bg.line.fill.background()

    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, OTA_GOLD)

    add_text_box(slide, "ご清聴ありがとうございました",
                 Inches(1.5), Inches(2.8), Inches(10.3), Inches(1.2),
                 font_size=34, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_text_box(slide, "大田区　●●部　●●課　担当：●●　TEL: 03-XXXX-XXXX",
                 Inches(1.5), Inches(4.5), Inches(10.3), Inches(0.6),
                 font_size=14, color=OTA_LIGHT, align=PP_ALIGN.CENTER)

    add_text_box(slide, "大田区",
                 Inches(10.5), Inches(0.3), Inches(2.5), Inches(0.7),
                 font_size=22, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)


def _add_footer(slide, title_text):
    """共通フッター"""
    add_rect(slide, 0, Inches(7.15), SLIDE_W, Inches(0.35), OTA_NAVY)
    add_text_box(slide, "大田区",
                 Inches(0.2), Inches(7.15), Inches(3.0), Inches(0.35),
                 font_size=11, color=WHITE)
    add_text_box(slide, title_text,
                 Inches(4.5), Inches(7.15), Inches(4.3), Inches(0.35),
                 font_size=11, color=WHITE, align=PP_ALIGN.CENTER)


def main():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    create_title_slide(prs)
    create_toc_slide(prs)
    create_content_slide(prs)
    create_two_column_slide(prs)
    create_closing_slide(prs)

    output = "ota_template.pptx"
    prs.save(output)
    print(f"テンプレートを保存しました: {output}")
    print(f"スライド数: {len(prs.slides)} 枚")


if __name__ == "__main__":
    main()
