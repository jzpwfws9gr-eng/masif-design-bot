import os
import arabic_reshaper
from bidi.algorithm import get_display
from PIL import Image, ImageDraw, ImageFont
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")
FONT_PATH = "NotoNaskhArabic-Bold.ttf"

def get_font(size):
    return ImageFont.truetype(FONT_PATH, size)

def ar(text):
    return get_display(arabic_reshaper.reshape(str(text)))

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "ارسل كذا:\n"
        "/group A\n"
        "إسبانيا 6\n"
        "السعودية 4\n"
        "الأوروغواي 1\n"
        "نيوزيلندا 0"
    )

async def group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lines = update.message.text.splitlines()
    group_name = lines[0].replace("/group", "").strip() or "A"

    teams = []
    for line in lines[1:5]:
        parts = line.rsplit(" ", 1)
        if len(parts) == 2:
            teams.append((parts[0].strip(), parts[1].strip()))

    img = Image.new("RGB", (1080, 1350), (13, 5, 38))
    draw = ImageDraw.Draw(img)

    gold = (222, 174, 55)
    purple = (67, 25, 123)
    purple2 = (42, 14, 86)
    dark = (13, 5, 38)
    white = (255, 255, 255)

    title_font = get_font(90)
    sub_font = get_font(42)
    group_font = get_font(105)
    row_font = get_font(72)
    num_font = get_font(72)
    small_font = get_font(40)

    # Header
    draw.rounded_rectangle((70, 55, 1010, 210), radius=45, fill=purple2, outline=gold, width=4)
    draw.text((540, 112), ar("مونديال المصيف 2026"), fill=gold, font=sub_font, anchor="mm")
    draw.text((540, 168), ar("ترتيب المجموعة"), fill=white, font=title_font, anchor="mm")

    # Group badge
    draw.rounded_rectangle((390, 245, 690, 355), radius=35, fill=gold)
    draw.text((540, 300), group_name.upper(), fill=dark, font=group_font, anchor="mm")

    # Labels
    draw.text((910, 415), ar("الترتيب"), fill=gold, font=small_font, anchor="mm")
    draw.text((540, 415), ar("المنتخب"), fill=gold, font=small_font, anchor="mm")
    draw.text((145, 415), ar("النقاط"), fill=gold, font=small_font, anchor="mm")

    y = 510
    for i in range(4):
        name, pts = teams[i] if i < len(teams) else ("-", "0")

        # Row shadow
        draw.rounded_rectangle((75, y - 52, 1015, y + 72), radius=34, fill=(8, 2, 25))

        # Row body
        draw.rounded_rectangle((65, y - 65, 1005, y + 55), radius=34, fill=purple)

        # Rank box right
        draw.rounded_rectangle((820, y - 65, 1005, y + 55), radius=34, fill=gold)
        draw.text((912, y - 3), str(i + 1), fill=dark, font=num_font, anchor="mm")

        # Points box left
        draw.rounded_rectangle((65, y - 65, 225, y + 55), radius=34, fill=purple2, outline=gold, width=4)
        draw.text((145, y - 3), str(pts), fill=white, font=num_font, anchor="mm")

        # Team name center-right
        draw.text((770, y - 3), ar(name), fill=white, font=row_font, anchor="rm")

        y += 150

    # Footer
    draw.line((160, 1145, 920, 1145), fill=gold, width=3)
    draw.text((540, 1205), ar("نقاط المجموعة"), fill=gold, font=small_font, anchor="mm")
    draw.text((540, 1260), "MONDIAL ALMASIF 2026", fill=gold, font=small_font, anchor="mm")

    path = "group.png"
    img.save(path)

    with open(path, "rb") as photo:
        await update.message.reply_photo(photo=photo)

def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("group", group))
    app.run_polling()

if __name__ == "__main__":
    main()