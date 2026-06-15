import os
import arabic_reshaper
from bidi.algorithm import get_display
from PIL import Image, ImageDraw, ImageFont
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")
FONT_PATH = "Cairo-Bold-1.ttf"

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

    img = Image.new("RGB", (1080, 1080), (18, 8, 45))
    draw = ImageDraw.Draw(img)

    title_font = get_font(76)
    group_font = get_font(90)
    row_font = get_font(66)
    pts_font = get_font(70)
    small_font = get_font(36)

    gold = (218, 172, 62)
    purple = (55, 24, 100)
    dark = (18, 8, 45)

    draw.text((540, 80), ar("ترتيب المجموعة"), fill="white", font=title_font, anchor="mm")

    draw.rounded_rectangle((390, 130, 690, 230), radius=28, fill=gold)
    draw.text((540, 180), group_name.upper(), fill=dark, font=group_font, anchor="mm")

    y = 320
    for i in range(4):
        name, pts = teams[i] if i < len(teams) else ("-", "0")

        draw.rounded_rectangle((70, y - 45, 1010, y + 55), radius=25, fill=purple)

        draw.rounded_rectangle((840, y - 45, 1010, y + 55), radius=25, fill=gold)
        draw.text((925, y), str(i + 1), fill=dark, font=pts_font, anchor="mm")

        draw.text((780, y), ar(name), fill="white", font=row_font, anchor="rm")

        draw.rounded_rectangle((70, y - 45, 200, y + 55), radius=22, outline=gold, width=4)
        draw.text((135, y), str(pts), fill="white", font=pts_font, anchor="mm")

        y += 130

    draw.text((540, 910), ar("النقاط"), fill=gold, font=small_font, anchor="mm")
    draw.text((540, 965), ar("مونديال المصيف 2026"), fill=gold, font=small_font, anchor="mm")

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