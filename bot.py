import os
from PIL import Image, ImageDraw, ImageFont
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")

def get_font(size):
    fonts = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]

    for path in fonts:
        try:
            return ImageFont.truetype(path, size)
        except:
            pass

    return ImageFont.load_default()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Send like this:\n"
        "/group A\n"
        "Spain 6\n"
        "Saudi Arabia 4\n"
        "Uruguay 1\n"
        "New Zealand 0"
    )

async def group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lines = update.message.text.splitlines()
    group_name = lines[0].replace("/group", "").strip() or "A"

    teams = []
    for line in lines[1:5]:
        parts = line.rsplit(" ", 1)
        if len(parts) == 2:
            name = parts[0].strip()
            pts = parts[1].strip()
            teams.append((name, pts))

    img = Image.new("RGB", (1080, 1350), (18, 8, 45))
    draw = ImageDraw.Draw(img)

    title_font = get_font(100)
    group_font = get_font(120)
    row_font = get_font(76)
    pts_font = get_font(78)
    small_font = get_font(44)

    gold = (218, 172, 62)
    purple = (55, 24, 100)
    dark = (18, 8, 45)
    box_dark = (30, 12, 65)

    draw.text((540, 95), "GROUP STANDINGS", fill="white", font=title_font, anchor="mm")

    draw.rounded_rectangle((390, 165, 690, 275), radius=35, fill=gold)
    draw.text((540, 220), group_name.upper(), fill=dark, font=group_font, anchor="mm")

    y = 405
    for i in range(4):
        name, pts = teams[i] if i < len(teams) else ("-", "0")

        draw.rounded_rectangle((70, y - 60, 1010, y + 70), radius=30, fill=purple)

        draw.rounded_rectangle((835, y - 60, 1010, y + 70), radius=30, fill=gold)
        draw.text((922, y), str(i + 1), fill=dark, font=pts_font, anchor="mm")

        draw.text((780, y), name, fill="white", font=row_font, anchor="rm")

        draw.rounded_rectangle((70, y - 60, 220, y + 70), radius=28, fill=box_dark, outline=gold, width=4)
        draw.text((145, y), pts, fill="white", font=pts_font, anchor="mm")

        y += 160

    draw.text((540, 1130), "POINTS", fill=gold, font=small_font, anchor="mm")
    draw.text((540, 1210), "MONDIAL ALMASIF 2026", fill=gold, font=small_font, anchor="mm")

    path = "group.png"
    img.save(path)

    await update.message.reply_photo(photo=open(path, "rb"))

def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("group", group))
    app.run_polling()

if __name__ == "__main__":
    main()