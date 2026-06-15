import os
from PIL import Image, ImageDraw, ImageFont
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")

def get_font(size):
    fonts = [
        "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
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
        "ارسل كذا:\n"
        "/group A\n"
        "اسبانيا 3\n"
        "السعودية 1\n"
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
            name = parts[0].strip()
            pts = parts[1].strip()
            teams.append((name, pts))

    img = Image.new("RGB", (1080, 1350), (18, 8, 45))
    draw = ImageDraw.Draw(img)

    title_font = get_font(78)
    group_font = get_font(88)
    row_font = get_font(54)
    pts_font = get_font(64)
    small_font = get_font(34)

    draw.text((540, 95), "ترتيب المجموعة", fill="white", font=title_font, anchor="mm")
    draw.rounded_rectangle((390, 160, 690, 265), radius=35, fill=(218, 172, 62))
    draw.text((540, 210), group_name, fill=(20, 10, 45), font=group_font, anchor="mm")

    y = 390
    for i in range(4):
        name, pts = teams[i] if i < len(teams) else ("-", "0")

        draw.rounded_rectangle((80, y - 55, 1000, y + 65), radius=28, fill=(55, 24, 100))
        draw.rounded_rectangle((850, y - 55, 1000, y + 65), radius=28, fill=(218, 172, 62))
        draw.text((925, y), str(i + 1), fill=(20, 10, 45), font=pts_font, anchor="mm")

        draw.text((780, y), name, fill="white", font=row_font, anchor="rm")

        draw.rounded_rectangle((80, y - 55, 210, y + 65), radius=25, fill=(30, 12, 65), outline=(218, 172, 62), width=3)
        draw.text((145, y), pts, fill="white", font=pts_font, anchor="mm")

        y += 155

    draw.text((540, 1110), "النقاط", fill=(218, 172, 62), font=small_font, anchor="mm")
    draw.text((540, 1190), "MONDIAL ALMASIF 2026", fill=(218, 172, 62), font=small_font, anchor="mm")

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