import os
import arabic_reshaper
from bidi.algorithm import get_display
from PIL import Image, ImageDraw, ImageFont
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")
FONT_PATH = "NotoNaskhArabic-Bold.ttf"
TEMPLATE_PATH = "group_template.png.png"

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

    img = Image.open(TEMPLATE_PATH).convert("RGB")
    draw = ImageDraw.Draw(img)

    group_font = get_font(95)
    team_font = get_font(78)
    num_font = get_font(76)

    dark = (20, 8, 45)
    white = (255, 255, 255)

    # حرف المجموعة
    draw.text((540, 407), str(group_name).upper(), fill=(222, 174, 55), font=group_font, anchor="mm")

    # الإحداثيات حسب القالب
    rows_y = [605, 755, 905, 1055]

    for i in range(4):
        name, pts = teams[i] if i < len(teams) else ("-", "0")
        y = rows_y[i]

        # النقاط يسار
        draw.text((143, y), str(pts), fill=white, font=num_font, anchor="mm")

        # اسم المنتخب وسط
        draw.text((690, y), ar(name), fill=white, font=team_font, anchor="rm")

        # المركز يمين
        draw.text((925, y), str(i + 1), fill=dark, font=num_font, anchor="mm")

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