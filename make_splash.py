# ═══════════════════════════════════════════════════════════════════
# إنشاء شاشة Splash لبرنامج Sunrise
# ═══════════════════════════════════════════════════════════════════
from PIL import Image, ImageDraw, ImageFont
import os
import arabic_reshaper
from bidi.algorithm import get_display

# ═══════════════════════════════════════════════════════════════════
# الإعدادات
# ═══════════════════════════════════════════════════════════════════
W, H = 600, 480

# الألوان
BG_COLOR = '#1A5276'       # أزرق داكن
GOLD = '#F1C40F'           # ذهبي
WHITE = '#FFFFFF'          # أبيض
GRAY = '#BDC3C7'           # رمادي فاتح

# الشعار
LOGO_FILE = 'm.png'
LOGO_SIZE = 130

# النصوص
PROGRAM_NAME_EN = "Sunrise System"
PROGRAM_NAME_AR = "نظام متكامل لإدارة المبيعات والخزينة"
LOADING_TEXT = "جارٍ التحميل..."

COMPANY_NAME = "شركة الزهراء للبرمجيات"
MANAGER_NAME = "المحاسب / حسام الدين سيد ربيع"
PHONE = "01003796984"


# ═══════════════════════════════════════════════════════════════════
# دالة معالجة النص العربي
# ═══════════════════════════════════════════════════════════════════
def ar(text):
    """تحويل النص العربي لعرضه بشكل صحيح"""
    if not text:
        return ''
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)


# ═══════════════════════════════════════════════════════════════════
# تحميل الخطوط
# ═══════════════════════════════════════════════════════════════════
def load_font(name, size):
    paths = [
        f"C:/Windows/Fonts/{name}",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
    ]
    for p in paths:
        try:
            return ImageFont.truetype(p, size)
        except:
            pass
    return ImageFont.load_default()


font_program    = load_font("arialbd.ttf", 34)   # اسم البرنامج
font_desc       = load_font("arial.ttf", 15)     # الوصف
font_loading    = load_font("arial.ttf", 14)     # جار التحميل
font_company    = load_font("arialbd.ttf", 14)   # اسم الشركة
font_manager    = load_font("arial.ttf", 12)     # اسم المسؤول
font_phone      = load_font("arialbd.ttf", 13)   # الهاتف


# ═══════════════════════════════════════════════════════════════════
# إنشاء الصورة
# ═══════════════════════════════════════════════════════════════════
img = Image.new('RGB', (W, H), BG_COLOR)
draw = ImageDraw.Draw(img)

# ─── الحدود الذهبية ───
draw.rectangle([(0, 0), (W - 1, H - 1)], outline=GOLD, width=3)
draw.rectangle([(8, 8), (W - 9, H - 9)], outline=GOLD, width=1)

# ═══════════════════════════════════════════════════════════════════
# 1) الشعار (في المنتصف)
# ═══════════════════════════════════════════════════════════════════
if os.path.exists(LOGO_FILE):
    try:
        logo = Image.open(LOGO_FILE).convert('RGBA')
        logo.thumbnail((LOGO_SIZE, LOGO_SIZE), Image.Resampling.LANCZOS)
        logo_x = (W - logo.width) // 2
        logo_y = 35
        img.paste(logo, (logo_x, logo_y), logo)
    except Exception as e:
        print(f"⚠️ فشل تحميل الشعار: {e}")
else:
    # نص بديل إذا لم يوجد شعار
    fallback = "☀️"
    bbox = draw.textbbox((0, 0), fallback, font=font_program)
    fw = bbox[2] - bbox[0]
    draw.text(((W - fw) / 2, 60), fallback, fill=GOLD, font=font_program)

# ═══════════════════════════════════════════════════════════════════
# 2) خط ذهبي فاصل أعلى
# ═══════════════════════════════════════════════════════════════════
y_sep1 = 190
draw.line([(80, y_sep1), (W - 80, y_sep1)], fill=GOLD, width=2)

# ═══════════════════════════════════════════════════════════════════
# 3) اسم البرنامج
# ═══════════════════════════════════════════════════════════════════
name_text = PROGRAM_NAME_EN
bbox = draw.textbbox((0, 0), name_text, font=font_program)
tw = bbox[2] - bbox[0]
draw.text(((W - tw) / 2, 205), name_text, fill=GOLD, font=font_program)

# ═══════════════════════════════════════════════════════════════════
# 4) الوصف العربي
# ═══════════════════════════════════════════════════════════════════
desc_text = ar(PROGRAM_NAME_AR)
bbox = draw.textbbox((0, 0), desc_text, font=font_desc)
dw = bbox[2] - bbox[0]
draw.text(((W - dw) / 2, 250), desc_text, fill=WHITE, font=font_desc)

# ═══════════════════════════════════════════════════════════════════
# 5) خط ذهبي فاصل أسفل
# ═══════════════════════════════════════════════════════════════════
y_sep2 = 285
draw.line([(80, y_sep2), (W - 80, y_sep2)], fill=GOLD, width=2)

# ═══════════════════════════════════════════════════════════════════
# 6) حالة التحميل
# ═══════════════════════════════════════════════════════════════════
loading_text = ar(LOADING_TEXT)
bbox = draw.textbbox((0, 0), loading_text, font=font_loading)
lw = bbox[2] - bbox[0]
draw.text(((W - lw) / 2, 305), loading_text, fill=WHITE, font=font_loading)

# ═══════════════════════════════════════════════════════════════════
# 7) بيانات الشركة (في الأسفل)
# ═══════════════════════════════════════════════════════════════════

# اسم الشركة
company_text = ar(COMPANY_NAME)
bbox = draw.textbbox((0, 0), company_text, font=font_company)
cw = bbox[2] - bbox[0]
draw.text(((W - cw) / 2, 380), company_text, fill=GOLD, font=font_company)

# اسم المسؤول
manager_text = ar(MANAGER_NAME)
bbox = draw.textbbox((0, 0), manager_text, font=font_manager)
mw = bbox[2] - bbox[0]
draw.text(((W - mw) / 2, 408), manager_text, fill=GRAY, font=font_manager)

# رقم الهاتف
phone_text = ar(f"📞 {PHONE}")
bbox = draw.textbbox((0, 0), phone_text, font=font_phone)
pw = bbox[2] - bbox[0]
draw.text(((W - pw) / 2, 432), phone_text, fill=GRAY, font=font_phone)

# ═══════════════════════════════════════════════════════════════════
# 8) الحفظ
# ═══════════════════════════════════════════════════════════════════
output_file = 'splash.png'
img.save(output_file, 'PNG')
print(f"✅ تم إنشاء شاشة Splash: {output_file}")
print(f"📐 الأبعاد: {W} × {H}")
print(f"📁 المكان: {os.path.abspath(output_file)}")
