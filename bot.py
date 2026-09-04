import telebot
from telebot import types

import pandas as pd
import requests
import os
import random
import re
import io
import json
import time
import calendar
import datetime as dt

import arabic_reshaper
from bidi.algorithm import get_display

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# =========================================================
# الإعدادات
# =========================================================

# التوكن الآن يُقرأ من متغير بيئة (Environment Variable)
# بدل ما يكون مكتوبًا صراحة داخل الكود.
#
# طريقة التشغيل:
#
#   Linux / Mac:
#       export BOT_TOKEN="ضع_التوكن_هنا"
#       python3 bot.py
#
#   Windows (PowerShell):
#       $env:BOT_TOKEN="ضع_التوكن_هنا"
#       python bot.py
#
#   أو أنشئ ملف باسم .env بجانب الكود يحتوي سطر واحد:
#       BOT_TOKEN=ضع_التوكن_هنا
#   (يتطلب تثبيت مكتبة python-dotenv: pip install python-dotenv)

try:

    from dotenv import load_dotenv

    load_dotenv()

except ImportError:

    pass

BOT_TOKEN = os.environ.get(
    "BOT_TOKEN"
)

if not BOT_TOKEN:

    raise SystemExit(
        "❌ لم يتم العثور على BOT_TOKEN.\n"
        "الرجاء تعيينه كمتغير بيئة قبل التشغيل، مثال:\n"
        "   export BOT_TOKEN=\"التوكن_الخاص_بك\"\n"
        "أو أنشئ ملف .env يحتوي على:\n"
        "   BOT_TOKEN=التوكن_الخاص_بك"
    )

ADMIN_ID = int(
    os.environ.get(
        "ADMIN_ID",
        "414840027"
    )
)

COMPANY_NAME = "شركة أبراج الماضونة للخدمات الهندسية والمساندة"

FILE_ID = os.environ.get(
    "SALARY_FILE_ID",
    "1BJgLh9gb580c4ZAw_kNCpmna2NvmlaaE"
)

DOWNLOAD_URL = (
    f"https://drive.google.com/uc?"
    f"export=download&id={FILE_ID}&confirm=t"
)

LOCAL_FILE = "current_salary_data.xlsx"

bot = telebot.TeleBot(BOT_TOKEN)


# =========================================================
# مصدر ملف الرواتب (قابل للتغيير من داخل البوت)
# =========================================================
#
# كل شهر ممكن يتغيّر ملف الرواتب (رابط جديد على Google
# Drive أو ملف يُرفع يدويًا)، فبدل تعديل الكود، نخزّن
# مصدر الملف بملف إعدادات صغير (file_config.json) بجانب
# الكود، ونوفر أمرين للإدارة:
#
#   /setlink <رابط أو ID من Google Drive>
#       لتحديد ملف جديد على Drive والمزامنة معه تلقائيًا
#
#   إرسال ملف Excel مباشرة للبوت (كمرفق مستند)
#       لاعتماده كملف الرواتب الحالي فورًا، بدون Drive
#
# =========================================================

CONFIG_FILE = "file_config.json"


def load_file_config():

    if os.path.exists(CONFIG_FILE):

        try:

            with open(
                CONFIG_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                return json.load(f)

        except Exception as e:

            print(
                f"⚠️ تعذرت قراءة ملف إعدادات المصدر: {e}"
            )

    return {}


def save_file_config(cfg):

    try:

        with open(
            CONFIG_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                cfg,
                f,
                ensure_ascii=False,
                indent=2
            )

    except Exception as e:

        print(
            f"⚠️ تعذرت كتابة ملف إعدادات المصدر: {e}"
        )


def build_download_url(file_id):

    return (
        f"https://drive.google.com/uc?"
        f"export=download&id={file_id}&confirm=t"
    )


# حالة مصدر الملف الحالية (تُحمّل من الإعدادات المحفوظة
# إن وجدت، وإلا تُستخدم القيم الافتراضية من متغيرات البيئة)

_saved_cfg = load_file_config()

FILE_STATE = {

    "source": _saved_cfg.get(
        "source",
        "drive"
    ),

    "file_id": _saved_cfg.get(
        "file_id",
        FILE_ID
    ),

    "download_url": _saved_cfg.get(
        "download_url",
        build_download_url(
            _saved_cfg.get(
                "file_id",
                FILE_ID
            )
        )
    ),

    "uploaded_filename": _saved_cfg.get(
        "uploaded_filename",
        None
    ),

    "updated_at": _saved_cfg.get(
        "updated_at",
        None
    )
}


def extract_drive_file_id(text):
    """
    يستخرج معرّف ملف Google Drive من رابط بأي صيغة شائعة،
    أو يعتبر النص نفسه هو المعرّف إذا كان بالشكل الصحيح.
    يرجع None إذا لم يستطع التعرف على معرّف صالح.
    """

    if not text:
        return None

    text = text.strip()

    # صيغة: .../file/d/FILE_ID/...
    m = re.search(
        r"/d/([a-zA-Z0-9_-]{15,})",
        text
    )

    if m:
        return m.group(1)

    # صيغة: ...?id=FILE_ID أو &id=FILE_ID
    m = re.search(
        r"[?&]id=([a-zA-Z0-9_-]{15,})",
        text
    )

    if m:
        return m.group(1)

    # نص عبارة عن المعرّف مباشرة (بدون رابط)
    if re.fullmatch(
        r"[a-zA-Z0-9_-]{15,}",
        text
    ):
        return text

    return None


def set_drive_source(file_id):

    FILE_STATE["source"] = "drive"
    FILE_STATE["file_id"] = file_id
    FILE_STATE["download_url"] = build_download_url(
        file_id
    )
    FILE_STATE["uploaded_filename"] = None
    FILE_STATE["updated_at"] = time.strftime(
        "%Y-%m-%d %H:%M"
    )

    save_file_config(FILE_STATE)


def set_upload_source(original_filename):

    FILE_STATE["source"] = "upload"
    FILE_STATE["uploaded_filename"] = original_filename
    FILE_STATE["updated_at"] = time.strftime(
        "%Y-%m-%d %H:%M"
    )

    save_file_config(FILE_STATE)



# =========================================================
# أيام الأسبوع
# =========================================================

DAY_MAP = {
    "أر": "الأربعاء",
    "ار": "الأربعاء",
    "خ": "الخميس",
    "ج": "الجمعة",
    "س": "السبت",
    "أح": "الأحد",
    "اح": "الأحد",
    "إث": "الإثنين",
    "اث": "الإثنين",
    "ث": "الثلاثاء"
}


# =========================================================
# أسماء الأشهر وأيام الأسبوع بالعربي (لبناء التاريخ الفعلي)
# =========================================================

ARABIC_MONTHS = {
    1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل",
    5: "مايو", 6: "يونيو", 7: "يوليو", 8: "أغسطس",
    9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر"
}

# date.weekday(): الإثنين=0 ... الأحد=6
ARABIC_WEEKDAYS_FULL = [
    "الإثنين", "الثلاثاء", "الأربعاء",
    "الخميس", "الجمعة", "السبت", "الأحد"
]


def get_report_period(df):
    """
    يحاول استخراج (السنة، الشهر) الفعليين من أعمدة تاريخ
    حقيقية داخل شيت الشبكة (يومية / حساب الاضافي)، بدلاً
    من الاعتماد على شهر مكتوب يدويًا بالكود.
    """

    if df is None:
        return None

    for col in df.columns:

        if isinstance(
            col,
            (dt.datetime, dt.date)
        ):

            return (col.year, col.month)

    return None


# =========================================================
# رسائل تحفيزية
# =========================================================

MOTIVATIONAL_TIPS = [

    "💡 تميّزك وإخلاصك في أداء مهامك اليومية هو الركيزة الأساسية لنجاح الشركة وبناء مسيرتك المهنية.",

    "💡 الانضباط والدقة في العمل يعكسان مهنيتك العالية ويصنعان الفارق في جودة خدماتنا أمام العملاء.",

    "💡 كل جهد تبذله بأمانة يرفع من اسم شركتنا ويسهم مباشرة في نمو ونجاح الفريق ككل.",

    "💡 التزامك بمعايير السلامة المهنية وجودة الأداء دليل على حرصك على تمثيل الشركة بأفضل صورة.",

    "💡 العمل بروح المسؤولية والإيجابية هو مفتاح التطور الوظيفي والتقدير المستمر داخل المؤسسة."
]


# =========================================================
# تنظيف الأرقام
# =========================================================

def clean_num(val):

    if pd.isna(val) or val is None:
        return 0.0

    s = str(val).strip()

    if s in [
        "",
        "nan",
        "None",
        "#REF!",
        "#VALUE!",
        "#N/A"
    ]:
        return 0.0

    # تحويل الأرقام العربية
    s = s.translate(
        str.maketrans(
            "٠١٢٣٤٥٦٧٨٩",
            "0123456789"
        )
    )

    try:

        clean_str = re.sub(
            r"[^\d\.\-]",
            "",
            s
        )

        return float(clean_str) if clean_str else 0.0

    except:

        return 0.0


# =========================================================
# تنظيف الاسم العربي
# =========================================================

def clean_arabic(text):

    if pd.isna(text) or text is None:
        return ""

    t = str(text).strip()

    if not t:
        return ""

    # إزالة التطويل وبعض العلامات
    t = re.sub(
        r"[ـ\-_\/\.]",
        "",
        t
    )

    # توحيد الحروف
    t = re.sub(
        r"[إأآا]",
        "ا",
        t
    )

    t = re.sub(
        r"[يى]",
        "ي",
        t
    )

    t = re.sub(
        r"ة",
        "ه",
        t
    )

    t = re.sub(
        r"ؤ",
        "و",
        t
    )

    t = re.sub(
        r"ئ",
        "ي",
        t
    )

    # إزالة مسافات زائدة
    t = re.sub(
        r"\s+",
        " ",
        t
    ).strip().lower()

    # توحيد بعض الكلمات
    t = t.replace(
        "عبد ",
        "عبد"
    )

    t = t.replace(
        "ابو ",
        "ابو"
    )

    words = t.split()

    # أول 3 كلمات للمطابقة
    if len(words) >= 3:
        return "".join(words[:3])

    return "".join(words)


# =========================================================
# استخراج الأرقام فقط
# =========================================================

def get_digits_only(text):

    if pd.isna(text) or text is None:
        return ""

    s = str(text)

    s = s.translate(
        str.maketrans(
            "٠١٢٣٤٥٦٧٨٩",
            "0123456789"
        )
    )

    if s.endswith(".0"):
        s = s[:-2]

    return re.sub(
        r"[^\d]",
        "",
        s
    )


# =========================================================
# النص العربي داخل PDF
# =========================================================

def ar_txt(text):

    return get_display(
        arabic_reshaper.reshape(
            str(text)
        )
    )


# =========================================================
# تحميل ملف Excel
# =========================================================

def sync_data():

    if FILE_STATE["source"] == "upload":

        if os.path.exists(LOCAL_FILE):

            print(
                "ℹ️ المصدر الحالي ملف مرفوع يدويًا - "
                "لا حاجة للتحميل من Drive."
            )

            return True

        print(
            "❌ لا يوجد ملف محلي، ولم يتم رفع ملف بعد."
        )

        return False

    try:

        print(
            "⏳ جاري تحميل أحدث ملف Excel..."
        )

        response = requests.get(
            FILE_STATE["download_url"],
            timeout=30
        )

        if (
            response.status_code == 200
            and len(response.content) > 5000
        ):

            with open(
                LOCAL_FILE,
                "wb"
            ) as f:

                f.write(
                    response.content
                )

            print(
                "✅ تم تحديث ملف Excel."
            )

            return True

        print(
            "❌ لم يتم تحميل الملف بشكل صحيح."
        )

    except Exception as e:

        print(
            f"❌ خطأ أثناء تحميل الملف: {e}"
        )

    return False


# =========================================================
# قراءة شيت بيانات الموظفين
# =========================================================

def load_employee_sheet(file):

    try:

        xl = pd.ExcelFile(file)

        for sheet_name in xl.sheet_names:

            sheet_text = str(
                sheet_name
            )

            if not any(
                x in sheet_text
                for x in [
                    "بيانات",
                    "عمال",
                    "موظفين"
                ]
            ):
                continue

            raw = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=None,
                dtype=str
            )

            header_row = None

            # البحث عن صف العناوين
            for idx, row in raw.iterrows():

                values = [
                    normalize_header(x)
                    for x in row.values
                    if pd.notna(x)
                ]

                text = " ".join(values)

                if (
                    "الاسم" in text
                    and (
                        "الرقم الوطني" in text
                        or "الرقم الشخصي" in text
                        or "الكود" in text
                    )
                ):

                    header_row = idx
                    break

            if header_row is None:
                continue

            df = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=header_row,
                dtype=str
            )

            df.columns = [
                clean_col_name(c)
                for c in df.columns
            ]

            return df

    except Exception as e:

        print(
            f"❌ خطأ في قراءة شيت الموظفين: {e}"
        )

    return None


# =========================================================
# البحث عن الموظف
# =========================================================

def find_employee(
    df_emp,
    user_input
):

    if df_emp is None:
        return None

    query_digits = get_digits_only(
        user_input
    )

    query_text = clean_arabic(
        user_input
    )

    name_col = None
    nat_col = None
    dept_col = None
    job_col = None
    base_sal_col = None

    # -----------------------------------------------------
    # تحديد الأعمدة
    # -----------------------------------------------------

    for col in df_emp.columns:

        col_text = str(
            col
        ).strip()

        if col_text in [
            "الاسم",
            "اسم الموظف",
            "اسم"
        ]:

            name_col = col

        if (
            "الرقم الوطني" in col_text
            or "الرقم الشخصي" in col_text
            or col_text == "الكود"
        ):

            nat_col = col

        if col_text == "القسم":

            dept_col = col

        if col_text in [
            "المسمى",
            "المسمى الوظيفي",
            "الوظيفة"
        ]:

            job_col = col

        if (
            "الراتب الأساسي" in col_text
            or "الراتب الاساسي" in col_text
        ):

            base_sal_col = col

    # -----------------------------------------------------
    # البحث
    # -----------------------------------------------------

    for _, row in df_emp.iterrows():

        # الاسم
        r_name = ""

        if name_col is not None:

            value = row.get(
                name_col,
                ""
            )

            if pd.notna(value):

                r_name = str(
                    value
                ).strip()

        # الرقم الوطني
        r_nat = ""

        if nat_col is not None:

            value = row.get(
                nat_col,
                ""
            )

            r_nat = get_digits_only(
                value
            )

        clean_name = clean_arabic(
            r_name
        )

        matched = False

        # مطابقة الرقم
        if (
            query_digits
            and r_nat
            and query_digits == r_nat
        ):

            matched = True

        # مطابقة الاسم
        elif (
            query_text
            and clean_name
            and len(query_text) > 3
            and (
                query_text in clean_name
                or clean_name in query_text
            )
        ):

            matched = True

        if not matched:
            continue

        # القسم
        dept = "العمليات"

        if dept_col is not None:

            value = row.get(
                dept_col,
                ""
            )

            if pd.notna(value):

                dept = str(
                    value
                ).strip()

        # المسمى
        job = "موظف"

        if job_col is not None:

            value = row.get(
                job_col,
                ""
            )

            if pd.notna(value):

                job = str(
                    value
                ).strip()

        # الراتب الأساسي (احتياطي في حال عدم توفره
        # بشيت الرواتب الشهري)
        base_sal_fallback = 0.0

        if base_sal_col is not None:

            value = row.get(
                base_sal_col,
                ""
            )

            if pd.notna(value):

                base_sal_fallback = clean_num(
                    value
                )

        return {
            "name": r_name,
            "job": job,
            "dept": dept,
            "nat_id": (
                r_nat
                if r_nat
                else user_input
            ),
            "clean_name": clean_name,
            "base_sal_fallback": base_sal_fallback
        }

    return None


# =========================================================
# تطبيع نص العناوين (إزالة التطويل والمسافات الزائدة)
# بعض الشيتات الحقيقية تحتوي على "الاســـــــم" بتطويل
# مما يمنع أي مطابقة نصية مباشرة لكلمة "الاسم"
# =========================================================

def normalize_header(text):

    if text is None:
        return ""

    s = str(text)

    # إزالة حرف التطويل (ـ)
    s = s.replace("ـ", "")

    s = re.sub(r"\s+", " ", s).strip()

    return s


def clean_col_name(col):

    import datetime as _dt

    # أعمدة التاريخ (أيام الشهر) تبقى كما هي كي تستمر
    # دوال قراءة الأيام بالتعرف عليها ككائن تاريخ
    if isinstance(col, (_dt.datetime, _dt.date, pd.Timestamp)):
        return col

    return normalize_header(col)


# =========================================================
# استخراج رقم اليوم (01-31) من قيمة عمود
# تدعم: "01" / "1" / تاريخ Excel كامل (datetime)
# / نص تاريخ مثل "2026-07-01 00:00:00"
# =========================================================

def parse_day_number(val):

    if val is None:
        return None

    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass

    # تاريخ حقيقي (datetime / Timestamp / date)
    import datetime as _dt

    if isinstance(val, (_dt.datetime, _dt.date)):
        return val.day

    if isinstance(val, pd.Timestamp):
        return val.day

    s = str(val).strip()

    if not s or s.lower() in ["nan", "none", "nat"]:
        return None

    # تحويل الأرقام العربية
    s = s.translate(
        str.maketrans(
            "٠١٢٣٤٥٦٧٨٩",
            "0123456789"
        )
    )

    # رقم يوم مباشر: "1" .. "31" أو "01" .. "31"
    if re.fullmatch(r"\d{1,2}", s):

        n = int(s)

        if 1 <= n <= 31:
            return n

        return None

    # نص تاريخ مثل "2026-07-01" أو "2026-07-01 00:00:00"
    m = re.match(
        r"^\d{4}-\d{2}-(\d{2})",
        s
    )

    if m:
        return int(m.group(1))

    # نص تاريخ مثل "01/07/2026" أو "01-07-2026"
    m = re.match(
        r"^(\d{1,2})[/-]\d{1,2}[/-]\d{2,4}",
        s
    )

    if m:
        n = int(m.group(1))

        if 1 <= n <= 31:
            return n

    return None


# =========================================================
# البحث عن شيت فيه الأيام 01 - 31
# =========================================================

def load_grid_sheet(
    file,
    sheet_keywords
):

    try:

        xl = pd.ExcelFile(file)

        for sheet_name in xl.sheet_names:

            sheet_text = str(
                sheet_name
            )

            if not any(
                keyword in sheet_text
                for keyword in sheet_keywords
            ):
                continue

            print(
                f"🔎 قراءة الشيت: {sheet_name}"
            )

            raw = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=None,
                dtype=str
            )

            header_row = None

            # -------------------------------------------------
            # البحث عن صف العناوين
            # يجب أن يحتوي على الاسم و01 و02 و03
            # -------------------------------------------------

            for idx, row in raw.iterrows():

                values = []

                for x in row.values:

                    if pd.notna(x):

                        value = normalize_header(
                            x
                        )

                        values.append(
                            value
                        )

                has_name = any(
                    x in [
                        "الاسم",
                        "اسم",
                        "اسم الموظف"
                    ]
                    for x in values
                )

                # نستخرج أرقام الأيام من كل خلايا الصف
                # (تدعم "01"/"1" وأيضاً التواريخ الكاملة
                # التي يخزنها Excel لأعمدة الأيام)
                day_numbers = set()

                for x in row.values:

                    d = parse_day_number(x)

                    if d is not None:
                        day_numbers.add(d)

                has_01 = 1 in day_numbers
                has_02 = 2 in day_numbers
                has_03 = 3 in day_numbers

                if (
                    has_name
                    and has_01
                    and has_02
                    and has_03
                ):

                    header_row = idx
                    break

            if header_row is None:

                print(
                    f"⚠️ لم يتم العثور على صف عناوين في {sheet_name}"
                )

                continue

            # -------------------------------------------------
            # قراءة البيانات بعد صف العناوين
            # -------------------------------------------------

            df = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=header_row,
                dtype=str
            )

            # تنظيف أسماء الأعمدة (مع الحفاظ على أعمدة
            # التاريخ ككائنات تاريخ لضمان عمل parse_day_number)
            df.columns = [
                clean_col_name(c)
                for c in df.columns
            ]

            print(
                f"✅ تم العثور على صف العناوين: {header_row + 1}"
            )

            print(
                "📌 الأعمدة:",
                list(df.columns)[:35]
            )

            return df

    except Exception as e:

        print(
            f"❌ خطأ في قراءة الشيت: {e}"
        )

    return None


# =========================================================
# البحث عن صف الموظف داخل شيت اليومية / الإضافي
# =========================================================

def find_employee_row(
    df,
    target_clean_name
):

    if df is None or df.empty:
        return None

    # -----------------------------------------------------
    # تحديد عمود الاسم
    # -----------------------------------------------------

    name_column = None

    for col in df.columns:

        col_text = str(
            col
        ).strip()

        if col_text in [
            "الاسم",
            "اسم",
            "اسم الموظف"
        ]:

            name_column = col
            break

    # إذا لم نجد اسم
    if name_column is None:

        if len(df.columns) > 0:

            name_column = df.columns[0]

    if name_column is None:
        return None

    # -----------------------------------------------------
    # البحث
    # -----------------------------------------------------

    for _, row in df.iterrows():

        value = row.get(
            name_column,
            ""
        )

        if pd.isna(value):
            continue

        employee_name = str(
            value
        ).strip()

        if not employee_name:
            continue

        clean_name = clean_arabic(
            employee_name
        )

        if not clean_name:
            continue

        if (
            target_clean_name in clean_name
            or clean_name in target_clean_name
        ):

            print(
                f"✅ تم العثور على الموظف: {employee_name}"
            )

            return row

    print(
        "⚠️ لم يتم العثور على الموظف داخل الشيت."
    )

    return None


# =========================================================
# استخراج بيانات الأيام 01 - 31
# =========================================================

def extract_grid_data(
    df,
    target_clean_name
):

    days_data = {}

    if df is None or df.empty:

        return days_data

    employee_row = find_employee_row(
        df,
        target_clean_name
    )

    if employee_row is None:

        # إنشاء 01 - 31 فارغة
        for day in range(1, 32):

            days_data[
                f"{day:02d}"
            ] = ""

        return days_data

    # -----------------------------------------------------
    # قراءة الأيام
    # -----------------------------------------------------

    for day in range(1, 32):

        day_str = f"{day:02d}"

        column = None

        for col in df.columns:

            # يدعم "01"/"1" وأيضاً أعمدة التاريخ الكاملة
            # (datetime) التي يستخدمها Excel لأيام الشهر
            if parse_day_number(col) == day:

                column = col
                break

        if column is None:

            days_data[day_str] = ""
            continue

        value = employee_row.get(
            column,
            ""
        )

        if pd.isna(value):

            value = ""

        value = str(
            value
        ).strip()

        if value.lower() in [
            "nan",
            "none"
        ]:

            value = ""

        days_data[
            day_str
        ] = value

    return days_data


# =========================================================
# قراءة شيت الرواتب
# =========================================================

def load_salary_sheet(file):

    try:

        xl = pd.ExcelFile(file)

        for sheet_name in xl.sheet_names:

            if not any(
                x in str(sheet_name)
                for x in [
                    "ملخص رواتب",
                    "رواتب",
                    "راتب"
                ]
            ):
                continue

            raw = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=None,
                dtype=str
            )

            header_row = None

            for idx, row in raw.iterrows():

                values = [
                    normalize_header(x)
                    for x in row.values
                    if pd.notna(x)
                ]

                text = " ".join(values)

                if (
                    "الاسم" in text
                    and "الراتب" in text
                ):

                    header_row = idx
                    break

            if header_row is None:
                continue

            df = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=header_row,
                dtype=str
            )

            df.columns = [
                normalize_header(c)
                for c in df.columns
            ]

            print(
                f"💰 تم استخدام شيت الرواتب: {sheet_name}"
            )

            return df

    except Exception as e:

        print(
            f"❌ خطأ في شيت الرواتب: {e}"
        )

    return None


# =========================================================
# الحصول على بيانات الراتب
# =========================================================

def get_salary_data(
    df_sal,
    employee
):

    result = {}

    if df_sal is None:
        return result

    name_col = None

    for col in df_sal.columns:

        if str(col).strip() in [
            "الاسم",
            "اسم الموظف",
            "اسم"
        ]:

            name_col = col
            break

    if name_col is None:

        name_col = df_sal.columns[0]

    target_name = employee[
        "clean_name"
    ]

    for _, row in df_sal.iterrows():

        value = row.get(
            name_col,
            ""
        )

        if pd.isna(value):
            continue

        row_name = str(
            value
        ).strip()

        clean_name = clean_arabic(
            row_name
        )

        if (
            target_name in clean_name
            or clean_name in target_name
        ):

            result = row.to_dict()

            print(
                f"💰 تم العثور على راتب الموظف: {row_name}"
            )

            break

    return result


# =========================================================
# جلب كل بيانات الموظف
# =========================================================

def fetch_employee_data(
    user_input
):

    # -----------------------------------------------------
    # التأكد من وجود الملف
    # -----------------------------------------------------

    if not os.path.exists(
        LOCAL_FILE
    ):

        sync_data()

    if not os.path.exists(
        LOCAL_FILE
    ):

        return None

    # -----------------------------------------------------
    # شيت الموظفين
    # -----------------------------------------------------

    df_emp = load_employee_sheet(
        LOCAL_FILE
    )

    if df_emp is None:

        print(
            "❌ لم يتم العثور على شيت بيانات الموظفين."
        )

        return None

    employee = find_employee(
        df_emp,
        user_input
    )

    if employee is None:

        print(
            "❌ الموظف غير موجود."
        )

        return None

    # -----------------------------------------------------
    # شيت الرواتب
    # -----------------------------------------------------

    df_sal = load_salary_sheet(
        LOCAL_FILE
    )

    sal_data = get_salary_data(
        df_sal,
        employee
    )

    # -----------------------------------------------------
    # شيت اليومية
    # -----------------------------------------------------

    df_att = load_grid_sheet(
        LOCAL_FILE,
        [
            "يومية",
            "يومية د",
            "حضور"
        ]
    )

    att_dict = extract_grid_data(
        df_att,
        employee["clean_name"]
    )

    # -----------------------------------------------------
    # شيت الإضافي
    # -----------------------------------------------------

    df_ot = load_grid_sheet(
        LOCAL_FILE,
        [
            "حساب الاضافي",
            "حساب الإضافي",
            "الاضافي",
            "الإضافي"
        ]
    )

    ot_dict = extract_grid_data(
        df_ot,
        employee["clean_name"]
    )

    # -----------------------------------------------------
    # تحديد الشهر/السنة الفعليين من التواريخ الحقيقية
    # داخل الشيت (بدل الاعتماد على شهر مكتوب يدويًا)
    # -----------------------------------------------------

    period = (
        get_report_period(df_att)
        or get_report_period(df_ot)
    )

    if period:

        report_year, report_month = period

        days_in_month = calendar.monthrange(
            report_year,
            report_month
        )[1]

    else:

        now = dt.datetime.now()

        report_year = now.year
        report_month = now.month

        days_in_month = 31

    report_month_name = ARABIC_MONTHS.get(
        report_month,
        str(report_month)
    )

    # -----------------------------------------------------
    # حساب الدوام
    # -----------------------------------------------------

    calc_work_days = 0
    calc_leaves = 0
    calc_absent = 0
    total_ot_days = 0

    att_lines = []
    ot_lines = []
    daily_rows = []

    # -----------------------------------------------------
    # أيام الشهر الفعلية (حسب عدد أيام الشهر الحقيقي)
    # -----------------------------------------------------

    for day in range(1, days_in_month + 1):

        day_str = f"{day:02d}"

        # ================================================
        # الدوام
        # ================================================

        status = str(
            att_dict.get(
                day_str,
                ""
            )
        ).strip()

        # تنظيف
        status = status.replace(
            " ",
            ""
        )

        # --------------------------------
        # د = دوام
        # --------------------------------

        if status == "د":

            txt_status = "دوام فعلي ✅"

            calc_work_days += 1

        # --------------------------------
        # م = إجازة
        # --------------------------------

        elif status == "م":

            txt_status = "إجازة 🏖"

            calc_leaves += 1

        # --------------------------------
        # غ = غياب
        # --------------------------------

        elif status == "غ":

            txt_status = "غياب ❌"

            calc_absent += 1

        # --------------------------------
        # خلية فارغة
        # --------------------------------

        elif status == "":

            txt_status = "غير مسجل ⚠️"

        # --------------------------------
        # أي حالة أخرى
        # --------------------------------

        else:

            txt_status = (
                f"حالة ({status})"
            )

        att_lines.append(
            f"🗓 `{day_str}`: {txt_status}"
        )

        # ================================================
        # الإضافي
        # ================================================

        ot_val = clean_num(
            ot_dict.get(
                day_str,
                0
            )
        )

        if ot_val > 0:

            total_ot_days += 1

            ot_txt = (
                f"⏱ **{ot_val:g} ساعات إضافي**"
            )

        else:

            ot_txt = (
                "لا يوجد إضافي (0 س) ➖"
            )

        ot_lines.append(
            f"🗓 `{day_str}`: {ot_txt}"
        )

        # ================================================
        # سجل يومي مُهيكل (للاستخدام بجدول الـ PDF)
        # ================================================

        status_short_map = {
            "د": "دوام",
            "م": "إجازة",
            "غ": "غياب"
        }

        status_short = status_short_map.get(
            status,
            (
                "-"
                if status == ""
                else status
            )
        )

        try:

            day_date = dt.date(
                report_year,
                report_month,
                day
            )

            date_str = day_date.strftime(
                "%d/%m/%Y"
            )

            weekday_name = ARABIC_WEEKDAYS_FULL[
                day_date.weekday()
            ]

        except ValueError:

            date_str = f"{day_str}/{report_month:02d}"
            weekday_name = ""

        daily_rows.append({
            "day": day,
            "date": date_str,
            "weekday": weekday_name,
            "status": status_short,
            "ot_hours": ot_val
        })

    # =====================================================
    # البيانات المالية
    # =====================================================

    basic_sal = clean_num(
        sal_data.get(
            "الراتب الاساسي",
            sal_data.get(
                "الراتب الأساسي",
                sal_data.get(
                    "الراتب",
                    0
                )
            )
        )
    )

    # احتياطي: إن لم يوجد الراتب الأساسي بشيت الرواتب
    # الشهري، نستخدم القيمة من شيت بيانات العمال (المصدر
    # الأساسي لبيانات الموظف الثابتة)
    if basic_sal == 0:

        basic_sal = employee.get(
            "base_sal_fallback",
            0.0
        )

    # -----------------------------------------------------
    # ساعات الإضافي
    # -----------------------------------------------------

    ot_n_hrs = clean_num(
        sal_data.get(
            "عدد ساعات إضافي عادي",
            sal_data.get(
                "عدد ساعات الاضافي العادي",
                0
            )
        )
    )

    ot_h_hrs = clean_num(
        sal_data.get(
            "عدد ساعات إضافي جمع",
            sal_data.get(
                "عدد ساعات إضافي عطل",
                0
            )
        )
    )

    # -----------------------------------------------------
    # قيمة الإضافي
    # -----------------------------------------------------

    ot_n_val = clean_num(
        sal_data.get(
            "استحقاق العادي",
            sal_data.get(
                "استحقاق اضافي عادي",
                0
            )
        )
    )

    ot_h_val = clean_num(
        sal_data.get(
            "استحقاق الجمع",
            sal_data.get(
                "استحقاق إضافي جمع",
                sal_data.get(
                    "استحقاق العطل",
                    0
                )
            )
        )
    )

    # -----------------------------------------------------
    # مجموع الإضافي
    # -----------------------------------------------------

    ot_total = clean_num(
        sal_data.get(
            "مجموع الاضافي",
            sal_data.get(
                "مجموع الإضافي",
                ot_n_val + ot_h_val
            )
        )
    )

    # إذا كان مجموع الإضافي غير موجود
    if ot_total == 0:

        ot_total = (
            ot_n_val
            + ot_h_val
        )

    # -----------------------------------------------------
    # المكافآت
    # -----------------------------------------------------

    bonus = clean_num(
        sal_data.get(
            "مكافأة",
            sal_data.get(
                "مكافآت",
                sal_data.get(
                    "حوافز",
                    0
                )
            )
        )
    )

    # -----------------------------------------------------
    # الضمان
    # -----------------------------------------------------

    ssc = clean_num(
        sal_data.get(
            "الضمان",
            sal_data.get(
                "الضمان الاجتماعي",
                0
            )
        )
    )

    # -----------------------------------------------------
    # الحسميات
    # -----------------------------------------------------

    deductions = clean_num(
        sal_data.get(
            "حسم",
            sal_data.get(
                "الحسم",
                sal_data.get(
                    "حسومات",
                    0
                )
            )
        )
    )

    # -----------------------------------------------------
    # السلف
    # -----------------------------------------------------

    advances = clean_num(
        sal_data.get(
            "السلف",
            sal_data.get(
                "السلفة",
                0
            )
        )
    )

    # =====================================================
    # صافي الراتب
    # نعتمد القيمة الجاهزة من شيت الرواتب مباشرة (وهي
    # القيمة المدققة فعلياً من الشركة) بدلاً من إعادة حسابها،
    # لضمان تطابق 100% مع الشيت. إن لم تتوفر، نحسبها يدوياً.
    # =====================================================

    net_salary_raw = sal_data.get(
        "صافي الراتب",
        None
    )

    if (
        net_salary_raw is not None
        and str(net_salary_raw).strip() not in [
            "", "nan", "None"
        ]
    ):

        net_salary = round(
            clean_num(net_salary_raw),
            2
        )

    else:

        net_salary = round(
            basic_sal
            + ot_total
            + bonus
            - ssc
            - advances
            - deductions,
            2
        )

    # =====================================================
    # تحديث بيانات الموظف
    # =====================================================

    employee.update({

        "basic_sal": basic_sal,

        "ot_n_hrs": ot_n_hrs,

        "ot_h_hrs": ot_h_hrs,

        "ot_n_val": ot_n_val,

        "ot_h_val": ot_h_val,

        "ot_total": ot_total,

        "bonus": bonus,

        "ssc": ssc,

        "deductions": deductions,

        "advances": advances,

        "net_salary": net_salary,

        "work_days": calc_work_days,

        "leaves": calc_leaves,

        "absent": calc_absent,

        "ot_days": total_ot_days,

        "att_report": "\n".join(
            att_lines
        ),

        "ot_report": "\n".join(
            ot_lines
        ),

        "daily_rows": daily_rows,

        "report_year": report_year,

        "report_month": report_month,

        "report_month_name": report_month_name,

        "days_in_month": days_in_month
    })

    return employee


# =========================================================
# الخطوط العربية
# =========================================================
#
# نبحث أولاً عن خط Amiri المرفق مع البوت (يعمل على أي
# نظام تشغيل - Windows/Linux/Mac)، وإن لم يوجد نجرب بعض
# مسارات الخطوط الشائعة على النظام كحل احتياطي.
# =========================================================

_FONTS_REGISTERED = False
_FONT_REGULAR_NAME = "Helvetica"
_FONT_BOLD_NAME = "Helvetica-Bold"


def register_arabic_fonts():

    global _FONTS_REGISTERED
    global _FONT_REGULAR_NAME
    global _FONT_BOLD_NAME

    if _FONTS_REGISTERED:
        return (
            _FONT_REGULAR_NAME,
            _FONT_BOLD_NAME
        )

    base_dir = os.path.dirname(
        os.path.abspath(__file__)
    )

    candidates = [

        # الخط المرفق مع المشروع (الأفضل - يعمل بكل مكان)
        (
            os.path.join(
                base_dir, "fonts", "Amiri-Regular.ttf"
            ),
            os.path.join(
                base_dir, "fonts", "Amiri-Bold.ttf"
            )
        ),

        # Windows
        (
            "C:/Windows/Fonts/tahoma.ttf",
            "C:/Windows/Fonts/tahomabd.ttf"
        ),

        # Linux (خطوط عربية شائعة قد تكون مثبتة)
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
        ),

        (
            "/usr/share/fonts/truetype/kacst/KacstOne.ttf",
            "/usr/share/fonts/truetype/kacst/KacstOne.ttf"
        ),

        # macOS
        (
            "/System/Library/Fonts/Supplemental/Tahoma.ttf",
            "/System/Library/Fonts/Supplemental/Tahoma Bold.ttf"
        )
    ]

    for regular_path, bold_path in candidates:

        try:

            if not os.path.exists(regular_path):
                continue

            pdfmetrics.registerFont(
                TTFont("Arabic", regular_path)
            )

            if os.path.exists(bold_path):

                pdfmetrics.registerFont(
                    TTFont("Arabic-Bold", bold_path)
                )

            else:

                pdfmetrics.registerFont(
                    TTFont("Arabic-Bold", regular_path)
                )

            _FONT_REGULAR_NAME = "Arabic"
            _FONT_BOLD_NAME = "Arabic-Bold"
            _FONTS_REGISTERED = True

            print(
                f"✅ تم تحميل الخط العربي: {regular_path}"
            )

            break

        except Exception as e:

            print(
                f"⚠️ تعذر تحميل الخط {regular_path}: {e}"
            )

            continue

    if not _FONTS_REGISTERED:

        print(
            "⚠️ لم يتم العثور على أي خط عربي - "
            "قد لا يظهر النص العربي بشكل صحيح بملف الـ PDF. "
            "ضع ملفي Amiri-Regular.ttf و Amiri-Bold.ttf "
            "داخل مجلد fonts/ بجانب bot.py."
        )

    return (
        _FONT_REGULAR_NAME,
        _FONT_BOLD_NAME
    )


# =========================================================
# اختصار اسم اليوم (للجدول اليومي المضغوط)
# =========================================================

SHORT_WEEKDAY = {
    "الإثنين": "إث",
    "الثلاثاء": "ث",
    "الأربعاء": "أر",
    "الخميس": "خ",
    "الجمعة": "ج",
    "السبت": "س",
    "الأحد": "أح"
}


# =========================================================
# ألوان حالة الدوام (لتلوين خلايا الجدول اليومي)
# =========================================================

STATUS_COLORS = {
    "دوام": colors.HexColor("#e8f5e9"),
    "إجازة": colors.HexColor("#fff8e1"),
    "غياب": colors.HexColor("#ffebee"),
    "-": colors.HexColor("#f5f5f5")
}

STATUS_TEXT_COLORS = {
    "دوام": colors.HexColor("#1b5e20"),
    "إجازة": colors.HexColor("#e65100"),
    "غياب": colors.HexColor("#b71c1c"),
    "-": colors.HexColor("#757575")
}


# =========================================================
# بناء جدول يومي واحد (نصف الشهر) - عمود فرعي
# =========================================================

def build_half_month_table(
    rows,
    font_regular,
    font_bold
):

    table_data = [[
        ar_txt("إضافي"),
        ar_txt("الحالة"),
        ar_txt("اليوم"),
        ar_txt("التاريخ"),
        ar_txt("#")
    ]]

    style_commands = [

        (
            "FONTNAME",
            (0, 0),
            (-1, -1),
            font_regular
        ),

        (
            "FONTNAME",
            (0, 0),
            (-1, 0),
            font_bold
        ),

        (
            "FONTSIZE",
            (0, 0),
            (-1, -1),
            7.5
        ),

        (
            "ALIGN",
            (0, 0),
            (-1, -1),
            "CENTER"
        ),

        (
            "VALIGN",
            (0, 0),
            (-1, -1),
            "MIDDLE"
        ),

        (
            "BACKGROUND",
            (0, 0),
            (-1, 0),
            colors.HexColor("#1b365d")
        ),

        (
            "TEXTCOLOR",
            (0, 0),
            (-1, 0),
            colors.white
        ),

        (
            "GRID",
            (0, 0),
            (-1, -1),
            0.4,
            colors.HexColor("#cccccc")
        ),

        (
            "TOPPADDING",
            (0, 0),
            (-1, -1),
            3
        ),

        (
            "BOTTOMPADDING",
            (0, 0),
            (-1, -1),
            3
        )
    ]

    for i, row in enumerate(rows, start=1):

        weekday_short = SHORT_WEEKDAY.get(
            row["weekday"],
            row["weekday"]
        )

        ot_text = (
            f'{row["ot_hours"]:g}'
            if row["ot_hours"] > 0
            else "-"
        )

        table_data.append([
            ot_text,
            ar_txt(row["status"]),
            ar_txt(weekday_short),
            row["date"][:5],
            str(row["day"])
        ])

        bg = STATUS_COLORS.get(
            row["status"],
            colors.white
        )

        fg = STATUS_TEXT_COLORS.get(
            row["status"],
            colors.black
        )

        style_commands.append((
            "BACKGROUND",
            (1, i),
            (1, i),
            bg
        ))

        style_commands.append((
            "TEXTCOLOR",
            (1, i),
            (1, i),
            fg
        ))

        if row["ot_hours"] > 0:

            style_commands.append((
                "TEXTCOLOR",
                (0, i),
                (0, i),
                colors.HexColor("#c62828")
            ))

            style_commands.append((
                "FONTNAME",
                (0, i),
                (0, i),
                font_bold
            ))

    t = Table(
        table_data,
        colWidths=[32, 55, 30, 48, 20]
    )

    t.setStyle(
        TableStyle(style_commands)
    )

    return t


# =========================================================
# إنشاء PDF
# =========================================================

def generate_professional_pdf(
    data
):

    pdf_buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=40,
        bottomMargin=40
    )

    elements = []

    # -----------------------------------------------------
    # الخط
    # -----------------------------------------------------

    font_regular, font_bold = register_arabic_fonts()

    period_label = (
        f'{data.get("report_month_name", "")} '
        f'{data.get("report_year", "")}'
    )

    # -----------------------------------------------------
    # Styles
    # -----------------------------------------------------

    title_style = ParagraphStyle(
        "TitleArabic",
        fontName=font_bold,
        fontSize=16,
        alignment=1,
        textColor=colors.white,
        leading=20
    )

    subtitle_style = ParagraphStyle(
        "SubTitle",
        fontName=font_regular,
        fontSize=11,
        alignment=1,
        textColor=colors.white,
        leading=15
    )

    section_title_style = ParagraphStyle(
        "SectionTitle",
        fontName=font_bold,
        fontSize=11,
        alignment=2,
        textColor=colors.HexColor("#1b365d"),
        spaceAfter=6,
        spaceBefore=4
    )

    # -----------------------------------------------------
    # رأس الصفحة (بانر ملوّن باسم الشركة والفترة)
    # -----------------------------------------------------

    header_table = Table(
        [
            [Paragraph(ar_txt(COMPANY_NAME), title_style)],
            [Paragraph(
                ar_txt(
                    f"قسيمة كشف راتب شهري - شهر {period_label}"
                ),
                subtitle_style
            )]
        ],
        colWidths=[519]
    )

    header_table.setStyle(
        TableStyle([

            (
                "BACKGROUND",
                (0, 0),
                (-1, -1),
                colors.HexColor("#1b365d")
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, 0),
                12
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, 0),
                4
            ),

            (
                "TOPPADDING",
                (0, 1),
                (-1, 1),
                2
            ),

            (
                "BOTTOMPADDING",
                (0, 1),
                (-1, 1),
                12
            )
        ])
    )

    elements.append(header_table)

    elements.append(Spacer(1, 18))

    # -----------------------------------------------------
    # بيانات الموظف
    # -----------------------------------------------------

    emp_info_data = [

        [
            ar_txt(data["name"]),
            ar_txt("اسم الموظف:")
        ],

        [
            ar_txt(f'{data["dept"]} / {data["job"]}'),
            ar_txt("القسم والمسمى:")
        ],

        [
            data["nat_id"],
            ar_txt("الرقم الوطني / الكود:")
        ]
    ]

    t_info = Table(
        emp_info_data,
        colWidths=[369, 150]
    )

    t_info.setStyle(
        TableStyle([

            (
                "FONTNAME",
                (0, 0),
                (-1, -1),
                font_regular
            ),

            (
                "FONTNAME",
                (1, 0),
                (1, -1),
                font_bold
            ),

            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                10.5
            ),

            (
                "ALIGN",
                (0, 0),
                (0, -1),
                "RIGHT"
            ),

            (
                "ALIGN",
                (1, 0),
                (1, -1),
                "RIGHT"
            ),

            (
                "TEXTCOLOR",
                (1, 0),
                (1, -1),
                colors.HexColor("#1b365d")
            ),

            (
                "BACKGROUND",
                (0, 0),
                (-1, -1),
                colors.HexColor("#f8f9fa")
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                7
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                7
            ),

            (
                "LINEBELOW",
                (0, 0),
                (-1, -2),
                0.5,
                colors.white
            )
        ])
    )

    elements.append(t_info)

    elements.append(Spacer(1, 16))

    # -----------------------------------------------------
    # البيانات المالية
    # -----------------------------------------------------

    elements.append(
        Paragraph(
            ar_txt("البيانات المالية:"),
            section_title_style
        )
    )

    finance_data = [

        [
            ar_txt("القيمة (د.أ)"),
            ar_txt("الاقتطاعات"),
            ar_txt("القيمة (د.أ)"),
            ar_txt("الاستحقاقات")
        ],

        [
            f'{data["ssc"]:.2f}',
            ar_txt("الضمان الاجتماعي"),
            f'{data["basic_sal"]:.2f}',
            ar_txt("الراتب الأساسي")
        ],

        [
            f'{data["advances"]:.2f}',
            ar_txt("السلف المسحوبة"),
            f'{data["ot_total"]:.2f}',
            ar_txt("إجمالي العمل الإضافي")
        ],

        [
            f'{data["deductions"]:.2f}',
            ar_txt("اقتطاعات وحسومات"),
            f'{data["bonus"]:.2f}',
            ar_txt("مكافآت وحوافز")
        ]
    ]

    total_deductions = (
        data["ssc"]
        + data["advances"]
        + data["deductions"]
    )

    total_earnings = (
        data["basic_sal"]
        + data["ot_total"]
        + data["bonus"]
    )

    finance_data.append([
        f"{total_deductions:.2f}",
        ar_txt("إجمالي الاقتطاعات"),
        f"{total_earnings:.2f}",
        ar_txt("إجمالي الاستحقاقات")
    ])

    t_finance = Table(
        finance_data,
        colWidths=[75, 175, 75, 194]
    )

    t_finance.setStyle(
        TableStyle([

            (
                "FONTNAME",
                (0, 0),
                (-1, -1),
                font_regular
            ),

            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                9.5
            ),

            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "CENTER"
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "BACKGROUND",
                (2, 0),
                (3, 0),
                colors.HexColor("#2e7d32")
            ),

            (
                "TEXTCOLOR",
                (2, 0),
                (3, 0),
                colors.white
            ),

            (
                "FONTNAME",
                (2, 0),
                (3, 0),
                font_bold
            ),

            (
                "BACKGROUND",
                (0, 0),
                (1, 0),
                colors.HexColor("#c62828")
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (1, 0),
                colors.white
            ),

            (
                "FONTNAME",
                (0, 0),
                (1, 0),
                font_bold
            ),

            (
                "BACKGROUND",
                (0, -1),
                (-1, -1),
                colors.HexColor("#e0e0e0")
            ),

            (
                "FONTNAME",
                (0, -1),
                (-1, -1),
                font_bold
            ),

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.75,
                colors.HexColor("#cccccc")
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                6
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                6
            )
        ])
    )

    elements.append(t_finance)

    elements.append(Spacer(1, 12))

    # -----------------------------------------------------
    # صافي الراتب
    # -----------------------------------------------------

    net_data = [[
        f'{data["net_salary"]:.2f} JOD',
        ar_txt("صافي الراتب المستحق للصرف:")
    ]]

    t_net = Table(
        net_data,
        colWidths=[150, 369]
    )

    t_net.setStyle(
        TableStyle([

            (
                "FONTNAME",
                (0, 0),
                (-1, -1),
                font_bold
            ),

            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                14
            ),

            (
                "ALIGN",
                (1, 0),
                (1, 0),
                "RIGHT"
            ),

            (
                "ALIGN",
                (0, 0),
                (0, 0),
                "CENTER"
            ),

            (
                "BACKGROUND",
                (0, 0),
                (-1, -1),
                colors.HexColor("#e8f5e9")
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (-1, -1),
                colors.HexColor("#1b5e20")
            ),

            (
                "BOX",
                (0, 0),
                (-1, -1),
                1.5,
                colors.HexColor("#4caf50")
            ),

            (
                "PADDING",
                (0, 0),
                (-1, -1),
                11
            )
        ])
    )

    elements.append(t_net)

    elements.append(Spacer(1, 18))

    # -----------------------------------------------------
    # ملخص الدوام
    # -----------------------------------------------------

    elements.append(
        Paragraph(
            ar_txt("ملخص سجل الدوام:"),
            section_title_style
        )
    )

    att_summary_data = [[
        ar_txt(f'أيام الغياب: {data["absent"]}'),
        ar_txt(f'الإجازات: {data["leaves"]}'),
        ar_txt(f'أيام الدوام الفعلي: {data["work_days"]}'),
        ar_txt(f'ساعات الإضافي: {data["ot_n_hrs"] + data["ot_h_hrs"]:g}')
    ]]

    t_att = Table(
        att_summary_data,
        colWidths=[130, 130, 130, 129]
    )

    t_att.setStyle(
        TableStyle([

            (
                "FONTNAME",
                (0, 0),
                (-1, -1),
                font_regular
            ),

            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                9.5
            ),

            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "CENTER"
            ),

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor("#dddddd")
            ),

            (
                "BACKGROUND",
                (0, 0),
                (-1, -1),
                colors.HexColor("#fafafa")
            ),

            (
                "PADDING",
                (0, 0),
                (-1, -1),
                8
            )
        ])
    )

    elements.append(t_att)

    elements.append(Spacer(1, 40))

    # -----------------------------------------------------
    # التوقيع
    # -----------------------------------------------------

    sig_data = [[
        ar_txt("توقيع المحاسب / الإدارة"),
        ar_txt("توقيع الموظف بالاستلام")
    ]]

    t_sig = Table(
        sig_data,
        colWidths=[230, 230]
    )

    t_sig.setStyle(
        TableStyle([

            (
                "FONTNAME",
                (0, 0),
                (-1, -1),
                font_regular
            ),

            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "CENTER"
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (-1, -1),
                colors.gray
            ),

            (
                "LINEABOVE",
                (0, 0),
                (0, 0),
                1,
                colors.black
            ),

            (
                "LINEABOVE",
                (1, 0),
                (1, 0),
                1,
                colors.black
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                10
            )
        ])
    )

    elements.append(t_sig)

    # =======================================================
    # الصفحة الثانية: كشف الدوام والإضافي التفصيلي اليومي
    # (التاريخ + اسم اليوم + الحالة + ساعات الإضافي)
    # كلها بنفس الصفحة كما هو مطلوب
    # =======================================================

    elements.append(PageBreak())

    page2_title_style = ParagraphStyle(
        "Page2Title",
        fontName=font_bold,
        fontSize=13,
        alignment=1,
        textColor=colors.HexColor("#1b365d"),
        spaceAfter=4
    )

    page2_sub_style = ParagraphStyle(
        "Page2Sub",
        fontName=font_regular,
        fontSize=9.5,
        alignment=1,
        textColor=colors.HexColor("#555555"),
        spaceAfter=14
    )

    elements.append(
        Paragraph(
            ar_txt(
                "كشف سجل الدوام وساعات العمل الإضافي اليومي"
            ),
            page2_title_style
        )
    )

    elements.append(
        Paragraph(
            ar_txt(
                f'{data["name"]}  -  شهر {period_label}'
            ),
            page2_sub_style
        )
    )

    daily_rows = data.get("daily_rows", [])

    half = (len(daily_rows) + 1) // 2

    first_half = daily_rows[:half]
    second_half = daily_rows[half:]

    table_first = build_half_month_table(
        first_half,
        font_regular,
        font_bold
    )

    table_second = (
        build_half_month_table(
            second_half,
            font_regular,
            font_bold
        )
        if second_half
        else None
    )

    if table_second is not None:

        # الجدول الأول (أيام 1 - N) يظهر على اليمين
        # والثاني على اليسار، بما يوافق اتجاه القراءة العربي
        combined = Table(
            [[table_second, table_first]],
            colWidths=[195, 195]
        )

        combined.setStyle(
            TableStyle([

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP"
                ),

                (
                    "ALIGN",
                    (0, 0),
                    (-1, -1),
                    "CENTER"
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                )
            ])
        )

        elements.append(combined)

    else:

        elements.append(table_first)

    elements.append(Spacer(1, 16))

    # -----------------------------------------------------
    # دليل الألوان (Legend)
    # -----------------------------------------------------

    legend_style = ParagraphStyle(
        "Legend",
        fontName=font_regular,
        fontSize=8.5,
        alignment=1,
        textColor=colors.HexColor("#555555")
    )

    elements.append(
        Paragraph(
            ar_txt(
                "دوام = حضور فعلي   |   إجازة = يوم إجازة   |   "
                "غياب = يوم غياب   |   - = غير مسجل بالنظام"
            ),
            legend_style
        )
    )

    # -----------------------------------------------------
    # إنشاء PDF
    # -----------------------------------------------------

    doc.build(
        elements
    )

    pdf_buffer.seek(0)

    return pdf_buffer



# =========================================================
# /start
# =========================================================

@bot.message_handler(
    commands=["start"]
)
def send_welcome(message):

    text = (
        f"مرحباً بك في النظام الآلي لـ "
        f"**{COMPANY_NAME}**.\n\n"

        "الرجاء إرسال **الرقم الوطني** "
        "أو **اسم الموظف** للاطلاع على:\n\n"

        "💰 كشف الراتب\n"
        "📅 سجل الدوام\n"
        "⏱ ساعات الإضافي\n"
        "📄 قسيمة الراتب PDF"
    )

    bot.reply_to(
        message,
        text,
        parse_mode="Markdown"
    )


# =========================================================
# /sync
# =========================================================

@bot.message_handler(
    commands=["sync"]
)
def handle_sync(message):

    if (
        message.from_user.id
        != ADMIN_ID
    ):

        bot.reply_to(
            message,
            "⛔ هذا الأمر مخصص للإدارة فقط."
        )

        return

    bot.reply_to(
        message,
        "⏳ جاري تحديث البيانات..."
    )

    if sync_data():

        bot.reply_to(
            message,
            "✅ تمت مزامنة ملف الرواتب بنجاح."
        )

    else:

        bot.reply_to(
            message,
            "❌ فشلت عملية المزامنة."
        )


# =========================================================
# /setlink - تحديد رابط Google Drive لملف الرواتب
# =========================================================

@bot.message_handler(
    commands=["setlink"]
)
def handle_setlink(message):

    if (
        message.from_user.id
        != ADMIN_ID
    ):

        bot.reply_to(
            message,
            "⛔ هذا الأمر مخصص للإدارة فقط."
        )

        return

    # النص بعد الأمر "/setlink "
    parts = message.text.split(
        maxsplit=1
    )

    if len(parts) < 2 or not parts[1].strip():

        bot.reply_to(
            message,
            "📎 **تحديد ملف رواتب جديد من Google Drive**\n\n"
            "أرسل الأمر مع الرابط أو المعرّف، مثال:\n\n"
            "`/setlink https://drive.google.com/file/d/XXXXX/view`\n\n"
            "أو المعرّف مباشرة:\n"
            "`/setlink XXXXXXXXXXXXXXXXXXXXX`\n\n"
            "⚠️ تأكد أن مشاركة الملف على Drive مفعّلة "
            "لأي شخص لديه الرابط (Anyone with the link).",
            parse_mode="Markdown"
        )

        return

    link_text = parts[1].strip()

    file_id = extract_drive_file_id(
        link_text
    )

    if not file_id:

        bot.reply_to(
            message,
            "❌ لم أستطع التعرف على رابط أو معرّف "
            "Google Drive صحيح من النص المُرسل.\n"
            "تأكد من نسخ الرابط كاملاً."
        )

        return

    bot.reply_to(
        message,
        "⏳ جاري التحقق من الرابط وتحميل الملف..."
    )

    old_state = dict(FILE_STATE)

    set_drive_source(file_id)

    if sync_data():

        try:

            sheet_count = len(
                pd.ExcelFile(
                    LOCAL_FILE
                ).sheet_names
            )

            bot.reply_to(
                message,
                "✅ تم تحديث مصدر الملف بنجاح وتحميله.\n"
                f"📄 عدد الشيتات بالملف: {sheet_count}\n\n"
                "يمكنك الآن تجربة استعلام عن أي موظف "
                "للتأكد أن كل شيء مطابق."
            )

        except Exception as e:

            # الملف تحمّل لكنه غير صالح كإكسل - نرجع للحالة القديمة
            save_file_config(old_state)

            FILE_STATE.update(
                old_state
            )

            bot.reply_to(
                message,
                "⚠️ تم تحميل الملف لكنه لا يبدو ملف Excel "
                "صالح، وتم التراجع عن الرابط الجديد.\n"
                f"تفاصيل: {e}"
            )

    else:

        # فشل التحميل - نرجع للحالة القديمة
        save_file_config(old_state)

        FILE_STATE.update(
            old_state
        )

        bot.reply_to(
            message,
            "❌ فشل تحميل الملف من الرابط المُرسل.\n"
            "تأكد أن مشاركة الملف مفعّلة لأي شخص لديه "
            "الرابط، وأعد المحاولة."
        )


# =========================================================
# استقبال ملف Excel مباشرة من الإدارة
# =========================================================

@bot.message_handler(
    content_types=["document"]
)
def handle_document_upload(message):

    if (
        message.from_user.id
        != ADMIN_ID
    ):

        bot.reply_to(
            message,
            "⛔ رفع ملف الرواتب مخصص للإدارة فقط."
        )

        return

    file_name = (
        message.document.file_name
        or ""
    )

    if not file_name.lower().endswith(
        (".xlsx", ".xlsm")
    ):

        bot.reply_to(
            message,
            "⚠️ الرجاء إرسال ملف Excel بصيغة "
            ".xlsx أو .xlsm فقط."
        )

        return

    bot.reply_to(
        message,
        "⏳ جاري استلام الملف وحفظه..."
    )

    try:

        file_info = bot.get_file(
            message.document.file_id
        )

        downloaded = bot.download_file(
            file_info.file_path
        )

        with open(
            LOCAL_FILE,
            "wb"
        ) as f:

            f.write(
                downloaded
            )

        # التحقق أن الملف صالح فعلاً كإكسل قبل اعتماده
        sheet_count = len(
            pd.ExcelFile(
                LOCAL_FILE
            ).sheet_names
        )

        set_upload_source(
            file_name
        )

        bot.reply_to(
            message,
            "✅ تم استلام واعتماد الملف بنجاح كمصدر "
            "حالي للرواتب.\n"
            f"📎 اسم الملف: {file_name}\n"
            f"📄 عدد الشيتات: {sheet_count}\n\n"
            "ملاحظة: أمر /sync لن يستبدل هذا الملف بعد "
            "الآن - أرسل ملف Excel جديد كل شهر لتحديثه، "
            "أو استخدم /setlink للعودة لمصدر Drive."
        )

    except Exception as e:

        bot.reply_to(
            message,
            "❌ حدث خطأ أثناء حفظ الملف، أو أن الملف "
            "غير صالح.\n"
            f"تفاصيل: {e}"
        )


# =========================================================
# /source - عرض مصدر ملف الرواتب الحالي
# =========================================================

@bot.message_handler(
    commands=["source"]
)
def handle_source_status(message):

    if (
        message.from_user.id
        != ADMIN_ID
    ):

        bot.reply_to(
            message,
            "⛔ هذا الأمر مخصص للإدارة فقط."
        )

        return

    if FILE_STATE["source"] == "upload":

        text = (
            "📎 **مصدر الملف الحالي: رفع يدوي**\n\n"
            f"اسم الملف: {FILE_STATE['uploaded_filename']}\n"
            f"آخر تحديث: {FILE_STATE['updated_at']}\n\n"
            "لتحديثه: أرسل ملف Excel جديد بنفس التنسيق.\n"
            "أو استخدم /setlink للتحويل لمصدر Drive."
        )

    else:

        text = (
            "🔗 **مصدر الملف الحالي: رابط Google Drive**\n\n"
            f"معرّف الملف: `{FILE_STATE['file_id']}`\n"
            f"آخر تحديث: {FILE_STATE['updated_at']}\n\n"
            "لتحديثه: استخدم /setlink برابط جديد، "
            "أو /sync لإعادة تحميل نفس الرابط.\n"
            "أو أرسل ملف Excel مباشرة لاعتماده كمصدر يدوي."
        )

    bot.reply_to(
        message,
        text,
        parse_mode="Markdown"
    )



# =========================================================
# زر PDF
# =========================================================

@bot.callback_query_handler(
    func=lambda call:
        call.data.startswith("pdf_")
)
def handle_pdf_callback(call):

    user_id = call.data.replace(
        "pdf_",
        "",
        1
    )

    bot.answer_callback_query(
        call.id,
        "⏳ جاري إنشاء قسيمة الراتب..."
    )

    data = fetch_employee_data(
        user_id
    )

    if data is None:

        bot.send_message(
            call.message.chat.id,
            "⚠️ تعذر العثور على بيانات الموظف."
        )

        return

    try:

        pdf_file = generate_professional_pdf(
            data
        )

        filename = (
            f"Salary_Slip_"
            f"{data['report_month']:02d}_"
            f"{data['report_year']}_"
            f"{data['name']}.pdf"
        )

        bot.send_document(
            call.message.chat.id,
            pdf_file,
            visible_file_name=filename,
            caption=(
                f"📄 **قسيمة راتب شهر "
                f"{data['report_month_name']} "
                f"{data['report_year']}**\n"
                f"👤 الموظف: {data['name']}"
            ),
            parse_mode="Markdown"
        )

    except Exception as e:

        print(
            f"❌ خطأ إنشاء PDF: {e}"
        )

        bot.send_message(
            call.message.chat.id,
            "❌ حدث خطأ أثناء إنشاء قسيمة PDF."
        )


# =========================================================
# الاستعلام عن الموظف
# =========================================================

@bot.message_handler(
    func=lambda message: True,
    content_types=["text"]
)
def handle_query(message):

    user_input = message.text.strip()

    if not user_input:
        return

    print(
        f"🔎 استعلام جديد: {user_input}"
    )

    data = fetch_employee_data(
        user_input
    )

    if data is None:

        bot.reply_to(
            message,
            "⚠️ لم يتم العثور على الرقم أو الاسم.\n\n"
            "يرجى التأكد من الرقم الوطني أو كتابة الاسم بشكل صحيح."
        )

        return

    # =====================================================
    # التقرير
    # =====================================================

    total_ot_hours = (
        data["ot_n_hrs"]
        + data["ot_h_hrs"]
    )

    report = f"""
🏢 **{COMPANY_NAME}**

📋 **كشف حساب الراتب وسجل الدوام والإضافي المفصل**

══════════════════════════

👤 **بيانات الموظف:**

• **الاسم:** {data['name']}
• **القسم:** {data['dept']}
• **المسمى:** {data['job']}
• **الرقم الوطني:** `{data['nat_id']}`

💰 **البيانات المالية والاقتطاعات:**

• الراتب الأساسي: **{data['basic_sal']:.2f}** د.أ
• إجمالي مستحقات الإضافي: **{data['ot_total']:.2f}** د.أ
• مكافآت وحوافز: **{data['bonus']:.2f}** د.أ
• اقتطاع الضمان الاجتماعي: **{data['ssc']:.2f}** د.أ
• مجموع السلف: **{data['advances']:.2f}** د.أ
• اقتطاعات وحسومات: **{data['deductions']:.2f}** د.أ

──────────────────────────

💵 **صافي الراتب المستحق للصرف:
{data['net_salary']:.2f} دينار أردني**

══════════════════════════

⏱ **ملخص العمل الإضافي:**

• إضافي أيام عادية: **{data['ot_n_hrs']:g}** ساعة = **{data['ot_n_val']:.2f}** د.أ
• إضافي عطل وأعياد: **{data['ot_h_hrs']:g}** ساعة = **{data['ot_h_val']:.2f}** د.أ
• إجمالي ساعات الإضافي: **{total_ot_hours:g}** ساعة

📊 **ملخص الدوام:**

• أيام الدوام الفعلي: **{data['work_days']}** يوم
• الإجازات: **{data['leaves']}** يوم
• أيام الغياب: **{data['absent']}** يوم
• عدد أيام إنجاز الإضافي: **{data['ot_days']}** يوم

══════════════════════════

⏱ **كشف ساعات العمل الإضافي التفصيلي (01 - {data['days_in_month']:02d}):**

──────────────────────────

{data['ot_report']}

══════════════════════════

📅 **كشف سجل الدوام والغياب اليومي (01 - {data['days_in_month']:02d}):**

──────────────────────────

{data['att_report']}

══════════════════════════

{random.choice(MOTIVATIONAL_TIPS)}
"""

    # =====================================================
    # زر PDF
    # =====================================================

    markup = types.InlineKeyboardMarkup()

    markup.add(
        types.InlineKeyboardButton(
            "📥 تحميل قسيمة الراتب PDF",
            callback_data=(
                f"pdf_{data['nat_id']}"
            )
        )
    )

    bot.reply_to(
        message,
        report,
        parse_mode="Markdown",
        reply_markup=markup
    )


# =========================================================
# تشغيل البوت
# =========================================================

if __name__ == "__main__":

    print(
        "---------------------------------------------"
    )

    print(
        "🚀 جاري بدء نظام الرواتب..."
    )

    print(
        "⏳ جاري تنزيل أحدث ملف Excel..."
    )

    sync_data()

    print(
        "🤖 البوت متصل وجاهز للاستعلام..."
    )

    print(
        "---------------------------------------------"
    )

    bot.infinity_polling(
        timeout=15,
        long_polling_timeout=10
    )