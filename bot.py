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
# أدوات التنظيف والمطابقة
# =========================================================

import html as _html

AUTO_SYNC_MINUTES = int(
    os.environ.get("AUTO_SYNC_MINUTES", "30")
)

_DIGITS_MAP = str.maketrans(
    "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹",
    "01234567890123456789"
)


def is_blank(v):

    if v is None:
        return True

    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass

    return str(v).strip().lower() in (
        "", "nan", "none", "nat"
    )


def clean_num(val):
    """
    تحويل أي قيمة خلية إلى رقم.
    تدعم: أرقام عربية، فواصل الآلاف، والأرقام السالبة
    بين أقواس (50) أو بإشارة ناقص.
    """

    if is_blank(val):
        return 0.0

    s = str(val).strip()

    if s in ("-", "—", "–") or s.startswith("#"):
        return 0.0

    s = s.translate(_DIGITS_MAP)
    s = s.replace("٫", ".").replace("٬", ",")

    negative = s.startswith("(") and s.endswith(")")

    s = re.sub(r"[^\d\.\-]", "", s)

    try:

        num = float(s) if s not in ("", "-", ".") else 0.0

    except ValueError:

        return 0.0

    if negative:
        num = -abs(num)

    return round(num, 4)


def get_digits_only(text):

    if is_blank(text):
        return ""

    s = str(text).strip().translate(_DIGITS_MAP)

    if s.endswith(".0"):
        s = s[:-2]

    return re.sub(r"[^\d]", "", s)


def norm_ar(text):
    """
    تطبيع عام للنص العربي (للمطابقة فقط):
    إزالة التطويل والتشكيل، توحيد الهمزات والياء والتاء
    المربوطة، وتحويل علامات الترقيم إلى مسافات.
    """

    if is_blank(text):
        return ""

    t = str(text).strip().lower().translate(_DIGITS_MAP)

    t = re.sub(r"[\u0640\u064B-\u065F\u0670]", "", t)

    t = re.sub(r"[إأآٱ]", "ا", t)

    t = (
        t.replace("ى", "ي")
        .replace("ة", "ه")
        .replace("ؤ", "و")
        .replace("ئ", "ي")
    )

    t = t.replace("_", " ")

    t = re.sub(r"[^\w\s%]", " ", t)

    return re.sub(r"\s+", " ", t).strip()


def normalize_header(text):

    if text is None:
        return ""

    s = str(text).replace("ـ", "")

    return re.sub(r"\s+", " ", s).strip()


# =========================================================
# الأسماء: تقسيم لكلمات (tokens) بدل الاعتماد على أول 3 كلمات
# =========================================================
#
# المشكلة القديمة: كان البوت يدمج أول 3 كلمات فقط من الاسم
# ثم يبحث بطريقة "نص داخل نص"، فإذا اختلف عدد الأسماء
# بين شيت الموظفين وشيت الرواتب (مثلاً "أمين الحجوج" مقابل
# "أمين محمد الحجوج") لا يجد صف الراتب، فيرجع الراتب
# الأساسي فقط بدون أي بنود أخرى.
# =========================================================

def name_tokens(text):

    t = norm_ar(text)

    t = re.sub(r"\bعبد\s+", "عبد", t)
    t = re.sub(r"\bابو\s+", "ابو", t)

    tokens = []

    for w in t.split():

        if len(w) > 3 and w.startswith("ال") and w != "الله":
            w = w[2:]

        tokens.append(w)

    return tokens


def query_score(q_tokens, e_tokens):
    """درجة تطابق ما كتبه المستخدم مع اسم موظف."""

    if not q_tokens or not e_tokens:
        return 0

    if q_tokens == e_tokens:
        return 100

    qs, es = set(q_tokens), set(e_tokens)

    if qs == es:
        return 95

    if qs <= es:
        return 85 - min(len(es - qs), 10)

    if all(
        len(q) >= 2 and any(e.startswith(q) for e in e_tokens)
        for q in q_tokens
    ):
        return max(40, 70 - abs(len(e_tokens) - len(q_tokens)))

    return 0


def sheet_name_score(e_tokens, r_tokens):
    """درجة تطابق اسم الموظف مع اسم في شيت آخر."""

    if not e_tokens or not r_tokens:
        return 0

    if e_tokens == r_tokens:
        return 100

    es, rs = set(e_tokens), set(r_tokens)

    if es == rs:
        return 95

    # اسم الموظف (بشيت الموظفين) موجود بالكامل داخل اسم أطول
    # في هذا الشيت: الحالة الأرجح (الكشف فيه الاسم الكامل)
    if len(es) >= 2 and es <= rs:
        return 85 - min(len(rs - es), 10)

    # عكس ذلك: اسم الشيت أقصر من اسم الموظف (أقل ترجيحاً)
    if len(rs) >= 2 and rs <= es:
        return 70 - min(len(es - rs), 10)

    return 0


# =========================================================
# التعرف على الأعمدة بالمعنى (وليس بالاسم الحرفي)
# =========================================================
#
# المشكلة القديمة: كانت الأعمدة تُقرأ بأسماء حرفية ثابتة
# مثل "الضمان" و"السلف" و"حسم"، فأي اختلاف بسيط في عنوان
# العمود بالكشف (مثلاً "اقتطاع الضمان الاجتماعي") يجعل
# القيمة 0 بصمت. الآن كل عمود يُقيَّم بقواعد مرنة.
# =========================================================

def _has(h, *words):
    return all(w in h for w in words)


def _any(h, words):
    return any(w in h for w in words)


_TOTAL_WORDS = ["مجموع", "اجمالي", "الاجمالي", "كلي"]
_HOURS_WORDS = ["ساعات", "ساعه"]
_HOLIDAY_WORDS = ["جمع", "عطل", "عيد", "اعياد"]
_VALUE_WORDS = ["استحقاق", "قيمه", "مبلغ", "اجر", "مستحقات"]


def r_id(h):

    if h in (
        "الرقم الوطني", "رقم وطني", "الرقم الشخصي",
        "رقم شخصي", "الرقم الوظيفي", "رقم الموظف",
        "الكود", "كود", "رقم الهويه"
    ):
        return 100

    if _any(h, [
        "الرقم الوطني", "رقم وطني",
        "الرقم الشخصي", "رقم شخصي"
    ]):
        return 80

    return 0


def r_name(h):

    if h in ("الاسم", "اسم", "الموظف", "العامل"):
        return 100

    if h.startswith("اسم ") and len(h.split()) <= 3:
        return 90

    return 0


def r_dept(h):
    return 100 if h in (
        "القسم", "قسم", "الاداره", "المشروع", "الموقع"
    ) else 0


def r_job(h):
    return 100 if h in (
        "المسمي الوظيفي", "المسمي", "الوظيفه", "المهنه"
    ) else 0


def r_net(h):

    if _any(h, ["صافي", "الصافي"]):
        return 100

    if "المستحق" in h and _any(h, ["راتب", "صرف"]):
        return 80

    return 0


def r_ot_total(h):

    if (
        _any(h, _TOTAL_WORDS)
        and _any(h, ["اضافي", "اضافه"])
        and not _any(h, _HOURS_WORDS)
        and not _any(h, ["عادي"] + _HOLIDAY_WORDS)
    ):
        return 100

    return 0


def r_ot_n_hrs(h):

    if (
        _any(h, _HOURS_WORDS)
        and "عادي" in h
        and not _any(h, _VALUE_WORDS)
    ):
        return 90

    return 0


def r_ot_h_hrs(h):

    if (
        _any(h, _HOURS_WORDS)
        and _any(h, _HOLIDAY_WORDS)
        and not _any(h, _VALUE_WORDS)
    ):
        return 90

    return 0


def r_ot_n_val(h):

    if (
        _any(h, _VALUE_WORDS)
        and "عادي" in h
        and not _any(h, _HOURS_WORDS)
    ):
        return 90

    return 0


def r_ot_h_val(h):

    if (
        _any(h, _VALUE_WORDS)
        and _any(h, _HOLIDAY_WORDS)
        and not _any(h, _HOURS_WORDS)
    ):
        return 90

    return 0


def r_basic(h):

    if _any(h, [
        "صافي", "ضمان", "مجموع", "اجمالي",
        "استحقاق", "ساعات", "ساعه"
    ]):
        return 0

    if h in ("الراتب الاساسي", "راتب اساسي", "الاساسي"):
        return 100

    if "اساسي" in h:
        return 80

    if h in ("الراتب", "راتب", "الراتب الشهري"):
        return 60

    return 0


def r_ssc(h):

    if "ضمان" not in h:
        return 0

    if _any(h, [
        "شركه", "صاحب", "مساهمه", "منشاه", "نسبه",
        "خاضع", "شامل", "للضمان", "راتب الضمان",
        "اجر الضمان"
    ]):
        return 0

    if h in (
        "الضمان", "ضمان", "الضمان الاجتماعي",
        "ضمان اجتماعي", "اقتطاع الضمان"
    ):
        return 100

    return 70


def r_advances(h):

    if "سلف" not in h:
        return 0

    if _any(h, ["رصيد", "متبقي", "باقي"]):
        return 0

    if h in (
        "السلف", "السلفه", "سلف", "سلفه",
        "مجموع السلف", "اجمالي السلف"
    ):
        return 100

    return 80


def r_deductions(h):

    if not _any(h, ["حسم", "خصم", "جزاء", "غرام", "اقتطاع"]):
        return 0

    if _any(h, ["ضمان", "سلف"] + _HOURS_WORDS):
        return 0

    is_total = _any(h, _TOTAL_WORDS)

    # "إجمالي الاقتطاعات" يشمل الضمان والسلف أيضاً،
    # فلا يُحسب ضمن الحسومات حتى لا يتكرر الخصم
    if is_total and "اقتطاع" in h:
        return 0

    return 100 if is_total else 70


def r_bonus(h):

    if _any(h, ["حسم", "خصم"]):
        return 0

    return 90 if _any(h, [
        "مكاف", "حافز", "حوافز", "علاوه", "علاوات"
    ]) else 0


def r_allow(h):

    if "بدل" in h and not _any(
        h, ["حسم", "خصم"] + _HOURS_WORDS
    ):
        return 70

    return 0


# ترتيب المعالجة مهم: العمود يُحجز لأول حقل يطابقه
FIELD_RULES = [
    ("id", r_id, False),
    ("name", r_name, False),
    ("net", r_net, False),
    ("ot_total", r_ot_total, False),
    ("ot_n_hrs", r_ot_n_hrs, False),
    ("ot_h_hrs", r_ot_h_hrs, False),
    ("ot_n_val", r_ot_n_val, False),
    ("ot_h_val", r_ot_h_val, False),
    ("basic", r_basic, False),
    ("ssc", r_ssc, False),
    ("advances", r_advances, False),
    ("deductions", r_deductions, True),
    ("bonus", r_bonus, False),
    ("allow", r_allow, True),
    ("dept", r_dept, False),
    ("job", r_job, False)
]

RULES = {name: fn for name, fn, _ in FIELD_RULES}


def resolve_fields(df):
    """
    يرجع قاموساً: اسم الحقل -> عمود (أو قائمة أعمدة للحقول
    المتعددة مثل الحسومات والبدلات).
    """

    cols = list(df.columns)

    keys = {c: norm_ar(c) for c in cols}

    used = set()

    result = {}

    for field, fn, multi in FIELD_RULES:

        scored = []

        for i, c in enumerate(cols):

            if c in used or str(c).startswith("__col"):
                continue

            s = fn(keys[c])

            if s > 0:
                scored.append((s, i, c))

        if not scored:
            continue

        if multi:

            top = [x for x in scored if x[0] >= 100]

            # إن وُجد عمود "مجموع" نعتمده وحده، وإلا نجمع الأعمدة
            chosen = top[:1] if top else sorted(
                scored, key=lambda x: x[1]
            )

            result[field] = [c for _, _, c in chosen]

            used.update(result[field])

        else:

            best = max(scored, key=lambda x: (x[0], -x[1]))

            result[field] = best[2]

            used.add(best[2])

    return result


# =========================================================
# قراءة ملف Excel مرة واحدة (مع كاش حسب تاريخ الملف)
# =========================================================

_WB_CACHE = {"sig": None, "sheets": {}}


def load_workbook_raw():

    if not os.path.exists(LOCAL_FILE):
        return {}

    st = os.stat(LOCAL_FILE)

    sig = (st.st_mtime, st.st_size)

    if _WB_CACHE["sig"] == sig and _WB_CACHE["sheets"]:
        return _WB_CACHE["sheets"]

    try:

        sheets = pd.read_excel(
            LOCAL_FILE,
            sheet_name=None,
            header=None,
            dtype=str
        )

    except Exception as e:

        print(f"❌ تعذرت قراءة ملف Excel: {e}")

        return {}

    _WB_CACHE["sig"] = sig
    _WB_CACHE["sheets"] = sheets

    return sheets


def find_sheet_header(raw, predicate, max_rows=60):

    for idx in range(min(len(raw), max_rows)):

        cells = [
            normalize_header(x)
            for x in raw.iloc[idx].values
            if not is_blank(x)
        ]

        if cells and predicate(cells):
            return idx

    return None


def table_from_raw(raw, header_row):

    header = [
        "" if is_blank(x) else normalize_header(x)
        for x in raw.iloc[header_row].values
    ]

    cols, seen = [], {}

    for i, h in enumerate(header):

        if not h:
            h = f"__col{i}"

        if h in seen:

            seen[h] += 1
            h = f"{h}__{seen[h]}"

        else:

            seen[h] = 0

        cols.append(h)

    df = raw.iloc[header_row + 1:].copy()

    df.columns = cols

    return df.reset_index(drop=True)


# =========================================================
# تحميل ملف Excel من Drive
# =========================================================

_LAST_SYNC_ATTEMPT = {"t": 0.0}


def sync_data():

    if FILE_STATE["source"] == "upload":

        if os.path.exists(LOCAL_FILE):

            print(
                "ℹ️ المصدر الحالي ملف مرفوع يدويًا - "
                "لا حاجة للتحميل من Drive."
            )

            return True

        print("❌ لا يوجد ملف محلي، ولم يتم رفع ملف بعد.")

        return False

    _LAST_SYNC_ATTEMPT["t"] = time.time()

    try:

        print("⏳ جاري تحميل أحدث ملف Excel...")

        response = requests.get(
            FILE_STATE["download_url"],
            timeout=30
        )

        # ملف xlsx الحقيقي يبدأ بـ PK (zip). هذا يمنع حفظ
        # صفحة HTML (صفحة تأكيد/خطأ من Drive) على أنها ملف
        if (
            response.status_code == 200
            and len(response.content) > 5000
            and response.content[:2] == b"PK"
        ):

            tmp = LOCAL_FILE + ".tmp"

            with open(tmp, "wb") as f:
                f.write(response.content)

            os.replace(tmp, LOCAL_FILE)

            print("✅ تم تحديث ملف Excel.")

            return True

        print("❌ لم يتم تحميل الملف بشكل صحيح.")

    except Exception as e:

        print(f"❌ خطأ أثناء تحميل الملف: {e}")

    return False


def maybe_auto_sync():
    """
    مزامنة تلقائية صامتة إذا كان الملف المحلي أقدم من
    AUTO_SYNC_MINUTES دقيقة (لضمان أن البوت يقرأ آخر
    نسخة من الكشف دون الحاجة لأمر /sync يدوي).
    """

    if AUTO_SYNC_MINUTES <= 0:
        return

    if FILE_STATE["source"] != "drive":
        return

    now = time.time()

    if now - _LAST_SYNC_ATTEMPT["t"] < 300:
        return

    if (
        os.path.exists(LOCAL_FILE)
        and now - os.path.getmtime(LOCAL_FILE)
        < AUTO_SYNC_MINUTES * 60
    ):
        return

    sync_data()


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
# شيت بيانات الموظفين + البحث عن الموظف
# =========================================================

def load_employee_sheet(book):

    for sheet_name, raw in book.items():

        if not any(
            k in norm_ar(sheet_name)
            for k in ("بيانات", "عمال", "موظفين")
        ):
            continue

        def pred(cells):

            keys = [norm_ar(c) for c in cells]

            return (
                any(r_name(k) for k in keys)
                and any(r_id(k) for k in keys)
            )

        hr = find_sheet_header(raw, pred)

        if hr is None:
            continue

        return table_from_raw(raw, hr), sheet_name

    return None, None


def find_employees(df_emp, user_input):
    """
    يرجع قائمة مرشحين مرتبة بالأفضل. البحث بالرقم الوطني
    مطابقة تامة، وبالاسم مطابقة كلمات (ترتيب الكلمات غير مهم
    ولا يشترط كتابة الاسم الكامل).
    """

    fields = resolve_fields(df_emp)

    name_col = fields.get("name")
    id_col = fields.get("id")
    dept_col = fields.get("dept")
    job_col = fields.get("job")
    basic_col = fields.get("basic")

    q_digits = get_digits_only(user_input)
    q_tokens = name_tokens(user_input)

    is_id_query = len(q_digits) >= 4 and bool(
        re.fullmatch(
            r"[\d\s\-]+",
            str(user_input).translate(_DIGITS_MAP).strip()
        )
    )

    found = []

    for _, row in df_emp.iterrows():

        r_name_txt = (
            "" if name_col is None or is_blank(row.get(name_col))
            else str(row.get(name_col)).strip()
        )

        r_nat = (
            get_digits_only(row.get(id_col))
            if id_col is not None else ""
        )

        score = 0

        if is_id_query:

            if r_nat and q_digits == r_nat:
                score = 100

        elif (
            r_name_txt
            and len(norm_ar(user_input).replace(" ", "")) >= 3
        ):

            score = query_score(
                q_tokens,
                name_tokens(r_name_txt)
            )

        if score <= 0:
            continue

        def cell(col, default):

            if col is None or is_blank(row.get(col)):
                return default

            return str(row.get(col)).strip()

        found.append({
            "score": score,
            "name": r_name_txt,
            "job": cell(job_col, "موظف"),
            "dept": cell(dept_col, "العمليات"),
            "nat_id": r_nat if r_nat else "-",
            "id_digits": r_nat,
            "tokens": name_tokens(r_name_txt),
            "base_sal_fallback": (
                clean_num(row.get(basic_col))
                if basic_col is not None else 0.0
            )
        })

    found.sort(key=lambda x: -x["score"])

    return found


def pick_candidates(cands):
    """
    يقرر هل النتيجة واضحة (موظف واحد) أم ملتبسة.
    يرجع (موظف أو None، قائمة الملتبسين).
    """

    if not cands:
        return None, []

    if len(cands) == 1:
        return cands[0], []

    if cands[0]["score"] >= 100 and cands[1]["score"] < 100:
        return cands[0], []

    return None, cands


# =========================================================
# البحث عن صف الموظف داخل أي شيت
# =========================================================

def find_employee_row(df, emp):
    """يرجع (الصف أو None، وصف طريقة المطابقة)."""

    if df is None or df.empty:
        return None, "الشيت فارغ"

    fields = resolve_fields(df)

    name_col = fields.get("name")
    id_col = fields.get("id")

    if name_col is None and id_col is None:

        # آخر حل: أول عمود (سلوك النسخة القديمة)
        name_col = df.columns[0]

    emp_id = emp.get("id_digits", "")

    if id_col is not None and emp_id:

        ids = df[id_col].map(get_digits_only)

        hit = df[ids == emp_id]

        if len(hit):
            return hit.iloc[0], "مطابقة بالرقم الوطني"

    if name_col is None:
        return None, "لا يوجد عمود اسم أو رقم وطني"

    scored = []

    for idx, v in df[name_col].items():

        if is_blank(v):
            continue

        s = sheet_name_score(emp["tokens"], name_tokens(v))

        if s > 0:
            scored.append((s, idx))

    if not scored:
        return None, "الاسم غير موجود في هذا الشيت"

    scored.sort(key=lambda x: (-x[0], x[1]))

    best = scored[0][0]

    tops = [i for s, i in scored if s == best]

    if len(tops) > 1 and best < 95:

        return None, (
            "عدة أسماء متشابهة (تطابق جزئي) - "
            "تم تجاهلها لتجنب خلط رواتب موظفين"
        )

    how = (
        "مطابقة اسم كاملة" if best >= 95
        else "مطابقة اسم جزئية (اسم أقصر داخل اسم أطول)"
    )

    return df.loc[tops[0]], how


# =========================================================
# استخراج رقم اليوم (01-31) من عنوان عمود
# =========================================================

def parse_day_number(val):

    if is_blank(val):
        return None

    import datetime as _dt

    if isinstance(val, (_dt.datetime, _dt.date, pd.Timestamp)):
        return val.day

    s = str(val).strip().translate(_DIGITS_MAP)

    if re.fullmatch(r"\d{1,2}", s):

        n = int(s)

        return n if 1 <= n <= 31 else None

    m = re.match(r"^\d{4}-\d{2}-(\d{2})", s)

    if m:
        return int(m.group(1))

    m = re.match(r"^(\d{1,2})[/-]\d{1,2}[/-]\d{2,4}", s)

    if m:

        n = int(m.group(1))

        return n if 1 <= n <= 31 else None

    return None


# =========================================================
# استخراج الشهر والسنة الفعليين
# =========================================================

_MONTH_NAMES = [
    (1, "يناير"), (2, "فبراير"), (3, "مارس"), (4, "ابريل"),
    (5, "مايو"), (6, "يونيو"), (7, "يوليو"), (8, "اغسطس"),
    (9, "سبتمبر"), (10, "اكتوبر"), (11, "نوفمبر"),
    (12, "ديسمبر"),
    (1, "كانون الثاني"), (2, "شباط"), (3, "اذار"),
    (4, "نيسان"), (5, "ايار"), (6, "حزيران"), (7, "تموز"),
    (8, "اب"), (9, "ايلول"), (10, "تشرين الاول"),
    (11, "تشرين الثاني"), (12, "كانون الاول")
]

_MONTH_NAMES.sort(key=lambda x: -len(x[1]))


def period_from_text(text):

    raw = str(text).translate(_DIGITS_MAP)

    m = re.search(
        r"(?<!\d)(\d{1,2})\s*[/\-.]\s*(20\d{2})(?!\d)", raw
    )

    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(2)), int(m.group(1))

    m = re.search(
        r"(?<!\d)(20\d{2})\s*[/\-.]\s*(\d{1,2})(?!\d)", raw
    )

    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))

    t = " " + norm_ar(raw) + " "

    ym = re.search(r"(?<!\d)(20\d{2})(?!\d)", t)

    year = int(ym.group(1)) if ym else dt.datetime.now().year

    for month, name in _MONTH_NAMES:

        if f" {name} " in t:
            return year, month

    return None


def detect_period(raw, header_row, sheet_name):

    for x in raw.iloc[header_row].values:

        if is_blank(x):
            continue

        m = re.match(
            r"^(\d{4})-(\d{2})-\d{2}",
            str(x).strip()
        )

        if m:
            return int(m.group(1)), int(m.group(2))

    texts = [str(sheet_name)]

    for i in range(header_row):

        texts += [
            str(x) for x in raw.iloc[i].values
            if not is_blank(x)
        ]

    return period_from_text(" ".join(texts))


# =========================================================
# شيتات الشبكة (اليومية / الإضافي)
# =========================================================

def load_grid_sheets(book, keywords, exclude=()):

    results = []

    kws = [norm_ar(k) for k in keywords]
    exs = [norm_ar(k) for k in exclude]

    for sheet_name, raw in book.items():

        sn = norm_ar(sheet_name)

        if not any(k in sn for k in kws):
            continue

        if any(k in sn for k in exs):
            continue

        def pred(cells):

            keys = [norm_ar(c) for c in cells]

            if not any(r_name(k) for k in keys):
                return False

            days = {parse_day_number(c) for c in cells}

            return {1, 2, 3} <= days

        hr = find_sheet_header(raw, pred)

        if hr is None:

            print(f"⚠️ لم يتم العثور على صف عناوين في {sheet_name}")

            continue

        results.append({
            "sheet": sheet_name,
            "df": table_from_raw(raw, hr),
            "period": detect_period(raw, hr, sheet_name)
        })

    return results


def extract_grid_data(grids, emp):
    """
    يرجع (قاموس الأيام، اسم الشيت، وصف المطابقة).
    يجرّب كل شيت مطابق للاسم حتى يجد الموظف.
    """

    last_how = "لا يوجد شيت مطابق"

    for g in grids:

        row, how = find_employee_row(g["df"], emp)

        last_how = how

        if row is None:
            continue

        days = {}

        for day in range(1, 32):

            value = ""

            for col in g["df"].columns:

                if parse_day_number(col) == day:

                    value = row.get(col, "")

                    break

            days[f"{day:02d}"] = (
                "" if is_blank(value) else str(value).strip()
            )

        return days, g["sheet"], how

    return {f"{d:02d}": "" for d in range(1, 32)}, None, last_how


# =========================================================
# شيت الرواتب
# =========================================================

def load_salary_sheets(book):
    """
    يرجع كل الشيتات المرشحة للرواتب مرتبة بعدد الأعمدة
    المفهومة (الأكثر أولاً)، حتى لا يُختار شيت خاطئ.
    """

    cands = []

    for sheet_name, raw in book.items():

        sn = norm_ar(sheet_name)

        if not any(
            k in sn for k in ("راتب", "رواتب", "مسير")
        ):
            continue

        def pred(cells):

            keys = [norm_ar(c) for c in cells]

            return any(r_name(k) for k in keys) and any(
                r_basic(k) or r_net(k) or r_ssc(k)
                or r_ot_total(k) or r_advances(k)
                for k in keys
            )

        hr = find_sheet_header(raw, pred)

        if hr is None:
            continue

        df = table_from_raw(raw, hr)

        fields = resolve_fields(df)

        score = len(fields) + (
            3 if "ملخص" in sn else 0
        )

        cands.append((score, sheet_name, df, fields))

    cands.sort(key=lambda x: -x[0])

    return cands


# =========================================================
# جلب كل بيانات الموظف
# =========================================================

ATT_WORK = {"د", "دوام"}
ATT_LEAVE = {"م", "اجازه"}
ATT_ABSENT = {"غ", "غياب"}


def fetch_employee_data(
    user_input,
    choose_first=False
):

    maybe_auto_sync()

    if not os.path.exists(LOCAL_FILE):
        sync_data()

    book = load_workbook_raw()

    if not book:
        return None

    df_emp, emp_sheet = load_employee_sheet(book)

    if df_emp is None:

        print("❌ لم يتم العثور على شيت بيانات الموظفين.")

        return None

    cands = find_employees(df_emp, user_input)

    employee, ambiguous = pick_candidates(cands)

    if employee is None and ambiguous:

        if not choose_first:
            return {"ambiguous": ambiguous[:8]}

        employee = ambiguous[0]

    if employee is None:

        print("❌ الموظف غير موجود.")

        return None

    dbg = [
        f"الموظف: {employee['name']}",
        f"الرقم الوطني: {employee['id_digits'] or 'غير متوفر'}",
        f"شيت الموظفين: {emp_sheet}"
    ]

    warnings = []

    # -----------------------------------------------------
    # شيت الرواتب
    # -----------------------------------------------------

    sal_row, sal_fields, sal_sheet = None, {}, None

    sal_cands = load_salary_sheets(book)

    if not sal_cands:
        dbg.append("⚠️ لا يوجد شيت رواتب مطابق في الملف")

    for _, sname, sdf, sfields in sal_cands:

        row, how = find_employee_row(sdf, employee)

        dbg.append(f"شيت الرواتب [{sname}]: {how}")

        if row is not None:

            sal_row, sal_fields, sal_sheet = row, sfields, sname

            break

    sal_found = sal_row is not None

    def val(field):

        col = sal_fields.get(field)

        if not sal_found or col is None:
            return 0.0

        if isinstance(col, list):
            return sum(clean_num(sal_row.get(c)) for c in col)

        return clean_num(sal_row.get(col))

    if sal_found:

        dbg.append("ربط الأعمدة (الحقل ← العمود ← القيمة):")

        for field, col in sal_fields.items():

            if field in ("id", "name", "dept", "job"):
                continue

            cols = col if isinstance(col, list) else [col]

            for c in cols:

                dbg.append(
                    f"  {field} ← {c} ← {sal_row.get(c)}"
                )

        missing = [
            f for f in (
                "basic", "ssc", "advances", "deductions",
                "ot_total", "net"
            ) if f not in sal_fields
        ]

        if missing:

            dbg.append(
                "⚠️ حقول لم يتم العثور على عمود لها: "
                + ", ".join(missing)
            )

            print(
                f"⚠️ أعمدة غير معروفة بشيت الرواتب: {missing}"
            )

    else:

        warnings.append("salary_row_missing")

        print(
            f"⚠️ لم يُعثر على صف راتب للموظف: {employee['name']}"
        )

    # -----------------------------------------------------
    # اليومية والإضافي
    # -----------------------------------------------------

    att_grids = load_grid_sheets(
        book,
        ["يومية", "حضور"],
        exclude=["اضافي", "اضافه"]
    )

    ot_grids = load_grid_sheets(
        book,
        ["اضافي", "اضافه"]
    )

    att_dict, att_sheet, att_how = extract_grid_data(
        att_grids, employee
    )

    ot_dict, ot_sheet, ot_how = extract_grid_data(
        ot_grids, employee
    )

    dbg.append(f"شيت الدوام [{att_sheet}]: {att_how}")
    dbg.append(f"شيت الإضافي [{ot_sheet}]: {ot_how}")

    period = None

    for g in att_grids + ot_grids:

        if g["period"]:

            period = g["period"]

            break

    if period:

        report_year, report_month = period

    else:

        now = dt.datetime.now()

        report_year, report_month = now.year, now.month

        warnings.append("period_guessed")

    days_in_month = calendar.monthrange(
        report_year, report_month
    )[1]

    report_month_name = ARABIC_MONTHS.get(
        report_month, str(report_month)
    )

    # -----------------------------------------------------
    # الدوام والإضافي اليومي
    # -----------------------------------------------------

    calc_work_days = calc_leaves = calc_absent = 0

    total_ot_days = 0

    ot_daily_total = 0.0

    daily_rows = []

    for day in range(1, days_in_month + 1):

        day_str = f"{day:02d}"

        status = str(att_dict.get(day_str, "")).strip()

        code = norm_ar(status).replace(" ", "")

        if code in ATT_WORK:

            status_short = "دوام"
            calc_work_days += 1

        elif code in ATT_LEAVE:

            status_short = "إجازة"
            calc_leaves += 1

        elif code in ATT_ABSENT:

            status_short = "غياب"
            calc_absent += 1

        elif code == "":

            status_short = "-"

        else:

            status_short = status

        ot_val = clean_num(ot_dict.get(day_str, 0))

        if ot_val > 0:

            total_ot_days += 1
            ot_daily_total += ot_val

        try:

            day_date = dt.date(report_year, report_month, day)

            date_str = day_date.strftime("%d/%m/%Y")

            weekday_name = ARABIC_WEEKDAYS_FULL[day_date.weekday()]

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

    # -----------------------------------------------------
    # البيانات المالية (من الكشف مباشرة)
    # -----------------------------------------------------

    basic_sal = val("basic")

    if basic_sal == 0:
        basic_sal = employee.get("base_sal_fallback", 0.0)

    ot_n_hrs = val("ot_n_hrs")
    ot_h_hrs = val("ot_h_hrs")
    ot_n_val = val("ot_n_val")
    ot_h_val = val("ot_h_val")

    ot_total = val("ot_total")

    if ot_total == 0:
        ot_total = ot_n_val + ot_h_val

    bonus = val("bonus") + val("allow")

    ssc = val("ssc")
    advances = val("advances")
    deductions = val("deductions")

    calc_net = round(
        basic_sal + ot_total + bonus
        - ssc - advances - deductions,
        2
    )

    net_col = sal_fields.get("net")

    net_in_sheet = (
        sal_found
        and net_col is not None
        and not is_blank(sal_row.get(net_col))
    )

    other_adj = 0.0

    if not sal_found:

        net_salary = None

    elif net_in_sheet:

        # صافي الراتب الرسمي من الكشف هو المرجع
        net_salary = round(clean_num(sal_row.get(net_col)), 2)

        diff = round(net_salary - calc_net, 2)

        if abs(diff) >= 0.01:

            # فرق بين مجموع البنود المعروضة وصافي الكشف:
            # لا نخفيه، نعرضه كبند مستقل ليتطابق الإجمالي
            other_adj = diff

            dbg.append(
                f"⚠️ صافي الكشف {net_salary:.2f} يختلف عن "
                f"المحسوب {calc_net:.2f} (فرق {diff:+.2f}) "
                "← غالباً هناك عمود استحقاق/اقتطاع غير معروف"
            )

            print(
                f"⚠️ فرق بالصافي للموظف {employee['name']}: {diff:+.2f}"
            )

    else:

        net_salary = calc_net

        dbg.append(
            "ℹ️ لا يوجد عمود صافي بالكشف - تم الحساب: "
            "الأساسي + الإضافي + المكافآت − الضمان − السلف − الحسومات"
        )

    ot_sheet_hrs = ot_n_hrs + ot_h_hrs

    if abs(ot_sheet_hrs - ot_daily_total) > 0.01:

        dbg.append(
            f"ℹ️ ساعات الإضافي بشيت الرواتب = {ot_sheet_hrs:g} "
            f"بينما مجموع الإضافي اليومي = {ot_daily_total:g}"
        )

    dbg.append(
        f"الحساب: {basic_sal:.2f} + {ot_total:.2f} + {bonus:.2f} "
        f"− {ssc:.2f} − {advances:.2f} − {deductions:.2f} "
        f"= {calc_net:.2f}"
    )

    employee.update({

        "sal_found": sal_found,
        "sal_sheet": sal_sheet,
        "other_adj": other_adj,
        "warnings": warnings,
        "debug": dbg,

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
        "ot_daily_total": ot_daily_total,

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
            ar_txt("مكافآت وحوافز وبدلات")
        ]
    ]

    adj = data.get("other_adj", 0.0)

    if adj >= 0.01:

        finance_data.append([
            "",
            "",
            f"{adj:.2f}",
            ar_txt("بنود أخرى في الكشف")
        ])

    elif adj <= -0.01:

        finance_data.append([
            f"{-adj:.2f}",
            ar_txt("بنود أخرى في الكشف"),
            "",
            ""
        ])

    total_deductions = (
        data["ssc"]
        + data["advances"]
        + data["deductions"]
        + max(-adj, 0.0)
    )

    total_earnings = (
        data["basic_sal"]
        + data["ot_total"]
        + data["bonus"]
        + max(adj, 0.0)
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
# مساعدات الرسائل
# =========================================================
#
# - تنسيق HTML بدل Markdown: أي اسم يحتوي _ أو * كان يكسر
#   رسالة Markdown ويمنع وصول الكشف.
# - تقسيم الرسالة الطويلة (حد تيليجرام 4096 حرف).
# - callback_data محدودة بـ 64 بايت، والاسم العربي قد
#   يتجاوزها، لذلك نستخدم الرقم الوطني أو رمزًا قصيرًا.
# =========================================================

from collections import OrderedDict

_PENDING = OrderedDict()


def make_token(key):

    tok = "".join(
        random.choices("abcdefghijklmnopqrstuvwxyz0123456789", k=8)
    )

    _PENDING[tok] = key

    while len(_PENDING) > 2000:
        _PENDING.popitem(last=False)

    return tok


def employee_key(emp):

    if emp.get("id_digits") and len(emp["id_digits"]) <= 40:
        return emp["id_digits"]

    return make_token(emp["name"])


def resolve_callback_key(raw):

    if raw.startswith("t:"):
        return _PENDING.get(raw[2:])

    return raw


def split_text(text, limit=3900):

    chunks, cur = [], ""

    for line in text.split("\n"):

        if len(cur) + len(line) + 1 > limit and cur:

            chunks.append(cur)
            cur = ""

        cur += line + "\n"

    if cur.strip():
        chunks.append(cur)

    return chunks


def fmt_money(x):

    return f"{x:.2f}"


def build_report_text(d):

    e = _html.escape

    period = f"{d['report_month_name']} {d['report_year']}"

    L = [
        f"🏢 <b>{e(COMPANY_NAME)}</b>",
        "",
        "📋 <b>كشف حساب الراتب وسجل الدوام والإضافي المفصل</b>",
        f"🗓 <b>الشهر:</b> {e(period)}",
        "",
        "══════════════════════════",
        "",
        "👤 <b>بيانات الموظف:</b>",
        "",
        f"• <b>الاسم:</b> {e(d['name'])}",
        f"• <b>القسم:</b> {e(d['dept'])}",
        f"• <b>المسمى:</b> {e(d['job'])}",
        f"• <b>الرقم الوطني:</b> <code>{e(d['nat_id'])}</code>",
        ""
    ]

    if d["sal_found"]:

        L += [
            "💰 <b>البيانات المالية والاقتطاعات:</b>",
            "",
            f"• الراتب الأساسي: <b>{fmt_money(d['basic_sal'])}</b> د.أ",
            f"• إجمالي مستحقات الإضافي: <b>{fmt_money(d['ot_total'])}</b> د.أ",
            f"• مكافآت وحوافز وبدلات: <b>{fmt_money(d['bonus'])}</b> د.أ",
            f"• اقتطاع الضمان الاجتماعي: <b>{fmt_money(d['ssc'])}</b> د.أ",
            f"• مجموع السلف: <b>{fmt_money(d['advances'])}</b> د.أ",
            f"• اقتطاعات وحسومات: <b>{fmt_money(d['deductions'])}</b> د.أ"
        ]

        if abs(d["other_adj"]) >= 0.01:

            sign = "+" if d["other_adj"] > 0 else "−"

            L.append(
                f"• بنود أخرى في الكشف: <b>{sign}"
                f"{fmt_money(abs(d['other_adj']))}</b> د.أ"
            )

        L += [
            "",
            "──────────────────────────",
            "",
            "💵 <b>صافي الراتب المستحق للصرف:</b>",
            f"<b>{fmt_money(d['net_salary'])} دينار أردني</b>"
        ]

    else:

        L += [
            "⚠️ <b>تنبيه:</b> لم يتم العثور على سجل راتب هذا الموظف "
            f"في كشف رواتب شهر {e(period)}.",
            "لذلك لا يمكن عرض الاستحقاقات والاقتطاعات وصافي الراتب، "
            "يرجى مراجعة الإدارة / المحاسب."
        ]

    total_ot = d["ot_n_hrs"] + d["ot_h_hrs"]

    L += [
        "",
        "══════════════════════════",
        "",
        "⏱ <b>ملخص العمل الإضافي:</b>",
        "",
        f"• إضافي أيام عادية: <b>{d['ot_n_hrs']:g}</b> ساعة = "
        f"<b>{fmt_money(d['ot_n_val'])}</b> د.أ",
        f"• إضافي عطل وأعياد: <b>{d['ot_h_hrs']:g}</b> ساعة = "
        f"<b>{fmt_money(d['ot_h_val'])}</b> د.أ",
        f"• إجمالي ساعات الإضافي: <b>{total_ot:g}</b> ساعة",
        "",
        "📊 <b>ملخص الدوام:</b>",
        "",
        f"• أيام الدوام الفعلي: <b>{d['work_days']}</b> يوم",
        f"• الإجازات: <b>{d['leaves']}</b> يوم",
        f"• أيام الغياب: <b>{d['absent']}</b> يوم",
        f"• عدد أيام إنجاز الإضافي: <b>{d['ot_days']}</b> يوم",
        "",
        "══════════════════════════",
        "",
        f"⏱ <b>كشف ساعات العمل الإضافي التفصيلي "
        f"(01 - {d['days_in_month']:02d}):</b>",
        "",
        "──────────────────────────",
        ""
    ]

    for r in d["daily_rows"]:

        ot_txt = (
            f"⏱ <b>{r['ot_hours']:g} ساعات إضافي</b>"
            if r["ot_hours"] > 0
            else "لا يوجد إضافي (0 س) ➖"
        )

        L.append(f"🗓 <code>{r['day']:02d}</code>: {ot_txt}")

    L += [
        "",
        "══════════════════════════",
        "",
        f"📅 <b>كشف سجل الدوام والغياب اليومي "
        f"(01 - {d['days_in_month']:02d}):</b>",
        "",
        "──────────────────────────",
        ""
    ]

    status_txt = {
        "دوام": "دوام فعلي ✅",
        "إجازة": "إجازة 🏖",
        "غياب": "غياب ❌",
        "-": "غير مسجل ⚠️"
    }

    for r in d["daily_rows"]:

        txt = status_txt.get(
            r["status"], f"حالة ({e(str(r['status']))})"
        )

        L.append(f"🗓 <code>{r['day']:02d}</code>: {txt}")

    L += [
        "",
        "══════════════════════════",
        "",
        random.choice(MOTIVATIONAL_TIPS)
    ]

    return "\n".join(L)


def send_report(chat_id, data, reply_to=None):

    text = build_report_text(data)

    chunks = split_text(text)

    markup = None

    if data["sal_found"]:

        markup = types.InlineKeyboardMarkup()

        markup.add(
            types.InlineKeyboardButton(
                "📥 تحميل قسيمة الراتب PDF",
                callback_data=f"pdf_{employee_key(data)}"
            )
        )

    for i, chunk in enumerate(chunks):

        last = i == len(chunks) - 1

        kwargs = dict(
            parse_mode="HTML",
            reply_markup=markup if last else None
        )

        if i == 0 and reply_to:

            bot.send_message(
                chat_id,
                chunk,
                reply_to_message_id=reply_to,
                **kwargs
            )

        else:

            bot.send_message(chat_id, chunk, **kwargs)


# =========================================================
# زر PDF
# =========================================================

@bot.callback_query_handler(
    func=lambda call: call.data.startswith("pdf_")
)
def handle_pdf_callback(call):

    key = resolve_callback_key(
        call.data.replace("pdf_", "", 1)
    )

    bot.answer_callback_query(
        call.id,
        "⏳ جاري إنشاء قسيمة الراتب..."
    )

    if not key:

        bot.send_message(
            call.message.chat.id,
            "⚠️ انتهت صلاحية الزر، أعد إرسال الاسم أو الرقم "
            "الوطني ثم اضغط الزر مرة أخرى."
        )

        return

    data = fetch_employee_data(key, choose_first=True)

    if data is None or data.get("ambiguous"):

        bot.send_message(
            call.message.chat.id,
            "⚠️ تعذر العثور على بيانات الموظف."
        )

        return

    if not data["sal_found"]:

        bot.send_message(
            call.message.chat.id,
            "⚠️ لا يوجد سجل راتب لهذا الموظف في الكشف الحالي، "
            "لذلك لا يمكن إصدار قسيمة."
        )

        return

    try:

        pdf_file = generate_professional_pdf(data)

        filename = (
            f"Salary_Slip_{data['report_month']:02d}_"
            f"{data['report_year']}_{data['name']}.pdf"
        )

        bot.send_document(
            call.message.chat.id,
            pdf_file,
            visible_file_name=filename,
            caption=(
                f"📄 <b>قسيمة راتب شهر "
                f"{_html.escape(data['report_month_name'])} "
                f"{data['report_year']}</b>\n"
                f"👤 الموظف: {_html.escape(data['name'])}"
            ),
            parse_mode="HTML"
        )

    except Exception as e:

        print(f"❌ خطأ إنشاء PDF: {e}")

        bot.send_message(
            call.message.chat.id,
            "❌ حدث خطأ أثناء إنشاء قسيمة PDF."
        )


# =========================================================
# اختيار موظف عند تشابه الأسماء (للإدارة)
# =========================================================

@bot.callback_query_handler(
    func=lambda call: call.data.startswith("pick_")
)
def handle_pick_callback(call):

    key = resolve_callback_key(
        call.data.replace("pick_", "", 1)
    )

    bot.answer_callback_query(call.id)

    if not key:

        bot.send_message(
            call.message.chat.id,
            "⚠️ انتهت صلاحية الاختيار، أعد البحث."
        )

        return

    data = fetch_employee_data(key, choose_first=True)

    if data is None or data.get("ambiguous"):

        bot.send_message(
            call.message.chat.id,
            "⚠️ تعذر العثور على بيانات الموظف."
        )

        return

    send_report(call.message.chat.id, data)


# =========================================================
# /debug - تشخيص حساب راتب موظف (للإدارة فقط)
# =========================================================

@bot.message_handler(commands=["debug"])
def handle_debug(message):

    if message.from_user.id != ADMIN_ID:

        bot.reply_to(message, "⛔ هذا الأمر مخصص للإدارة فقط.")

        return

    parts = message.text.split(maxsplit=1)

    if len(parts) < 2 or not parts[1].strip():

        bot.reply_to(
            message,
            "الاستخدام: /debug <اسم الموظف أو رقمه الوطني>\n"
            "يعرض كيف تم ربط أعمدة الكشف بقيم الراتب."
        )

        return

    data = fetch_employee_data(parts[1].strip(), choose_first=True)

    if data is None:

        bot.reply_to(message, "⚠️ الموظف غير موجود.")

        return

    text = "🔧 تشخيص الراتب\n\n" + "\n".join(data["debug"])

    for chunk in split_text(text):

        bot.send_message(message.chat.id, chunk)


# =========================================================
# الاستعلام عن الموظف
# =========================================================

@bot.message_handler(
    func=lambda message: True,
    content_types=["text"]
)
def handle_query(message):

    user_input = message.text.strip()

    if not user_input or user_input.startswith("/"):
        return

    print(f"🔎 استعلام جديد: {user_input}")

    data = fetch_employee_data(user_input)

    if data is None:

        bot.reply_to(
            message,
            "⚠️ لم يتم العثور على الرقم أو الاسم.\n\n"
            "يرجى التأكد من الرقم الوطني أو كتابة الاسم بشكل صحيح."
        )

        return

    if data.get("ambiguous"):

        # لا نعرض قائمة أسماء الموظفين لغير الإدارة (خصوصية)
        if message.from_user.id != ADMIN_ID:

            bot.reply_to(
                message,
                "⚠️ الاسم مشترك بين أكثر من موظف.\n"
                "الرجاء إرسال الرقم الوطني أو الاسم الكامل."
            )

            return

        markup = types.InlineKeyboardMarkup()

        for emp in data["ambiguous"]:

            markup.add(
                types.InlineKeyboardButton(
                    f"{emp['name']} ({emp['nat_id']})",
                    callback_data=f"pick_{employee_key(emp)}"
                )
            )

        bot.reply_to(
            message,
            "وُجد أكثر من موظف مطابق، اختر الموظف المطلوب:",
            reply_markup=markup
        )

        return

    send_report(
        message.chat.id,
        data,
        reply_to=message.message_id
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