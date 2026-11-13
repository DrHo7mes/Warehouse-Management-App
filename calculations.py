import math
from datetime import datetime, timedelta
import jdatetime


def get_today_jalali_str():
    """برگرداندن تاریخ امروز به شمسی برای نمایش در بالای برنامه"""
    now = jdatetime.datetime.now()
    weekdays = ["دوشنبه", "سه‌شنبه", "چهارشنبه", "پنج‌شنبه", "جمعه", "شنبه", "یکشنبه"]
    # در jdatetime شنبه=0 است
    j_weekdays = ["شنبه", "یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنج‌شنبه", "جمعه"]
    day_name = j_weekdays[now.weekday()]
    return f"{day_name}، {now.strftime('%Y/%m/%d')}"


def iso_to_jalali(iso_str, include_time=False):
    """تبدیل تاریخ استاندارد دیتابیس به تاریخ شمسی خوانا"""
    if not iso_str:
        return "-"
    try:
        clean_str = str(iso_str).strip()
        if len(clean_str) >= 19:
            dt = datetime.strptime(clean_str[:19], "%Y-%m-%d %H:%M:%S")
        else:
            dt = datetime.strptime(clean_str[:10], "%Y-%m-%d")
        jdt = jdatetime.datetime.fromgregorian(datetime=dt)
        if include_time:
            return jdt.strftime("%Y/%m/%d - %H:%M")
        return jdt.strftime("%Y/%m/%d")
    except Exception:
        return str(iso_str)


def jalali_to_gregorian_str(jalali_str):
    """تبدیل تاریخ شمسی وارد شده توسط کاربر (مثلاً 1405/08/12) به میلادی برای فیلتر دیتابیس"""
    if not jalali_str or not jalali_str.strip():
        return None
    try:
        parts = jalali_str.strip().replace("-", "/").split("/")
        if len(parts) != 3:
            return None
        jy, jm, jd = int(parts[0]), int(parts[1]), int(parts[2])
        gdate = jdatetime.date(jy, jm, jd).togregorian()
        return gdate.strftime("%Y-%m-%d")
    except Exception:
        return None


def calculate_depletion_info(stock, daily_usage, critical_threshold=7, warning_threshold=14):
    """
    محاسبه دقیق تعداد روزهای باقی‌مانده، تاریخ اتمام موجودی به شمسی و وضعیت هشدار
    طبق قوانین مرحله ششم و هشتم
    """
    try:
        stock = float(stock)
    except (ValueError, TypeError):
        stock = 0.0

    try:
        daily_usage = float(daily_usage)
    except (ValueError, TypeError):
        daily_usage = 0.0

    # قانون ۲: اگر موجودی صفر باشد
    if stock <= 0:
        return {
            "days_left": 0,
            "days_display": "۰ روز",
            "depletion_date": "تمام شده",
            "status_code": "OUT_OF_STOCK",
            "status_text": "اتمام موجودی",
            "color_bg": "#dcdde1",
            "color_fg": "#2f3640",
            "sort_priority": 0
        }

    # قانون ۱: اگر مصرف روزانه صفر باشد
    if daily_usage <= 0:
        return {
            "days_left": None,
            "days_display": "نامشخص (مصرف صفر)",
            "depletion_date": "پیش‌بینی امکان‌پذیر نیست",
            "status_code": "NO_USAGE",
            "status_text": "بدون مصرف روزانه",
            "color_bg": "#d1ecf1",
            "color_fg": "#0c5460",
            "sort_priority": 999999
        }

    # فرمول محاسبه با گرد کردن رو به بالا (قانون ۸)
    raw_days = stock / daily_usage
    days_left = math.ceil(raw_days)

    today_dt = datetime.now()
    depletion_dt = today_dt + timedelta(days=days_left)
    depletion_jalali = jdatetime.date.fromgregorian(date=depletion_dt.date()).strftime("%Y/%m/%d")

    # تعیین وضعیت بر اساس روزهای باقی‌مانده (قوانین ۳، ۴ و ۵)
    if days_left < int(critical_threshold):
        status_code = "CRITICAL"
        status_text = "هشدار بحرانی"
        color_bg = "#f8d7da"
        color_fg = "#721c24"
    elif days_left <= int(warning_threshold):
        status_code = "WARNING"
        status_text = "کمبود موجودی"
        color_bg = "#fff3cd"
        color_fg = "#856404"
    else:
        status_code = "NORMAL"
        status_text = "وضعیت عادی"
        color_bg = "#d4edda"
        color_fg = "#155724"

    return {
        "days_left": days_left,
        "days_display": f"{days_left} روز",
        "depletion_date": depletion_jalali,
        "status_code": status_code,
        "status_text": status_text,
        "color_bg": color_bg,
        "color_fg": color_fg,
        "sort_priority": days_left
    }


def enrich_item_with_calculations(item_dict, critical_threshold=7, warning_threshold=14):
    """افزودن اطلاعات محاسباتی و تاریخ شمسی به دیکشنری یک کالا"""
    info = calculate_depletion_info(
        item_dict.get("stock", 0),
        item_dict.get("daily_usage", 0),
        critical_threshold,
        warning_threshold
    )
    enriched = dict(item_dict)
    enriched.update(info)
    enriched["updated_at_jalali"] = iso_to_jalali(item_dict.get("updated_at"), include_time=True)
    return enriched