import io
import os
import re
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont, ImageOps
from PyPDF2 import PdfReader, PdfWriter

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    ARABIC_LIBS_AVAILABLE = True
except ModuleNotFoundError:
    arabic_reshaper = None
    get_display = None
    ARABIC_LIBS_AVAILABLE = False


st.set_page_config(page_title="نظام بطاقات الطلاب", page_icon="🪪", layout="wide")

DPI = 300
A4_WIDTH = int(8.27 * DPI)
A4_HEIGHT = int(11.69 * DPI)
BACKGROUND_COLOR = "white"
DEFAULT_EXCEL_PATH = Path(__file__).with_name("students.xlsx")
SELECTION_COLUMN = "تحديد"
MANUAL_NAME_COLUMN = "الاسم"
MANUAL_SERIAL_COLUMN = "الرقم المتسلسل"


def make_safe_filename(text: str) -> str:
    text = str(text).strip()
    text = re.sub(r'[<>:"/\\|?*]+', "_", text)
    text = re.sub(r"\s+", "_", text)
    return text


@st.cache_data(show_spinner=False)
def read_excel_file(file_source) -> pd.DataFrame:
    return pd.read_excel(file_source)


def save_excel_file(df: pd.DataFrame, output_target=None):
    if isinstance(output_target, (str, Path)):
        output_path = Path(output_target)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(output_path, index=False)
        return output_path

    output_buffer = io.BytesIO()
    df.to_excel(output_buffer, index=False)
    output_buffer.seek(0)
    return output_buffer.getvalue()


def with_selection_column(df: pd.DataFrame) -> pd.DataFrame:
    prepared_df = df.copy().fillna("")
    if SELECTION_COLUMN not in prepared_df.columns:
        prepared_df.insert(0, SELECTION_COLUMN, True)
    else:
        prepared_df[SELECTION_COLUMN] = prepared_df[SELECTION_COLUMN].fillna(True).astype(bool)
    return prepared_df


def without_selection_column(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop(columns=[SELECTION_COLUMN], errors="ignore").copy()


@st.cache_data(show_spinner=False)
def reshape_arabic_text(text: str) -> str:
    if not ARABIC_LIBS_AVAILABLE:
        return str(text)
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)


def load_font(font_bytes: bytes | None, size: int, fallback_candidates: list[str] | None = None):
    fallback_candidates = fallback_candidates or []

    if font_bytes:
        try:
            return ImageFont.truetype(io.BytesIO(font_bytes), size)
        except Exception:
            pass

    for candidate in fallback_candidates:
        if os.path.exists(candidate):
            try:
                return ImageFont.truetype(candidate, size)
            except Exception:
                continue

    try:
        return ImageFont.truetype("arial.ttf", size)
    except Exception:
        return ImageFont.load_default()


def draw_right(draw: ImageDraw.ImageDraw, text: str, font, right_x: int, y: int, color):
    bbox = draw.textbbox((0, 0), text, font=font)
    width = bbox[2] - bbox[0]
    x = right_x - width
    draw.text((x, y), text, font=font, fill=color)


def hex_to_rgb(hex_color: str):
    hex_color = hex_color.strip().lstrip("#")
    if len(hex_color) != 6:
        return (255, 255, 255)
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


def generate_barcode_image(serial: str, write_text: bool, module_height: int = 18):
    barcode_buffer = io.BytesIO()
    barcode = Code128(serial, writer=ImageWriter())
    barcode.write(
        barcode_buffer,
        options={
            "write_text": write_text,
            "module_height": module_height,
            "dpi": 300,
        },
    )
    barcode_buffer.seek(0)
    return Image.open(barcode_buffer).convert("RGBA")


def create_resized_barcode(serial: str, settings: dict) -> Image.Image:
    barcode_img = generate_barcode_image(
        serial,
        write_text=settings["write_text_under_barcode"],
        module_height=settings["barcode_module_height"],
    )
    return barcode_img.resize((settings["barcode_width"], settings["barcode_height"]))


def open_image_uploaded(uploaded_file) -> Image.Image:
    return Image.open(uploaded_file).convert("RGBA")


def create_front_card(
    template_img: Image.Image,
    name: str,
    serial: str,
    settings: dict,
    arabic_font_bytes: bytes | None,
    latin_font_bytes: bytes | None,
) -> Image.Image:
    base = template_img.copy().convert("RGBA")
    draw = ImageDraw.Draw(base)

    name_font = load_font(
        arabic_font_bytes,
        settings["name_font_size"],
        fallback_candidates=[
            "C:/Windows/Fonts/arial.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ],
    )
    serial_font = load_font(
        latin_font_bytes,
        settings["serial_font_size"],
        fallback_candidates=[
            "C:/Windows/Fonts/arial.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ],
    )

    display_name = reshape_arabic_text(name) if settings["reshape_arabic"] else str(name)
    serial = str(serial)

    if settings["draw_name"]:
        draw_right(
            draw,
            display_name,
            name_font,
            settings["name_right_x"],
            settings["name_y"],
            settings["name_color"],
        )

    if settings["draw_serial_text"]:
        draw_right(
            draw,
            serial,
            serial_font,
            settings["serial_right_x"],
            settings["serial_y"],
            settings["serial_color"],
        )

    if settings["draw_barcode"]:
        barcode_img = create_resized_barcode(serial, settings)
        base.paste(barcode_img, (settings["barcode_x"], settings["barcode_y"]), barcode_img)

    return base.convert("RGB")


def create_back_card(back_template_img: Image.Image | None) -> Image.Image | None:
    if back_template_img is None:
        return None
    return back_template_img.copy().convert("RGB")


def save_card(image: Image.Image, output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(output_path)


def fit_image_in_box(img: Image.Image, box_w: int, box_h: int) -> Image.Image:
    return ImageOps.contain(img, (box_w, box_h))


def open_local_image_as_rgb(path: Path) -> Image.Image:
    img = Image.open(path)
    if img.mode in ("RGBA", "LA"):
        bg = Image.new("RGB", img.size, "white")
        alpha = img.getchannel("A")
        bg.paste(img, mask=alpha)
        return bg
    return img.convert("RGB")


def create_pages_from_images(
    image_files: list[Path],
    cols: int,
    rows: int,
    page_margin_x: int,
    page_margin_y: int,
    gap_x: int,
    gap_y: int,
) -> list[Image.Image]:
    pages = []
    cards_per_page = cols * rows

    usable_width = A4_WIDTH - (2 * page_margin_x) - ((cols - 1) * gap_x)
    usable_height = A4_HEIGHT - (2 * page_margin_y) - ((rows - 1) * gap_y)
    cell_w = usable_width // cols
    cell_h = usable_height // rows

    for start in range(0, len(image_files), cards_per_page):
        batch = image_files[start:start + cards_per_page]
        page = Image.new("RGB", (A4_WIDTH, A4_HEIGHT), BACKGROUND_COLOR)

        for i, img_path in enumerate(batch):
            row = i // cols
            col = i % cols

            x = page_margin_x + col * (cell_w + gap_x)
            y = page_margin_y + row * (cell_h + gap_y)

            img = open_local_image_as_rgb(img_path)
            fitted = fit_image_in_box(img, cell_w, cell_h)

            paste_x = x + (cell_w - fitted.width) // 2
            paste_y = y + (cell_h - fitted.height) // 2
            page.paste(fitted, (paste_x, paste_y))

        pages.append(page)

    return pages


def save_pages_as_pdf(pages: list[Image.Image], output_pdf: Path):
    if not pages:
        return

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    pages[0].save(
        output_pdf,
        "PDF",
        resolution=DPI,
        save_all=True,
        append_images=pages[1:],
    )


def merge_pdfs_alternating(pdf1_path: Path, pdf2_path: Path, output_path: Path):
    reader1 = PdfReader(str(pdf1_path))
    reader2 = PdfReader(str(pdf2_path))
    writer = PdfWriter()

    max_pages = max(len(reader1.pages), len(reader2.pages))
    for i in range(max_pages):
        if i < len(reader1.pages):
            writer.add_page(reader1.pages[i])
        if i < len(reader2.pages):
            writer.add_page(reader2.pages[i])

    with open(output_path, "wb") as f:
        writer.write(f)


def zip_folder(folder_path: Path, zip_path: Path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                full_path = Path(root) / file
                rel_path = full_path.relative_to(folder_path)
                zf.write(full_path, rel_path)


def get_excel_source_key(excel_source) -> str:
    if isinstance(excel_source, Path):
        timestamp = excel_source.stat().st_mtime if excel_source.exists() else 0
        return f"path:{excel_source.resolve()}:{timestamp}"
    return f"upload:{excel_source.name}:{excel_source.size}"


def load_editor_dataframe(excel_source) -> pd.DataFrame:
    source_key = get_excel_source_key(excel_source)
    if st.session_state.get("editor_source_key") != source_key:
        st.session_state["editor_df"] = with_selection_column(read_excel_file(excel_source))
        st.session_state["editor_source_key"] = source_key
    return st.session_state["editor_df"].copy()


def load_manual_dataframe() -> pd.DataFrame:
    if "manual_editor_df" not in st.session_state:
        st.session_state["manual_editor_df"] = with_selection_column(
            pd.DataFrame([{MANUAL_NAME_COLUMN: "", MANUAL_SERIAL_COLUMN: ""}])
        )
    return st.session_state["manual_editor_df"].copy()


def build_system(
    df: pd.DataFrame,
    front_template_file,
    back_template_file,
    arabic_font_file,
    latin_font_file,
    settings: dict,
    sheet_settings: dict,
) -> dict:
    required_columns = [settings["name_column"], settings["serial_column"]]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"الأعمدة التالية غير موجودة في ملف الإكسل: {missing}")

    temp_root = Path(tempfile.mkdtemp(prefix="student_card_system_"))
    fronts_dir = temp_root / "front_cards"
    backs_dir = temp_root / "back_cards"
    barcodes_dir = temp_root / "barcodes"
    pdf_dir = temp_root / "pdf"

    fronts_dir.mkdir(parents=True, exist_ok=True)
    backs_dir.mkdir(parents=True, exist_ok=True)
    barcodes_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    arabic_font_bytes = arabic_font_file.getvalue() if arabic_font_file else None
    latin_font_bytes = latin_font_file.getvalue() if latin_font_file else None
    front_template_img = open_image_uploaded(front_template_file)
    back_template_img = open_image_uploaded(back_template_file) if back_template_file else None

    created_fronts = []
    created_backs = []
    created_barcodes = []

    for index, row in df.iterrows():
        name = str(row[settings["name_column"]]).strip()
        serial = str(row[settings["serial_column"]]).strip()
        safe_serial = make_safe_filename(serial)

        front_card = create_front_card(
            template_img=front_template_img,
            name=name,
            serial=serial,
            settings=settings,
            arabic_font_bytes=arabic_font_bytes,
            latin_font_bytes=latin_font_bytes,
        )
        front_path = fronts_dir / f"{index + 1}_{safe_serial}.png"
        save_card(front_card, front_path)
        created_fronts.append(front_path)

        if settings["draw_barcode"]:
            barcode_image = create_resized_barcode(serial, settings).convert("RGB")
            barcode_path = barcodes_dir / f"{index + 1}_{safe_serial}.png"
            save_card(barcode_image, barcode_path)
            created_barcodes.append(barcode_path)

        if back_template_img is not None:
            back_card = create_back_card(back_template_img)
            back_path = backs_dir / f"{index + 1}_{safe_serial}.png"
            save_card(back_card, back_path)
            created_backs.append(back_path)

    front_pdf = pdf_dir / "front_cards_a4.pdf"
    back_pdf = pdf_dir / "back_cards_a4.pdf"
    merged_pdf = pdf_dir / "merged_front_back.pdf"

    front_pages = create_pages_from_images(
        created_fronts,
        cols=sheet_settings["cols"],
        rows=sheet_settings["rows"],
        page_margin_x=sheet_settings["page_margin_x"],
        page_margin_y=sheet_settings["page_margin_y"],
        gap_x=sheet_settings["gap_x"],
        gap_y=sheet_settings["gap_y"],
    )
    save_pages_as_pdf(front_pages, front_pdf)

    if created_backs:
        back_pages = create_pages_from_images(
            created_backs,
            cols=sheet_settings["cols"],
            rows=sheet_settings["rows"],
            page_margin_x=sheet_settings["page_margin_x"],
            page_margin_y=sheet_settings["page_margin_y"],
            gap_x=sheet_settings["gap_x"],
            gap_y=sheet_settings["gap_y"],
        )
        save_pages_as_pdf(back_pages, back_pdf)
        merge_pdfs_alternating(front_pdf, back_pdf, merged_pdf)

    all_outputs_zip = temp_root / "student_cards_outputs.zip"
    zip_folder(temp_root, all_outputs_zip)

    return {
        "temp_root": temp_root,
        "fronts_dir": fronts_dir,
        "backs_dir": backs_dir,
        "barcodes_dir": barcodes_dir,
        "front_pdf": front_pdf if front_pdf.exists() else None,
        "back_pdf": back_pdf if back_pdf.exists() else None,
        "merged_pdf": merged_pdf if merged_pdf.exists() else None,
        "zip_file": all_outputs_zip,
        "front_count": len(created_fronts),
        "back_count": len(created_backs),
        "barcode_count": len(created_barcodes),
    }


st.title("🪪 نظام موحد لإنشاء بطاقات الطلاب")
st.caption("رفع الإكسل + القالب + الخطوط + إخراج PNG و PDF من واجهة ويب واحدة")

if not ARABIC_LIBS_AVAILABLE:
    st.warning("مكتبات دعم العربية غير مثبتة في هذه البيئة. ثبّت الحزم من requirements.txt للحصول على تشكيل وعرض عربي صحيح.")

with st.expander("فكرة النظام", expanded=True):
    st.markdown(
        """
        هذا التطبيق يدمج 3 مراحل في نظام واحد:
        1. إنشاء بطاقات الطلاب من ملف Excel وقالب صورة.
        2. تجميع البطاقات داخل صفحات A4 للطباعة.
        3. دمج ملف الواجهة والخلفية بالتناوب داخل PDF واحد.
        """
    )

left, right = st.columns([1, 1])

with left:
    st.subheader("1) الملفات المطلوبة")
    input_mode = st.radio(
        "مصدر البيانات",
        ["Excel", "إدخال مباشر"],
        horizontal=True,
    )
    use_local_excel = False
    if input_mode == "Excel" and DEFAULT_EXCEL_PATH.exists():
        use_local_excel = st.checkbox(
            "استخدام ملف students.xlsx المحلي",
            value=True,
            help="عند التفعيل سيتم تحميل الملف المحلي ويمكنك حفظ التعديلات عليه مباشرة من داخل البرنامج.",
        )
        st.caption(f"الملف المحلي المتاح: {DEFAULT_EXCEL_PATH.name}")

    excel_file = None
    if input_mode == "Excel":
        excel_file = st.file_uploader("ملف Excel", type=["xlsx", "xls"])
    else:
        st.caption("يمكنك كتابة الأسماء والأرقام المتسلسلة مباشرة من داخل الجدول بدون ملف Excel.")
    front_template_file = st.file_uploader("قالب الوجه الأمامي", type=["png", "jpg", "jpeg"])
    back_template_file = st.file_uploader("قالب الخلفية (اختياري)", type=["png", "jpg", "jpeg"])
    arabic_font_file = st.file_uploader("خط عربي (اختياري)", type=["ttf", "otf"])
    latin_font_file = st.file_uploader("خط لاتيني/إنجليزي (اختياري)", type=["ttf", "otf"])

with right:
    st.subheader("2) معاينة البيانات وتعديل الشيت")
    excel_source = DEFAULT_EXCEL_PATH if input_mode == "Excel" and use_local_excel and DEFAULT_EXCEL_PATH.exists() else excel_file

    if input_mode == "إدخال مباشر":
        try:
            editor_df = load_manual_dataframe()
            df_preview = st.data_editor(
                editor_df,
                use_container_width=True,
                num_rows="dynamic",
                key="student_sheet_editor",
            ).fillna("")
            st.session_state["manual_editor_df"] = with_selection_column(df_preview)

            st.write(f"عدد السجلات: {len(df_preview)}")
            available_columns = list(df_preview.columns)

            editor_col1, editor_col2 = st.columns(2)
            manual_excel_bytes = save_excel_file(without_selection_column(df_preview))
            editor_col1.download_button(
                "تنزيل الإدخال كملف Excel",
                data=manual_excel_bytes,
                file_name="manual_students.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            if editor_col2.button("إعادة ضبط الجدول اليدوي", use_container_width=True):
                st.session_state["manual_editor_df"] = with_selection_column(
                    pd.DataFrame([{MANUAL_NAME_COLUMN: "", MANUAL_SERIAL_COLUMN: ""}])
                )
                st.rerun()
        except Exception as e:
            st.error(f"تعذر تجهيز الإدخال اليدوي: {e}")
            available_columns = []
            df_preview = None
    elif excel_source is not None:
        try:
            editor_df = load_editor_dataframe(excel_source)
            df_preview = st.data_editor(
                editor_df,
                use_container_width=True,
                num_rows="dynamic",
                key="student_sheet_editor",
            ).fillna("")
            st.session_state["editor_df"] = with_selection_column(df_preview)

            st.write(f"عدد السجلات: {len(df_preview)}")
            available_columns = list(df_preview.columns)

            editor_col1, editor_col2 = st.columns(2)
            cleaned_df_preview = without_selection_column(df_preview)
            if isinstance(excel_source, Path):
                if editor_col1.button("حفظ التعديلات في students.xlsx", use_container_width=True):
                    save_excel_file(cleaned_df_preview, excel_source)
                    read_excel_file.clear()
                    st.session_state["editor_source_key"] = get_excel_source_key(excel_source)
                    st.session_state["editor_df"] = with_selection_column(cleaned_df_preview)
                    st.success("تم حفظ التعديلات داخل students.xlsx بنجاح.")
            else:
                edited_excel_bytes = save_excel_file(cleaned_df_preview)
                editor_col1.download_button(
                    "تنزيل نسخة Excel المعدلة",
                    data=edited_excel_bytes,
                    file_name=f"edited_{excel_source.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            if editor_col2.button("إعادة تحميل الشيت", use_container_width=True):
                read_excel_file.clear()
                st.session_state["editor_df"] = with_selection_column(read_excel_file(excel_source))
                st.rerun()
        except Exception as e:
            st.error(f"تعذر قراءة ملف الإكسل: {e}")
            available_columns = []
            df_preview = None
    else:
        available_columns = []
        df_preview = None
        st.info("ارفع ملف Excel أو فعّل الملف المحلي، أو اختر الإدخال المباشر لكتابة البيانات من داخل النظام.")

st.divider()

c1, c2, c3 = st.columns(3)
with c1:
    st.subheader("3) ربط الأعمدة")
    selectable_columns = [col for col in available_columns if col != SELECTION_COLUMN]
    if input_mode == "إدخال مباشر":
        default_name_options = selectable_columns if selectable_columns else [MANUAL_NAME_COLUMN]
        default_name_index = default_name_options.index(MANUAL_NAME_COLUMN) if MANUAL_NAME_COLUMN in default_name_options else 0
        default_serial_index = default_name_options.index(MANUAL_SERIAL_COLUMN) if MANUAL_SERIAL_COLUMN in default_name_options else min(1, len(default_name_options) - 1)
    else:
        default_name_options = selectable_columns if selectable_columns else ["Name"]
        default_name_index = 0
        default_serial_index = 1 if len(default_name_options) > 1 else 0

    name_column = st.selectbox(
        "عمود الاسم",
        options=default_name_options,
        index=default_name_index,
    )
    serial_column = st.selectbox(
        "عمود الرقم/السيريال",
        options=default_name_options,
        index=default_serial_index,
    )
    reshape_arabic = st.checkbox("تفعيل تشكيل/معالجة العربية", value=True)
    st.caption("استخدم عمود `تحديد` داخل الجدول لاختيار الكروت التي تريد حفظها فقط.")

with c2:
    st.subheader("4) النصوص")
    draw_name = st.checkbox("كتابة الاسم", value=True)
    draw_serial_text = st.checkbox("كتابة الرقم", value=True)
    name_font_size = st.number_input("حجم خط الاسم", min_value=10, max_value=200, value=70)
    serial_font_size = st.number_input("حجم خط الرقم", min_value=10, max_value=120, value=34)
    name_color_hex = st.color_picker("لون الاسم", value="#FFFFFF")
    serial_color_hex = st.color_picker("لون الرقم", value="#FFFFFF")

with c3:
    st.subheader("5) الباركود")
    draw_barcode = st.checkbox("إضافة باركود", value=True)
    write_text_under_barcode = st.checkbox("كتابة الرقم أسفل الباركود", value=False)
    barcode_module_height = st.number_input("ارتفاع الباركود الداخلي", min_value=5, max_value=50, value=18)
    barcode_width = st.number_input("عرض الباركود", min_value=50, max_value=1000, value=430)
    barcode_height = st.number_input("ارتفاع الباركود", min_value=20, max_value=500, value=95)

st.divider()

p1, p2, p3 = st.columns(3)
with p1:
    st.subheader("6) مواقع الوجه الأمامي")
    name_right_x = st.number_input("نهاية الاسم يمينًا", min_value=0, max_value=5000, value=1100)
    name_y = st.number_input("موضع الاسم Y", min_value=0, max_value=5000, value=500)
    serial_right_x = st.number_input("نهاية الرقم يمينًا", min_value=0, max_value=5000, value=1100)
    serial_y = st.number_input("موضع الرقم Y", min_value=0, max_value=5000, value=790)

with p2:
    st.subheader("7) موضع الباركود")
    barcode_x = st.number_input("باركود X", min_value=0, max_value=5000, value=650)
    barcode_y = st.number_input("باركود Y", min_value=0, max_value=5000, value=870)

with p3:
    st.subheader("8) تخطيط صفحات A4")
    cols = st.number_input("عدد الأعمدة", min_value=1, max_value=10, value=2)
    rows = st.number_input("عدد الصفوف", min_value=1, max_value=10, value=5)
    page_margin_x = st.number_input("الهامش الأفقي", min_value=0, max_value=500, value=80)
    page_margin_y = st.number_input("الهامش الرأسي", min_value=0, max_value=500, value=80)
    gap_x = st.number_input("المسافة الأفقية", min_value=0, max_value=500, value=40)
    gap_y = st.number_input("المسافة الرأسية", min_value=0, max_value=500, value=30)

settings = {
    "name_column": name_column,
    "serial_column": serial_column,
    "reshape_arabic": reshape_arabic,
    "draw_name": draw_name,
    "draw_serial_text": draw_serial_text,
    "name_font_size": int(name_font_size),
    "serial_font_size": int(serial_font_size),
    "name_right_x": int(name_right_x),
    "name_y": int(name_y),
    "serial_right_x": int(serial_right_x),
    "serial_y": int(serial_y),
    "name_color": hex_to_rgb(name_color_hex),
    "serial_color": hex_to_rgb(serial_color_hex),
    "draw_barcode": draw_barcode,
    "write_text_under_barcode": write_text_under_barcode,
    "barcode_module_height": int(barcode_module_height),
    "barcode_width": int(barcode_width),
    "barcode_height": int(barcode_height),
    "barcode_x": int(barcode_x),
    "barcode_y": int(barcode_y),
}

sheet_settings = {
    "cols": int(cols),
    "rows": int(rows),
    "page_margin_x": int(page_margin_x),
    "page_margin_y": int(page_margin_y),
    "gap_x": int(gap_x),
    "gap_y": int(gap_y),
}

run = st.button("🚀 إنشاء النظام والملفات", type="primary", use_container_width=True)

if run:
    if front_template_file is None:
        st.error("ارفع قالب الوجه الأمامي أولًا.")
    else:
        try:
            if input_mode == "إدخال مباشر":
                df = st.session_state.get("manual_editor_df")
                if df is None:
                    df = load_manual_dataframe()
            else:
                if excel_source is None:
                    st.error("ارفع ملف Excel أو فعّل الملف المحلي أولًا.")
                    st.stop()
                df = st.session_state.get("editor_df")
                if df is None:
                    df = with_selection_column(read_excel_file(excel_source))

            df = with_selection_column(df).fillna("")
            selected_df = df[df[SELECTION_COLUMN].fillna(False)].copy()
            selected_df = without_selection_column(selected_df)
            selected_df = selected_df[
                selected_df[name_column].astype(str).str.strip().ne("")
                & selected_df[serial_column].astype(str).str.strip().ne("")
            ].copy()

            if selected_df.empty:
                st.error("لا توجد سجلات محددة وصالحة للتنفيذ. فعّل `تحديد` أمام الكروت المطلوبة واكتب الاسم والرقم المتسلسل.")
                st.stop()

            with st.spinner("جاري إنشاء البطاقات والملفات..."):
                result = build_system(
                    df=selected_df,
                    front_template_file=front_template_file,
                    back_template_file=back_template_file,
                    arabic_font_file=arabic_font_file,
                    latin_font_file=latin_font_file,
                    settings=settings,
                    sheet_settings=sheet_settings,
                )

            st.success("تم تنفيذ النظام بنجاح")

            a, b, c, d = st.columns(4)
            a.metric("عدد بطاقات الوجه", result["front_count"])
            b.metric("عدد بطاقات الخلفية", result["back_count"])
            c.metric("عدد الباركودات", result["barcode_count"])
            d.metric("بطاقات لكل صفحة", int(cols) * int(rows))

            st.subheader("التحميل")
            with open(result["zip_file"], "rb") as f:
                st.download_button(
                    "تحميل كل المخرجات ZIP",
                    data=f.read(),
                    file_name="student_cards_outputs.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

            d1, d2, d3 = st.columns(3)
            if result["front_pdf"] and result["front_pdf"].exists():
                with open(result["front_pdf"], "rb") as f:
                    d1.download_button(
                        "تحميل PDF الوجه",
                        data=f.read(),
                        file_name="front_cards_a4.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

            if result["back_pdf"] and result["back_pdf"].exists():
                with open(result["back_pdf"], "rb") as f:
                    d2.download_button(
                        "تحميل PDF الخلفية",
                        data=f.read(),
                        file_name="back_cards_a4.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

            if result["merged_pdf"] and result["merged_pdf"].exists():
                with open(result["merged_pdf"], "rb") as f:
                    d3.download_button(
                        "تحميل PDF المدمج",
                        data=f.read(),
                        file_name="merged_front_back.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

            preview_files = sorted(result["fronts_dir"].glob("*.png"))[:6]
            if preview_files:
                st.subheader("معاينة أول 6 بطاقات")
                preview_cols = st.columns(3)
                for i, file in enumerate(preview_files):
                    with preview_cols[i % 3]:
                        st.image(str(file), caption=file.name, use_container_width=True)
        except Exception as e:
            st.exception(e)

st.divider()
st.subheader("تشغيل محلي")
st.code(
    """pip install streamlit pandas pillow python-barcode openpyxl PyPDF2 arabic-reshaper python-bidi
streamlit run student_card_system.py""",
    language="bash",
)

st.info(
    "يمكنك الآن تعديل بيانات الشيت من داخل البرنامج مباشرة، ثم الحفظ إلى الملف المحلي أو تنزيل نسخة معدلة عند استخدام ملف مرفوع."
)
