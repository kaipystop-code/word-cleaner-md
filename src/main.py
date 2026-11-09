"""
WORD清洗MD 主入口程式 (main.py)
用途：批次掃描 input/ 目錄下的 .docx 檔案，轉換並清洗為高品質 Markdown 格式，輸出至 output/ 目錄。
特色：
1. 完整支援 Word 表格轉換為標準 Markdown 表格。
2. 自動抽取內嵌圖片至 output/images/ 並建立 Markdown 相對連結。
3. 智慧辨識並轉譯 Word 勾選方塊 ([x] / [ ])。
4. 清洗目錄錨點 (TOC) 雜訊、多餘空行與階層清單排版。
"""

import re
import sys
import logging
from pathlib import Path
from typing import Callable, Dict, Any

import mammoth
from bs4 import BeautifulSoup
import markdownify

# 配置標準日誌格式
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)


def create_image_handler(images_dir: Path, rel_prefix: str, file_stem: str) -> Callable:
    """
    建立自訂圖片抽取處理器，將 Word 內嵌圖片存檔並回傳相對 URL
    
    :param images_dir: 圖片存檔目錄 (實體路徑)
    :param rel_prefix: Markdown 中引用的相對路徑前綴
    :param file_stem: 來源檔案名稱 (用於命名圖片)
    :return: mammoth 所需的圖片轉換函式
    """
    image_counter = 0

    def handle_image(image) -> Dict[str, str]:
        nonlocal image_counter
        image_counter += 1

        # 決定副檔名
        ext = image.content_type.split("/")[-1].lower()
        if ext in ("jpeg", "jpg"):
            ext = "jpg"
        elif ext not in ("png", "gif", "svg", "webp"):
            ext = "png"

        image_filename = f"{file_stem}_img_{image_counter:02d}.{ext}"
        image_filepath = images_dir / image_filename

        try:
            with image.open() as src, open(image_filepath, "wb") as dst:
                dst.write(src.read())
            logger.debug("已抽取圖片: %s", image_filename)
        except OSError as err:
            logger.error("儲存圖片 %s 失敗: %s", image_filename, str(err))

        return {"src": f"{rel_prefix}/{image_filename}"}

    return mammoth.images.inline(handle_image)


def clean_html(raw_html: str) -> str:
    """
    使用 BeautifulSoup 前置清洗 HTML 結構
    - 轉譯勾選方塊為 [x] 與 [ ]
    - 扁平化表格內被誤設的標題標籤 (h1~h6)，避免破壞表格語意
    - 解除無內容的純巢狀清單 (li 只包覆 ul/ol 而無直接文字)，避免多重符號 (* + -)
    - 清理無文字內容的空段落與空標籤
    
    :param raw_html: mammoth 轉換後的原始 HTML 字串
    :return: 清洗後的 HTML 字串
    """
    soup = BeautifulSoup(raw_html, "html.parser")

    # 1. 處理核取方塊 (Checkbox)
    for checkbox in soup.find_all("input", type="checkbox"):
        if checkbox.has_attr("checked"):
            checkbox.replace_with("[x] ")
        else:
            checkbox.replace_with("[ ] ")

    # 2. 避免表格內出現 <h1>~<h6> 標籤破壞 Markdown 表格語意
    for table in soup.find_all("table"):
        for heading in table.find_all(["h1", "h2", "h3", "h4", "h5", "h6"]):
            heading.name = "strong"

    # 3. 解除巢狀空清單項目 (消除 * + - 多餘符號堆疊)
    for li in soup.find_all("li"):
        direct_text = "".join(li.find_all(string=True, recursive=False)).strip()
        child_lists = li.find_all(["ul", "ol"], recursive=False)
        if not direct_text and child_lists:
            li.unwrap()

    # 4. 清理純空白段落
    for p in soup.find_all("p"):
        if not p.text.strip() and not p.find_all(["img", "input"]):
            p.decompose()

    return str(soup)


def clean_markdown_text(raw_text: str) -> str:
    """
    後置文字清洗器：清理 Markdown 語法符號雜訊、Word 目錄錨點與多餘空行
    
    :param raw_text: 由 markdownify 產出的原始 Markdown 內容
    :return: 結構化、乾淨的 Markdown 文本
    """
    # 1. 移除無效的 Word TOC 錨點標籤 (例如 <a id="_Toc..."></a>)
    cleaned = re.sub(r'<a\s+id="_Toc\d+"></a>', '', raw_text)

    # 2. 修復被 Markdown 反斜線過度跳脫的符號 (如 QA30\-01 -> QA30-01, 8\.5 -> 8.5)
    cleaned = re.sub(r'\\([_#\-\.])', r'\1', cleaned)

    # 3. 清除行首殘留的無意義巢狀符號組合 (如 '* + - 1.' 簡化為 '1.')
    cleaned = re.sub(r'^\s*([*+-]\s+){1,}(?=\d+\.|\w|\[)', '', cleaned, flags=re.MULTILINE)

    # 4. 清理空清單符號階層
    cleaned = re.sub(r'(?:^\s*[-*+]\s*\n)+', '', cleaned, flags=re.MULTILINE)

    # 5. 行層級處理：去尾端空白與壓縮連續空行
    lines = cleaned.splitlines()
    processed_lines = []
    consecutive_empty = 0

    for line in lines:
        stripped = line.rstrip()
        if not stripped:
            consecutive_empty += 1
            if consecutive_empty <= 1:
                processed_lines.append("")
        else:
            consecutive_empty = 0
            processed_lines.append(stripped)

    return "\n".join(processed_lines).strip() + "\n"


def convert_docx_to_markdown(docx_path: Path, output_dir: Path) -> Path:
    """
    將單一 Word (.docx) 檔案轉換並清洗為 Markdown 檔案
    
    :param docx_path: Word 來源檔案路徑
    :param output_dir: 輸出目標目錄
    :return: 產出的 Markdown 檔案路徑
    """
    logger.info("開始處理檔案: %s", docx_path.name)
    file_stem = docx_path.stem

    # 準備圖片目錄
    images_dir = output_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)
    image_handler = create_image_handler(images_dir, rel_prefix="images", file_stem=file_stem)

    try:
        # 1. 使用 mammoth 轉譯為語意化 HTML
        with open(docx_path, "rb") as docx_file:
            conversion_result = mammoth.convert_to_html(
                docx_file,
                convert_image=image_handler
            )

        # 2. 前置 HTML 結構清洗
        cleaned_html = clean_html(conversion_result.value)

        # 3. 轉譯為 Markdown (GFM 風格表格與標題)
        raw_md = markdownify.markdownify(
            cleaned_html,
            heading_style="ATX",
            sub_symbol="",
            sup_symbol=""
        )

        # 4. 後置 Markdown 文字清洗
        final_md = clean_markdown_text(raw_md)

        # 5. 寫入輸出檔案
        output_file = output_dir / f"{file_stem}.md"
        output_file.write_text(final_md, encoding="utf-8")

        logger.info("成功產出 Markdown 文件: %s", output_file.name)
        if conversion_result.messages:
            logger.debug("Mammoth 轉換提示訊息數: %d", len(conversion_result.messages))

        return output_file
    except Exception as exc:
        logger.error("處理檔案 '%s' 失敗: %s", docx_path.name, str(exc), exc_info=True)
        raise exc


def batch_process(input_dir: Path, output_dir: Path) -> int:
    """
    批次掃描並轉換 input 目錄下的所有 Word 檔案
    
    :param input_dir: 來源目錄
    :param output_dir: 輸出目錄
    :return: 成功轉換的檔案總數
    """
    docx_files = sorted(list(input_dir.glob("*.docx")))
    if not docx_files:
        logger.warning("在 '%s' 目錄下未找到任何 .docx 檔案。", input_dir)
        logger.info("請將欲轉換的 Word 檔案放入該目錄後重新執行。")
        return 0

    logger.info("找到 %d 個 Word 檔案，準備開始批次轉換...", len(docx_files))
    success_count = 0

    for docx_file in docx_files:
        try:
            convert_docx_to_markdown(docx_file, output_dir)
            success_count += 1
        except Exception:
            logger.error("檔案 '%s' 轉換失敗，跳過繼續處理下一筆。", docx_file.name)
            continue

    logger.info("處理完畢！成功: %d / 總數: %d", success_count, len(docx_files))
    return success_count


def main() -> None:
    """主執行進入點"""
    base_dir = Path(__file__).resolve().parent.parent
    input_dir = base_dir / "input"
    output_dir = base_dir / "output"

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    logger.info("=== WORD清洗MD 工具啟動 ===")
    logger.info("來源目錄: %s", input_dir)
    logger.info("輸出目錄: %s", output_dir)

    batch_process(input_dir, output_dir)


if __name__ == "__main__":
    main()
