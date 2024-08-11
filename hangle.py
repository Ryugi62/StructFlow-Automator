import sys
import os
import time
import win32clipboard
import win32com.client as win32
import argparse
from PIL import Image
from io import BytesIO
import base64
import win32gui
import traceback

def set_clipboard_text(data):
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, data)
        win32clipboard.CloseClipboard()
    except Exception as e:
        print(f"Error setting text to clipboard: {e}")

def set_clipboard_image(image_data):
    try:
        if isinstance(image_data, str) and image_data.startswith("data:image"):
            base64_data = image_data.split(",")[1]
            image_data = base64.b64decode(base64_data)
            image = Image.open(BytesIO(image_data))
        else:
            image = Image.open(image_data)

        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()
        print(f"Image set to clipboard: {image_data}")
    except Exception as e:
        print(f"Error setting image to clipboard: {e}")

def convert_emf_to_png(emf_path):
    try:
        # EMF 파일을 PNG로 변환
        image = Image.open(emf_path)
        png_path = emf_path.replace('.emf', '.png')
        image.save(png_path)
        return png_path
    except Exception as e:
        print(f"Error converting EMF to PNG: {e}")
        return None

def set_clipboard_emf(file_path):
    try:
        # EMF 파일을 PNG로 변환한 후 클립보드에 설정
        png_path = convert_emf_to_png(file_path)
        if png_path:
            set_clipboard_image(png_path)
    except Exception as e:
        print(f"Error setting EMF to clipboard: {e}")

def read_file_content(file_path):
    file_path = file_path.strip('"')
    print(f"Attempting to read file: {file_path}")

    if not os.path.exists(file_path):
        print(f"File does not exist: {file_path}")
        return None, None

    _, ext = os.path.splitext(file_path)
    print(f"File extension: {ext}")

    if ext.lower() == ".txt":
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                content = file.read().strip()
                if "년월" in file_path:
                    content = content.replace("년 ", ".").replace("월", "")
                print(f"Successfully read text file with UTF-8 encoding. Content: {content}")
                return content, "text"
        except UnicodeDecodeError:
            try:
                with open(file_path, "r", encoding="euc-kr") as file:
                    content = file.read().strip()
                    if "년월" in file_path:
                        content = content.replace("년 ", ".").replace("월", "")
                    print(f"Successfully read text file with EUC-KR encoding. Content: {content}")
                    return content, "text"
            except Exception as e:
                print(f"Failed to read text file with both UTF-8 and EUC-KR encodings: {str(e)}")
                return None, None
    elif ext.lower() in [".png", ".jpg", ".jpeg", ".bmp", ".gif"]:
        if file_path.startswith("data:image"):
            print("Detected base64 encoded image")
            return file_path, "base64_image"
        else:
            print(f"Detected image file: {file_path}")
            return file_path, "image"
    elif ext.lower() == ".emf":
        print(f"Detected EMF file: {file_path}")
        return file_path, "emf"
    else:
        print(f"Unsupported file type: {ext}")
        return None, None

def paste_clipboard_content(hwp):
    try:
        time.sleep(1)
        hwp.HAction.GetDefault("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
        hwp.HAction.Execute("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
        print("Content pasted from clipboard.")
    except Exception as e:
        print(f"Error during paste: {e}")

def replace_text_with_clipboard(hwp, search_text, replacement_type):
    try:
        occurrences = 0
        while True:
            hwp.MovePos(2)
            hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            findReplace = hwp.HParameterSet.HFindReplace
            findReplace.ReplaceString = ""
            findReplace.FindString = search_text
            findReplace.IgnoreReplaceString = 0
            findReplace.IgnoreFindString = 0
            findReplace.Direction = 0
            findReplace.WholeWordOnly = 0
            findReplace.UseWildCards = 0
            findReplace.SeveralWords = 0
            findReplace.AllWordForms = 0
            findReplace.MatchCase = 0
            findReplace.ReplaceMode = 0
            findReplace.ReplaceStyle = ""
            findReplace.FindStyle = ""
            findReplace.FindTextInPicture = 0
            findReplace.FindRegExp = 0
            findReplace.FindJaso = 0
            findReplace.HanjaFromHangul = 0
            findReplace.IgnoreMessage = 1
            findReplace.FindType = 1

            result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            if not result:
                break
            paste_clipboard_content(hwp)
            occurrences += 1

        return occurrences > 0
    except Exception as e:
        print(f"Error: {e}")
        return False

def save_as_hwp_and_pdf(hwp, original_path):
    try:
        directory, filename = os.path.split(original_path)
        filename_without_ext, ext = os.path.splitext(filename)

        new_filename = f"{filename_without_ext}_수정완료{ext}"
        new_filepath = os.path.join(directory, new_filename)

        hwp.SaveAs(new_filepath, "HWP", "")
        print(f"HWP file saved: {new_filepath}")

        pdf_filepath = f"{os.path.splitext(new_filepath)[0]}.pdf"
        hwp.SaveAs(pdf_filepath, "PDF")
        print(f"PDF file saved: {pdf_filepath}")

    except Exception as e:
        print(f"Error saving file: {e}")

def check_document_state(hwp):
    try:
        doc_count = hwp.XHwpDocuments.Count
        if (doc_count > 0):
            print(f"Documents are open. (Open document count: {doc_count})")
        else:
            print("No documents open.")
    except Exception as e:
        print(f"Error checking document state: {e}")

def parse_arguments():
    parser = argparse.ArgumentParser(description="Replace text in a HWP file with clipboard content.")
    parser.add_argument("file_path", help="Path to the HWP file")
    parser.add_argument("replacements", nargs="+", help="Pairs of search text and replacement file path")

    args = parser.parse_args()

    replacements = [(args.replacements[i], args.replacements[i + 1]) for i in range(0, len(args.replacements), 2)]
    return args.file_path, replacements

def main():
    file_path, replacements = parse_arguments()
    print(f"HWP file path: {file_path}")
    print("Replacements:")
    for search_text, replacement_file_path in replacements:
        print(f"  Search text: '{search_text}', Replacement file: '{replacement_file_path}'")

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(file_path)
        print("File opened:", file_path)

        check_document_state(hwp)

        for search_text, replacement_file_path in replacements:
            content, content_type = read_file_content(replacement_file_path)
            if content is not None:
                if content_type == "text":
                    set_clipboard_text(content)
                elif content_type in ["image", "base64_image"]:
                    set_clipboard_image(content)
                elif content_type == "emf":
                    set_clipboard_emf(content)
                if replace_text_with_clipboard(hwp, search_text.strip('"'), content_type):
                    print(f"Replaced all instances of '{search_text}' with content from {replacement_file_path}")
                else:
                    print(f"Failed to replace '{search_text}'")
            else:
                print(f"Failed to read content from {replacement_file_path}")

        save_as_hwp_and_pdf(hwp, file_path)

    except Exception as e:
        print(f"Error: {e}")
        print(f"Error type: {type(e).__name__}")
        print("Traceback:")
        print(traceback.format_exc())
    finally:
        if "hwp" in locals():
            hwp.Quit()
        print("HWP application closed.")

if __name__ == "__main__":
    main()
