import sys
import os
import time
import win32clipboard
import win32com.client as win32
import argparse
from PIL import Image
from io import BytesIO

def set_clipboard_text(data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, data)
    win32clipboard.CloseClipboard()

def set_clipboard_image(image_path):
    image = Image.open(image_path)
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def read_file_content(file_path):
    _, ext = os.path.splitext(file_path)
    if ext.lower() == '.txt':
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read(), 'text'
    elif ext.lower() in ['.png', '.jpg', '.jpeg', '.bmp', '.gif']:
        return file_path, 'image'
    else:
        print(f"Unsupported file type: {ext}")
        return None, None

def paste_clipboard_content(hwp):
    try:
        time.sleep(0.5)
        hwp.HAction.GetDefault("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
        hwp.HAction.Execute("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
    except Exception as e:
        print(f"Error during paste: {e}")

def replace_text_with_clipboard(hwp, search_text, replacement_type):
    try:
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
        
        if result:
            paste_clipboard_content(hwp)
            return True
        else:
            print(f"'{search_text}' not found.")
            return False
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
        if doc_count > 0:
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

    if len(args.replacements) % 2 != 0:
        parser.error("Replacements should be in pairs of search text and replacement file path.")

    replacements = [(args.replacements[i], args.replacements[i+1]) for i in range(0, len(args.replacements), 2)]
    return args.file_path, replacements

def main():
    file_path, replacements = parse_arguments()

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        
        hwp.Open(file_path)
        
        print("File opened:", file_path)

        check_document_state(hwp)

        for search_text, replacement_file_path in replacements:
            content, content_type = read_file_content(replacement_file_path.strip('"'))
            if content:
                if content_type == 'text':
                    set_clipboard_text(content)
                elif content_type == 'image':
                    set_clipboard_image(content)
                if replace_text_with_clipboard(hwp, search_text.strip('"'), content_type):
                    print(f"Replaced '{search_text}' with content from {replacement_file_path}")
                else:
                    print(f"Failed to replace '{search_text}'")
            else:
                print(f"Failed to read content from {replacement_file_path}")

        save_as_hwp_and_pdf(hwp, file_path)

    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'hwp' in locals():
            hwp.Quit()
        print("HWP application closed.")

if __name__ == "__main__":
    main()
