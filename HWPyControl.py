import win32com.client as win32
import sys
import os
import win32clipboard
from PIL import Image
import io

class HWPyControl:
    def __init__(self):
        self.hwp = None

    def open_hwp(self, file_path):
        try:
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            self.hwp.Open(file_path)
            print(f"HWP file opened successfully: {file_path}")
        except Exception as e:
            print(f"Failed to open HWP file: {e}")
            sys.exit(1)

    def save_hwp(self, file_path):
        try:
            self.hwp.SaveAs(file_path, "HWP", "")
            print(f"HWP file saved successfully: {file_path}")
        except Exception as e:
            print(f"Failed to save HWP file: {e}")
            sys.exit(1)

    def save_pdf(self, file_path):
        try:
            self.hwp.SaveAs(file_path, "PDF")
            print(f"PDF file saved successfully: {file_path}")
        except Exception as e:
            print(f"Failed to save PDF file: {e}")
            sys.exit(1)

    def close_hwp(self):
        if self.hwp:
            try:
                self.hwp.Quit()
                print("HWP file closed successfully.")
            except Exception as e:
                print(f"Failed to close HWP: {e}")
            finally:
                self.hwp = None

    def find_text(self, text, instance=1):
        try:
            self.hwp.HAction.Run("MoveDocBegin")
            find_replace_set = self.hwp.HParameterSet.HFindReplace
            find_replace_set.FindString = text
            find_replace_set.IgnoreMessage = 1
            self.hwp.HAction.Execute("FindDlg", find_replace_set.HSet)
            for _ in range(instance):
                self.hwp.HAction.Execute("RepeatFind", find_replace_set.HSet)
            self.hwp.HAction.Run("MoveDown")
            print(f"Found text: '{text}' (instance: {instance})")
            return True
        except Exception as e:
            print(f"Failed to find text: {e}")
            return False

    def replace_text(self, search_text, replace_text):
        try:
            self.hwp.HAction.Run("MoveDocBegin")
            find_replace_set = self.hwp.HParameterSet.HFindReplace
            find_replace_set.FindString = search_text
            find_replace_set.ReplaceString = replace_text
            find_replace_set.IgnoreMessage = 1
            self.hwp.HAction.Execute("AllReplace", find_replace_set.HSet)
            print(f"Replaced '{search_text}' with '{replace_text}'")
        except Exception as e:
            print(f"Failed to replace text: {e}")

    def insert_text(self, text):
        if text.strip():
            self._copy_to_clipboard(text.strip())
            self.hwp.HAction.Run("Paste")
            print(f"Inserted text: {text.strip()}")

    def insert_image(self, img_path, size=None):
        if not os.path.exists(img_path):
            print(f"Image not found: {img_path}")
            return
        self._copy_image_to_clipboard(img_path, size)
        self.hwp.HAction.Run("Paste")
        print(f"Inserted image from {img_path}")

    def find_and_insert_image(self, search_text, img_path, size=None, instance=1):
        if self.find_text(search_text, instance):
            self.insert_image(img_path, size)
            print(f"Inserted image after finding '{search_text}' (instance: {instance})")
        else:
            print(f"Failed to insert image: Text '{search_text}' not found")

    def add_table_row(self, rows):
        self.hwp.HAction.GetDefault("TableInsertRowColumn", self.hwp.HParameterSet.HTableInsertLine.HSet)
        self.hwp.HParameterSet.HTableInsertLine.Side = 3
        self.hwp.HParameterSet.HTableInsertLine.Count = rows
        self.hwp.HAction.Execute("TableInsertRowColumn", self.hwp.HParameterSet.HTableInsertLine.HSet)
        print(f"Added {rows} row(s) to the table")

    def move_to_top(self):
        self.hwp.HAction.Run("MoveDocBegin")
        print("Moved to the top of the document")

    def move_up(self):
        self.hwp.HAction.Run("MoveUp")
        print("Moved cursor up")

    def move_down(self):
        self.hwp.HAction.Run("MoveDown")
        print("Moved cursor down")

    def move_left(self):
        self.hwp.HAction.Run("MoveLeft")
        print("Moved cursor left")

    def move_right(self):
        self.hwp.HAction.Run("MoveRight")
        print("Moved cursor right")

    def press_esc(self):
        self.hwp.HAction.Run("Cancel")
        print("Pressed ESC key")

    def _copy_to_clipboard(self, text):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
        except Exception as e:
            print(f"Failed to copy text to clipboard: {e}")

    def _copy_image_to_clipboard(self, image_path, size=None):
        img = Image.open(image_path)
        if size:
            img = img.resize(size, Image.Resampling.LANCZOS)
        output = io.BytesIO()
        img.convert("RGB").save(output, format="BMP")
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()

def main():
    hwp_control = HWPyControl()
    
    current_dir = os.getcwd()
    print(f"현재 작업 디렉토리: {current_dir}")

    # 한글 파일 열기
    input_file = os.path.join(current_dir, "StructFlow-Automator-Private", "References", "default.hwp")
    hwp_control.open_hwp(input_file)
    
    # 문서 최상단으로 이동
    hwp_control.move_to_top()
    
    # 커서 이동 테스트
    hwp_control.move_down()
    hwp_control.move_right()
    hwp_control.move_up()
    hwp_control.move_left()
    
    # ESC 키 테스트
    hwp_control.press_esc()
    
    # 텍스트 찾기 후 이미지 삽입
    image_file = os.path.join(current_dir, "StructFlow-Automator-Private", "References", "image.png")
    hwp_control.find_and_insert_image(r"{{이미지}}", image_file, (100, 100))
    
    # 텍스트 교체
    hwp_control.replace_text("원본 텍스트", "교체할 텍스트")
    
    # 텍스트 삽입
    hwp_control.insert_text("삽입할 텍스트")
    
    # 표에 행 추가
    hwp_control.add_table_row(2)
    
    # 파일 저장 (HWP)
    output_hwp = os.path.join(current_dir, "StructFlow-Automator-Private", "References", "output.hwp")
    hwp_control.save_hwp(output_hwp)
    
    # 파일 저장 (PDF)
    output_pdf = os.path.join(current_dir, "StructFlow-Automator-Private", "References", "output.pdf")
    hwp_control.save_pdf(output_pdf)
    
    # 한글 종료
    hwp_control.close_hwp()

    print(f"모든 작업이 완료되었습니다. 파일이 {current_dir}에 저장되었습니다.")

if __name__ == "__main__":
    main()