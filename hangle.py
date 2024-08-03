import sys
import os
import time
import win32clipboard
import win32com.client as win32

def get_clipboard_data():
    win32clipboard.OpenClipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        elif win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_BITMAP):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_BITMAP)
        elif win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
        elif win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_HDROP):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
        else:
            data = None
    except:
        data = None
    finally:
        win32clipboard.CloseClipboard()
    return data

def paste_clipboard_content(hwp):
    try:
        time.sleep(0.5)  # Add a small delay to ensure clipboard data is ready
        hwp.HAction.GetDefault("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
        hwp.HAction.Execute("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
        print("클립보드 내용을 붙여넣었습니다.")
    except Exception as e:
        print(f"붙여넣기 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()

def replace_text_with_clipboard(hwp, search_text):
    try:
        hwp.HAction.GetDefault("FindDlg", hwp.HParameterSet.HFindReplace.HSet)
        findReplace = hwp.HParameterSet.HFindReplace
        findReplace.MatchCase = 0
        findReplace.AllWordForms = 0
        findReplace.SeveralWords = 0
        findReplace.UseWildCards = 0
        findReplace.WholeWordOnly = 0
        findReplace.AutoSpell = 1
        findReplace.Direction = 0
        findReplace.IgnoreFindString = 0
        findReplace.IgnoreReplaceString = 0
        findReplace.FindString = ""
        findReplace.ReplaceString = ""
        findReplace.IgnoreMessage = 1
        findReplace.HanjaFromHangul = 0
        findReplace.FindJaso = 0
        findReplace.FindRegExp = 0
        findReplace.FindTextInPicture = 0
        findReplace.FindStyle = ""
        findReplace.ReplaceStyle = ""
        hwp.HAction.Execute("FindDlg", hwp.HParameterSet.HFindReplace.HSet)

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
            clipboard_content = get_clipboard_data()
            if clipboard_content:
                paste_clipboard_content(hwp)
                print(f"'{search_text}'를 찾아 클립보드 내용으로 대체했습니다.")
                return True
            else:
                print("클립보드에 유효한 데이터가 없습니다.")
                return False
        else:
            print(f"'{search_text}'를 찾을 수 없습니다.")
            return False
    except Exception as e:
        print(f"오류 발생: {e}")
        print(f"오류 타입: {type(e).__name__}")
        print("오류 상세:")
        import traceback
        traceback.print_exc()
        return False

def save_as_hwp_and_pdf(hwp, original_path):
    try:
        directory, filename = os.path.split(original_path)
        filename_without_ext, ext = os.path.splitext(filename)

        new_filename = f"{filename_without_ext}_수정완료{ext}"
        new_filepath = os.path.join(directory, new_filename)

        hwp.SaveAs(new_filepath, "HWP", "")
        print(f"HWP 파일이 저장되었습니다: {new_filepath}")

        pdf_filepath = f"{os.path.splitext(new_filepath)[0]}.pdf"
        hwp.SaveAs(pdf_filepath, "PDF")
        print(f"PDF 파일이 저장되었습니다: {pdf_filepath}")

    except Exception as e:
        print(f"파일 저장 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()

def check_document_state(hwp):
    try:
        doc_count = hwp.XHwpDocuments.Count
        if doc_count > 0:
            print(f"문서가 열려 있습니다. (열린 문서 수: {doc_count})")
        else:
            print("열린 문서가 없습니다.")
    except Exception as e:
        print(f"문서 상태 확인 중 오류 발생: {e}")

def main():
    if len(sys.argv) < 3:
        print("사용법: python hangle.py <파일 경로> <검색할 텍스트>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    search_text = sys.argv[2]

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        
        hwp.Open(file_path)
        
        print("파일이 열렸습니다:", file_path)

        check_document_state(hwp)

        if replace_text_with_clipboard(hwp, search_text):
            save_as_hwp_and_pdf(hwp, file_path)

    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if 'hwp' in locals():
            hwp.Quit()
        print("한글 애플리케이션이 종료되었습니다.")

if __name__ == "__main__":
    main()
