//
//  ViewController.swift
//  excel-example
//
//  Created by GwakDoyoung on 01/06/2018.
//  Copyright © 2018 GwakDoyoung. All rights reserved.
//

import UIKit
import QuickLook

class ViewController: UIViewController {
    
    var url: URL? = nil
    
    let preview = QLPreviewController()

    override func viewDidLoad() {
        super.viewDidLoad()
        
        // 저장할 경로 설정(앱 내의 도큐먼트 폴더에 filename.xlsx 이름의 파일)
        let fileManager = FileManager.default
        let documentsURL = fileManager.urls(for: .documentDirectory, in: .userDomainMask)[0]
        url = documentsURL.appendingPathComponent("filename").appendingPathExtension("xlsx")
        let path = NSString(string: url!.path).fileSystemRepresentation
        print("path:", url?.path)
        
        
        
        /* 새 워크북을 만들고 워크시트를 하나 추가하기 */
        let workbook = new_workbook(path)
        let worksheet = workbook_add_worksheet(workbook, nil);
        
        
        // 0컬럼부터 5컬럼까지 13.7 너비 적용
        worksheet_set_column(worksheet, 0, 5, 13.7, nil);
        

        /* 0 row, 4 column에 "이름" 문자열 넣기 */
        worksheet_write_string(worksheet, 0, 4, "이름", nil);
        /* 1 row, 2 column에 1.25123 숫자 넣기 */
        worksheet_write_number(worksheet, 1, 2, 1.25123, nil);
        
        
        /* 포멧 생성 */
        let my_format = workbook_add_format(workbook);
        format_set_align(my_format, UInt8(LXW_ALIGN_RIGHT.rawValue));
        format_set_align(my_format, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue));
        format_set_bold(my_format);
        format_set_font_size(my_format, 20)
        format_set_font_name(my_format, "Arial")
        /* 포멧이 적용된 셀 생성 */
        worksheet_write_string(worksheet, 0, 1, "타이틀", my_format);
        

        /* 3 row, 0 columm부터 3 row, 5 column 까지 셀 합치기 */
        worksheet_merge_range(worksheet, 3, 0, 3, 5, "합쳐진 셀", my_format);
        
        
        /* 워크북 닫기 */
        workbook_close(workbook);
        
        
        
        
    
        
        // 저장된 xlsx 파일 미리보기
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.5) {
            self.preview.dataSource = self
            self.preview.currentPreviewItemIndex = 0
            self.present(self.preview, animated: true)
        }
    
        
        // 저장된 xlsx 파일 내보내기
        //self.exportFile(url: url)
        
    }
}

extension ViewController: QLPreviewControllerDataSource {
    func numberOfPreviewItems(in controller: QLPreviewController) -> Int {
        return 1
    }

    func previewController(_ controller: QLPreviewController, previewItemAt index: Int) -> QLPreviewItem {
        return url! as QLPreviewItem
    }
}


extension UIViewController {
    func exportFile(url: URL) {
        let activityVC = UIActivityViewController(activityItems: [url], applicationActivities: [])
        activityVC.excludedActivityTypes = [
            .message,
            .mail,
            .print,
            // .copyToPasteboard
        ]
        activityVC.completionWithItemsHandler = {
            (activity, success, items, error) in
            if success {
                // 파일 내보내기 성공
                self.showSimpleAlert(message: "파일 내보내기 성공")
            } else if let e = error {
                self.showSimpleAlert(message: "내보내기 실패\n\(e)")
            } else {
                // 내보내기 취소
            }
        }
        
        present(activityVC, animated: true)
    }
}

extension UIViewController {
    func showSimpleAlert(message: String?) {
        self.showSimpleAlert(title: "알림", message: message)
    }
    
    func showSimpleAlert(title: String?, message: String?) {
        let alert = UIAlertController(title: title, message: message, preferredStyle: .alert)
        alert.addAction(UIAlertAction(title: "확인", style: .default, handler: nil))
        self.present(alert, animated: true, completion: nil)
    }
}

