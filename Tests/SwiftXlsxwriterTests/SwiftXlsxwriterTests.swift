import XCTest
@testable import SwiftXlsxwriter

import ExcelWriter

final class SwiftXlsxwriterTests: XCTestCase {
    func testExample() {
        // This is an example of a functional test case.
        // Use XCTAssert and related functions to verify your tests produce the correct
        // results.
        XCTAssertEqual(SwiftXlsxwriter().text, "Hello, World!")
    }
    
    func testWb() {
        let wb = Workbook(file: "./testtest.xlsx")
        let sheet: Worksheet? = wb.addWorksheet(sheetName: "sheet1")
        XCTAssertNotNil(sheet)
        
        var row = 1
        var col = 0

        sheet!.writeString(row: row, col: col, content: "你好", format: nil)
        var fmt = DateFormatter()
        fmt.dateFormat = "MM月dd日"
        let cal = Calendar.current
        let now = Date()
        var comp = DateComponents()



        var header = ["地区", "日期",]

        for d in 1 ... 7 {
            comp.day = d
            if let nextDay = cal.date(byAdding: comp, to: now) {
                header.append(fmt.string(from: nextDay))
            }
        }

        print(header)

        row = 3
        for c in 0 ..< header.count  {
            sheet!.writeString(row: row, col: c, content: header[c], format: nil)
        }
//
//        var type = ["天气状况", "最低气温", "最高气温", "云量"]
//        var city = ["成都", "绵阳", "德阳", "雅安"]
//
//        let format = workbook_add_format(wb)
//        format_set_border(format,  1)
//        format_set_align(format, 2)
//        format_set_align(format, 10)
//
//
//        let eleNum: UInt32 = 3
//        row += 1
//        city.forEach { cityName in
//            print("------------")
//            print(row)
//
//            worksheet_merge_range(worksheet, row, 0, (row + eleNum), 0, cityName, format)
//
//            for r in 0 ..< 4 {
//                worksheet_write_string(worksheet, row, 1, type[Int(r)], format)
//                for c in 2 ..< 9 {
//                   worksheet_write_string(worksheet, row, UInt16(c), "\(c)", format)
//                }
//                row += 1
//                print(row)
//            }
//            print("该组生成完毕，为下一行准备: \(row)")
//        }
//
//
//
//
//        let rt1 = workbook_close(wb)
        // print("rt: \( lxw_strerror(rt)); \( lxw_strerror(rt1))")
    }

    static var allTests = [
        ("testExample", testExample),
        ("testWb", testWb),
    ]
}
