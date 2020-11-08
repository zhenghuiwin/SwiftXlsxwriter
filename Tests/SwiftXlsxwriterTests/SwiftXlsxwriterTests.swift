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

        sheet!.writeString(row: row, col: 0, content: "你好", format: nil)
        let fmt = DateFormatter()
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

        var type = ["天气状况", "最低气温", "最高气温", "云量"]
        let city = ["成都", "绵阳", "德阳", "雅安"]

        let format = wb.addFormat()
        XCTAssertNotNil(format)
        format!.setBorder(type: .thin)
        format!.setAlign(type: .center)
        format!.setAlign(type: .verticalCenter)

        let eleNum = 3
        row += 1
        city.forEach { cityName in
            print("------------")
            print(row)

            sheet!.mergeRange(firstRow: row, firstCol: 0, lastRow: (row + eleNum), lastCol: 0, content: cityName, format: format!)

            for r in 0 ..< 4 {
                sheet!.writeString(row: row, col: 1, content: type[Int(r)], format: format!)
                for c in 2 ..< 9 {
                    sheet!.writeString(row: row, col: c, content: "\(c)", format: format!)
                }
                row += 1
                print(row)
            }
            print("该组生成完毕，为下一行准备: \(row)")
        } // city.forEach
    }

    static var allTests = [
        ("testExample", testExample),
        ("testWb", testWb),
    ]
}
