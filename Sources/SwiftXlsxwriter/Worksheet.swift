//
//  File.swift
//  
//
//  Created by zhenghuiwin on 2020/3/21.
//

import Foundation
import ExcelWriter

public struct Worksheet {
    private let ws: UnsafeMutablePointer<lxw_worksheet>
    
    public init(_ ws: UnsafeMutablePointer<lxw_worksheet>) {
        self.ws = ws
    }
    
    //  worksheet_write_string(worksheet, row, UInt16(c), header[c], nil)
    public func writeString(row: Int,  col: Int, content: String, format: Format?) {
        worksheet_write_string(ws, UInt32(row), UInt16(col), content, format?.rawValue)
    }
    
    // worksheet_write_number(worksheet, 1, 0, 123.456, format);
    
    public func writeNumber(row: Int,  col: Int, number: Double, format: Format?) {
        worksheet_write_number(ws, UInt32(row), UInt16(col), number, format?.rawValue);
    }
    
    // worksheet_merge_range(worksheet, row, 0, (row + eleNum), 0, cityName, format)
    public func mergeRange(firstRow: Int, firstCol: Int, lastRow: Int, lastCol: Int, content: String, format: Format) {
        worksheet_merge_range(ws, UInt32(firstRow), UInt16(firstCol), UInt32(lastRow), UInt16(lastCol), content, format.rawValue)
    }
    
    // sheet.setColumn(firstCol: 0, lastCol: 10, width: 15, format: nil)
    public func setColumn(firstCol: UInt16, lastCol: UInt16, width: Double, format: Format?) {
        worksheet_set_column(ws, firstCol, lastCol, width, format?.rawValue)
    }
}
